import base64
import hashlib
import json
import math
import os
import re
import subprocess
import tempfile
import tkinter as tk
import webbrowser
from datetime import datetime
from io import BytesIO
from pathlib import Path
from tkinter import messagebox, ttk
from urllib.error import URLError
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse
from urllib.request import Request, urlopen


ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "moderator_config.json"
REJECTIONS_PATH = ROOT / "moderator_rejections.json"
STATUS_HEADER = "Статус модерации"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]
FIELDS = {
    "type": ["что именно вы публикуете"],
    "name": ["название стиля"],
    "author": ["ник автора"],
    "description": ["краткое описание стиля"],
    "version": ["версия стиля"],
    "tags": ["тэги", "теги"],
    "preview": ["превью"],
    "css": ["стиль, прикрипите", "стиль, прикрепите", "файл с рабочим стилем"],
    "email": ["электронная почта", "email"],
}


def load_config():
    if not CONFIG_PATH.exists():
        raise RuntimeError(
            "Не найден moderator_config.json. Скопируйте "
            "moderator_config.example.json в moderator_config.json и укажите "
            "ID таблицы и точное имя вкладки с ответами."
        )
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    if not config.get("spreadsheet_id") or "PASTE_" in config["spreadsheet_id"]:
        raise RuntimeError("Укажите spreadsheet_id в moderator_config.json.")
    if not config.get("sheet_name"):
        raise RuntimeError("Укажите sheet_name в moderator_config.json.")
    return config


def quote_sheet(name):
    return "'" + name.replace("'", "''") + "'"


def column_name(index):
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def find_drive_ids(value):
    value = value or ""
    ids = re.findall(r"(?:/d/|[?&]id=)([A-Za-z0-9_-]{10,})", value)
    if not ids and re.fullmatch(r"[A-Za-z0-9_-]{20,}", value.strip()):
        ids = [value.strip()]
    return list(dict.fromkeys(ids))


def find_onedrive_url(value):
    match = re.search(r"https://[^\s<>\"']+", value or "", re.IGNORECASE)
    if not match:
        return None
    url = match.group(0).rstrip(".,);]")
    host = (urlparse(url).hostname or "").casefold()
    if (
        host == "1drv.ms"
        or host == "onedrive.live.com"
        or host.endswith(".sharepoint.com")
    ):
        return url
    return None


def parse_tags(value):
    return [tag.strip() for tag in value.splitlines() if tag.strip()]


def submission_key(row, status_column):
    submission = [value for index, value in enumerate(row) if index != status_column]
    serialized = json.dumps(submission, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(serialized.encode("utf-8")).hexdigest()


def local_timestamp():
    return datetime.now().astimezone().isoformat(timespec="seconds")


def slug_from_url(value):
    parts = [part for part in urlparse(value or "").path.split("/") if part]
    try:
        themes_index = parts.index("themes")
        return parts[themes_index + 1]
    except (ValueError, IndexError):
        return ""


class ModeratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Legion7TV — модерация стилей")
        self.root.geometry("1200x760")
        self.config = load_config()
        self.repo = (ROOT / self.config.get("repository_path", ".")).resolve()
        self.credentials_path = ROOT / self.config.get(
            "credentials_file", "credentials.json"
        )
        self.token_path = ROOT / self.config.get("token_file", "token.json")
        self.sheets = None
        self.drive = None
        self.headers = []
        self.columns = {}
        self.status_column = 0
        self.submissions = []
        self.rejected_submissions = self._load_rejections()
        self.selected = None
        self.file_cache = {}
        self.preview_image = None
        self._build_ui()
        self.root.after(200, self.refresh)

    @staticmethod
    def _load_rejections():
        if not REJECTIONS_PATH.exists():
            return set()
        try:
            state = json.loads(REJECTIONS_PATH.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise RuntimeError(
                f"Не удалось прочитать список отклонённых заявок "
                f"{REJECTIONS_PATH}: {exc}"
            ) from exc
        rejected = state.get("rejected_submissions") if isinstance(state, dict) else None
        if not isinstance(rejected, list) or any(
            not isinstance(key, str) for key in rejected
        ):
            raise ValueError(
                f"Некорректный формат файла отклонённых заявок: {REJECTIONS_PATH}"
            )
        return set(rejected)

    @staticmethod
    def _save_rejections(rejected):
        state = {"rejected_submissions": sorted(rejected)}
        content = (json.dumps(state, ensure_ascii=False, indent=2) + "\n").encode(
            "utf-8"
        )
        ModeratorApp._atomic_write(REJECTIONS_PATH, content)

    def _build_ui(self):
        toolbar = ttk.Frame(self.root, padding=8)
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="Обновить заявки", command=self.refresh).pack(
            side="left"
        )
        ttk.Button(
            toolbar,
            text="Git add + commit + push",
            command=self.commit_and_push,
        ).pack(side="right")
        self.status = tk.StringVar(value="Подключение к Google...")
        ttk.Label(toolbar, textvariable=self.status).pack(side="left", padx=12)

        body = ttk.Panedwindow(self.root, orient="horizontal")
        body.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        left = ttk.Frame(body, padding=4)
        right = ttk.Frame(body, padding=4)
        body.add(left, weight=1)
        body.add(right, weight=3)

        ttk.Label(left, text="Новые заявки").pack(anchor="w", pady=(0, 5))
        self.listbox = tk.Listbox(left, exportselection=False)
        self.listbox.pack(fill="both", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self._select_submission)
        ttk.Button(
            left, text="Отклонить заявку", command=self.reject_submission
        ).pack(anchor="w", pady=(6, 0))

        ttk.Label(right, text="Данные заявки").pack(anchor="w", pady=(0, 5))
        self.details = tk.Text(right, height=10, wrap="word", state="disabled")
        self.details.pack(fill="x")
        self.preview_label = ttk.Label(
            right,
            text="Выберите заявку для предпросмотра",
            anchor="center",
            wraplength=700,
        )
        self.preview_label.pack(fill="x", pady=(8, 0))
        ttk.Label(right, text="Содержимое CSS (просмотр, не выполнение)").pack(
            anchor="w", pady=(10, 5)
        )
        self.css_view = tk.Text(right, height=6, wrap="none", state="disabled")
        self.css_view.pack(fill="both", expand=True)

        buttons = ttk.Frame(right)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="Открыть превью PNG", command=self.open_preview).pack(
            side="left"
        )
        ttk.Button(buttons, text="Копировать CSS", command=self.copy_css).pack(
            side="left", padx=6
        )
        ttk.Button(
            buttons,
            text="Сгенерировать тему и добавить в каталог",
            command=self.approve,
        ).pack(side="right")

    def _services(self):
        if self.sheets and self.drive:
            return
        try:
            from google.auth.transport.requests import Request
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
            from googleapiclient.discovery import build
        except ImportError as exc:
            raise RuntimeError(
                "Не установлены Google API библиотеки. Запустите:\n"
                "py -m pip install -r requirements-moderator.txt"
            ) from exc

        if not self.credentials_path.exists():
            raise RuntimeError(
                f"Не найден файл OAuth credentials:\n{self.credentials_path}\n"
                "Скачайте OAuth-клиент типа Desktop из Google Cloud Console."
            )

        credentials = None
        if self.token_path.exists():
            credentials = Credentials.from_authorized_user_file(
                str(self.token_path), SCOPES
            )
        if credentials and not credentials.has_scopes(SCOPES):
            credentials = None
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        if not credentials or not credentials.valid:
            flow = InstalledAppFlow.from_client_secrets_file(
                str(self.credentials_path), SCOPES
            )
            credentials = flow.run_local_server(port=0)
        self.token_path.write_text(credentials.to_json(), encoding="utf-8")
        self.sheets = build(
            "sheets", "v4", credentials=credentials, cache_discovery=False
        )
        self.drive = build(
            "drive", "v3", credentials=credentials, cache_discovery=False
        )

    def _sheet_range(self, cells):
        return f"{quote_sheet(self.config['sheet_name'])}!{cells}"

    def _ensure_status_column(self, headers):
        normalized = [header.strip().casefold() for header in headers]
        if STATUS_HEADER.casefold() in normalized:
            return normalized.index(STATUS_HEADER.casefold()), headers
        index = len(headers)
        self.sheets.spreadsheets().values().update(
            spreadsheetId=self.config["spreadsheet_id"],
            range=self._sheet_range(f"{column_name(index + 1)}1"),
            valueInputOption="RAW",
            body={"values": [[STATUS_HEADER]]},
        ).execute()
        headers.append(STATUS_HEADER)
        return index, headers

    def _find_columns(self, headers):
        normalized = [" ".join(value.casefold().split()) for value in headers]
        columns = {}
        for field, fragments in FIELDS.items():
            matches = []
            for index, header in enumerate(normalized):
                if any(fragment in header for fragment in fragments):
                    matches.append(index)
            if field in ("tags", "email", "version") and not matches:
                columns[field] = None
            elif not matches:
                raise ValueError(
                    f"В таблице не найден столбец для поля «{field}».\n"
                    "Проверьте заголовки вопросов в первой строке таблицы."
                )
            else:
                columns[field] = matches[0]
        return columns

    @staticmethod
    def _cell(row, index):
        if index is None or index >= len(row):
            return ""
        return str(row[index]).strip()

    def refresh(self):
        try:
            self.status.set("Читаю ответы...")
            self.root.update_idletasks()
            self._services()
            response = (
                self.sheets.spreadsheets()
                .values()
                .get(
                    spreadsheetId=self.config["spreadsheet_id"],
                    range=self._sheet_range("A:ZZ"),
                )
                .execute()
            )
            rows = response.get("values", [])
            if not rows:
                self.status.set("В таблице пока нет ответов.")
                return

            headers = list(rows[0])
            self.status_column, headers = self._ensure_status_column(headers)
            self.columns = self._find_columns(headers)
            self.headers = headers
            self.submissions = []
            for sheet_row, row in enumerate(rows[1:], start=2):
                review_status = self._cell(row, self.status_column).casefold()
                if review_status and review_status not in ("новая", "new"):
                    continue
                key = submission_key(row, self.status_column)
                if key in self.rejected_submissions:
                    continue
                item = {
                    field: self._cell(row, column)
                    for field, column in self.columns.items()
                }
                item["row"] = sheet_row
                item["key"] = key
                self.submissions.append(item)

            self.listbox.delete(0, tk.END)
            for item in self.submissions:
                title = item["name"] or "(без названия)"
                author = item["author"] or "автор не указан"
                self.listbox.insert(
                    tk.END, f"Строка {item['row']} · {title} — {author}"
                )
            self.selected = None
            self._set_text(self.details, "")
            self._set_text(self.css_view, "")
            self._set_preview(None)
            self.status.set(f"Новых заявок: {len(self.submissions)}")
        except Exception as exc:
            self.status.set("Не удалось загрузить заявки.")
            messagebox.showerror("Ошибка загрузки", str(exc), parent=self.root)

    @staticmethod
    def _set_text(widget, value):
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert("1.0", value)
        widget.configure(state="disabled")

    def _select_submission(self, _event=None):
        selection = self.listbox.curselection()
        if not selection:
            self.selected = None
            return
        self.selected = self.submissions[selection[0]]
        item = self.selected
        tags = parse_tags(item["tags"])
        details = [
            f"Тип: {item['type'] or 'не указан'}",
            f"Название: {item['name'] or 'не указано'}",
            f"Автор: {item['author'] or 'не указан'}",
            f"Почта: {item['email'] or 'не указана'}",
            f"Версия: {item['version'] or '1.0.0'}",
            f"Теги: {', '.join(tags) if tags else 'не указаны'}",
            "Лицензия при публикации: MIT",
            f"Описание: {item['description'] or 'не указано'}",
            f"Превью: {item['preview'] or 'файл не найден'}",
            f"CSS: {item['css'] or 'файл не найден'}",
        ]
        self._set_text(self.details, "\n\n".join(details))
        self._set_text(self.css_view, "")
        self.show_preview()

    def _set_preview(self, image, message=""):
        self.preview_image = image
        self.preview_label.configure(image="", text=message)
        if image is not None:
            width, height = image.width(), image.height()
            scale = max(1, math.ceil(width / 700), math.ceil(height / 280))
            if scale > 1:
                image = image.subsample(scale, scale)
            self.preview_image = image
            self.preview_label.configure(image=image, text="")

    def reject_submission(self):
        selection = self.listbox.curselection()
        if not selection:
            messagebox.showinfo("Выберите заявку", "Сначала выберите заявку.")
            return
        index = selection[0]
        item = self.submissions[index]
        if not messagebox.askyesno(
            "Отклонить заявку",
            "Отклонить заявку и больше не показывать её в программе? "
            "В таблице будет изменён только статус модерации.",
            parent=self.root,
        ):
            return
        rejected = self.rejected_submissions | {item["key"]}
        try:
            self._save_rejections(rejected)
        except Exception as exc:
            self.status.set("Не удалось отклонить заявку.")
            messagebox.showerror(
                "Не удалось сохранить отклонение", str(exc), parent=self.root
            )
            return
        self.rejected_submissions = rejected
        review_status_error = None
        try:
            self._set_review_status(item["row"], "Отклонено — локально")
        except Exception as exc:
            review_status_error = exc
        self.submissions.pop(index)
        self.selected = None
        self.listbox.delete(index)
        self._set_text(self.details, "")
        self._set_text(self.css_view, "")
        self._set_preview(None, "Заявка отклонена и больше не будет показана.")
        if review_status_error:
            self.status.set("Отклонено локально; статус в таблице не обновлён.")
            messagebox.showwarning(
                "Локально отклонено, статус не обновлён",
                "Заявка скрыта в этой программе, но статус в Google Таблице "
                "не удалось изменить. Обновите статус вручную или удалите её "
                "запись из moderator_rejections.json, чтобы повторить попытку.\n\n"
                f"Подробности: {review_status_error}",
                parent=self.root,
            )
        else:
            self.status.set(
                f"Новых заявок: {len(self.submissions)} "
                "(отклонение записано в таблицу)"
            )
            messagebox.showinfo(
                "Заявка отклонена",
                "В таблице установлен статус «Отклонено — локально». "
                "Ответ автора и файлы в облаке не удалены.",
                parent=self.root,
            )

    def _download_onedrive_file(self, url, max_size):
        encoded_url = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii")
        share_url = (
            "https://api.onedrive.com/v1.0/shares/u!"
            f"{encoded_url.rstrip('=')}/root/content"
        )
        parsed = urlparse(url)
        query = dict(parse_qsl(parsed.query, keep_blank_values=True))
        query["download"] = "1"
        direct_url = urlunparse(parsed._replace(query=urlencode(query)))

        last_error = None
        for download_url in (share_url, direct_url):
            try:
                request = Request(download_url, headers={"User-Agent": "Mozilla/5.0"})
                with urlopen(request, timeout=30) as response:
                    size = response.headers.get("Content-Length")
                    if size and int(size) > max_size:
                        raise ValueError(
                            f"Файл из OneDrive превышает допустимый размер "
                            f"{max_size // (1024 * 1024)} МБ."
                        )
                    content = response.read(max_size + 1)
                    if len(content) > max_size:
                        raise ValueError(
                            f"Файл из OneDrive превышает допустимый размер "
                            f"{max_size // (1024 * 1024)} МБ."
                        )
                    return content
            except ValueError:
                raise
            except (OSError, URLError) as exc:
                last_error = exc
        raise RuntimeError(f"Не удалось скачать файл из OneDrive: {last_error}")

    def _download_answer_file(self, field, extension):
        if not self.selected:
            raise ValueError("Сначала выберите заявку.")
        value = self.selected[field]
        ids = find_drive_ids(value)
        onedrive_url = find_onedrive_url(value)
        if onedrive_url:
            file_id = None
            cache_key = f"onedrive:{onedrive_url}"
        elif len(ids) == 1:
            file_id = ids[0]
            cache_key = file_id
        else:
            raise ValueError(
                f"Для поля «{field}» ожидался один файл Google Drive или "
                "ссылка общего доступа OneDrive."
            )
        if cache_key in self.file_cache:
            name, content = self.file_cache[cache_key]
        elif file_id is None:
            max_size = 12 * 1024 * 1024 if extension == ".png" else 2 * 1024 * 1024
            content = self._download_onedrive_file(onedrive_url, max_size)
            name = f"OneDrive файл{extension}"
            self.file_cache[cache_key] = (name, content)
        else:
            from googleapiclient.http import MediaIoBaseDownload

            metadata = (
                self.drive.files()
                .get(fileId=file_id, fields="id,name,size")
                .execute()
            )
            name = metadata.get("name", "")
            if Path(name).suffix.casefold() != extension:
                raise ValueError(
                    f"Ожидался файл {extension}, получен «{name or 'без имени'}»."
                )
            max_size = 12 * 1024 * 1024 if extension == ".png" else 2 * 1024 * 1024
            size = int(metadata.get("size", 0))
            if size > max_size:
                raise ValueError(
                    f"Файл «{name}» превышает допустимый размер "
                    f"{max_size // (1024 * 1024)} МБ."
                )
            request = self.drive.files().get_media(fileId=file_id)
            stream = BytesIO()
            downloader = MediaIoBaseDownload(stream, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            content = stream.getvalue()
            self.file_cache[cache_key] = (name, content)

        if Path(name).suffix.casefold() != extension:
            raise ValueError(
                f"Ожидался файл {extension}, получен «{name or 'без имени'}»."
            )
        if extension == ".png" and not content.startswith(b"\x89PNG\r\n\x1a\n"):
            raise ValueError(f"Файл «{name}» имеет расширение PNG, но не является PNG.")
        if extension == ".css":
            content.decode("utf-8-sig")
        return name, content

    def open_preview(self):
        try:
            _, content = self._download_answer_file("preview", ".png")
            directory = Path(tempfile.gettempdir()) / "Legion7TV-moderator"
            directory.mkdir(parents=True, exist_ok=True)
            path = directory / "preview.png"
            path.write_bytes(content)
            if hasattr(os, "startfile"):
                os.startfile(str(path))
            else:
                webbrowser.open(path.as_uri())
        except Exception as exc:
            messagebox.showerror("Не удалось открыть превью", str(exc), parent=self.root)

    def show_preview(self):
        try:
            _, content = self._download_answer_file("preview", ".png")
            image = tk.PhotoImage(data=base64.b64encode(content))
            self._set_preview(image)
        except Exception as exc:
            self._set_preview(None, f"Не удалось показать превью:\n{exc}")

    def copy_css(self):
        try:
            _, content = self._download_answer_file("css", ".css")
            css = content.decode("utf-8-sig")
            self._set_text(self.css_view, css)
            self.root.clipboard_clear()
            self.root.clipboard_append(css)
            self.status.set("CSS скопирован в буфер обмена.")
        except Exception as exc:
            messagebox.showerror("Не удалось скопировать CSS", str(exc), parent=self.root)

    def _ask_identity(self, is_update):
        dialog = tk.Toplevel(self.root)
        dialog.title(
            "ID и папка существующей темы" if is_update else "ID и папка новой темы"
        )
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        frame = ttk.Frame(dialog, padding=12)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text="ID темы и имя папки в themes/ (одинаковые):",
        ).grid(
            row=0, column=0, sticky="w", pady=4
        )
        identity = tk.StringVar()
        identity_entry = ttk.Entry(frame, textvariable=identity, width=38)
        identity_entry.grid(row=0, column=1, pady=4)
        result = {}

        def accept():
            theme_id = identity.get().strip()
            result["id"] = theme_id
            result["folder"] = theme_id
            dialog.destroy()

        def cancel():
            dialog.destroy()

        buttons = ttk.Frame(frame)
        buttons.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(buttons, text="Отмена", command=cancel).pack(side="right")
        ttk.Button(buttons, text="Продолжить", command=accept).pack(
            side="right", padx=6
        )
        identity_entry.focus_set()
        self.root.wait_window(dialog)
        return result or None

    def _ask_issue_number(self):
        try:
            last_commit = self._run_git("log", "-1", "--format=%s")
        except Exception as exc:
            messagebox.showerror(
                "Не удалось получить последний commit",
                str(exc),
                parent=self.root,
            )
            return None

        dialog = tk.Toplevel(self.root)
        dialog.title("Commit и push")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        frame = ttk.Frame(dialog, padding=12)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text=f"Последний commit: {last_commit or '(без названия)'}",
            wraplength=560,
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        ttk.Label(frame, text="Номер issue для commit-сообщения #n:").grid(
            row=1, column=0, sticky="w", pady=4
        )
        issue_number = tk.StringVar()
        entry = ttk.Entry(frame, textvariable=issue_number, width=20)
        entry.grid(row=1, column=1, pady=4)
        result = {}

        def accept():
            result["number"] = issue_number.get().strip()
            dialog.destroy()

        buttons = ttk.Frame(frame)
        buttons.grid(row=2, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(buttons, text="Отмена", command=dialog.destroy).pack(side="right")
        ttk.Button(buttons, text="Продолжить", command=accept).pack(
            side="right", padx=6
        )
        entry.focus_set()
        self.root.wait_window(dialog)
        return result.get("number")

    def _run_git(self, *arguments):
        try:
            result = subprocess.run(
                ["git", *arguments],
                cwd=self.repo,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=120,
                check=False,
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise RuntimeError(f"Не удалось выполнить git {' '.join(arguments)}: {exc}") from exc
        if result.returncode:
            details = result.stderr.strip() or result.stdout.strip()
            raise RuntimeError(
                f"git {' '.join(arguments)} завершился с ошибкой"
                + (f":\n{details}" if details else ".")
            )
        return result.stdout.strip()

    def commit_and_push(self):
        issue_number = self._ask_issue_number()
        if not issue_number:
            return
        if not re.fullmatch(r"[1-9][0-9]*", issue_number):
            messagebox.showerror(
                "Некорректный номер",
                "Введите положительный номер issue, например 42.",
                parent=self.root,
            )
            return
        try:
            changes = self._run_git("status", "--short")
        except Exception as exc:
            messagebox.showerror("Git недоступен", str(exc), parent=self.root)
            return
        if not changes:
            messagebox.showinfo(
                "Нет изменений", "Нет изменений для добавления и commit.", parent=self.root
            )
            return
        if not messagebox.askyesno(
            "Подтвердите публикацию",
            "Будут выполнены git add ., git commit и git push. "
            "В commit попадут все текущие изменения в папке репозитория.\n\n"
            f"Сообщение commit: #{issue_number}\n\n{changes}",
            parent=self.root,
        ):
            return

        try:
            self.status.set("Обновляю timestamp, добавляю изменения и создаю commit...")
            self.root.update_idletasks()
            self._update_catalog_timestamp()
            self._run_git("add", ".")
            self._run_git("commit", "-m", f"#{issue_number}")
        except Exception as exc:
            self.status.set("Результат git add/commit не подтверждён.")
            messagebox.showerror(
                "Не удалось завершить git add/commit",
                "Проверьте git status и git log перед повторным запуском.\n\n"
                f"{exc}",
                parent=self.root,
            )
            return

        self.status.set("Отправляю commit в GitHub...")
        self.root.update_idletasks()
        try:
            self._run_git("push")
        except Exception as exc:
            self.status.set("Commit создан; результат push не подтверждён.")
            messagebox.showerror(
                "Commit создан, push не подтверждён",
                f"Commit #{issue_number} создан, но git push не удалось подтвердить.\n"
                f"Проверьте состояние репозитория.\n\n{exc}",
                parent=self.root,
            )
            return

        self.status.set(f"Commit #{issue_number} отправлен.")
        messagebox.showinfo(
            "Готово",
            f"Изменения отправлены: git add ., commit #{issue_number}, git push.",
            parent=self.root,
        )

    def _read_catalog(self):
        catalog_path = self.repo / "catalog.json"
        if not catalog_path.exists():
            raise FileNotFoundError(f"Не найден каталог: {catalog_path}")
        catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
        if not isinstance(catalog.get("themes"), list):
            raise ValueError("В catalog.json ожидается массив `themes`.")
        return catalog_path, catalog

    def _update_catalog_timestamp(self):
        catalog_path, catalog = self._read_catalog()
        catalog["updatedAt"] = local_timestamp()
        content = (json.dumps(catalog, ensure_ascii=False, indent=2) + "\n").encode(
            "utf-8"
        )
        self._atomic_write(catalog_path, content)

    @staticmethod
    def _atomic_write(path, content):
        path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="wb", dir=path.parent, delete=False
        ) as temp_file:
            temp_file.write(content)
            temp_path = Path(temp_file.name)
        try:
            os.replace(temp_path, path)
        except Exception:
            temp_path.unlink(missing_ok=True)
            raise

    def _approve_local(self, item, theme_id, folder, is_update):
        catalog_path, catalog = self._read_catalog()
        themes = catalog["themes"]
        matches = [theme for theme in themes if theme.get("id") == theme_id]
        if is_update and len(matches) != 1:
            raise ValueError(
                "Для обновления ID должен совпасть ровно с одной записью каталога."
            )
        if not is_update and matches:
            raise ValueError("Этот ID уже есть в каталоге; для него выберите обновление.")
        if folder != theme_id:
            raise ValueError("ID темы и имя папки в themes/ должны совпадать.")

        target = self.repo / "themes" / folder
        if is_update:
            existing_folder = slug_from_url(matches[0].get("css", ""))
            if existing_folder != folder:
                raise ValueError(
                    f"Для этой темы в каталоге уже используется папка "
                    f"«{existing_folder or 'не определена'}»."
                )
            if not target.is_dir():
                raise FileNotFoundError(f"Папка темы не найдена: {target}")
        elif target.exists():
            raise FileExistsError(
                f"Папка уже существует: {target}. Выберите другое имя."
            )

        _, css_bytes = self._download_answer_file("css", ".css")
        _, preview_bytes = self._download_answer_file("preview", ".png")
        tags = parse_tags(item["tags"])
        required = {
            "название": item["name"],
            "автор": item["author"],
        }
        missing = [label for label, value in required.items() if not value]
        if missing:
            raise ValueError(
                "В заявке не заполнены обязательные поля: " + ", ".join(missing)
            )
        if len(tags) > 3:
            raise ValueError(
                f"В заявке {len(tags)} тегов, разрешено не больше трёх. "
                "Попросите автора исправить заявку."
            )

        raw_type = item["type"].casefold()
        if ("обнов" in raw_type or "update" in raw_type) != is_update:
            raise ValueError("Тип заявки изменился. Обновите список перед одобрением.")

        css_url = (
            "https://cdn.jsdelivr.net/gh/mykhailkot-pixel/"
            f"Legion7TV-themes@main/themes/{folder}/style.css"
        )
        preview_url = (
            "https://cdn.jsdelivr.net/gh/mykhailkot-pixel/"
            f"Legion7TV-themes@main/themes/{folder}/preview.png"
        )
        entry = {
            "id": theme_id,
            "name": item["name"],
            "author": item["author"],
            "version": item["version"] or "1.0.0",
            "description": item["description"],
            "tags": tags,
            "license": "MIT",
            "css": css_url,
            "preview": preview_url,
        }

        if is_update:
            themes[themes.index(matches[0])] = entry
        else:
            themes.append(entry)
        catalog["updatedAt"] = local_timestamp()
        catalog_bytes = (
            json.dumps(catalog, ensure_ascii=False, indent=2) + "\n"
        ).encode("utf-8")

        created_directory = not target.exists()
        target.mkdir(parents=True, exist_ok=True)
        css_path = target / "style.css"
        preview_path = target / "preview.png"
        old_css = css_path.read_bytes() if css_path.exists() else None
        old_preview = preview_path.read_bytes() if preview_path.exists() else None
        old_catalog = catalog_path.read_bytes()
        try:
            self._atomic_write(css_path, css_bytes)
            self._atomic_write(preview_path, preview_bytes)
            self._atomic_write(catalog_path, catalog_bytes)
        except Exception:
            if old_css is None:
                css_path.unlink(missing_ok=True)
            else:
                self._atomic_write(css_path, old_css)
            if old_preview is None:
                preview_path.unlink(missing_ok=True)
            else:
                self._atomic_write(preview_path, old_preview)
            self._atomic_write(catalog_path, old_catalog)
            if created_directory:
                target.rmdir()
            raise

    def _set_review_status(self, row, value):
        self.sheets.spreadsheets().values().update(
            spreadsheetId=self.config["spreadsheet_id"],
            range=self._sheet_range(
                f"{column_name(self.status_column + 1)}{row}"
            ),
            valueInputOption="RAW",
            body={"values": [[value]]},
        ).execute()

    def approve(self):
        if not self.selected:
            messagebox.showinfo("Выберите заявку", "Сначала выберите заявку.")
            return
        item = self.selected
        raw_type = item["type"].casefold()
        is_update = "обнов" in raw_type or "update" in raw_type
        if not is_update and not ("релиз" in raw_type or "release" in raw_type):
            messagebox.showerror(
                "Неизвестный тип",
                "В поле «Что именно вы публикуете» укажите «Первый релиз» "
                "или «Обновление».",
                parent=self.root,
            )
            return
        try:
            identity = self._ask_identity(is_update)
            if not identity:
                return
            theme_id = identity["id"]
            folder = identity["folder"]
            if not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", theme_id):
                raise ValueError("ID должен содержать только строчные латинские буквы, цифры и дефисы.")

            question = (
                f"Добавить обновление темы «{item['name']}»?"
                if is_update
                else f"Одобрить тему «{item['name']}»?"
            )
            if not messagebox.askyesno(
                "Подтверждение", question, parent=self.root
            ):
                return
            self.status.set("Проверяю файлы и обновляю локальный каталог...")
            self.root.update_idletasks()
            self._approve_local(item, theme_id, folder, is_update)
            try:
                self._set_review_status(item["row"], "Одобрено — локально")
            except Exception as exc:
                messagebox.showwarning(
                    "Локально одобрено, статус не обновлён",
                    "Файлы и catalog.json уже обновлены, но статус в Google Таблице "
                    "не удалось изменить.\nНе одобряйте эту заявку повторно; "
                    "отметьте её вручную.\n\n"
                    f"Подробности: {exc}",
                    parent=self.root,
                )
            else:
                messagebox.showinfo(
                    "Готово",
                    "Файлы записаны в themes/ и catalog.json обновлён.\n"
                    "Изменения пока только локальные — отправьте их в GitHub "
                    "после проверки.",
                    parent=self.root,
                )
            self.refresh()
        except Exception as exc:
            self.status.set("Одобрение не выполнено.")
            messagebox.showerror("Не удалось одобрить заявку", str(exc), parent=self.root)


def main():
    root = tk.Tk()
    root.withdraw()
    try:
        app = ModeratorApp(root)
    except Exception as exc:
        messagebox.showerror("Не удалось запустить модератор", str(exc))
        root.destroy()
        return
    root.deiconify()
    root.mainloop()


if __name__ == "__main__":
    main()
