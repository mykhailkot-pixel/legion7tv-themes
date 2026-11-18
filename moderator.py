import json
import os
import re
import tempfile
import tkinter as tk
import webbrowser
from datetime import date
from io import BytesIO
from pathlib import Path
from tkinter import messagebox, ttk
from urllib.parse import urlparse


ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "moderator_config.json"
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
        self.root.geometry("1100x720")
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
        self.selected = None
        self.file_cache = {}
        self._build_ui()
        self.root.after(200, self.refresh)

    def _build_ui(self):
        toolbar = ttk.Frame(self.root, padding=8)
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="Обновить заявки", command=self.refresh).pack(
            side="left"
        )
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

        ttk.Label(right, text="Данные заявки").pack(anchor="w", pady=(0, 5))
        self.details = tk.Text(right, height=13, wrap="word", state="disabled")
        self.details.pack(fill="x")
        ttk.Label(right, text="Содержимое CSS (просмотр, не выполнение)").pack(
            anchor="w", pady=(10, 5)
        )
        self.css_view = tk.Text(right, wrap="none", state="disabled")
        self.css_view.pack(fill="both", expand=True)

        buttons = ttk.Frame(right)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(
            buttons, text="Открыть превью PNG", command=self.open_preview
        ).pack(side="left")
        ttk.Button(
            buttons, text="Показать CSS", command=self.show_css
        ).pack(side="left", padx=6)
        ttk.Button(
            buttons,
            text="Одобрить и добавить в каталог",
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
                item = {
                    field: self._cell(row, column)
                    for field, column in self.columns.items()
                }
                item["row"] = sheet_row
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
        tags = [tag.strip() for tag in item["tags"].splitlines() if tag.strip()]
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

    def _download_answer_file(self, field, extension):
        if not self.selected:
            raise ValueError("Сначала выберите заявку.")
        ids = find_drive_ids(self.selected[field])
        if len(ids) != 1:
            raise ValueError(
                f"Для поля «{field}» ожидался один файл, найдено: {len(ids)}."
            )
        file_id = ids[0]
        if file_id in self.file_cache:
            name, content = self.file_cache[file_id]
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
            self.file_cache[file_id] = (name, content)

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

    def show_css(self):
        try:
            _, content = self._download_answer_file("css", ".css")
            self._set_text(self.css_view, content.decode("utf-8-sig"))
        except Exception as exc:
            messagebox.showerror("Не удалось прочитать CSS", str(exc), parent=self.root)

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

        ttk.Label(frame, text="ID темы (например, uwu-fluff):").grid(
            row=0, column=0, sticky="w", pady=4
        )
        theme_id = tk.StringVar()
        id_entry = ttk.Entry(frame, textvariable=theme_id, width=38)
        id_entry.grid(row=0, column=1, pady=4)
        ttk.Label(frame, text="Папка в themes/ (например, uwu-fluff):").grid(
            row=1, column=0, sticky="w", pady=4
        )
        folder = tk.StringVar()
        ttk.Entry(frame, textvariable=folder, width=38).grid(
            row=1, column=1, pady=4
        )
        result = {}

        def accept():
            result["id"] = theme_id.get().strip()
            result["folder"] = folder.get().strip()
            dialog.destroy()

        def cancel():
            dialog.destroy()

        buttons = ttk.Frame(frame)
        buttons.grid(row=2, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(buttons, text="Отмена", command=cancel).pack(side="right")
        ttk.Button(buttons, text="Продолжить", command=accept).pack(
            side="right", padx=6
        )
        id_entry.focus_set()
        self.root.wait_window(dialog)
        return result or None

    def _read_catalog(self):
        catalog_path = self.repo / "catalog.json"
        if not catalog_path.exists():
            raise FileNotFoundError(f"Не найден каталог: {catalog_path}")
        catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
        if not isinstance(catalog.get("themes"), list):
            raise ValueError("В catalog.json ожидается массив `themes`.")
        return catalog_path, catalog

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
        tags = [tag.strip() for tag in item["tags"].splitlines() if tag.strip()]
        required = {
            "название": item["name"],
            "автор": item["author"],
            "описание": item["description"],
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
        catalog["updatedAt"] = date.today().isoformat()
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
            if not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", folder):
                raise ValueError("Имя папки должно содержать только строчные латинские буквы, цифры и дефисы.")

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
