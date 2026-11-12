from __future__ import annotations

import argparse
import csv
import hashlib
import re
import sqlite3
import sys
import tkinter as tk
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Any, Iterable


def resource_path(filename: str | Path) -> str:
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    file_path = Path(filename)
    if file_path.is_absolute():
        return str(file_path)
    return str((base_dir / file_path.name).resolve())


try:
    from PIL import Image, ImageTk # type: ignore
except ImportError:
    Image = None
    ImageTk = None

try:
    from openpyxl import load_workbook  # type: ignore[import-not-found]
except ImportError:  # pragma: no cover
    load_workbook = None

DEFAULT_MIN_AGE = 14
DEFAULT_MAX_AGE = 15
DEFAULT_ADMIN_USER = "admin"
DEFAULT_ADMIN_PASSWORD = "admin123"

NAME_KEYWORDS = {
    "фио",
    "fio",
    "full_name",
    "full name",
    "student",
    "name",
    "имя",
    "ученик",
    "полное имя",
}

BIRTHDAY_KEYWORDS = {
    "дата рождения",
    "birthday",
    "birth_date",
    "birth date",
    "date of birth",
    "dob",
    "дата рожд",
    "дата р",
    "др",
    "д.р.",
    "дата рождения ученика",
}

CLASS_KEYWORDS = {
    "класс",
    "class",
    "grade",
    "class number",
    "класс обучения",
    "класс ученика",
    "класс и буква",
}


def normalize_key(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = text.replace(".", " ").replace("-", " ").replace("/", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def find_matching_key(row: dict[str, Any], keywords: set[str]) -> str | None:
    for key in row.keys():
        normalized = normalize_key(key)
        if normalized in keywords:
            return str(key)
    for key in row.keys():
        normalized = normalize_key(key)
        for keyword in keywords:
            if keyword in normalized or normalized in keyword:
                return str(key)
    return None


def parse_date_value(value: Any) -> date | None:
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    if isinstance(value, (int, float)):
        if value > 0:
            try:
                return date(1899, 12, 30) + timedelta(days=float(value))
            except (OverflowError, ValueError):
                pass

    text = str(value).strip()
    if not text:
        return None

    text = text.split(" ")[0]
    patterns = (
        "%d.%m.%Y",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d.%m.%y",
        "%d-%m-%y",
        "%d/%m/%y",
        "%Y.%m.%d",
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d.%m.%Y %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
    )
    for pattern in patterns:
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            continue

    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        pass

    return None


def calculate_age(birthday: date, today: date | None = None) -> int | None:
    if birthday is None:
        return None
    today = today or date.today()
    if birthday > today:
        return None
    return today.year - birthday.year - ((today.month, today.day) < (birthday.month, birthday.day))


def normalize_class_value(value: Any) -> str | None:
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    text = re.sub(r"(?i)\b(class|grade|класс)\b", "", text)
    text = text.replace("_", "").replace("-", "").replace(" ", "")
    text = text.replace(".", "").replace("/", "")
    text = text.replace("№", "")

    match = re.search(r"(\d+)([A-Za-zА-Яа-яЁё]+)?", text)
    if not match:
        return text or None

    number = match.group(1)
    letter = match.group(2)
    if letter:
        return f"{number}{letter.upper()}"
    return number


def detect_student_rows(
    rows: Iterable[dict[str, Any]],
    today: date | None = None,
    min_age: int = DEFAULT_MIN_AGE,
    max_age: int = DEFAULT_MAX_AGE,
) -> list[dict[str, Any]]:
    today = today or date.today()
    result: list[dict[str, Any]] = []

    for row in rows:
        if not isinstance(row, dict):
            continue

        name_key = find_matching_key(row, NAME_KEYWORDS)
        birth_key = find_matching_key(row, BIRTHDAY_KEYWORDS)
        if not name_key or not birth_key:
            continue

        name = row.get(name_key)
        birthday_raw = row.get(birth_key)
        if name is None or birthday_raw is None:
            continue

        birthday = parse_date_value(birthday_raw)
        if birthday is None:
            continue

        age = calculate_age(birthday, today)
        if age is None or not min_age <= age <= max_age:
            continue

        class_key = find_matching_key(row, CLASS_KEYWORDS)
        class_value = row.get(class_key) if class_key else None
        normalized_class = normalize_class_value(class_value) if class_value is not None else None

        result.append(
            {
                "name": str(name).strip(),
                "birthday": birthday.isoformat(),
                "age": age,
                "class_name": normalized_class or "—",
            }
        )

    return result


def iter_rows_from_excel(path: Path):
    if load_workbook is None:
        return []

    workbook = load_workbook(path, read_only=True, data_only=True)
    tables: list[tuple[str, list[dict[str, Any]]]] = []

    for sheet in workbook.worksheets:
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            continue

        header = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
        data_rows: list[dict[str, Any]] = []

        for values in rows[1:]:
            row: dict[str, Any] = {}
            for idx, key in enumerate(header):
                if idx >= len(values):
                    break
                row[key] = values[idx]
            if any(value is not None and str(value).strip() for value in row.values()):
                data_rows.append(row)

        tables.append((sheet.title, data_rows))

    return tables


def iter_rows_from_csv(path: Path):
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file, delimiter=";" if path.suffix.lower() == ".csv" else ",")
        rows = list(reader)
    return [(path.stem, rows)]


def iter_rows_from_txt(path: Path):
    with path.open("r", encoding="utf-8-sig") as file:
        lines = [line.strip() for line in file if line.strip()]

    if not lines:
        return []

    entries: list[dict[str, Any]] = []
    for line in lines:
        parts = re.split(r"\s{2,}|\t+|;|,", line)
        if len(parts) < 2:
            continue
        name, birthday = parts[0], parts[1]
        entries.append({"ФИО": name, "Дата рождения": birthday})
    return [(path.stem, entries)]


def find_student_files(folder: str | Path) -> list[Path]:
    root = Path(folder)
    if not root.exists():
        return []

    extensions = {".csv", ".xlsx", ".xlsm", ".xls", ".txt", ".tsv"}
    files = [path for path in root.rglob("*") if path.is_file() and path.suffix.lower() in extensions]
    return sorted(files)


def scan_path(
    path: str | Path,
    today: date | None = None,
    min_age: int = DEFAULT_MIN_AGE,
    max_age: int = DEFAULT_MAX_AGE,
) -> list[dict[str, Any]]:
    source = Path(path)
    if source.is_file():
        files = [source]
    else:
        files = find_student_files(source)

    matches: list[dict[str, Any]] = []
    for file_path in files:
        if file_path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
            tables = iter_rows_from_excel(file_path)
        elif file_path.suffix.lower() in {".txt", ".tsv"}:
            tables = iter_rows_from_txt(file_path)
        else:
            tables = iter_rows_from_csv(file_path)

        for sheet_name, rows in tables:
            for item in detect_student_rows(rows, today=today, min_age=min_age, max_age=max_age):
                matches.append(
                    {
                        "file": str(file_path),
                        "sheet": sheet_name,
                        "name": item["name"],
                        "birthday": item["birthday"],
                        "age": item["age"],
                        "class_name": item.get("class_name", "—"),
                    }
                )

    return matches


def scan_directory(folder: str | Path, today: date | None = None, min_age: int = DEFAULT_MIN_AGE, max_age: int = DEFAULT_MAX_AGE) -> list[dict[str, Any]]:
    return scan_path(folder, today=today, min_age=min_age, max_age=max_age)


def build_report(matches: list[dict[str, Any]], output_path: str | Path | None = None) -> str:
    rows = [
        ["Файл", "Лист", "Имя", "Класс", "Дата рождения", "Возраст"],
        *[
            [item["file"], item["sheet"], item["name"], item.get("class_name", "—"), item["birthday"], str(item["age"])]
            for item in matches
        ],
    ]

    if output_path:
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        with out.open("w", encoding="utf-8", newline="") as file:
            writer = csv.writer(file, delimiter=";")
            writer.writerows(rows)
        return str(out)

    return "\n".join("; ".join(row) for row in rows)


DB_PATH = Path(__file__).with_name("edutrack.db")
BACKUP_DIR = Path(__file__).resolve().parent / "backups"


def normalize_student_name(name: str | None) -> str:
    cleaned = re.sub(r"\s+", " ", str(name or "").strip())
    return cleaned


def validate_student_record(name: str | None, birthday: str | Any, class_name: str | None = None) -> tuple[str, str, str]:
    cleaned_name = normalize_student_name(name)
    if not cleaned_name:
        raise ValueError("Имя ученика не может быть пустым.")

    raw_birthday = str(birthday or "").strip()
    parsed_birthday = parse_date_value(raw_birthday)
    if parsed_birthday is None:
        raise ValueError(f"Некорректная дата рождения: {birthday!r}")
    if parsed_birthday > date.today():
        raise ValueError(f"Дата рождения не может быть позже сегодняшней даты: {parsed_birthday.isoformat()}")

    normalized_class = str(class_name or "—").strip() or "—"
    return cleaned_name, parsed_birthday.isoformat(), normalized_class


def get_db_connection() -> sqlite3.Connection:
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def ensure_database() -> None:
    connection = get_db_connection()
    try:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                birth_date TEXT NOT NULL,
                current_class TEXT,
                status TEXT DEFAULT 'Активен',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(full_name, birth_date)
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS student_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL,
                year TEXT,
                old_class TEXT,
                new_class TEXT,
                status TEXT,
                reason TEXT,
                changed_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(student_id) REFERENCES students(id)
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                action TEXT NOT NULL,
                details TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        integrity = validate_database_integrity(connection)
        if not integrity["ok"]:
            details = "; ".join(integrity["issues"][:5])
            connection.execute(
                "INSERT INTO audit_log (action, details) VALUES ('database_integrity', ?)",
                (details,),
            )
        connection.commit()
    finally:
        connection.close()

    ensure_employee_accounts()


def validate_database_integrity(connection: sqlite3.Connection | None = None) -> dict[str, Any]:
    should_close = connection is None
    local_connection = connection or get_db_connection()
    issues: list[str] = []

    try:
        result = local_connection.execute("PRAGMA integrity_check").fetchone()
        if result and result[0] != "ok":
            issues.append(f"PRAGMA integrity_check: {result[0]}")
    except sqlite3.DatabaseError as exc:
        issues.append(f"Ошибка проверки целостности: {exc}")

    duplicate_rows = local_connection.execute(
        """
        SELECT full_name, birth_date, COUNT(*) AS total
        FROM students
        GROUP BY full_name, birth_date
        HAVING COUNT(*) > 1
        """
    ).fetchall()
    if duplicate_rows:
        issues.append(f"Обнаружены дубли: {len(duplicate_rows)} комбинаций ФИО+дата рождения")

    invalid_dates = local_connection.execute(
        "SELECT id, full_name, birth_date FROM students WHERE birth_date IS NULL OR birth_date = ''"
    ).fetchall()
    for row in invalid_dates:
        issues.append(f"Пустая дата рождения у записи #{row['id']} ({row['full_name']})")

    future_dates = local_connection.execute(
        "SELECT id, full_name, birth_date FROM students WHERE birth_date IS NOT NULL AND birth_date > date('now')"
    ).fetchall()
    for row in future_dates:
        issues.append(f"Будущая дата рождения у записи #{row['id']} ({row['full_name']})")

    orphan_history = local_connection.execute(
        """
        SELECT h.id, h.student_id
        FROM student_history h
        LEFT JOIN students s ON s.id = h.student_id
        WHERE s.id IS NULL
        """
    ).fetchall()
    if orphan_history:
        issues.append(f"Обнаружены ссылки на несуществующих учеников в истории: {len(orphan_history)}")
    if should_close:
        local_connection.close()

    return {"ok": not issues, "issues": issues, "duplicates": len(duplicate_rows)}


def hash_password(password: str) -> str:
    salt = b"edutrack-v1"
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200000).hex()


def verify_password(password: str, stored_hash: str) -> bool:
    return hash_password(password) == stored_hash


def log_action(action: str, details: str, username: str | None = None) -> None:
    ensure_database()
    connection = get_db_connection()
    try:
        connection.execute(
            "INSERT INTO audit_log (action, details) VALUES (?, ?)",
            (action, f"{username or 'system'}: {details}"),
        )
        connection.commit()
    finally:
        connection.close()


def authenticate_employee(username: str, password: str) -> bool:
    ensure_database()
    connection = get_db_connection()
    try:
        row = connection.execute(
            "SELECT password_hash FROM users WHERE username = ?",
            (str(username or "").strip(),),
        ).fetchone()
        if row is None:
            return False
        return verify_password(password, row["password_hash"])
    finally:
        connection.close()


def ensure_employee_accounts() -> None:
    connection = get_db_connection()
    try:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                role TEXT DEFAULT 'employee',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        user_count = connection.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        if user_count == 0:
            connection.execute(
                "INSERT INTO users (username, password_hash, role) VALUES (?, ?, 'admin')",
                (DEFAULT_ADMIN_USER, hash_password(DEFAULT_ADMIN_PASSWORD)),
            )
        connection.commit()
    finally:
        connection.close()


def read_setting(key: str, default: str | None = None) -> str | None:
    connection = get_db_connection()
    try:
        row = connection.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        return row["value"] if row else default
    finally:
        connection.close()


def write_setting(key: str, value: str) -> None:
    connection = get_db_connection()
    try:
        connection.execute(
            "INSERT INTO settings (key, value, updated_at) VALUES (?, ?, datetime('now')) ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = datetime('now')",
            (key, value),
        )
        connection.commit()
    finally:
        connection.close()


def xor_encrypt(data: bytes, password: str) -> bytes:
    key = hashlib.sha256(password.encode("utf-8")).digest()
    return bytes(byte ^ key[index % len(key)] for index, byte in enumerate(data))


def xor_decrypt(data: bytes, password: str) -> bytes:
    return xor_encrypt(data, password)


def backup_database(output_path: str | Path | None = None, encrypt: bool = False, password: str | None = None) -> str:
    ensure_database()
    if output_path is None:
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = BACKUP_DIR / f"edutrack_backup_{timestamp}.db"
    else:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

    source = sqlite3.connect(DB_PATH)
    try:
        destination = sqlite3.connect(str(output_path))
        try:
            with source:
                source.backup(destination)
        finally:
            destination.close()
    finally:
        source.close()

    if encrypt:
        secret = password or read_setting("backup_password", "edutrack-default-key")
        encrypted_bytes = xor_encrypt(output_path.read_bytes(), str(secret))
        encrypted_path = Path(str(output_path) + ".enc")
        encrypted_path.write_bytes(encrypted_bytes)
        output_path.unlink(missing_ok=True)
        return str(encrypted_path)

    return str(output_path)


def restore_database(backup_path: str | Path, password: str | None = None) -> str:
    backup_file = Path(backup_path)
    if not backup_file.exists():
        raise FileNotFoundError(f"Резервная копия не найдена: {backup_file}")

    if backup_file.suffix.lower() == ".enc":
        secret = password or read_setting("backup_password", "edutrack-default-key")
        decrypted_bytes = xor_decrypt(backup_file.read_bytes(), str(secret))
        temp_path = backup_file.with_suffix(".db")
        temp_path.write_bytes(decrypted_bytes)
        backup_file = temp_path

    connection = get_db_connection()
    try:
        connection.execute("PRAGMA quick_check")
    finally:
        connection.close()

    safe_backup = backup_file.resolve()
    target = DB_PATH.resolve()
    source = sqlite3.connect(str(safe_backup))
    try:
        destination = sqlite3.connect(str(target))
        try:
            with source:
                source.backup(destination)
        finally:
            destination.close()
    finally:
        source.close()

    ensure_database()
    return str(target)


def apply_student_batch(
    records: list[dict[str, Any]],
    operation: str,
    target_class: str | None = None,
    reason: str = "",
) -> dict[str, Any]:
    valid_operations = {"enroll", "withdraw", "transfer"}
    if operation not in valid_operations:
        return {"success": False, "updated_count": 0, "errors": [f"Неверная операция: {operation}"]}
    if operation == "transfer" and not target_class:
        return {"success": False, "updated_count": 0, "errors": ["Для перевода нужно указать новый класс."]}

    connection = get_db_connection()
    updated_count = 0
    errors: list[str] = []
    try:
        connection.execute("BEGIN IMMEDIATE")
        for item in records:
            try:
                name = str(item.get("name", "")).strip()
                birthday = str(item.get("birthday", "")).strip()
                validated_name, validated_birthday, _ = validate_student_record(name, birthday)
            except ValueError as exc:
                errors.append(str(exc))
                continue

            row = connection.execute(
                "SELECT id, current_class, status FROM students WHERE full_name = ? AND birth_date = ?",
                (validated_name, validated_birthday),
            ).fetchone()
            if row is None:
                errors.append(f"Ученик не найден: {validated_name} ({validated_birthday})")
                continue

            student_id = row["id"]
            old_class = row["current_class"] or "—"
            if operation == "enroll":
                new_class = target_class or old_class or "—"
                status = "Активен"
            elif operation == "withdraw":
                new_class = "—"
                status = "Выбыл"
            else:
                new_class = str(target_class).strip() or old_class or "—"
                status = "Активен"

            connection.execute(
                "UPDATE students SET current_class = ?, status = ?, updated_at = datetime('now') WHERE id = ?",
                (new_class, status, student_id),
            )
            connection.execute(
                "INSERT INTO student_history (student_id, year, old_class, new_class, status, reason, changed_at) VALUES (?, ?, ?, ?, ?, ?, datetime('now'))",
                (student_id, str(date.today().year), old_class, new_class, status, reason or operation),
            )
            updated_count += 1

        connection.execute(
            "INSERT INTO audit_log (action, details) VALUES (?, ?)",
            (f"batch_{operation}", f"Обновлено {updated_count} записей. Причина: {reason or 'без комментария'}"),
        )
        connection.commit()
        return {"success": True, "updated_count": updated_count, "errors": errors}
    except Exception as exc:
        connection.rollback()
        return {"success": False, "updated_count": 0, "errors": [str(exc)]}
    finally:
        connection.close()


def mass_update_students(records: list[dict[str, Any]], operation: str, target_class: str | None = None, reason: str = "") -> dict[str, Any]:
    return apply_student_batch(records, operation=operation, target_class=target_class, reason=reason)


def persist_scan_results(matches: list[dict[str, Any]]) -> int:
    if not matches:
        return 0

    ensure_database()
    saved_count = 0

    connection = get_db_connection()
    try:
        connection.execute("BEGIN IMMEDIATE")
        try:
            seen_keys: set[tuple[str, str]] = set()
            for item in matches:
                name = str(item.get("name", "")).strip()
                birthday = str(item.get("birthday", "")).strip()
                new_class = str(item.get("class_name", "—")).strip() or "—"

                try:
                    normalized_name, normalized_birthday, normalized_class = validate_student_record(name, birthday, new_class)
                except ValueError:
                    connection.execute(
                        "INSERT INTO audit_log (action, details) VALUES ('scan_skip', ?)",
                        (f"Пропущена некорректная запись: {name or 'неизвестно'} / {birthday or '—'}",),
                    )
                    continue

                student_key = (normalized_name, normalized_birthday)
                if student_key in seen_keys:
                    connection.execute(
                        "INSERT INTO audit_log (action, details) VALUES ('scan_skip', ?)",
                        (f"Пропущен дубль в пакете: {normalized_name} / {normalized_birthday}",),
                    )
                    continue
                seen_keys.add(student_key)

                row = connection.execute(
                    "SELECT id, current_class FROM students WHERE full_name = ? AND birth_date = ?",
                    (normalized_name, normalized_birthday),
                ).fetchone()

                if row is None:
                    cursor = connection.execute(
                        """
                        INSERT INTO students (full_name, birth_date, current_class, status, created_at, updated_at)
                        VALUES (?, ?, ?, 'Активен', datetime('now'), datetime('now'))
                        """,
                        (normalized_name, normalized_birthday, normalized_class),
                    )
                    student_id = cursor.lastrowid
                    connection.execute(
                        "INSERT INTO student_history (student_id, year, old_class, new_class, status, reason, changed_at) VALUES (?, ?, ?, ?, 'Активен', 'Поступление/обновление по сканированию', datetime('now'))",
                        (student_id, str(date.today().year), "—", normalized_class),
                    )
                    saved_count += 1
                    continue

                student_id = row["id"]
                old_class = row["current_class"] or "—"
                if old_class != normalized_class:
                    connection.execute(
                        "INSERT INTO student_history (student_id, year, old_class, new_class, status, reason, changed_at) VALUES (?, ?, ?, ?, 'Активен', 'Автоматическое обновление класса', datetime('now'))",
                        (student_id, str(date.today().year), old_class, normalized_class),
                    )

                connection.execute(
                    "UPDATE students SET current_class = ?, status = 'Активен', updated_at = datetime('now') WHERE id = ?",
                    (normalized_class, student_id),
                )
                saved_count += 1

            connection.execute(
                "INSERT INTO audit_log (action, details) VALUES ('scan_results', ?)",
                (f"Сохранено {saved_count} записей в реестр {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",),
            )
            connection.commit()
        except Exception:
            connection.rollback()
            raise
    finally:
        connection.close()

    return saved_count


class StudentFinderApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        ensure_database()
        self.current_user = DEFAULT_ADMIN_USER
        if not self._show_login_dialog():
            self.destroy()
            return

        self.title("EduTrack")
        self.geometry("1180x780")
        self.minsize(1010, 660)
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except Exception:
            pass
        self.results: list[dict[str, Any]] = []

        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), foreground="#ffffff", background="#1f3c88")
        self.style.map("Primary.TButton", background=[("active", "#16386d")])
        self.style.configure("Secondary.TButton", background="#e5edf8", foreground="#17365a")
        self.style.configure("Quiet.TButton", background="#f2f5fb", foreground="#17365a")
        self.style.configure("Card.TFrame", background="#ffffff")
        self.style.configure("Header.TFrame", background="#edf2fb")
        self.style.configure("Panel.TFrame", background="#f8fafd")
        self.style.configure("Body.TLabel", font=("Segoe UI", 10), background="#ffffff", foreground="#1f2937")
        self.style.configure("Filter.TLabel", font=("Segoe UI", 10, "bold"), background="#ffffff", foreground="#1b2a41")
        self.style.configure("Stat.TLabel", font=("Segoe UI", 11, "bold"), background="#ffffff", foreground="#1f3c88")
        self.style.configure("TNotebook", background="#edf2fb")
        self.style.configure("TNotebook.Tab", padding=(16, 8), background="#dfe7f5", foreground="#17365a", font=("Segoe UI", 10, "bold"))
        self.style.map("TNotebook.Tab", background=[("selected", "#ffffff"), ("active", "#eaf0ff")], foreground=[("selected", "#17365a"), ("active", "#17365a")])
        self.style.configure("Treeview", background="#ffffff", fieldbackground="#ffffff", foreground="#111827")
        self.style.configure("Treeview.Heading", background="#e2ebff", foreground="#1b2a41", font=("Segoe UI", 10, "bold"))
        self.style.map("Treeview", background=[("selected", "#dfeaff")], foreground=[("selected", "#111827")])

        self.configure(bg="#edf2ff")

        self.background_image = self._load_background_image(opacity=0.30)
        self.background_label = tk.Label(self, image=self.background_image, bg="#edf2ff", bd=0, highlightthickness=0)
        self.background_label.place(x=0, y=0, relwidth=1, relheight=1)
        self.background_label.lower()
        self.background_label.configure(cursor="arrow")

        header = ttk.Frame(self, padding=(22, 18, 22, 14), style="Header.TFrame")
        header.pack(fill="x")

        logo_path = Path(resource_path("image.png"))
        if Image is not None and ImageTk is not None:
            self.logo_photo = self._load_background_image(opacity=0.30, max_size=(110, 110), source_path=logo_path)
        else:
            self.logo_photo = tk.PhotoImage(file=str(logo_path))

        logo_label = tk.Label(header, image=self.logo_photo, bg="#edf2fb", bd=0, highlightthickness=0)
        logo_label.image = self.logo_photo
        logo_label.pack(side="left", padx=(0, 18))

        title_wrap = tk.Frame(header, bg="#edf2fb")
        title_wrap.pack(side="left", fill="y")
        tk.Label(title_wrap, text="EduTrack", font=("Segoe UI", 22, "bold"), bg="#edf2fb", fg="#0f172a").pack(anchor="w")
        tk.Label(title_wrap, text="Школьный реестр учащихся и мониторинг обучения", font=("Segoe UI", 10), bg="#edf2fb", fg="#475569").pack(anchor="w")

        main_card = ttk.Frame(self, padding=18, style="Card.TFrame")
        main_card.pack(fill="both", expand=True)

        notebook = ttk.Notebook(main_card)
        notebook.pack(fill="both", expand=True)

        registry_tab = ttk.Frame(notebook, padding=16, style="Card.TFrame")
        analytics_tab = ttk.Frame(notebook, padding=16, style="Card.TFrame")
        processes_tab = ttk.Frame(notebook, padding=16, style="Card.TFrame")
        notebook.add(registry_tab, text="Реестр")
        notebook.add(analytics_tab, text="Аналитика")
        notebook.add(processes_tab, text="Процессы")

        self.path_var = tk.StringVar()
        ttk.Label(registry_tab, text="Путь к папке или файлу:", style="Body.TLabel").pack(anchor="w")

        path_row = ttk.Frame(registry_tab, style="Card.TFrame")
        path_row.pack(fill="x", pady=(8, 12))
        self.path_entry = ttk.Entry(path_row)
        self.path_entry.configure(textvariable=self.path_var, width=80)
        self.path_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(path_row, text="Обзор", style="Secondary.TButton", command=self.select_path).pack(side="left", padx=(10, 0))

        filter_row = ttk.Frame(registry_tab, style="Card.TFrame")
        filter_row.pack(fill="x", pady=(0, 12))

        ttk.Label(filter_row, text="Возраст от", style="Filter.TLabel").pack(side="left")
        self.min_age_var = tk.StringVar(value=str(DEFAULT_MIN_AGE))
        age_min_spin = ttk.Spinbox(filter_row, from_=1, to=30, textvariable=self.min_age_var, width=5)
        age_min_spin.pack(side="left", padx=(8, 18))

        ttk.Label(filter_row, text="до", style="Filter.TLabel").pack(side="left")
        self.max_age_var = tk.StringVar(value=str(DEFAULT_MAX_AGE))
        age_max_spin = ttk.Spinbox(filter_row, from_=1, to=30, textvariable=self.max_age_var, width=5)
        age_max_spin.pack(side="left", padx=(8, 18))
        ttk.Label(filter_row, text="лет", style="Filter.TLabel").pack(side="left")

        buttons = ttk.Frame(registry_tab, style="Card.TFrame")
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Найти учеников", style="Primary.TButton", command=self.run_search).pack(side="left")
        ttk.Button(buttons, text="Экспорт CSV", style="Secondary.TButton", command=self.export_report).pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="Массовое зачисление", style="Quiet.TButton", command=lambda: self._bulk_action("enroll")).pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="Массовый перевод", style="Quiet.TButton", command=lambda: self._bulk_action("transfer")).pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="Массовое выбытие", style="Quiet.TButton", command=lambda: self._bulk_action("withdraw")).pack(side="left", padx=(8, 0))

        self.status_var = tk.StringVar(value="Готов к формированию списков и аналитики по учащимся.")
        ttk.Label(registry_tab, textvariable=self.status_var, style="Body.TLabel").pack(anchor="w", pady=(12, 8))

        table_panel = ttk.Frame(registry_tab, style="Panel.TFrame", padding=10)
        table_panel.pack(fill="both", expand=True)

        columns = ("file", "sheet", "name", "class_name", "birthday", "age")
        self.result_tree = ttk.Treeview(table_panel, columns=columns, show="headings", height=14)
        self.result_tree.heading("file", text="Файл")
        self.result_tree.heading("sheet", text="Лист")
        self.result_tree.heading("name", text="ФИО")
        self.result_tree.heading("class_name", text="Класс")
        self.result_tree.heading("birthday", text="Дата рождения")
        self.result_tree.heading("age", text="Возраст")
        self.result_tree.column("file", width=220, anchor="w")
        self.result_tree.column("sheet", width=90, anchor="center")
        self.result_tree.column("name", width=230, anchor="w")
        self.result_tree.column("class_name", width=90, anchor="center")
        self.result_tree.column("birthday", width=140, anchor="center")
        self.result_tree.column("age", width=80, anchor="center")
        self.result_tree.pack(side="left", fill="both", expand=True)

        table_scroll = ttk.Scrollbar(table_panel, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=table_scroll.set)
        table_scroll.pack(side="right", fill="y")

        analytics_panel = ttk.Frame(analytics_tab, style="Panel.TFrame", padding=16)
        analytics_panel.pack(fill="both", expand=True)

        stat_row = ttk.Frame(analytics_panel, style="Panel.TFrame")
        stat_row.pack(fill="x")

        self.total_label = tk.Label(stat_row, text="Всего: 0", font=("Segoe UI", 16, "bold"), bg="#f8fafd", fg="#1f3c88")
        self.total_label.pack(side="left", padx=(0, 18))
        self.range_label = tk.Label(stat_row, text="Диапазон: 14–15", font=("Segoe UI", 11), bg="#f8fafd", fg="#334155")
        self.range_label.pack(side="left")

        self.analytics_text = tk.Text(analytics_panel, wrap="word", height=15, font=("Segoe UI", 10), bg="#ffffff", fg="#0f172a", relief="flat", borderwidth=0)
        self.analytics_text.pack(fill="both", expand=True, pady=(18, 0))
        self.analytics_text.configure(state="disabled")
        self._refresh_analytics()

        process_scroll = tk.Scrollbar(processes_tab, orient="vertical")
        process_canvas = tk.Canvas(processes_tab, bg="#f8fafd", highlightthickness=0, yscrollcommand=process_scroll.set)
        process_scroll.configure(command=process_canvas.yview)
        process_scroll.pack(side="right", fill="y")
        process_canvas.pack(side="left", fill="both", expand=True)

        process_inner = tk.Frame(process_canvas, bg="#f8fafd", padx=18, pady=16)
        process_canvas.create_window((0, 0), window=process_inner, anchor="nw")
        process_inner.bind("<Configure>", lambda event: process_canvas.configure(scrollregion=process_canvas.bbox("all")))

        self._add_process_section(
            process_inner,
            "# 1. Жизненный цикл ученика",
            [
                ("Жизненный цикл записи", None),
                ("1. Поступление - сотрудник создаёт карточку поступающего или импортирует данные из документов.", "Новая запись"),
                ("2. Зачисление - после подтверждения зачисления приказом ученик автоматически становится действующим учащимся.", "Активен"),
                ("3. Обучение - ученик числится в классе, участвует в ежегодном тестировании и получает историю результатов.", "Активен"),
                ("4. Перевод или окончание обучения - система получает информацию о соответствующем приказе и меняет статус.", "Выбыл"),
                ("5. Архив - карточка исключается из действующего списка, но сохраняется в архиве с историей.", "Архив"),
            ],
            accent="#1f3c88",
        )

        self._add_process_section(
            process_inner,
            "# 2. Откуда брать информацию",
            [
                ("Источник | Назначение", None),
                ("Ручное добавление | Зачисление отдельных учеников", "Внутренний"),
                ("Импорт Excel | Массовое добавление поступающих", "Массовый"),
                ("Интеграция с учётной системой школы | Автоматическое обновление сведений", "Авто"),
                ("В перспективе третий вариант наиболее интересный. Если в учреждении уже используется электронная система учёта контингента, EduRegistry может получать из неё сведения через API или регулярный обмен файлами.", None),
                ("Важно: автоматическое изменение статуса должно происходить только после подтверждения официального основания, а не просто из-за отсутствия ученика в очередной выгрузке.", None),
            ],
            accent="#1d5a8c",
        )

        self._add_process_section(
            process_inner,
            "# 3. Какую структуру базы данных сделать",
            [
                ("Students - карточки учеников", "Основная"),
                ("- Уникальный внутренний ID", None),
                ("- ФИО", None),
                ("- Дата рождения", None),
                ("- Текущий класс", None),
                ("- Дата зачисления", None),
                ("- Текущий статус", None),
                ("EnrollmentHistory - история обучения", "История"),
                ("- ID ученика", None),
                ("- Класс и учебный год", None),
                ("- Дата зачисления", None),
                ("- Дата выбытия", None),
                ("- Причина изменения статуса", None),
                ("- Номер и дата приказа", None),
                ("Testing - тестирования", "Аналитика"),
                ("- ID ученика", None),
                ("- Дата тестирования", None),
                ("- Учебный год", None),
                ("- Результаты", None),
                ("- Статус прохождения", None),
                ("AuditLog - журнал изменений", "Журнал"),
                ("- Кто изменил запись", None),
                ("- Что именно изменил", None),
                ("- Когда произошло изменение", None),
                ("- Старое и новое значение", None),
            ],
            accent="#2f6f4f",
        )

        self._add_process_section(
            process_inner,
            "# 3.1. Модель реестра и принципы работы",
            [
                ("1. Хранит всех когда-либо зачисленных учеников.", "Реестр"),
                ("2. Автоматически поддерживает актуальность классов.", "Класс"),
                ("3. Позволяет регистрировать новых поступающих.", "Поступление"),
                ("4. Архивирует выпускников и выбывших.", "Архив"),
                ("5. Формирует списки для ежегодного тестирования.", "Тестирование"),
                ("6. Ведёт историю каждого ученика.", "История"),
            ],
            accent="#1f3c88",
        )

        self._add_process_section(
            process_inner,
            "# 3.2. Технологический стек",
            [
                ("Python 3.13 | Основной язык разработки", "Python"),
                ("Tkinter / ttk | Интерфейс приложения", "UI"),
                ("SQLite | База данных и хранение реестра", "DB"),
                ("openpyxl | Импорт Excel-данных", "Excel"),
                ("Pillow | Работа с изображениями и иконками", "Image"),
                ("PyInstaller | Сборка исполняемого файла EXE", "EXE"),
            ],
            accent="#2d5b9a",
        )

        self._add_process_section(
            process_inner,
            "# 3.3. Первая версия: шесть модулей",
            [
                ("Главная | Статистика контингента и общий мониторинг", "Главная"),
                ("Ученики | Поиск, карточки, редактирование и связанная история", "Ученики"),
                ("Поступление | Добавление новых учеников и первичная регистрация", "Поступление"),
                ("Тестирование | Формирование списков и результатов ежегодных проверок", "Тестирование"),
                ("Архив | Выпускники, выбывшие и сохранённая история обучения", "Архив"),
                ("Настройки | Бэкапы, импорт, экспорт и административные параметры", "Настройки"),
            ],
            accent="#2f6f4f",
        )

        self._add_process_section(
            process_inner,
            "# 3.4. Система учебных годов",
            [
                ("2026–2027: ученик числится в 8А.", "Учебный год"),
                ("2027–2028: автоматически переводится в 9А после подтверждения перехода.", "Переход"),
                ("2028–2029: становится выпускником после подтверждения окончания обучения.", "Выпуск"),
                ("При этом ID ученика остаётся неизменным на протяжении всего периода работы системы.", "ID"),
            ],
            accent="#a75f2a",
        )

        self._add_process_section(
            process_inner,
            "# 4. Как автоматизировать процессы",
            [
                ("Автоматическое зачисление - создание карточки из поступающих с последующим подтверждением.", "Авто"),
                ("Автоматический перевод - при переходе в следующий класс обновляется класс, но не создаётся новая карточка.", "Переход"),
                ("Автоматическая архивация - после подтверждения выбытия ученик перемещается в архив.", "Архив"),
                ("Ежегодное обновление - подготовка списков к новому учебному году и проверка актуальности данных.", "Год"),
                ("Контроль изменений - уведомления о новых поступлениях, выбытиях и изменениях в реестре.", "Контроль"),
            ],
            accent="#a75f2a",
        )

    def _show_login_dialog(self) -> bool:
        dialog = tk.Toplevel(self)
        dialog.title("Авторизация EduTrack")
        dialog.geometry("340x180")
        dialog.transient(self)
        dialog.grab_set()

        tk.Label(dialog, text="Логин", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=18, pady=(18, 4))
        username_var = tk.StringVar(value=DEFAULT_ADMIN_USER)
        tk.Entry(dialog, textvariable=username_var, width=28).pack(anchor="w", padx=18, fill="x")

        tk.Label(dialog, text="Пароль", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=18, pady=(12, 4))
        password_var = tk.StringVar()
        tk.Entry(dialog, textvariable=password_var, show="*", width=28).pack(anchor="w", padx=18, fill="x")

        def attempt_login() -> None:
            username = username_var.get().strip()
            password = password_var.get()
            if authenticate_employee(username, password):
                self.current_user = username
                dialog.destroy()
                return
            messagebox.showerror("Ошибка доступа", "Неверный логин или пароль.")

        tk.Button(dialog, text="Войти", command=attempt_login, width=20).pack(pady=(18, 0))
        self.wait_window(dialog)
        return self.current_user is not None and self.current_user != ""

    def _confirm_dangerous_action(self, action_text: str) -> bool:
        return messagebox.askyesno("Подтверждение операции", f"Вы действительно хотите выполнить: {action_text}?")

    def _bulk_action(self, operation: str) -> None:
        selected = self.result_tree.selection()
        if not selected:
            messagebox.showwarning("Нет выбора", "Сначала выберите строки в таблице.")
            return

        records = []
        for item_id in selected:
            values = self.result_tree.item(item_id, "values")
            if len(values) < 6:
                continue
            records.append({"name": values[2], "birthday": values[4], "class_name": values[3]})

        if not records:
            messagebox.showwarning("Нет данных", "Выделенные строки не содержат корректных записей.")
            return

        action_label = {"enroll": "зачисление", "transfer": "перевод", "withdraw": "выбытие"}[operation]
        if not self._confirm_dangerous_action(action_label):
            return

        target_class = None
        if operation == "transfer":
            target_class = simpledialog.askstring("Перевод в класс", "Укажите новый класс:", parent=self)
            if target_class is None:
                return

        result = apply_student_batch(records, operation=operation, target_class=target_class, reason=f"{action_label} через интерфейс ({self.current_user})")
        if not result["success"]:
            messagebox.showerror("Ошибка операции", "; ".join(result["errors"]))
            return

        self.status_var.set(f"Выполнена массовая операция: {action_label}. Обновлено {result['updated_count']} записей.")
        messagebox.showinfo("Готово", f"Операция выполнена: {result['updated_count']} записей.")
        self.run_search()

    def _status_badge(self, parent: tk.Misc, text: str, bg: str, fg: str = "#ffffff") -> None:
        label = tk.Label(parent, text=text, bg=bg, fg=fg, font=("Segoe UI", 9, "bold"), padx=8, pady=2, bd=0)
        label.pack(side="right", padx=(8, 0))

    def _add_process_section(self, parent: tk.Misc, title: str, lines: list[tuple[str, str | None]], accent: str = "#1f3c88") -> None:
        card = tk.Frame(parent, bg="#ffffff", padx=16, pady=12, highlightbackground=accent, highlightthickness=2)
        card.pack(fill="x", pady=(0, 12))

        title_row = tk.Frame(card, bg="#ffffff")
        title_row.pack(fill="x", pady=(0, 8))
        tk.Label(title_row, text=title, bg="#ffffff", fg="#17365a", font=("Segoe UI", 12, "bold"), anchor="w").pack(side="left")

        for line, badge in lines[:1]:
            if badge:
                self._status_badge(title_row, badge, {"Новая запись": "#dfeeff", "Активен": "#dff7e8", "Выбыл": "#fde7e7", "Архив": "#f3e8ff", "Основная": "#dfe9ff", "История": "#dff3f4", "Аналитика": "#f9e7cf", "Журнал": "#efe7ff", "Авто": "#dfeaf9", "Массовый": "#dfeaff", "Внутренний": "#dff0ff", "Переход": "#e8f8e3", "Год": "#fceae6", "Контроль": "#e5e7eb"}.get(badge, "#eaf0ff"), fg="#17365a")
                if line.strip():
                    row = tk.Frame(card, bg="#ffffff")
                    row.pack(fill="x", pady=2)
                    tk.Label(row, text=line, bg="#ffffff", fg="#1f2937", justify="left", anchor="w", font=("Segoe UI", 10), wraplength=900).pack(anchor="w")
                    continue

        for line, badge in lines[1:]:
            row = tk.Frame(card, bg="#ffffff")
            row.pack(fill="x", pady=2)
            tk.Label(row, text="•", bg="#ffffff", fg=accent, font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 8))
            txt = tk.Label(row, text=line, bg="#ffffff", fg="#1f2937", justify="left", anchor="w", font=("Segoe UI", 10), wraplength=900)
            txt.pack(side="left")
            if badge:
                self._status_badge(row, badge, {"Новая запись": "#dfeeff", "Активен": "#dff7e8", "Выбыл": "#fde7e7", "Архив": "#f3e8ff", "Основная": "#dfe9ff", "История": "#dff3f4", "Аналитика": "#f9e7cf", "Журнал": "#efe7ff", "Авто": "#dfeaf9", "Массовый": "#dfeaff", "Внутренний": "#dff0ff", "Переход": "#e8f8e3", "Год": "#fceae6", "Контроль": "#e5e7eb"}.get(badge, "#eaf0ff"), fg="#17365a")

    def _refresh_analytics(self) -> None:
        count = len(self.results)
        min_age = self.min_age_var.get() if self.min_age_var.get() else DEFAULT_MIN_AGE
        max_age = self.max_age_var.get() if self.max_age_var.get() else DEFAULT_MAX_AGE
        self.total_label.config(text=f"Всего: {count}")
        self.range_label.config(text=f"Диапазон: {min_age}–{max_age}")

        summary = (
            "Аналитика по реестру\n\n"
            f"- Найдено записей: {count}\n"
            f"- Порог отбора: {min_age}–{max_age} лет\n"
            "- Состояние: школа в режиме актуализации данных\n\n"
            "Ключевые показатели:\n"
            "• Контроль за зачислением и переводом\n"
            "• Ежегодное обновление списков\n"
            "• Архивация выбытия и завершения обучения\n"
            "• Отслеживание истории результатов\n"
        )
        self.analytics_text.configure(state="normal")
        self.analytics_text.delete("1.0", tk.END)
        self.analytics_text.insert("1.0", summary)
        self.analytics_text.configure(state="disabled")

    def _build_school_process_text(self) -> str:
        return (
            "# 1. Жизненный цикл ученика\n\n"
            "Жизненный цикл записи\n"
            "1. Поступление — сотрудник создаёт карточку поступающего или импортирует данные из документов. Новая запись\n"
            "2. Зачисление — после подтверждения зачисления приказом ученик автоматически становится действующим учащимся. Активен\n"
            "3. Обучение — ученик числится в классе, участвует в ежегодном тестировании, получает историю результатов. Активен\n"
            "4. Перевод или окончание обучения - система получает информацию о соответствующем приказе и меняет статус. Завершил обучение / Выбыл\n"
            "5. Архив - карточка исключается из действующего списка, но сохраняется в архиве с историей. Архивная запись\n\n"
            "# 2. Откуда брать информацию\n\n"
            "Источник | Назначение\n"
            "Ручное добавление | Зачисление отдельных учеников\n"
            "Импорт Excel | Массовое добавление поступающих\n"
            "Интеграция с учётной системой школы | Автоматическое обновление сведений\n\n"
            "В перспективе третий вариант наиболее интересный. Если в учреждении уже используется электронная система учёта контингента, EduRegistry может получать из неё сведения через API или регулярный обмен файлами.\n\n"
            "Важно: автоматическое изменение статуса должно происходить только после подтверждения официального основания, а не просто из-за отсутствия ученика в очередной выгрузке.\n\n"
            "# 3. Какую структуру базы данных сделать\n\n"
            "Students - карточки учеников\n"
            "- Уникальный внутренний ID\n"
            "- ФИО\n"
            "- Дата рождения\n"
            "- Текущий класс\n"
            "- Дата зачисления\n"
            "- Текущий статус\n\n"
            "EnrollmentHistory - история обучения\n"
            "- ID ученика\n"
            "- Класс и учебный год\n"
            "- Дата зачисления\n"
            "- Дата выбытия\n"
            "- Причина изменения статуса\n"
            "- Номер и дата приказа\n\n"
            "Testing - тестирования\n"
            "- ID ученика\n"
            "- Дата тестирования\n"
            "- Учебный год\n"
            "- Результаты\n"
            "- Статус прохождения\n\n"
            "AuditLog - журнал изменений\n"
            "- Кто изменил запись\n"
            "- Что именно изменил\n"
            "- Когда произошло изменение\n"
            "- Старое и новое значение\n\n"
            "# 4. Как автоматизировать процессы\n\n"
            "Автоматическое зачисление - создание карточки из поступающих с последующим подтверждением.\n"
            "Автоматический перевод - при переходе в следующий класс обновляется класс, но не создаётся новая карточка.\n"
            "Автоматическая архивация - после подтверждения выбытия ученик перемещается в архив.\n"
            "Ежегодное обновление - подготовка списков к новому учебному году и проверка актуальности данных.\n"
            "Контроль изменений - уведомления о новых поступлениях, выбытиях и изменениях в реестре.\n"
        )

    def _load_background_image(self, opacity: float, max_size: tuple[int, int] | None = None, source_path: str | Path | None = None) -> Any:
        image_path = Path(source_path) if source_path is not None else Path(resource_path("image.png"))

        if Image is None or ImageTk is None:
            return tk.PhotoImage(file=str(image_path))

        with Image.open(image_path) as img:
            rgba = img.convert("RGBA")
            if max_size is not None:
                rgba.thumbnail(max_size, Image.Resampling.LANCZOS)
            alpha = rgba.getchannel("A")
            new_alpha = alpha.point(lambda p: int(p * opacity))
            rgba.putalpha(new_alpha)
            return ImageTk.PhotoImage(rgba)

    def select_path(self) -> None:
        path = filedialog.askdirectory(title="Выберите папку с файлами")
        if path:
            self.path_var.set(path)
            return

        file = filedialog.askopenfilename(
            title="Выберите Excel/CSV/TXT файл",
            filetypes=[
                ("Все поддерживаемые файлы", "*.xlsx *.xls *.xlsm *.csv *.txt *.tsv"),
                ("Excel", "*.xlsx *.xls *.xlsm"),
                ("CSV/TXT", "*.csv *.txt *.tsv"),
            ],
        )
        if file:
            self.path_var.set(file)

    def run_search(self) -> None:
        source = self.path_var.get().strip()
        if not source:
            messagebox.showwarning("Путь не выбран", "Выберите папку или файл.")
            return

        try:
            min_age = int(self.min_age_var.get())
            max_age = int(self.max_age_var.get())
        except ValueError:
            messagebox.showwarning("Некорректный возраст", "Возраст должен быть числом.")
            return

        if min_age > max_age:
            messagebox.showwarning("Некорректный диапазон", "Минимальный возраст не может быть больше максимального.")
            return

        matches = scan_path(source, min_age=min_age, max_age=max_age)
        self.results = matches
        persisted = persist_scan_results(matches)
        self._refresh_analytics()

        self.result_tree.delete(*self.result_tree.get_children())
        if not matches:
            self.status_var.set(f"По пути {source} подходящих учеников {min_age}–{max_age} лет не найдено.")
            self._set_result_text("Подходящие ученики не найдены.")
            return

        self.status_var.set(
            f"Найдено {len(matches)} учеников в диапазоне {min_age}–{max_age} лет. В реестр сохранено {persisted} записей."
        )
        for item in matches:
            self.result_tree.insert(
                "",
                "end",
                values=(item["file"], item["sheet"], item["name"], item.get("class_name", "—"), item["birthday"], item["age"]),
            )

    def export_report(self) -> None:
        if not self.results:
            messagebox.showinfo("Нет данных", "Сначала выполните поиск.")
            return

        output_path = filedialog.asksaveasfilename(
            title="Сохранить CSV-отчёт",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not output_path:
            return

        path = build_report(self.results, output_path)
        messagebox.showinfo("Готово", f"Отчёт сохранён: {path}")

    def _set_result_text(self, text: str) -> None:
        self.result_tree.delete(*self.result_tree.get_children())
        if text and "Подходящие ученики не найдены" in text:
            self.result_tree.insert("", "end", values=("—", "—", "Подходящие ученики не найдены", "—", "—", "—"))
        elif text:
            self.result_tree.insert("", "end", values=("—", "—", text, "—", "—", "—"))


def run_cli() -> None:
    parser = argparse.ArgumentParser(description="Находит учеников по диапазону возраста в Excel/CSV/TXT таблицах.")
    parser.add_argument("path", nargs="?", default=None, help="Папка или файл с таблицами учеников")
    parser.add_argument("--folder", dest="folder_option", default=None, help="Папка с файлами, где лежат списки учеников")
    parser.add_argument("--output", default=None, help="Путь для CSV-отчёта")
    parser.add_argument("--min-age", type=int, default=DEFAULT_MIN_AGE, help="Минимальный возраст для отбора")
    parser.add_argument("--max-age", type=int, default=DEFAULT_MAX_AGE, help="Максимальный возраст для отбора")
    args = parser.parse_args()

    if args.min_age > args.max_age:
        raise SystemExit("--min-age не может быть больше --max-age")

    path_value = args.folder_option or args.path or "."
    source = Path(path_value)
    matches = scan_path(source, min_age=args.min_age, max_age=args.max_age)

    if not matches:
        print(f"По пути {source} подходящих учеников {args.min_age}–{args.max_age} лет не найдено.")
        return

    print(f"Найдено {len(matches)} учеников {args.min_age}–{args.max_age} лет:")
    for item in matches:
        klass = item.get("class_name", "—")
        print(f"- {item['name']} ({klass}) : {item['birthday']} ({item['age']} лет) — {item['file']}")

    if args.output:
        report_path = build_report(matches, args.output)
        print(f"\nОтчёт сохранён: {report_path}")


def main() -> None:
    if len(sys.argv) > 1:
        run_cli()
        return

    app = StudentFinderApp()
    app.mainloop()


if __name__ == "__main__":
    main()
