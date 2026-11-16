from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
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
DEFAULT_UPDATE_REPO = "https://github.com/c00l-k1D/EduTrack"
APP_VERSION = "1.0.0"

USER_AGREEMENT_TEXT = '''# ПОЛЬЗОВАТЕЛЬСКОЕ СОГЛАШЕНИЕ

## об использовании программного обеспечения EduTrack
**Редакция 1.0**

### 1. Общие положения
1.1. Настоящее Пользовательское соглашение (далее — Соглашение) определяет условия использования программного обеспечения EduTrack (далее — Программа).

1.2. EduTrack представляет собой специализированное настольное программное обеспечение, предназначенное для автоматизации учёта учащихся образовательных учреждений, управления сведениями о контингенте, организации тестирования и формирования отчётности.

1.3. Программа предназначена для использования уполномоченными сотрудниками образовательных учреждений в рамках выполнения ими должностных обязанностей.

1.4. Использование Программы означает согласие пользователя с условиями настоящего Соглашения.

1.5. Администрация Программы оставляет за собой право изменять настоящее Соглашение с публикацией актуальной редакции.

### 2. Назначение и функциональность Программы
2.1. Программа EduTrack предназначена для выполнения следующих задач:

- Формирование и ведение единого реестра учащихся.
- Регистрация новых поступающих.
- Учёт движения контингента учащихся.
- Изменение сведений об учащихся и их образовательном статусе.
- Формирование списков участников тестирования.
- Хранение истории обучения и результатов тестирования.
- Архивирование сведений о выпускниках и выбывших учащихся.
- Импорт и экспорт структурированных данных.
- Формирование статистических и аналитических отчётов.

2.2. Функциональные возможности Программы могут изменяться и дополняться в последующих версиях.

### 3. Права и обязанности пользователя
3.1. Пользователь имеет право:

- Использовать Программу по её прямому назначению.
- Вносить, редактировать и систематизировать сведения в пределах предоставленных полномочий.
- Формировать и экспортировать отчёты.
- Создавать резервные копии базы данных.
- Использовать доступные инструменты восстановления информации.
- Обращаться за технической поддержкой при обнаружении неисправностей.

3.2. Пользователь обязуется:

- Использовать Программу исключительно в законных целях.
- Вносить достоверную и актуальную информацию.
- Соблюдать установленные учреждением правила обработки персональных данных.
- Не предоставлять доступ к базе данных неуполномоченным лицам.
- Обеспечивать сохранность учётных данных и резервных копий.
- Не использовать сведения об учащихся в личных или коммерческих целях.
- Своевременно актуализировать информацию о зачислении, переводе и выбытии учащихся.
- Не предпринимать действий, способных нарушить работоспособность Программы или привести к потере информации.

### 4. Обработка и хранение персональных данных
4.1. В процессе эксплуатации EduTrack могут обрабатываться персональные данные учащихся, необходимые для выполнения задач образовательного учреждения.

4.2. К таким сведениям могут относиться:

- Фамилия, имя и отчество.
- Дата рождения.
- Сведения о классе и периоде обучения.
- Статус обучения.
- История зачисления и выбытия.
- Результаты предусмотренных учреждением тестирований.

4.3. Образовательное учреждение самостоятельно определяет правовые основания, цели, состав и порядок обработки персональных данных, а также выступает оператором персональных данных в случаях, предусмотренных применимым законодательством.

4.4. Пользователь обязан обеспечивать обработку персональных данных исключительно в рамках предоставленных полномочий и установленных целей.

4.5. Запрещается передавать, публиковать или иным образом распространять персональные данные учащихся без соответствующих правовых оснований.

4.6. Сведения должны храниться с соблюдением установленных требований к защите информации и срокам хранения персональных данных.

4.7. Настоящее Соглашение не заменяет отдельные документы, регулирующие обработку и защиту персональных данных образовательного учреждения.

### 5. Хранение информации и резервное копирование
5.1. EduTrack предусматривает локальное хранение информации на компьютере, на котором установлена Программа.

5.2. Ответственность за организацию надлежащего хранения базы данных и резервных копий несёт образовательное учреждение в пределах установленного распределения обязанностей.

5.3. Пользователь обязан соблюдать правила резервного копирования и не допускать несанкционированного доступа к файлам базы данных.

5.4. Перед выполнением операций, способных существенно изменить структуру или содержимое базы данных, рекомендуется создавать резервную копию.

5.5. Удаление, повреждение или потеря локальных файлов могут привести к невозможности восстановления отдельных сведений при отсутствии актуальной резервной копии.

### 6. Ограничения использования
6.1. Пользователю запрещается:

- Использовать Программу для незаконного сбора информации.
- Несанкционированно изменять или уничтожать сведения об учащихся.
- Передавать базу данных третьим лицам без соответствующих полномочий.
- Использовать Программу для распространения персональных данных.
- Намеренно нарушать работоспособность программного обеспечения.
- Использовать обнаруженные уязвимости для получения неправомерного доступа к информации.

6.2. Любые действия с информацией должны соответствовать назначению Программы и внутренним регламентам образовательного учреждения.

### 7. Ответственность сторон
7.1. Разработчик обеспечивает функционирование Программы в соответствии с реализованными возможностями и техническими характеристиками соответствующей версии.

7.2. Разработчик не несёт ответственности за последствия неправильного ввода информации, нарушения пользователем инструкций или несанкционированного вмешательства в работу Программы в пределах, допускаемых применимым законодательством.

7.3. Образовательное учреждение самостоятельно обеспечивает наличие законных оснований для обработки персональных данных, организацию доступа к ним и соблюдение требований законодательства.

7.4. Пользователь несёт ответственность за действия, совершённые с использованием предоставленных ему полномочий, в соответствии с применимыми нормами законодательства и внутренними документами учреждения.

7.5. Стороны освобождаются от ответственности в случаях и порядке, предусмотренных применимым законодательством.

### 8. Техническая поддержка и обновления
8.1. Разработчик вправе выпускать обновления, исправления и дополнительные функциональные возможности EduTrack.

8.2. Обновления могут включать:

- Исправление программных ошибок.
- Повышение стабильности и производительности.
- Улучшение пользовательского интерфейса.
- Изменение существующих и добавление новых функций.
- Улучшение механизмов защиты и хранения информации.

8.3. Перед установкой обновлений рекомендуется создавать резервную копию базы данных.

8.4. Условия предоставления технической поддержки определяются отдельно.

### 9. Заключительные положения
9.1. Настоящее Соглашение вступает в силу с момента его принятия пользователем.

9.2. Все вопросы, не урегулированные настоящим Соглашением, регулируются применимым законодательством.

9.3. Если отдельные положения настоящего Соглашения признаются недействительными, остальные положения сохраняют силу в допустимых законом пределах.

9.4. Актуальная редакция Соглашения должна быть доступна пользователям Программы.

---

### 10. Сведения о программном обеспечении
**Наименование:** EduTrack

**Тип программного обеспечения:** Настольная информационная система.

**Назначение:** Учёт учащихся и автоматизация процессов образовательного учреждения.

**Платформа:** Microsoft Windows.

**Разработчик:** Резванов Андрей Павлович

**Версия документа:** 1.0

**Дата редакции:** 16 ноября 2026 года.
'''

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
    ensure_database()
    connection = get_db_connection()
    try:
        row = connection.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        return row["value"] if row else default
    finally:
        connection.close()


def get_persistence_dir() -> Path:
    override = os.environ.get("EDUTRACK_STORAGE_DIR")
    if override:
        return Path(override)
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / "EduTrack"
    return Path.home() / ".edutrack"


def get_persistent_credentials_path() -> Path:
    return get_persistence_dir() / "credentials.json"


def save_persisted_credentials(username: str, password: str) -> Path:
    path = get_persistent_credentials_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps({"username": username, "password": password}, ensure_ascii=False), encoding="utf-8")
    return path


def load_persisted_credentials() -> dict[str, str]:
    path = get_persistent_credentials_path()
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    username = str(payload.get("username", "")).strip()
    password = str(payload.get("password", ""))
    if not username:
        return {}
    return {"username": username, "password": password}


def write_setting(key: str, value: str) -> None:
    ensure_database()
    connection = get_db_connection()
    try:
        connection.execute(
            "INSERT INTO settings (key, value, updated_at) VALUES (?, ?, datetime('now')) ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = datetime('now')",
            (key, value),
        )
        connection.commit()
    finally:
        connection.close()


def normalize_repo_url(value: str | None) -> str:
    cleaned = str(value or "").strip()
    if not cleaned:
        return DEFAULT_UPDATE_REPO

    cleaned = cleaned.split("#", 1)[0].split("?", 1)[0].strip()
    cleaned = cleaned.rstrip("/\\")
    if cleaned.lower().endswith(".git"):
        cleaned = cleaned[:-4]
    cleaned = cleaned.rstrip("/\\")

    if not cleaned.startswith(("http://", "https://")):
        cleaned = f"https://{cleaned}"
    return cleaned.rstrip("/\\")


def get_update_repo_url() -> str:
    return normalize_repo_url(read_setting("update_repo", DEFAULT_UPDATE_REPO))


def get_github_repo_from_url(repo_url: str | None) -> str:
    normalized = normalize_repo_url(repo_url)
    match = re.search(r"github\.com[:/]+([^/\s:]+)\/([^/\s/?#]+)", normalized, flags=re.IGNORECASE)
    if not match:
        raise ValueError(f"Не удалось распознать GitHub-репозиторий: {repo_url}")
    return f"{match.group(1)}/{match.group(2)}"


def fetch_latest_release(repo_url: str | None = None) -> dict[str, Any] | None:
    target_repo = normalize_repo_url(repo_url or get_update_repo_url())
    try:
        repo_slug = get_github_repo_from_url(target_repo)
    except ValueError:
        return None

    api_url = f"https://api.github.com/repos/{repo_slug}/releases/latest"
    request = urllib.request.Request(api_url, headers={"User-Agent": "EduTrack-Updater/1.0"})
    try:
        with urllib.request.urlopen(request, timeout=12) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.URLError:
        return None

    tag_name = str(payload.get("tag_name") or "").strip()
    if not tag_name:
        return None

    return {
        "tag_name": tag_name,
        "html_url": str(payload.get("html_url") or target_repo),
        "repo_url": target_repo,
        "repo_slug": repo_slug,
    }


def normalize_version(value: str | None) -> tuple[int, ...]:
    text = str(value or "").strip().lower().lstrip("v")
    if not text:
        return (0,)
    numbers = re.findall(r"\d+", text)
    if not numbers:
        return (0,)
    return tuple(int(part) for part in numbers)


def has_update_available(repo_url: str | None = None, current_version: str = APP_VERSION) -> dict[str, Any] | None:
    release = fetch_latest_release(repo_url)
    if release is None:
        return None
    latest_version = release["tag_name"]
    if normalize_version(latest_version) <= normalize_version(current_version):
        return None
    return {**release, "current_version": current_version, "latest_version": latest_version}


def update_employee_account(current_username: str, new_username: str | None = None, new_password: str | None = None) -> dict[str, Any]:
    normalized_current = (current_username or "").strip()
    if not normalized_current:
        return {"success": False, "error": "Текущий логин не задан."}

    connection = get_db_connection()
    try:
        row = connection.execute(
            "SELECT id, username, password_hash FROM users WHERE username = ?",
            (normalized_current,),
        ).fetchone()
        if row is None:
            return {"success": False, "error": "Учетная запись не найдена."}

        next_username = (new_username or row["username"]).strip()
        if not next_username:
            return {"success": False, "error": "Логин не может быть пустым."}

        password_hash = row["password_hash"]
        if new_password is not None and str(new_password).strip():
            password_hash = hash_password(str(new_password).strip())

        connection.execute(
            "UPDATE users SET username = ?, password_hash = ? WHERE id = ?",
            (next_username, password_hash, row["id"]),
        )
        connection.commit()
        return {"success": True, "username": next_username}
    finally:
        connection.close()


def run_git_update(repo_url: str | None = None) -> tuple[bool, str]:
    project_root = Path(__file__).resolve().parent
    git = shutil.which("git")
    if not git:
        return False, "Git не найден в системе."
    if not (project_root / ".git").exists():
        target = normalize_repo_url(repo_url or get_update_repo_url())
        webbrowser.open(target)
        return False, f"Открыта страница релиза: {target}"

    fetch = subprocess.run([git, "fetch", "--tags"], cwd=str(project_root), capture_output=True, text=True)
    if fetch.returncode != 0:
        return False, fetch.stderr.strip() or "Не удалось обновить теги репозитория."

    pull = subprocess.run([git, "pull", "--ff-only"], cwd=str(project_root), capture_output=True, text=True)
    if pull.returncode != 0:
        return False, pull.stderr.strip() or "Не удалось выполнить git pull."
    return True, "Обновление из GitHub применено успешно."


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
        self.current_user = ""
        self._modal_windows: list[tk.Toplevel] = []

        self.title("EduTrack")
        self.state("zoomed")
        self.attributes("-fullscreen", False)
        self.minsize(1180, 760)
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
        self.bind("<Unmap>", lambda _event: self._sync_modal_visibility())
        self.bind("<Map>", lambda _event: self._sync_modal_visibility())

        self.background_image = self._load_background_image(opacity=0.08)
        self.background_label = tk.Label(self, image=self.background_image, bg="#edf2ff", bd=0, highlightthickness=0)
        self.background_label.place(x=0, y=0, relwidth=1, relheight=1)
        self.background_label.lower()
        self.background_label.configure(cursor="arrow")

        self._configure_menu_bar()

        header = ttk.Frame(self, padding=(20, 14, 20, 12), style="Header.TFrame")
        header.pack(fill="x")

        logo_path = Path(resource_path("image.png"))
        if Image is not None and ImageTk is not None:
            self.logo_photo = self._load_background_image(opacity=0.30, max_size=(86, 86), source_path=logo_path)
        else:
            self.logo_photo = tk.PhotoImage(file=str(logo_path))

        logo_label = tk.Label(header, image=self.logo_photo, bg="#edf2fb", bd=0, highlightthickness=0)
        logo_label.image = self.logo_photo
        logo_label.pack(side="left", padx=(0, 12))

        title_wrap = tk.Frame(header, bg="#edf2fb")
        title_wrap.pack(side="left", fill="y")
        tk.Label(title_wrap, text="EduTrack", font=("Segoe UI", 22, "bold"), bg="#edf2fb", fg="#0f172a").pack(anchor="w")

        content_wrapper = tk.Frame(self, bg="#edf2ff")
        content_wrapper.pack(fill="both", expand=True, padx=18, pady=(0, 18))

        sidebar = tk.Frame(content_wrapper, bg="#eaf1ff", width=310, padx=14, pady=16)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        main_card = ttk.Frame(content_wrapper, padding=16, style="Card.TFrame")
        main_card.pack(side="right", fill="both", expand=True)

        account_card = tk.Frame(sidebar, bg="#ffffff", padx=14, pady=12, highlightbackground="#dfe7f2", highlightthickness=1)
        account_card.pack(fill="x", pady=(0, 12))
        tk.Label(account_card, text="Учетка", bg="#ffffff", fg="#1d4ed8", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.account_name_var = tk.StringVar(value=self.current_user or "admin")
        tk.Label(account_card, textvariable=self.account_name_var, bg="#ffffff", fg="#0f172a", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Button(account_card, text="Проверить обновления", command=self.check_for_updates).pack(fill="x", pady=(12, 0))
        ttk.Button(account_card, text="Репозиторий обновлений", command=self.change_update_repo).pack(fill="x", pady=(8, 0))

        nav_card = tk.Frame(sidebar, bg="#ffffff", padx=12, pady=12, highlightbackground="#dfe7f2", highlightthickness=1)
        nav_card.pack(fill="x", pady=(0, 12))
        tk.Label(nav_card, text="Навигация", bg="#ffffff", fg="#1f2937", font=("Segoe UI", 10, "bold")).pack(anchor="w")

        self.tab_frames: dict[str, tk.Widget] = {}
        self.tab_container = tk.Frame(main_card, bg="#ffffff")
        self.tab_container.pack(fill="both", expand=True)

        registry_tab = ttk.Frame(self.tab_container, padding=16, style="Card.TFrame")
        analytics_tab = ttk.Frame(self.tab_container, padding=16, style="Card.TFrame")
        processes_tab = ttk.Frame(self.tab_container, padding=16, style="Card.TFrame")
        agreement_tab = ttk.Frame(self.tab_container, padding=18, style="Card.TFrame")
        self.tab_frames = {"registry": registry_tab, "analytics": analytics_tab, "processes": processes_tab, "agreement": agreement_tab}

        self.path_var = tk.StringVar()
        ttk.Label(registry_tab, text="Путь к папке или файлу:", style="Body.TLabel").pack(anchor="w")

        path_row = ttk.Frame(registry_tab, style="Card.TFrame")
        path_row.pack(fill="x", pady=(8, 12))
        self.path_entry = ttk.Entry(path_row)
        self.path_entry.configure(textvariable=self.path_var, width=80)
        self.path_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(path_row, text="Обзор", style="Secondary.TButton", command=self.select_path).pack(side="left", padx=(10, 0))

        filter_row = ttk.Frame(registry_tab, style="Card.TFrame", padding=(14, 10))
        filter_row.pack(fill="x", pady=(0, 12))

        age_filter = tk.Frame(filter_row, bg="#edf4ff", highlightbackground="#dfeaff", highlightthickness=1)
        age_filter.pack(fill="x")
        tk.Label(age_filter, text="Фильтр по возрасту", bg="#edf4ff", fg="#1d4ed8", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=12, pady=(10, 0))

        age_controls = tk.Frame(age_filter, bg="#edf4ff")
        age_controls.pack(fill="x", padx=12, pady=(6, 12))

        ttk.Label(age_controls, text="от", style="Filter.TLabel").pack(side="left")
        self.min_age_var = tk.StringVar(value=str(DEFAULT_MIN_AGE))
        age_min_spin = ttk.Spinbox(age_controls, from_=1, to=30, textvariable=self.min_age_var, width=5)
        age_min_spin.pack(side="left", padx=(8, 18))

        ttk.Label(age_controls, text="до", style="Filter.TLabel").pack(side="left")
        self.max_age_var = tk.StringVar(value=str(DEFAULT_MAX_AGE))
        age_max_spin = ttk.Spinbox(age_controls, from_=1, to=30, textvariable=self.max_age_var, width=5)
        age_max_spin.pack(side="left", padx=(8, 18))
        ttk.Label(age_controls, text="лет", style="Filter.TLabel").pack(side="left")

        buttons = ttk.Frame(registry_tab, style="Card.TFrame")
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Найти учеников", style="Primary.TButton", command=self.run_search).pack(side="left")
        ttk.Button(buttons, text="Экспорт CSV", style="Secondary.TButton", command=self.export_report).pack(side="left", padx=(8, 0))

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

        for tab_name, title in (("registry", "Реестр"), ("analytics", "Аналитика"), ("processes", "Процессы")):
            btn = tk.Button(nav_card, text=title, bg="#f3f6fb", fg="#1f2937", bd=0, padx=14, pady=10, anchor="w", font=("Segoe UI", 10, "bold"), command=lambda name=tab_name: self._show_tab(name))
            btn.pack(fill="x", pady=4)

        agreement_box = tk.Frame(sidebar, bg="#ffffff", padx=12, pady=12, highlightbackground="#dfe7f2", highlightthickness=1)
        agreement_box.pack(fill="x", pady=(0, 12))
        agreement_label = tk.Label(agreement_box, text="Пользовательское соглашение", bg="#ffffff", fg="#1d4ed8", font=("Segoe UI", 10, "bold"), cursor="hand2")
        agreement_label.pack(anchor="w")
        agreement_label.bind("<Button-1>", lambda event: self._show_tab("agreement"))

        info_bottom = tk.Frame(sidebar, bg="#ffffff", padx=16, pady=12, highlightbackground="#dfe7f2", highlightthickness=1)
        info_bottom.pack(fill="x")
        tk.Label(info_bottom, image=self.logo_photo, bg="#ffffff", bd=0, highlightthickness=0).pack(anchor="center")
        tk.Label(info_bottom, text="EduTrack", bg="#ffffff", fg="#0f172a", font=("Segoe UI", 16, "bold")).pack(anchor="center")
        tk.Label(info_bottom, text="Школьный реестр учащихся", bg="#ffffff", fg="#475569", font=("Segoe UI", 9)).pack(anchor="center", pady=(2, 0))

        agreement_title = tk.Label(agreement_tab, text="Пользовательское соглашение", bg="#ffffff", fg="#0f172a", font=("Segoe UI", 20, "bold"), justify="left")
        agreement_title.pack(anchor="w", pady=(0, 12))

        agreement_scroll = tk.Scrollbar(agreement_tab)
        agreement_scroll.pack(side="right", fill="y")
        agreement_text = tk.Text(agreement_tab, wrap="word", width=110, height=32, bg="#f8fbff", bd=0, padx=12, pady=12, relief="flat", font=("Segoe UI", 10), yscrollcommand=agreement_scroll.set)
        agreement_text.pack(side="left", fill="both", expand=True)
        agreement_text.insert("end", USER_AGREEMENT_TEXT)
        agreement_text.configure(state="disabled")
        agreement_scroll.configure(command=agreement_text.yview)

        self._show_tab("registry")

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

        self.after(250, self.check_for_updates)

    def _show_tab(self, tab_name: str) -> None:
        for key, frame in self.tab_frames.items():
            if key == tab_name:
                frame.pack(fill="both", expand=True)
                frame.tkraise()
            else:
                frame.pack_forget()

    def _configure_menu_bar(self) -> None:
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        # Верхнее меню оставлено пустым, чтобы учетка была только в левой панели.

    def _center_dialog(self, dialog: tk.Toplevel, width: int, height: int) -> None:
        root_x = self.winfo_rootx()
        root_y = self.winfo_rooty()
        root_w = self.winfo_width()
        root_h = self.winfo_height()
        x = root_x + max(0, (root_w - width) // 2)
        y = root_y + max(0, (root_h - height) // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")

    def _sync_modal_visibility(self) -> None:
        if self.state() == "withdrawn":
            for modal in list(self._modal_windows):
                try:
                    modal.withdraw()
                except tk.TclError:
                    pass
            return

        for modal in list(self._modal_windows):
            try:
                modal.deiconify()
            except tk.TclError:
                pass

    def _show_modal_message(self, title: str, message: str, kind: str = "info") -> None:
        colors = {"info": ("#eff6ff", "#1d4ed8"), "warning": ("#fff7ed", "#c2410c"), "error": ("#fef2f2", "#b91c1c"), "success": ("#ecfdf5", "#15803d")}
        _, accent = colors.get(kind, colors["info"])

        dialog = tk.Toplevel(self)
        dialog.overrideredirect(True)
        dialog.transient(self)
        dialog.attributes("-topmost", True)
        dialog.configure(bg="#f3f4f6")
        dialog.protocol("WM_DELETE_WINDOW", lambda: close_dialog())
        self._modal_windows.append(dialog)
        self._center_dialog(dialog, 420, 180)

        body = tk.Frame(dialog, bg="#ffffff", padx=24, pady=18, highlightbackground="#dfe7f2", highlightthickness=1)
        body.pack(fill="both", expand=True)

        close_btn = tk.Button(body, text="×", font=("Segoe UI", 12, "bold"), bd=0, bg="#ffffff", fg="#475569", command=lambda: close_dialog())
        close_btn.pack(anchor="ne")

        tk.Label(body, text=title, justify="left", bg="#ffffff", fg="#1f2937", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 8))
        tk.Label(body, text=message, justify="left", wraplength=320, bg="#ffffff", fg="#1f2937", font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 16))
        ttk.Button(body, text="OK", command=lambda: close_dialog()).pack(anchor="e")

        def close_dialog() -> None:
            if dialog in self._modal_windows:
                self._modal_windows.remove(dialog)
            dialog.destroy()
            self._sync_modal_visibility()

        dialog.focus_set()
        dialog.wait_window(dialog)

    def _show_modal_confirm(self, title: str, message: str) -> bool:
        result = {"value": False}

        dialog = tk.Toplevel(self)
        dialog.overrideredirect(True)
        dialog.transient(self)
        dialog.attributes("-topmost", True)
        dialog.configure(bg="#f3f4f6")
        self._modal_windows.append(dialog)
        self._center_dialog(dialog, 430, 180)

        body = tk.Frame(dialog, bg="#ffffff", padx=24, pady=18, highlightbackground="#dfe7f2", highlightthickness=1)
        body.pack(fill="both", expand=True)

        close_btn = tk.Button(body, text="×", font=("Segoe UI", 12, "bold"), bd=0, bg="#ffffff", fg="#475569", command=lambda: close_dialog(False))
        close_btn.pack(anchor="ne")

        tk.Label(body, text=title, bg="#ffffff", fg="#1f2937", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 10))
        tk.Label(body, text=message, justify="left", wraplength=320, bg="#ffffff", fg="#1f2937", font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 16))

        buttons = tk.Frame(body, bg="#ffffff")
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Да", command=lambda: close_dialog(True)).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="Нет", command=lambda: close_dialog(False)).pack(side="right")

        def close_dialog(value: bool) -> None:
            result["value"] = value
            if dialog in self._modal_windows:
                self._modal_windows.remove(dialog)
            dialog.destroy()
            self._sync_modal_visibility()

        dialog.focus_set()
        dialog.wait_window(dialog)
        return result["value"]

    def _show_modal_prompt(self, title: str, prompt: str, initial: str = "", show: str = "") -> str | None:
        result = {"value": None}

        dialog = tk.Toplevel(self)
        dialog.overrideredirect(True)
        dialog.transient(self)
        dialog.attributes("-topmost", True)
        dialog.configure(bg="#f3f4f6")
        self._modal_windows.append(dialog)
        self._center_dialog(dialog, 420, 200)

        body = tk.Frame(dialog, bg="#ffffff", padx=20, pady=18, highlightbackground="#dfe7f2", highlightthickness=1)
        body.pack(fill="both", expand=True)

        close_btn = tk.Button(body, text="×", font=("Segoe UI", 12, "bold"), bd=0, bg="#ffffff", fg="#475569", command=lambda: close_dialog(None))
        close_btn.pack(anchor="ne")

        tk.Label(body, text=title, bg="#ffffff", fg="#1f2937", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 8))
        tk.Label(body, text=prompt, bg="#ffffff", fg="#1f2937", font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 8))
        value_var = tk.StringVar(value=initial)
        entry = tk.Entry(body, textvariable=value_var, show=show, font=("Segoe UI", 10), width=40)
        entry.pack(fill="x", pady=(0, 14))
        entry.focus_set()

        buttons = tk.Frame(body, bg="#ffffff")
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Сохранить", command=lambda: close_dialog(value_var.get())).pack(side="right")
        ttk.Button(buttons, text="Отмена", command=lambda: close_dialog(None)).pack(side="right", padx=(0, 8))

        def close_dialog(value: str | None) -> None:
            result["value"] = value
            if dialog in self._modal_windows:
                self._modal_windows.remove(dialog)
            dialog.destroy()
            self._sync_modal_visibility()

        dialog.focus_set()
        dialog.wait_window(dialog)
        return result["value"]

    def _show_path_picker(self, mode: str = "directory", start_path: str | None = None) -> str | None:
        result = {"value": None}
        dialog = tk.Toplevel(self)
        dialog.overrideredirect(True)
        dialog.transient(self)
        dialog.attributes("-topmost", True)
        dialog.configure(bg="#f3f4f6")
        self._modal_windows.append(dialog)
        self._center_dialog(dialog, 560, 420)

        body = tk.Frame(dialog, bg="#ffffff", padx=12, pady=12, highlightbackground="#dfe7f2", highlightthickness=1)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="Путь:", bg="#ffffff", fg="#1f2937", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        path_var = tk.StringVar(value=str(Path(start_path or self.path_var.get() or Path.cwd())))
        path_entry = tk.Entry(body, textvariable=path_var, font=("Segoe UI", 9), width=50)
        path_entry.pack(fill="x", pady=(4, 8))

        listbox_frame = tk.Frame(body, bg="#ffffff")
        listbox_frame.pack(fill="both", expand=True)
        listbox = tk.Listbox(listbox_frame, width=70, height=14, bg="#f8fafc", fg="#0f172a", font=("Segoe UI", 9), activestyle="none")
        listbox.pack(side="left", fill="both", expand=True)
        scroll = tk.Scrollbar(listbox_frame, orient="vertical", command=listbox.yview)
        scroll.pack(side="right", fill="y")
        listbox.configure(yscrollcommand=scroll.set)

        actions = tk.Frame(body, bg="#ffffff")
        actions.pack(fill="x", pady=(8, 0))

        def refresh_items() -> None:
            folder = Path(path_var.get())
            if not folder.exists() or not folder.is_dir():
                folder = Path.cwd()
                path_var.set(str(folder))
            listbox.delete(0, tk.END)
            entries = sorted(folder.iterdir(), key=lambda p: (not p.is_dir(), p.name.lower()))
            for child in entries:
                listbox.insert(tk.END, child.name)

        def close_dialog(value: str | None) -> None:
            result["value"] = value
            if dialog in self._modal_windows:
                self._modal_windows.remove(dialog)
            dialog.destroy()
            self._sync_modal_visibility()

        def item_from_selection() -> Path | None:
            if not listbox.curselection():
                return None
            selected_name = listbox.get(listbox.curselection()[0])
            return Path(path_var.get()) / selected_name

        def open_dir() -> None:
            item = item_from_selection()
            if item is None:
                return
            if item.is_dir():
                path_var.set(str(item))
                refresh_items()
            elif mode == "file":
                close_dialog(str(item))

        def choose_current() -> None:
            value = Path(path_var.get())
            if mode == "directory" and value.exists() and value.is_dir():
                close_dialog(str(value))
            elif mode == "file" and value.exists() and value.is_file():
                close_dialog(str(value))

        def go_up() -> None:
            path_var.set(str(Path(path_var.get()).parent))
            refresh_items()

        tk.Button(actions, text="↑", width=4, command=go_up).pack(side="left")
        tk.Button(actions, text="Открыть", command=open_dir).pack(side="left", padx=(8, 0))
        tk.Button(actions, text="Выбрать", command=choose_current).pack(side="right")
        tk.Button(actions, text="Отмена", command=lambda: close_dialog(None)).pack(side="right", padx=(0, 8))

        listbox.bind("<Double-Button-1>", lambda _event: open_dir())
        refresh_items()
        dialog.wait_window(dialog)
        return result["value"]

    def _show_login_dialog(self) -> bool:
        result = {"ok": False}
        saved = load_persisted_credentials()
        username_default = saved.get("username") or read_setting("current_user", DEFAULT_ADMIN_USER) or DEFAULT_ADMIN_USER
        password_default = saved.get("password", "")

        dialog = tk.Toplevel(self)
        dialog.overrideredirect(True)
        dialog.transient(self)
        dialog.attributes("-topmost", True)
        dialog.configure(bg="#f3f4f6")
        self._modal_windows.append(dialog)
        self._center_dialog(dialog, 380, 220)

        body = tk.Frame(dialog, bg="#ffffff", padx=24, pady=20, highlightbackground="#dfe7f2", highlightthickness=1)
        body.pack(fill="both", expand=True)

        close_btn = tk.Button(body, text="×", font=("Segoe UI", 12, "bold"), bd=0, bg="#ffffff", fg="#475569", command=lambda: close_dialog(False))
        close_btn.pack(anchor="ne")

        tk.Label(body, text="Логин", bg="#ffffff", fg="#1f2937", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 6))
        username_var = tk.StringVar(value=username_default)
        login_entry = tk.Entry(body, textvariable=username_var, width=28, font=("Segoe UI", 10))
        login_entry.pack(anchor="w", fill="x", pady=(0, 10))

        tk.Label(body, text="Пароль", bg="#ffffff", fg="#1f2937", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 6))
        password_var = tk.StringVar(value=password_default)
        password_entry = tk.Entry(body, textvariable=password_var, show="*", width=28, font=("Segoe UI", 10))
        password_entry.pack(anchor="w", fill="x", pady=(0, 12))

        def close_dialog(is_success: bool, username: str | None = None) -> None:
            if is_success and username:
                self.current_user = username
                write_setting("current_user", username)
                save_persisted_credentials(username, password_var.get())
            else:
                self.current_user = ""
            result["ok"] = is_success
            if dialog in self._modal_windows:
                self._modal_windows.remove(dialog)
            dialog.destroy()
            self._sync_modal_visibility()

        def attempt_login() -> None:
            username = username_var.get().strip()
            password = password_var.get()
            if authenticate_employee(username, password):
                close_dialog(True, username)
                return
            self._show_modal_message("Ошибка доступа", "Неверный логин или пароль.", kind="error")
            password_var.set("")

        ttk.Button(body, text="Войти", command=attempt_login).pack(anchor="e")
        login_entry.focus_set()
        dialog.wait_window(dialog)
        return result["ok"]

    def change_account_credentials(self) -> None:
        current_username = self.current_user or read_setting("current_user", DEFAULT_ADMIN_USER) or DEFAULT_ADMIN_USER
        username = self._show_modal_prompt("Смена логина", "Новый логин:", initial=current_username)
        if username is None:
            return

        password = self._show_modal_prompt("Смена пароля", "Новый пароль:", show="*")
        if password is None:
            return

        confirm = self._show_modal_prompt("Подтверждение", "Повторите новый пароль:", show="*")
        if confirm is None:
            return
        if password != confirm:
            self._show_modal_message("Ошибка", "Пароли не совпадают.", kind="error")
            return

        result = update_employee_account(current_username, new_username=username.strip() or None, new_password=password.strip() or None)
        if not result["success"]:
            self._show_modal_message("Ошибка", result["error"], kind="error")
            return

        self.current_user = result["username"]
        write_setting("current_user", self.current_user)
        self._show_modal_message("Готово", "Логин и пароль обновлены.", kind="success")

    def change_update_repo(self) -> None:
        repo_url = self._show_modal_prompt(
            "Репозиторий обновлений",
            "Укажите GitHub URL для проверки релизов:",
            initial=get_update_repo_url(),
        )
        if repo_url is None:
            return
        normalized = normalize_repo_url(repo_url)
        write_setting("update_repo", normalized)
        self.status_var.set(f"Репозиторий обновлений: {normalized}")
        self._show_modal_message("Готово", f"Репозиторий обновлений сохранён: {normalized}", kind="success")

    def _check_update_result(self, payload: dict[str, Any] | None, silent: bool = False) -> None:
        if payload is None:
            if not silent:
                self._show_modal_message("Обновления", "Проверка обновлений не завершилась или релизов не найдено.", kind="info")
            return

        if payload.get("available"):
            choice = self._show_modal_confirm(
                "Доступно обновление",
                f"Есть новая версия: {payload['latest_version']}\n\nОбновить автоматически из {payload['repo_url']}?",
            )
            if choice:
                ok, detail = run_git_update(payload["repo_url"])
                if ok:
                    self._show_modal_message("Обновление", detail, kind="success")
                    self.destroy()
                    return
                self._show_modal_message("Обновление", detail, kind="warning")
            return

        if not silent:
            self._show_modal_message("Обновления", f"У вас актуальная версия: {payload['current_version']}", kind="success")

    def check_for_updates(self) -> None:
        repo_url = get_update_repo_url()

        def worker() -> None:
            info = has_update_available(repo_url)
            self.after(0, lambda: self._check_update_result({**(info or {}), "available": bool(info), "repo_url": repo_url, "current_version": APP_VERSION} if info else {"available": False, "current_version": APP_VERSION, "repo_url": repo_url}, silent=False))

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

    def _confirm_dangerous_action(self, action_text: str) -> bool:
        return self._show_modal_confirm("Подтверждение операции", f"Вы действительно хотите выполнить: {action_text}?")

    def _bulk_action(self, operation: str) -> None:
        selected = self.result_tree.selection()
        if not selected:
            self._show_modal_message("Нет выбора", "Сначала выберите строки в таблице.", kind="warning")
            return

        records = []
        for item_id in selected:
            values = self.result_tree.item(item_id, "values")
            if len(values) < 6:
                continue
            records.append({"name": values[2], "birthday": values[4], "class_name": values[3]})

        if not records:
            self._show_modal_message("Нет данных", "Выделенные строки не содержат корректных записей.", kind="warning")
            return

        action_label = {"enroll": "зачисление", "transfer": "перевод", "withdraw": "выбытие"}[operation]
        if not self._confirm_dangerous_action(action_label):
            return

        target_class = None
        if operation == "transfer":
            target_class = self._show_modal_prompt("Перевод в класс", "Укажите новый класс:")
            if target_class is None:
                return

        result = apply_student_batch(records, operation=operation, target_class=target_class, reason=f"{action_label} через интерфейс ({self.current_user})")
        if not result["success"]:
            self._show_modal_message("Ошибка операции", "; ".join(result["errors"]), kind="error")
            return

        self.status_var.set(f"Выполнена массовая операция: {action_label}. Обновлено {result['updated_count']} записей.")
        self._show_modal_message("Готово", f"Операция выполнена: {result['updated_count']} записей.", kind="success")
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
            target_size = max_size or (1200, 1200)
            rgba.thumbnail(target_size, Image.Resampling.LANCZOS)
            alpha = rgba.getchannel("A")
            new_alpha = alpha.point(lambda p: int(p * opacity))
            rgba.putalpha(new_alpha)
            return ImageTk.PhotoImage(rgba)

    def select_path(self) -> None:
        current = self.path_var.get().strip() or str(Path.cwd())
        directory_path = filedialog.askdirectory(
            title="Выберите папку с файлами",
            initialdir=current,
            mustexist=True,
        )
        if directory_path:
            self.path_var.set(directory_path)
            return

        file_path = filedialog.askopenfilename(
            title="Выберите файл данных",
            initialdir=current,
            filetypes=[("Excel/CSV/TXT", "*.xlsx *.xls *.xlsm *.csv *.txt *.tsv"), ("Все файлы", "*.*")],
        )
        if file_path:
            self.path_var.set(file_path)

    def run_search(self) -> None:
        source = self.path_var.get().strip()
        if not source:
            self._show_modal_message("Путь не выбран", "Выберите папку или файл.", kind="warning")
            return

        try:
            min_age = int(self.min_age_var.get())
            max_age = int(self.max_age_var.get())
        except ValueError:
            self._show_modal_message("Некорректный возраст", "Возраст должен быть числом.", kind="warning")
            return

        if min_age > max_age:
            self._show_modal_message("Некорректный диапазон", "Минимальный возраст не может быть больше максимального.", kind="warning")
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
            self._show_modal_message("Нет данных", "Сначала выполните поиск.", kind="info")
            return

        default_name = "EduTrack_report.csv"
        output_path = self._show_modal_prompt("Сохранить CSV-отчёт", "Укажите путь к файлу:", initial=str(Path.cwd() / default_name))
        if not output_path:
            return

        path = build_report(self.results, output_path)
        self._show_modal_message("Готово", f"Отчёт сохранён: {path}", kind="success")

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
