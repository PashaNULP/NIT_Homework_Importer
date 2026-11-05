#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NIT² by FIZYK

Імпорт тем уроків і домашніх завдань з Excel у електронний журнал НІТ.

Безпечна поведінка:
- спочатку лише показує preview;
- нічого не записує без явного підтвердження;
- перед записом створює локальний JSON-backup поточного стану журналу;
- порожні клітинки Excel за замовчуванням НЕ стирають дані в НІТ.

API було відновлено з мережевих запитів вебзастосунку НІТ:
  GET  /api/v1/academic-years/active
  GET  /api/v1/journals/classes
  GET  /api/v1/journals/{id}/pages
  GET  /api/v1/training-plans
  PUT  /api/v1/training-plans/{id}

Окремий режим ``--lesson-report`` отримує кількість і дати уроків у всіх
журналах учителя. Він нічого не записує в НІТ і не потребує Excel-файлу.
"""

from __future__ import annotations

import argparse
import base64
import getpass
import json
import os
import re
import select
import secrets
import signal
import sys
import subprocess
import textwrap
import time
from dataclasses import dataclass
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Callable, Iterable
from urllib.parse import quote

from nit_backup import write_secure_backup

try:
    import requests
except ImportError:
    print("Не встановлено пакет requests. Виконай: python3 -m pip install requests openpyxl")
    raise SystemExit(2)

try:
    from openpyxl import load_workbook
except ImportError:
    print("Не встановлено пакет openpyxl. Виконай: python3 -m pip install requests openpyxl")
    raise SystemExit(2)


API_BASE = "https://api.nit.school/api/v1"
APP_ORIGIN = "https://app.nit.school"
AUTH_TOKEN_URL = "https://is.nit.school/connect/token"
AUTH_SCOPE = (
    "openid profile permission other_info identity.api nit.api router.api "
    "offline_access chat.api school.api olympiad.api"
)
KEYCHAIN_SERVICE = "NIT Homework Importer OAuth Session"
KEYCHAIN_ACCOUNT = "nit.school"
KEYCHAIN_FORMAT_PREFIX = "NIT2"
KEYCHAIN_CHUNK_SIZE = 120

# Персональні ID беремо з локального файлу, який не потрапляє в Git.
try:
    _LOCAL_CONFIG = json.loads(
        Path(__file__).with_name("nit_config.local.json").read_text(encoding="utf-8")
    )
except (OSError, ValueError, TypeError):
    _LOCAL_CONFIG = {}
DEFAULT_SCHOOL_ID = str(_LOCAL_CONFIG.get("school_id") or "").strip()
DEFAULT_USER_ID = str(_LOCAL_CONFIG.get("user_id") or "").strip()
DEFAULT_CLIENT_VERSION = "v1.10.2.10"

DATE_ALIASES = {
    "дата", "дата уроку", "дата заняття", "date", "day"
}
TOPIC_ALIASES = {
    "тема", "тема уроку", "тема заняття", "lessoncontent", "lesson content"
}
HOMEWORK_ALIASES = {
    "домашнє завдання", "домашнє", "дз", "д/з", "homework", "home work"
}
LESSON_NUM_ALIASES = {
    "урок", "номер уроку", "№ уроку", "номер заняття", "lesson", "period"
}


class NitError(RuntimeError):
    pass


@dataclass
class ExcelRow:
    excel_row: int
    topic: str | None
    homework: str | None
    day: date | None = None
    lesson_number: int | None = None


@dataclass
class Match:
    source: ExcelRow
    lesson: dict[str, Any]
    warning: str | None = None


@dataclass
class PlannedChange:
    lesson: dict[str, Any]
    source: ExcelRow
    new_topic: str
    new_homework: str
    topic_changed: bool
    homework_changed: bool
    warning: str | None = None

    @property
    def changed(self) -> bool:
        return self.topic_changed or self.homework_changed


@dataclass
class JournalLessonReport:
    semester: int
    journal: dict[str, Any]
    lessons: list[dict[str, Any]]


@dataclass
class AbsenceReminder:
    semester: int
    journal: dict[str, Any]
    journal_day: dict[str, Any]
    user_id: str
    student_name: str = ""


def compact(s: Any) -> str:
    if s is None:
        return ""
    return str(s).strip()


def normalize_header(value: Any) -> str:
    s = compact(value).lower().replace("’", "'")
    s = re.sub(r"\s+", " ", s)
    return s


def normalize_topic(value: Any) -> str:
    s = compact(value).lower().replace("’", "'")
    s = re.sub(r"[\s\.,;:!?«»\"'()\[\]{}№–—-]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def similarity(a: str, b: str) -> float:
    a1, b1 = normalize_topic(a), normalize_topic(b)
    if not a1 or not b1:
        return 1.0
    return SequenceMatcher(None, a1, b1).ratio()


def parse_date_value(value: Any) -> date | None:
    if value is None or compact(value) == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = compact(value)
    # ISO datetime/date
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00")).date()
    except ValueError:
        pass
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Не вдалося розпізнати дату: {value!r}")


def iso_day_from_lesson(lesson: dict[str, Any]) -> date:
    raw = lesson.get("day")
    if not raw:
        raise NitError(f"В уроці {lesson.get('id')} відсутнє поле day")
    return parse_date_value(raw)  # type: ignore[return-value]


def short(text: str, width: int = 105) -> str:
    text = compact(text).replace("\n", " ")
    return text if len(text) <= width else text[: width - 1] + "…"


def choose_from_menu(title: str, options: list[str]) -> int:
    if not options:
        raise NitError(f"Немає варіантів для вибору: {title}")
    print(f"\n{title}")
    for i, item in enumerate(options, 1):
        print(f"  {i}. {item}")
    while True:
        raw = input("Вибір: ").strip()
        try:
            idx = int(raw)
            if 1 <= idx <= len(options):
                return idx - 1
        except ValueError:
            pass
        print(f"Введи число від 1 до {len(options)}.")


def configure_interactive_action(args: argparse.Namespace) -> bool:
    """Налаштовує режим роботи, коли скрипт запущено без аргументів."""
    def choose_semester_scope() -> int | None:
        semester_choice = choose_from_menu(
            "За який семестр?",
            ["Обидва семестри", "1 семестр", "2 семестр"],
        )
        return None if semester_choice == 0 else semester_choice

    action = choose_from_menu(
        "Що бажаєте зробити?",
        [
            "Імпортувати теми та домашні завдання з Excel",
            "Переглянути кількість і дати уроків",
            "Підтвердити відсутності (червоні трикутники → «н»)",
            "Вийти",
        ],
    )
    if action == 0:
        return True
    if action == 1:
        args.lesson_report = True
        report_action = choose_from_menu(
            "Що зробити зі звітом?",
            [
                "Показати в терміналі",
                "Показати та зберегти в Excel",
                "Показати та зберегти в JSON",
            ],
        )
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if report_action == 1:
            args.report_output = Path(f"lessons_{stamp}.xlsx")
        elif report_action == 2:
            args.report_output = Path(f"lessons_{stamp}.json")
        args.semester = choose_semester_scope()
        return True
    if action == 2:
        args.confirm_absences = True
        args.semester = choose_semester_scope()
        return True

    print("До побачення.")
    return False


def get_token_from_clipboard() -> str:
    """macOS: читає Bearer token з буфера обміну, оминаючи ліміт довжини TTY."""
    if sys.platform != "darwin":
        raise NitError("Читання токена з буфера обміну автоматично підтримується лише на macOS.")

    print("\nСкопіюй Bearer token НІТ у буфер обміну (Cmd+C).")
    input("Коли скопіював — натисни Enter. Сам токен сюди НЕ вставляй: ")
    try:
        proc = subprocess.run(
            ["pbpaste"],
            capture_output=True,
            text=True,
            check=True,
        )
    except (OSError, subprocess.CalledProcessError) as e:
        raise NitError(f"Не вдалося прочитати буфер обміну через pbpaste: {e}") from e

    token = proc.stdout.strip()
    if not token:
        raise NitError("Буфер обміну порожній. Скопіюй token і запусти ще раз.")

    # Дозволяємо копіювати або чистий JWT, або значення виду 'Bearer eyJ...'.
    if token.lower().startswith("authorization:"):
        token = token.split(":", 1)[1].strip()

    print(f"✓ Токен зчитано з буфера обміну ({len(token)} символів). Сам token не показую.")
    return token


def request_oauth_session(
    form_data: dict[str, str],
    timeout: int = 30,
    previous_refresh_token: str | None = None,
) -> dict[str, Any]:
    """Виконує OAuth-запит і повертає нормалізовану сесію без пароля."""
    try:
        response = requests.post(
            AUTH_TOKEN_URL,
            data=form_data,
            headers={
                "Accept": "application/json",
                "Origin": APP_ORIGIN,
                "Referer": f"{APP_ORIGIN}/",
            },
            timeout=timeout,
        )
    except requests.RequestException as e:
        raise NitError(f"Не вдалося з’єднатися із сервером авторизації НІТ: {e}") from e

    if response.status_code in (400, 401, 403):
        raise NitError("НІТ відхилив авторизацію або поновлення сесії.")
    if not response.ok:
        raise NitError(f"Сервер авторизації НІТ повернув HTTP {response.status_code}")
    try:
        data = response.json()
    except ValueError as e:
        raise NitError("Сервер авторизації НІТ повернув неочікувану відповідь") from e
    token = compact(data.get("access_token")) if isinstance(data, dict) else ""
    if not token:
        raise NitError("У відповіді авторизації НІТ немає access_token")
    try:
        expires_in = max(int(data.get("expires_in") or 0), 0)
    except (TypeError, ValueError):
        expires_in = 0
    return {
        "access_token": token,
        "refresh_token": compact(data.get("refresh_token")) or previous_refresh_token,
        "expires_at": time.time() + expires_in if expires_in else None,
    }


def login_session_with_password(username: str, password: str, timeout: int = 30) -> dict[str, Any]:
    """Отримує OAuth-сесію через перевірений password flow НІТ."""
    username = username.strip()
    if not username or not password:
        raise NitError("Email і пароль не можуть бути порожніми.")
    return request_oauth_session(
        {
            "grant_type": "password",
            "scope": AUTH_SCOPE,
            "client_id": "nit",
            "username": username,
            "password": password,
        },
        timeout=timeout,
    )


def get_my_permissions(access_token: str, timeout: int = 30) -> dict[str, Any]:
    """Отримує власний userId і доступні школи без контекстних X-NIT заголовків."""
    token = compact(access_token)
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    if not token:
        raise NitError("НІТ не надав access token для визначення користувача.")
    try:
        response = requests.get(
            f"{API_BASE}/users/my-permissions",
            headers={
                "Accept": "application/json",
                "Authorization": f"Bearer {token}",
                "Origin": APP_ORIGIN,
                "Referer": f"{APP_ORIGIN}/",
                "X-NIT-CLIENT": "web",
                "X-NIT-CLIENT-VERSION": DEFAULT_CLIENT_VERSION,
                "X-NIT-LOCALE": "uk",
            },
            timeout=timeout,
        )
    except requests.RequestException as exc:
        raise NitError(f"Не вдалося визначити профіль користувача НІТ: {exc}") from exc
    if response.status_code in (401, 403):
        raise NitError("НІТ не дозволив отримати профіль поточного користувача.")
    if not response.ok:
        raise NitError(
            f"НІТ повернув HTTP {response.status_code} під час визначення профілю."
        )
    try:
        data = response.json()
    except ValueError as exc:
        raise NitError("НІТ повернув невідому відповідь профілю користувача.") from exc
    if not isinstance(data, dict):
        raise NitError("У відповіді НІТ немає профілю користувача.")
    return data


def login_with_password(username: str, password: str, timeout: int = 30) -> str:
    """Сумісна обгортка, яка повертає лише access token."""
    return str(login_session_with_password(username, password, timeout)["access_token"])


def refresh_oauth_session(refresh_token: str, timeout: int = 30) -> dict[str, Any]:
    if not refresh_token:
        raise NitError("У збереженій сесії немає refresh token")
    return request_oauth_session(
        {
            "grant_type": "refresh_token",
            "scope": AUTH_SCOPE,
            "client_id": "nit",
            "refresh_token": refresh_token,
        },
        timeout=timeout,
        previous_refresh_token=refresh_token,
    )


def _keychain_read_value(account: str) -> str | None:
    if sys.platform != "darwin":
        return None
    try:
        proc = subprocess.run(
            [
                "/usr/bin/security",
                "find-generic-password",
                "-a",
                account,
                "-s",
                KEYCHAIN_SERVICE,
                "-w",
            ],
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return None
    if proc.returncode != 0:
        return None
    return proc.stdout.rstrip("\r\n")


def _keychain_write_value(account: str, secret: str) -> None:
    """Безпечно записує коротке значення через системну утиліту Keychain.

    ``security -w`` отримує секрет через псевдотермінал. Завдяки цьому JSON із
    токенами не потрапляє ні до аргументів процесу, ні до історії shell.
    """
    if sys.platform != "darwin":
        return

    try:
        import pty
    except ImportError as e:
        raise NitError("Ця версія Python не підтримує безпечний запис у macOS Keychain") from e

    master_fd = -1
    child_pid = -1
    child_status: int | None = None
    output = bytearray()
    try:
        command = [
            "/usr/bin/security",
            "add-generic-password",
            "-U",
            "-a",
            account,
            "-s",
            KEYCHAIN_SERVICE,
            "-w",
        ]
        # pty.fork(), на відміну від Popen(openpty()), робить псевдотермінал
        # керівним TTY дочірнього процесу. Інакше ``security`` може відкрити
        # /dev/tty користувача, показати запити в Terminal і не отримати JSON,
        # який програма надсилає у створений нею PTY.
        child_pid, master_fd = pty.fork()
        if child_pid == 0:
            try:
                os.execv(command[0], command)
            finally:
                os._exit(127)

        prompts = (b"password data for new item:", b"retype password for new item:")
        sent_prompts: set[bytes] = set()
        deadline = time.monotonic() + 20
        while child_status is None:
            waited_pid, status = os.waitpid(child_pid, os.WNOHANG)
            if waited_pid == child_pid:
                child_status = status
                break
            if time.monotonic() >= deadline:
                os.kill(child_pid, signal.SIGKILL)
                _, child_status = os.waitpid(child_pid, 0)
                raise NitError("macOS Keychain не відповів протягом 20 секунд")
            readable, _, _ = select.select([master_fd], [], [], 0.2)
            if not readable:
                continue
            try:
                chunk = os.read(master_fd, 4096)
            except OSError:
                continue
            if not chunk:
                continue
            output.extend(chunk)
            lower_output = bytes(output).lower()
            for prompt in prompts:
                if prompt in lower_output and prompt not in sent_prompts:
                    os.write(master_fd, secret.encode("utf-8") + b"\n")
                    sent_prompts.add(prompt)

        return_code = os.waitstatus_to_exitcode(child_status)
    except NitError:
        raise
    except (OSError, subprocess.SubprocessError) as e:
        if child_pid > 0 and child_status is None:
            try:
                os.kill(child_pid, signal.SIGKILL)
                os.waitpid(child_pid, 0)
            except OSError:
                pass
        raise NitError(f"Не вдалося запустити macOS Keychain: {e}") from e
    finally:
        if master_fd >= 0:
            os.close(master_fd)

    if return_code != 0:
        message = bytes(output).decode("utf-8", errors="replace").strip()
        # Захист на випадок, якщо незвична версія security відобразить ввід.
        message = message.replace(secret, "[приховано]")
        detail = f": {message}" if message else ""
        raise NitError(f"Не вдалося зберегти сесію в macOS Keychain{detail}")


def _keychain_delete_value(account: str) -> bool:
    if sys.platform != "darwin":
        return False
    try:
        proc = subprocess.run(
            [
                "/usr/bin/security",
                "delete-generic-password",
                "-a",
                account,
                "-s",
                KEYCHAIN_SERVICE,
            ],
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return False
    return proc.returncode == 0


def _keychain_part_accounts(manifest: str | None, account: str = KEYCHAIN_ACCOUNT) -> list[str]:
    if not manifest:
        return []
    match = re.fullmatch(rf"{KEYCHAIN_FORMAT_PREFIX}:R:([0-9a-f]{{12}}):(\d{{1,3}})", manifest)
    if not match:
        return []
    nonce, raw_count = match.groups()
    count = int(raw_count)
    if not 1 <= count <= 256:
        return []
    return [f"{account}.session.{nonce}.{index:03d}" for index in range(count)]


def load_keychain_session(account: str = KEYCHAIN_ACCOUNT) -> dict[str, Any] | None:
    manifest = _keychain_read_value(account)
    part_accounts = _keychain_part_accounts(manifest, account)
    if part_accounts:
        parts = [_keychain_read_value(account) for account in part_accounts]
        if any(part is None for part in parts):
            return None
        try:
            encoded = "".join(part or "" for part in parts)
            refresh_token = base64.b64decode(
                encoded.encode("ascii"), altchars=b"-_", validate=True
            ).decode("utf-8")
        except (ValueError, UnicodeError):
            return None
        if not refresh_token:
            return None
        # Access token не зберігаємо: при запуску він безпечно поновиться через
        # refresh token. Це також значно зменшує обсяг секретів у Keychain.
        return {"access_token": "", "refresh_token": refresh_token, "expires_at": 0}

    # Сумісність із коротким записом попередньої версії.
    if not manifest:
        return None
    try:
        session = json.loads(manifest)
    except (TypeError, ValueError):
        return None
    if not isinstance(session, dict):
        return None
    return session if session.get("access_token") or session.get("refresh_token") else None


def save_keychain_session(session: dict[str, Any], account: str = KEYCHAIN_ACCOUNT) -> None:
    """Зберігає refresh token частинами у macOS Keychain.

    Інтерактивний ввід ``security -w`` має ліміт 128 символів, тому довге
    значення ділиться на короткі записи. Маніфест перемикається останнім, щоб
    перерваний запис не пошкодив попередню робочу сесію.
    """
    if sys.platform != "darwin":
        return
    refresh_token = compact(session.get("refresh_token"))
    if not refresh_token:
        raise NitError("НІТ не надав refresh token, тому цю сесію неможливо зберегти")

    encoded = base64.urlsafe_b64encode(refresh_token.encode("utf-8")).decode("ascii")
    chunks = [encoded[i : i + KEYCHAIN_CHUNK_SIZE] for i in range(0, len(encoded), KEYCHAIN_CHUNK_SIZE)]
    if not chunks or len(chunks) > 256:
        raise NitError("Refresh token має неочікуваний розмір і не може бути збережений")

    old_manifest = _keychain_read_value(account)
    old_accounts = _keychain_part_accounts(old_manifest, account)
    nonce = secrets.token_hex(6)
    manifest_account = account
    new_accounts = [
        f"{manifest_account}.session.{nonce}.{index:03d}" for index in range(len(chunks))
    ]
    try:
        for part_account, chunk in zip(new_accounts, chunks):
            _keychain_write_value(part_account, chunk)
        _keychain_write_value(
            manifest_account,
            f"{KEYCHAIN_FORMAT_PREFIX}:R:{nonce}:{len(chunks)}",
        )
    except NitError:
        for part_account in new_accounts:
            _keychain_delete_value(part_account)
        raise

    for part_account in old_accounts:
        _keychain_delete_value(part_account)


def delete_keychain_session(account: str = KEYCHAIN_ACCOUNT) -> bool:
    manifest = _keychain_read_value(account)
    part_accounts = _keychain_part_accounts(manifest, account)
    deleted = _keychain_delete_value(account)
    for part_account in part_accounts:
        _keychain_delete_value(part_account)
    return deleted


def valid_or_refreshed_session(
    session: dict[str, Any],
    save_session: Callable[[dict[str, Any]], None] | None = None,
    min_validity_seconds: int = 60,
) -> dict[str, Any] | None:
    token = compact(session.get("access_token"))
    expires_at = session.get("expires_at")
    try:
        still_valid = (
            expires_at is None
            or float(expires_at) > time.time() + max(0, min_validity_seconds)
        )
    except (TypeError, ValueError):
        still_valid = False
    if token and still_valid:
        return session
    refresh_token = compact(session.get("refresh_token"))
    if not refresh_token:
        return None
    print("→ Поновлюю сесію НІТ через refresh token...", flush=True)
    try:
        refreshed = refresh_oauth_session(refresh_token)
    except NitError:
        return None
    try:
        (save_session or save_keychain_session)(refreshed)
    except (NitError, OSError) as e:
        print(f"⚠ {e}. Поновлена сесія працюватиме до завершення програми.")
    print("✓ Сесію НІТ поновлено.")
    return refreshed


def get_auth_token(provided_token: str | None = None) -> str | None:
    """Повертає переданий token або показує безпечне меню авторизації."""
    if provided_token:
        return provided_token

    saved_session = load_keychain_session()
    options: list[tuple[str, str]] = []
    if saved_session:
        options.append(("saved", "Використати збережену сесію з macOS Keychain"))
    options.append(("login", "Увійти через email і пароль"))
    if sys.platform == "darwin":
        options.append(("clipboard", "Зчитати Bearer token із буфера обміну"))
    else:
        options.append(("manual", "Ввести Bearer token приховано"))
    if saved_session:
        options.append(("delete", "Видалити збережену сесію"))
    options.append(("cancel", "Скасувати"))

    choice = choose_from_menu("Авторизація НІТ:", [label for _, label in options])
    action = options[choice][0]
    if action == "saved":
        session = valid_or_refreshed_session(saved_session or {})
        if session:
            return compact(session.get("access_token"))
        print("Збережена сесія протермінована й не поновилася. Увійдіть повторно.")
        delete_keychain_session()
        return get_auth_token()
    if action == "login":
        username = input("Email НІТ: ").strip()
        password = getpass.getpass("Пароль НІТ (ввід прихований): ")
        session = login_session_with_password(username, password)
        if sys.platform == "darwin":
            try:
                save_keychain_session(session)
            except NitError as e:
                print(f"⚠ {e}. Поточний вхід усе одно працюватиме.")
            else:
                print("✓ Сесію безпечно збережено в macOS Keychain.")
        return compact(session.get("access_token"))
    if action == "clipboard":
        return get_token_from_clipboard()
    if action == "manual":
        return getpass.getpass("Встав Bearer token НІТ (ввід прихований): ").strip()
    if action == "delete":
        if delete_keychain_session():
            print("Збережену сесію видалено з macOS Keychain.")
        else:
            print("Збереженої сесії не знайдено.")
        return get_auth_token()
    return None


def maybe_pick_excel_file() -> Path:
    # Спочатку пробуємо нативний file picker через tkinter.
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askopenfilename(
            title="Оберіть Excel з темами та домашніми завданнями",
            filetypes=[("Excel", "*.xlsx"), ("Усі файли", "*")],
        )
        root.destroy()
        if selected:
            return Path(selected)
    except Exception:
        pass

    raw = input("Шлях до Excel-файлу (.xlsx): ").strip().strip('"').strip("'")
    if not raw:
        raise NitError("Файл не обрано.")
    return Path(raw).expanduser()


def find_header(ws) -> tuple[int, dict[str, int]]:
    """Повертає номер рядка заголовка і мапу canonical_name -> column_index."""
    for r in range(1, min(ws.max_row, 30) + 1):
        found: dict[str, int] = {}
        for c in range(1, ws.max_column + 1):
            h = normalize_header(ws.cell(r, c).value)
            if not h:
                continue
            if h in TOPIC_ALIASES:
                found["topic"] = c
            elif h in HOMEWORK_ALIASES:
                found["homework"] = c
            elif h in DATE_ALIASES:
                found["date"] = c
            elif h in LESSON_NUM_ALIASES:
                found["lesson"] = c
        if "topic" in found or "homework" in found:
            return r, found
    raise NitError(
        "Не знайшов у Excel заголовки «Тема уроку» / «Домашнє завдання». "
        "Перевір структуру файлу."
    )


def load_excel_rows(path: Path, sheet_name: str | None = None) -> tuple[str, list[ExcelRow], bool]:
    if not path.exists():
        raise NitError(f"Файл не знайдено: {path}")
    if path.suffix.lower() != ".xlsx":
        raise NitError("Поки підтримується формат .xlsx")

    wb = load_workbook(path, data_only=True)

    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise NitError(f"Аркуша {sheet_name!r} немає. Є: {', '.join(wb.sheetnames)}")
        ws = wb[sheet_name]
    else:
        # Шукаємо перший аркуш із потрібними заголовками.
        ws = None
        for candidate in wb.worksheets:
            try:
                find_header(candidate)
                ws = candidate
                break
            except NitError:
                continue
        if ws is None:
            raise NitError("Не знайшов аркуш із колонками теми/домашнього завдання.")

    header_row, cols = find_header(ws)
    rows: list[ExcelRow] = []
    has_date = "date" in cols

    for r in range(header_row + 1, ws.max_row + 1):
        topic = compact(ws.cell(r, cols["topic"]).value) if "topic" in cols else ""
        homework = compact(ws.cell(r, cols["homework"]).value) if "homework" in cols else ""
        raw_date = ws.cell(r, cols["date"]).value if "date" in cols else None
        raw_lesson = ws.cell(r, cols["lesson"]).value if "lesson" in cols else None

        if not topic and not homework and raw_date in (None, ""):
            continue

        parsed_day = None
        if raw_date not in (None, ""):
            try:
                parsed_day = parse_date_value(raw_date)
            except ValueError as e:
                raise NitError(f"Excel, рядок {r}: {e}") from e

        lesson_number = None
        if raw_lesson not in (None, ""):
            try:
                lesson_number = int(raw_lesson)
            except (TypeError, ValueError):
                raise NitError(f"Excel, рядок {r}: невалідний номер уроку {raw_lesson!r}")

        rows.append(
            ExcelRow(
                excel_row=r,
                day=parsed_day,
                lesson_number=lesson_number,
                topic=topic or None,
                homework=homework or None,
            )
        )

    if not rows:
        raise NitError("У вибраному Excel не знайдено рядків для імпорту.")

    return ws.title, rows, has_date


class NitClient:
    def __init__(
        self,
        token: str,
        school_id: str,
        user_id: str,
        client_version: str,
        timeout: int = 30,
    ):
        token = token.strip()
        if token.lower().startswith("bearer "):
            token = token[7:].strip()
        if not token:
            raise NitError("Порожній Bearer token.")

        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Accept": "application/json",
                "Content-Type": "application/json",
                "Authorization": f"Bearer {token}",
                "Origin": APP_ORIGIN,
                "Referer": f"{APP_ORIGIN}/",
                "X-NIT-CLIENT": "web",
                "X-NIT-CLIENT-VERSION": client_version,
                "X-NIT-LOCALE": "uk",
                "X-NIT-REQUESTED-SCHOOL-ID": school_id,
                "X-NIT-REQUESTED-USER-ID": user_id,
            }
        )

    def set_access_token(self, token: str) -> None:
        """Оновлює Bearer token для наступних API-запитів."""
        clean = token.strip()
        if clean.lower().startswith("bearer "):
            clean = clean[7:].strip()
        if not clean:
            raise NitError("НІТ не надав access token після поновлення сесії")
        self.session.headers["Authorization"] = f"Bearer {clean}"

    def get_school_teachers(self, include_inactive: bool = False) -> list[dict[str, Any]]:
        """Повертає довідник учителів школи без зміни будь-яких даних."""
        r = self.session.get(
            f"{API_BASE}/journals/teachers",
            params={"includeInactive": str(include_inactive).lower()},
            headers={"Alias": "journal-teachers"},
            timeout=self.timeout,
        )
        data = self._handle(r, "Отримання списку вчителів")
        if not isinstance(data, dict) or not isinstance(data.get("data"), list):
            raise NitError("НІТ повернув невідомий формат списку вчителів")
        return [item for item in data["data"] if isinstance(item, dict)]

    def get_school_journal_days(
        self,
        start_day: date,
        end_day: date,
        teacher_ids: list[str],
    ) -> list[dict[str, Any]]:
        """Читає уроки вибраних учителів за діапазон дат із пагінацією."""
        if end_day < start_day:
            raise NitError("Кінцева дата розкладу не може бути раніше початкової")
        ids = [str(item).strip() for item in teacher_ids if str(item).strip()]
        if not ids:
            return []

        take = 1000
        skip = 0
        result: list[dict[str, Any]] = []
        while True:
            params = {
                "take": str(take),
                "skip": str(skip),
                "noTotal": "false",
                "filter[logic]": "and",
                "filter[filters][0][field]": "date",
                "filter[filters][0][operator]": "gte",
                "filter[filters][0][value]": start_day.isoformat(),
                "filter[filters][1][field]": "date",
                "filter[filters][1][operator]": "lte",
                "filter[filters][1][value]": end_day.isoformat(),
                "filter[filters][2][field]": "type",
                "filter[filters][2][operator]": "eq",
                "filter[filters][2][value]": "0",
                "filter[filters][3][field]": "journalTeacher.id",
                "filter[filters][3][operator]": "in",
                "filter[filters][3][value]": ",".join(ids),
            }
            r = self.session.get(
                f"{API_BASE}/journal-days",
                params=params,
                headers={"Alias": "journal-days-all"},
                timeout=max(self.timeout, 60),
            )
            data = self._handle(r, "Отримання розкладу всіх учителів")
            if not isinstance(data, dict) or not isinstance(data.get("data"), list):
                raise NitError("НІТ повернув невідомий формат загального розкладу")
            batch = [item for item in data["data"] if isinstance(item, dict)]
            result.extend(batch)
            total = data.get("total")
            if len(batch) < take or (isinstance(total, int) and len(result) >= total):
                break
            skip += take
            if skip > 50000:
                raise NitError("Розклад містить забагато рядків для одного експорту")
        return result

    def get_teacher_diary(self, week_start: date) -> dict[str, Any]:
        """Отримує персональний тиждень у підтвердженому форматі teacher-diary."""
        weekdays = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
        months = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
        value = (
            f"{weekdays[week_start.weekday()]} {months[week_start.month - 1]} "
            f"{week_start.day:02d} {week_start.year:04d}"
        )
        r = self.session.get(
            f"{API_BASE}/dashboard/teacher-diary?date={quote(value, safe='')}",
            timeout=self.timeout,
        )
        data = self._handle(r, f"Розклад на тиждень {week_start.isoformat()}")
        if not isinstance(data, dict):
            raise NitError("НІТ повернув невідомий формат teacher-diary")
        return data

    def get_academic_years(self) -> dict[str, Any]:
        """Повна конфігурація років, включно з офіційними дзвінками."""
        r = self.session.get(f"{API_BASE}/academic-years?take=10&skip=0", timeout=self.timeout)
        data = self._handle(r, "Повна конфігурація навчальних років")
        if not isinstance(data, dict):
            raise NitError("НІТ повернув невідомий формат academic-years")
        return data

    def _handle(self, response: requests.Response, context: str) -> Any:
        if response.status_code in (401, 403):
            raise NitError(
                f"{context}: сервер повернув {response.status_code}. "
                "Ймовірно, Bearer token протермінований або не підходить."
            )
        if not response.ok:
            body = response.text[:1000]
            raise NitError(f"{context}: HTTP {response.status_code}: {body}")
        if not response.content:
            return None
        try:
            return response.json()
        except ValueError:
            return response.text

    def get_active_year(self) -> dict[str, Any]:
        r = self.session.get(
            f"{API_BASE}/academic-years/active",
            headers={"Alias": "academic-year-active"},
            timeout=self.timeout,
        )
        data = self._handle(r, "Отримання активного навчального року")
        if not isinstance(data, dict):
            raise NitError("Неочікувана відповідь academic-years/active")
        return data

    def get_journals(self, schedule_id: str) -> list[dict[str, Any]]:
        r = self.session.get(
            f"{API_BASE}/journals/classes",
            params={"scheduleId": schedule_id, "myClassesOnly": "false"},
            headers={"Alias": "journal-classes"},
            timeout=self.timeout,
        )
        data = self._handle(r, "Отримання списку журналів")
        if not isinstance(data, dict) or not isinstance(data.get("data"), list):
            raise NitError("Неочікувана відповідь journals/classes")

        result: list[dict[str, Any]] = []
        for klass in data["data"]:
            if not isinstance(klass, dict):
                continue
            for subject in klass.get("subjects") or []:
                if not isinstance(subject, dict) or not subject.get("isMyJournal"):
                    continue
                result.append(
                    {
                        "classId": klass.get("id"),
                        "className": klass.get("name") or f"{klass.get('number', '')}-{klass.get('letter', '')}",
                        "classNumber": klass.get("number") or 0,
                        "classLetter": klass.get("letter") or "",
                        "subject": subject.get("subject") or "",
                        "subjectId": subject.get("subjectId"),
                        "journalId": subject.get("journalId"),
                        "scheduleId": subject.get("scheduleId") or schedule_id,
                    }
                )

        result.sort(key=lambda x: (x["classNumber"], x["classLetter"], x["subject"]))
        return result

    def get_training_plans(self, journal_id: str, take: int = 500) -> list[dict[str, Any]]:
        params = {
            "take": str(take),
            "filter[logic]": "and",
            "filter[filters][0][field]": "journalId",
            "filter[filters][0][operator]": "eq",
            "filter[filters][0][value]": journal_id,
            "sort[0][field]": "day",
            "sort[0][dir]": "asc",
        }
        r = self.session.get(
            f"{API_BASE}/training-plans",
            params=params,
            headers={"Alias": "training-plans-all"},
            timeout=self.timeout,
        )
        data = self._handle(r, "Отримання уроків журналу")
        if not isinstance(data, dict) or not isinstance(data.get("data"), list):
            raise NitError("Неочікувана відповідь training-plans")

        lessons = data["data"]
        total = data.get("total")
        if isinstance(total, int) and len(lessons) < total:
            raise NitError(
                f"НІТ повідомляє total={total}, але API повернуло лише {len(lessons)} уроків. "
                "Імпорт зупинено, щоб не працювати з неповним журналом."
            )
        lessons.sort(key=lambda x: (compact(x.get("day")), int(x.get("lesson") or 0)))
        return lessons

    def get_journal_pages(self, journal_id: str) -> list[dict[str, Any]]:
        """Завантажує всі сторінки журналу (page=0..countPages)."""
        def get_page(page: int) -> dict[str, Any]:
            r = self.session.get(
                f"{API_BASE}/journals/{journal_id}/pages",
                params={"page": str(page)},
                headers={"Alias": "journal"},
                timeout=self.timeout,
            )
            data = self._handle(r, f"Отримання сторінки {page + 1} журналу")
            if not isinstance(data, dict) or not isinstance(data.get("journalDay"), list):
                raise NitError(
                    "Неочікувана відповідь journals/{id}/pages: "
                    "немає масиву journalDay"
                )
            return data

        first_page = get_page(0)
        raw_count_pages = first_page.get("countPages", 0)
        try:
            last_page = int(raw_count_pages)
        except (TypeError, ValueError) as e:
            raise NitError(f"НІТ повернув невалідне countPages={raw_count_pages!r}") from e
        if last_page < 0 or last_page > 1000:
            raise NitError(f"НІТ повернув підозріле countPages={last_page}")

        pages = [first_page]
        pages.extend(get_page(page) for page in range(1, last_page + 1))
        return pages

    def get_journal_users(self, journal_id: str) -> list[dict[str, Any]]:
        """Повертає учнів журналу в тому порядку, який задає НІТ."""
        r = self.session.get(
            f"{API_BASE}/journals/{journal_id}/users",
            headers={"Alias": "journal-users"},
            timeout=self.timeout,
        )
        data = self._handle(r, "Отримання списку учнів журналу")
        if not isinstance(data, dict) or not isinstance(data.get("data"), list):
            raise NitError("Неочікувана відповідь journals/{id}/users")
        return [item for item in data["data"] if isinstance(item, dict)]

    def get_journal_lessons(self, journal_id: str) -> list[dict[str, Any]]:
        """Повертає кожен урок з усіх сторінок журналу окремим записом.

        ``allJournalDates`` тут навмисно не використовується: НІТ зберігає там
        унікальні дати, тому два уроки одного предмета в один день зливаються.
        Фактичні колонки уроків містяться в ``journalDay`` і мають ``type=0``.
        """
        pages = self.get_journal_pages(journal_id)

        lessons: list[dict[str, Any]] = []
        for page_data in pages:
            for journal_day in page_data["journalDay"]:
                if not isinstance(journal_day, dict):
                    continue
                # JournalColumnType.Day у вебклієнті НІТ дорівнює 0. Інші
                # значення — тематичні, семестрові та користувацькі колонки.
                if journal_day.get("type") not in (0, "0", "Day", "day"):
                    continue
                raw_date = journal_day.get("date")
                try:
                    parsed = parse_date_value(raw_date)
                except ValueError as e:
                    raise NitError(f"НІТ повернув невалідну дату уроку {raw_date!r}") from e
                if parsed is not None:
                    lesson_number = journal_day.get("lesson")
                    if isinstance(lesson_number, float) and lesson_number.is_integer():
                        lesson_number = int(lesson_number)
                    lessons.append(
                        {
                            "journalDayId": journal_day.get("id"),
                            "day": parsed.isoformat(),
                            "lesson": lesson_number,
                            "index": journal_day.get("index"),
                            "teacherId": journal_day.get("teacherId"),
                            "isFilled": False,
                        }
                    )

        lessons.sort(
            key=lambda item: (
                iso_day_from_lesson(item),
                int(item.get("lesson") or 0),
                int(item.get("index") or 0),
            )
        )
        return lessons

    def create_absence_mark(self, reminder: AbsenceReminder) -> Any:
        """Підтверджує звичайну відсутність (`н`) для одного нагадування."""
        day = reminder.journal_day
        payload = {
            "classId": day.get("classId") or reminder.journal.get("classId"),
            "columnId": day.get("id"),
            "customMark": None,
            "description": None,
            "groupId": day.get("classGroupId") or reminder.journal.get("classGroupId"),
            "id": None,
            # Саме так поле та значення названі в реальному запиті НІТ.
            "absense": 0,
            "subject": day.get("subject") or reminder.journal.get("subject"),
            "userId": reminder.user_id,
            "scheduleId": None,
            "rating": None,
            "day": day.get("date"),
            "isVerbalFormed": None,
            "leveledMark": None,
            "letterMark": None,
        }
        if not payload["classId"] or not payload["columnId"] or not payload["userId"] or not payload["day"]:
            raise NitError("Неможливо підтвердити відсутність: бракує classId/columnId/userId/day")

        r = self.session.post(
            f"{API_BASE}/marks",
            json=payload,
            timeout=self.timeout,
        )
        return self._handle(r, "Підтвердження відсутності учня")

    def update_training_plan(self, lesson: dict[str, Any], topic: str, homework: str) -> None:
        lesson_id = compact(lesson.get("id"))
        journal_id = compact(lesson.get("journalId"))
        if not lesson_id or not journal_id:
            raise NitError("Неможливо оновити урок без id/journalId")

        attachment_ids: list[str] = []
        for attachment in lesson.get("homeworkAttachments") or []:
            if isinstance(attachment, str):
                attachment_ids.append(attachment)
                continue
            if not isinstance(attachment, dict):
                raise NitError(
                    f"Урок {lesson_id}: невідомий формат homeworkAttachments; "
                    "запис зупинено, щоб не втратити вкладення."
                )
            attachment_id = (
                attachment.get("id")
                or attachment.get("attachmentId")
                or attachment.get("fileId")
                or attachment.get("_id")
            )
            if not attachment_id:
                raise NitError(
                    f"Урок {lesson_id}: не вдалося визначити ID вкладення; "
                    "запис зупинено, щоб не втратити вкладення."
                )
            attachment_ids.append(str(attachment_id))

        payload = {
            "id": lesson_id,
            "journalId": journal_id,
            "teacherId": lesson.get("teacherId"),
            "urls": lesson.get("urls") or [],
            "lessonContent": topic,
            "homework": homework,
            "homeworkAttachmentIds": attachment_ids,
            "isDateCustom": bool(lesson.get("isDateCustom", False)),
            "homeworkDate": lesson.get("homeworkDate"),
            "isRequiredToSubmit": bool(lesson.get("isRequiredToSubmit", False)),
        }

        r = self.session.put(
            f"{API_BASE}/training-plans/{lesson_id}",
            json=payload,
            timeout=self.timeout,
        )
        self._handle(r, f"Оновлення уроку {lesson_id}")

    def create_training_plan(
        self,
        lesson: dict[str, Any],
        journal_id: str,
        topic: str,
        homework: str,
    ) -> dict[str, Any] | None:
        """Створює наступний порожній план так само, як офіційний вебклієнт НІТ."""
        journal_id = compact(journal_id)
        topic = compact(topic)
        if not journal_id:
            raise NitError("Неможливо створити урок без journalId")
        if not topic:
            raise NitError("Неможливо створити порожній план без теми уроку")

        day = iso_day_from_lesson(lesson)
        payload = {
            "id": None,
            "journalId": journal_id,
            "teacherId": lesson.get("teacherId"),
            "urls": [],
            "lessonContent": topic,
            "homework": homework,
            "homeworkAttachmentIds": [],
            "isDateCustom": False,
            "homeworkDate": f"{day.isoformat()}T00:00:00.000Z",
            "isRequiredToSubmit": False,
        }
        r = self.session.post(
            f"{API_BASE}/training-plans",
            json=payload,
            timeout=self.timeout,
        )
        data = self._handle(r, f"Створення плану уроку за {day.isoformat()}")
        return data if isinstance(data, dict) else None


def match_rows_to_lessons(rows: list[ExcelRow], lessons: list[dict[str, Any]], has_date: bool) -> list[Match]:
    if has_date:
        by_day: dict[date, list[dict[str, Any]]] = {}
        for lesson in lessons:
            by_day.setdefault(iso_day_from_lesson(lesson), []).append(lesson)

        def lesson_identity(lesson: dict[str, Any]) -> str:
            """ID плану є лише у заповнених уроків; для порожніх беремо ID колонки."""
            stable_id = compact(lesson.get("id")) or compact(lesson.get("journalDayId"))
            if stable_id:
                return stable_id
            return f"{iso_day_from_lesson(lesson).isoformat()}:{int(lesson.get('lesson') or 0)}"

        matches: list[Match] = []
        used_ids: set[str] = set()
        for row in rows:
            if row.day is None:
                raise NitError(
                    f"Excel, рядок {row.excel_row}: колонка «Дата» є, але дата порожня. "
                    "Заповни дату або прибери цей рядок."
                )
            candidates = [
                x for x in by_day.get(row.day, []) if lesson_identity(x) not in used_ids
            ]
            if row.lesson_number is not None:
                candidates = [x for x in candidates if int(x.get("lesson") or 0) == row.lesson_number]

            if not candidates:
                raise NitError(
                    f"Excel, рядок {row.excel_row}: у НІТ не знайдено урок на {row.day.isoformat()}"
                    + (f" (урок {row.lesson_number})" if row.lesson_number else "")
                )
            if len(candidates) > 1:
                raise NitError(
                    f"Excel, рядок {row.excel_row}: на {row.day.isoformat()} знайдено кілька уроків. "
                    "Додай колонку «Урок» із номером уроку в розкладі."
                )
            lesson = candidates[0]
            used_ids.add(lesson_identity(lesson))
            warning = None
            if row.topic and compact(lesson.get("lessonContent")):
                score = similarity(row.topic, compact(lesson.get("lessonContent")))
                if score < 0.55:
                    warning = f"теми суттєво відрізняються (схожість {score:.0%})"
            matches.append(Match(row, lesson, warning))
        return matches

    # Старий формат Excel без дати: зіставляємо строго за порядком.
    if len(rows) != len(lessons):
        raise NitError(
            "В Excel немає колонки «Дата», тому можливе лише послідовне зіставлення. "
            f"Але в Excel {len(rows)} рядків, а в НІТ {len(lessons)} уроків. "
            "Додай колонку «Дата» — тоді зіставлення буде надійним."
        )

    matches = []
    for row, lesson in zip(rows, lessons):
        warning = None
        if row.topic and compact(lesson.get("lessonContent")):
            score = similarity(row.topic, compact(lesson.get("lessonContent")))
            if score < 0.55:
                warning = f"теми суттєво відрізняються (схожість {score:.0%})"
        matches.append(Match(row, lesson, warning))
    return matches


def plan_changes(matches: list[Match], mode: str, allow_clear: bool = False) -> list[PlannedChange]:
    result: list[PlannedChange] = []
    for m in matches:
        old_topic = compact(m.lesson.get("lessonContent"))
        old_hw = compact(m.lesson.get("homework"))
        x_topic = compact(m.source.topic)
        x_hw = compact(m.source.homework)

        new_topic = old_topic
        new_hw = old_hw

        if mode == "homework":
            if x_hw or allow_clear:
                new_hw = x_hw
        elif mode == "fill":
            if not old_topic and (x_topic or allow_clear):
                new_topic = x_topic
            if not old_hw and (x_hw or allow_clear):
                new_hw = x_hw
        elif mode == "sync":
            if x_topic or allow_clear:
                new_topic = x_topic
            if x_hw or allow_clear:
                new_hw = x_hw
        else:
            raise NitError(f"Невідомий режим: {mode}")

        # Новий training-plan у НІТ не можна створити без теми. Тому для
        # порожньої календарної колонки беремо тему з Excel навіть у режимі
        # «лише домашнє завдання».
        if not m.lesson.get("id") and not old_topic and x_topic:
            new_topic = x_topic

        result.append(
            PlannedChange(
                lesson=m.lesson,
                source=m.source,
                new_topic=new_topic,
                new_homework=new_hw,
                topic_changed=new_topic != old_topic,
                homework_changed=new_hw != old_hw,
                warning=m.warning,
            )
        )
    return result


def print_preview(changes: list[PlannedChange], has_date: bool, only_changed: bool = False) -> None:
    print("\n" + "=" * 92)
    print("ПОПЕРЕДНІЙ ПЕРЕГЛЯД — НІЧОГО ЩЕ НЕ ЗАПИСАНО")
    if not has_date:
        print("⚠ Excel не має колонки «Дата»: рядки зіставлено з уроками НІТ ЗА ПОРЯДКОМ.")
    print("=" * 92)

    shown = 0
    warnings = 0
    for i, ch in enumerate(changes, 1):
        if only_changed and not ch.changed:
            continue
        shown += 1
        lesson = ch.lesson
        day = iso_day_from_lesson(lesson).strftime("%d.%m.%Y")
        lesson_no = lesson.get("lesson")
        status = "ЗМІНА" if ch.changed else "без змін"
        print(f"\n[{i:02d}] {day} | урок {lesson_no or '-'} | {status} | Excel рядок {ch.source.excel_row}")
        if ch.warning:
            warnings += 1
            print(f"  ⚠ {ch.warning}")

        old_topic = compact(lesson.get("lessonContent"))
        old_hw = compact(lesson.get("homework"))
        if ch.topic_changed:
            print(f"  Тема НІТ  : {short(old_topic) or '[порожньо]'}")
            print(f"  Тема Excel: {short(ch.new_topic) or '[очистити]'}")
        elif not only_changed:
            print(f"  Тема      : {short(old_topic) or '[порожньо]'}")

        if ch.homework_changed:
            print(f"  ДЗ НІТ    : {short(old_hw) or '[порожньо]'}")
            print(f"  ДЗ Excel  : {short(ch.new_homework) or '[очистити]'}")
        elif not only_changed:
            print(f"  ДЗ        : {short(old_hw) or '[порожньо]'}")

    changed_count = sum(1 for x in changes if x.changed)
    print("\n" + "-" * 92)
    print(f"Зіставлено уроків: {len(changes)} | Буде змінено: {changed_count} | Попереджень: {warnings}")
    if only_changed:
        print(f"Показано рядків: {shown}")


def make_backup(
    lessons: list[dict[str, Any]],
    journal: dict[str, Any],
    active_year: dict[str, Any],
    semester: int,
    directory: Path,
) -> Path:
    safe_name = re.sub(r"[^0-9A-Za-zА-Яа-яІіЇїЄєҐґ_-]+", "_", f"{journal['className']}_{journal['subject']}")
    content = {
        "createdAt": datetime.now().isoformat(timespec="seconds"),
        "yearStart": active_year.get("yearStart"),
        "yearEnd": active_year.get("yearEnd"),
        "semester": semester,
        "journal": journal,
        "lessons": lessons,
    }
    return write_secure_backup(content, directory, f"nit_backup_{safe_name}")


def collect_lesson_reports(
    client: NitClient,
    semester_schedules: dict[int, str],
) -> list[JournalLessonReport]:
    """Отримує уроки всіх журналів учителя для вказаних семестрів."""
    reports: list[JournalLessonReport] = []
    seen: set[tuple[int, str]] = set()

    for semester, schedule_id in sorted(semester_schedules.items()):
        journals = client.get_journals(schedule_id)
        print(f"→ {semester} семестр: знайдено моїх журналів — {len(journals)}")
        for journal in journals:
            journal_id = compact(journal.get("journalId"))
            if not journal_id or (semester, journal_id) in seen:
                continue
            seen.add((semester, journal_id))
            print(f"  Отримую уроки: {journal['className']} — {journal['subject']}...", flush=True)
            journal_lessons = client.get_journal_lessons(journal_id)
            training_plans = client.get_training_plans(journal_id)
            lessons = merge_journal_lessons_with_plans(journal_lessons, training_plans)
            reports.append(JournalLessonReport(semester, journal, lessons))

    reports.sort(
        key=lambda item: (
            item.semester,
            item.journal.get("classNumber") or 0,
            compact(item.journal.get("classLetter")),
            compact(item.journal.get("subject")),
        )
    )
    return reports


def merge_journal_lessons_with_plans(
    journal_lessons: list[dict[str, Any]],
    training_plans: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Поєднує повний календар журналу з наявними (заповненими) планами."""
    plans_by_day: dict[date, list[dict[str, Any]]] = {}
    for plan in training_plans:
        plans_by_day.setdefault(iso_day_from_lesson(plan), []).append(plan)
    for plans in plans_by_day.values():
        plans.sort(key=lambda item: int(item.get("lesson") or 0))

    lessons: list[dict[str, Any]] = []
    for journal_lesson in journal_lessons:
        lesson_day = iso_day_from_lesson(journal_lesson)
        candidates = plans_by_day.get(lesson_day) or []
        lesson_number = journal_lesson.get("lesson")
        plan = None
        if lesson_number is not None:
            for index, candidate in enumerate(candidates):
                if int(candidate.get("lesson") or 0) == int(lesson_number):
                    plan = candidates.pop(index)
                    break
        elif candidates:
            plan = candidates.pop(0)

        if plan is not None:
            merged = dict(journal_lesson)
            merged.update(plan)
            merged["journalDayId"] = journal_lesson.get("journalDayId")
            merged["index"] = journal_lesson.get("index")
            merged["isFilled"] = True
        else:
            merged = dict(journal_lesson)
            merged["id"] = None
            merged["isFilled"] = False
        lessons.append(merged)

    lessons.sort(key=lambda item: (iso_day_from_lesson(item), int(item.get("lesson") or 0)))
    return lessons


def print_lesson_reports(reports: list[JournalLessonReport]) -> None:
    print("\n" + "=" * 92)
    print("УРОКИ У ВСІХ МОЇХ ЖУРНАЛАХ")
    print("=" * 92)

    total = 0
    current_semester: int | None = None
    for report in reports:
        if report.semester != current_semester:
            current_semester = report.semester
            print(f"\n{current_semester} семестр")

        journal = report.journal
        total += len(report.lessons)
        filled_count = sum(1 for lesson in report.lessons if lesson.get("isFilled"))
        print(
            f"\n{journal['className']} — {journal['subject']}: "
            f"{len(report.lessons)} уроків "
            f"(заповнено: {filled_count}, не заповнено: {len(report.lessons) - filled_count})"
        )
        if not report.lessons:
            print("  Дат немає")
            continue

        dates = []
        for lesson in report.lessons:
            day = iso_day_from_lesson(lesson).strftime("%d.%m.%Y")
            lesson_no = lesson.get("lesson")
            status = "✓" if lesson.get("isFilled") else "○"
            label = f"{day} (урок {lesson_no})" if lesson_no else day
            dates.append(f"{status} {label}")
        print(textwrap.fill(", ".join(dates), width=92, initial_indent="  ", subsequent_indent="  "))

    print("\n" + "-" * 92)
    print(f"Журналів: {len(reports)} | Уроків разом: {total}")


def save_lesson_reports(
    reports: list[JournalLessonReport],
    output_path: Path,
    active_year: dict[str, Any],
) -> Path:
    """Зберігає звіт у XLSX або JSON відповідно до розширення файлу."""
    output_path = output_path.expanduser()
    if not output_path.is_absolute():
        output_path = Path.cwd() / output_path
    if output_path.suffix.lower() not in {".xlsx", ".json"}:
        raise NitError("Файл звіту повинен мати розширення .xlsx або .json")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if output_path.suffix.lower() == ".json":
        payload = {
            "createdAt": datetime.now().isoformat(timespec="seconds"),
            "yearStart": active_year.get("yearStart"),
            "yearEnd": active_year.get("yearEnd"),
            "journals": [
                {
                    "semester": report.semester,
                    "className": report.journal.get("className"),
                    "subject": report.journal.get("subject"),
                    "journalId": report.journal.get("journalId"),
                    "lessonCount": len(report.lessons),
                    "filledLessonCount": sum(1 for lesson in report.lessons if lesson.get("isFilled")),
                    "lessons": [
                        {
                            "date": iso_day_from_lesson(lesson).isoformat(),
                            "lessonNumber": lesson.get("lesson"),
                            "journalIndex": lesson.get("index"),
                            "journalDayId": lesson.get("journalDayId"),
                            "lessonId": lesson.get("id"),
                            "isFilled": bool(lesson.get("isFilled")),
                        }
                        for lesson in report.lessons
                    ],
                }
                for report in reports
            ],
        }
        output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return output_path

    from openpyxl import Workbook
    from openpyxl.styles import Font

    workbook = Workbook()
    lessons_sheet = workbook.active
    lessons_sheet.title = "Уроки"
    lessons_sheet.append(
        [
            "Семестр",
            "Клас",
            "Предмет",
            "Дата",
            "Номер уроку",
            "Індекс у журналі",
            "Заповнено",
            "ID журналу",
            "ID колонки журналу",
            "ID плану уроку",
        ]
    )
    for cell in lessons_sheet[1]:
        cell.font = Font(bold=True)

    for report in reports:
        for lesson in report.lessons:
            lessons_sheet.append(
                [
                    report.semester,
                    report.journal.get("className"),
                    report.journal.get("subject"),
                    iso_day_from_lesson(lesson).strftime("%d.%m.%Y"),
                    lesson.get("lesson"),
                    lesson.get("index"),
                    "Так" if lesson.get("isFilled") else "Ні",
                    report.journal.get("journalId"),
                    lesson.get("journalDayId"),
                    lesson.get("id"),
                ]
            )

    summary_sheet = workbook.create_sheet("Підсумок")
    summary_sheet.append(
        ["Семестр", "Клас", "Предмет", "Усього уроків", "Заповнено", "Не заповнено"]
    )
    for cell in summary_sheet[1]:
        cell.font = Font(bold=True)
    for report in reports:
        summary_sheet.append(
            [
                report.semester,
                report.journal.get("className"),
                report.journal.get("subject"),
                len(report.lessons),
                sum(1 for lesson in report.lessons if lesson.get("isFilled")),
                sum(1 for lesson in report.lessons if not lesson.get("isFilled")),
            ]
        )

    for sheet in (lessons_sheet, summary_sheet):
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column_cells in sheet.columns:
            width = min(max(len(compact(cell.value)) for cell in column_cells) + 2, 45)
            sheet.column_dimensions[column_cells[0].column_letter].width = width

    workbook.save(output_path)
    return output_path


def collect_absence_reminders(
    client: NitClient,
    semester_schedules: dict[int, str],
) -> list[AbsenceReminder]:
    """Збирає непідтверджені нагадування про відсутність з усіх журналів."""
    result: list[AbsenceReminder] = []
    seen: set[tuple[int, str, str, str]] = set()

    for semester, schedule_id in sorted(semester_schedules.items()):
        journals = client.get_journals(schedule_id)
        print(f"→ {semester} семестр: перевіряю журналів — {len(journals)}")
        for journal in journals:
            journal_id = compact(journal.get("journalId"))
            if not journal_id:
                continue
            print(f"  Нагадування: {journal['className']} — {journal['subject']}...", flush=True)
            pages = client.get_journal_pages(journal_id)
            students = client.get_journal_users(journal_id)
            student_names = {
                compact(student.get("userId")): (
                    compact(student.get("fullName"))
                    or " ".join(
                        part
                        for part in (
                            compact(student.get("lastName")),
                            compact(student.get("firstName")),
                            compact(student.get("middleName")),
                        )
                        if part
                    )
                )
                for student in students
                if compact(student.get("userId"))
            }

            days_by_id: dict[str, dict[str, Any]] = {}
            existing_marks: set[tuple[str, str]] = set()
            raw_reminders: list[dict[str, Any]] = []
            for page in pages:
                for journal_day in page.get("journalDay") or []:
                    if isinstance(journal_day, dict) and journal_day.get("id"):
                        days_by_id[str(journal_day["id"])] = journal_day
                for mark in page.get("marks") or []:
                    if not isinstance(mark, dict):
                        continue
                    column_id = compact(mark.get("columnId") or mark.get("journalDayId"))
                    user_id = compact(mark.get("userId"))
                    if column_id and user_id:
                        existing_marks.add((column_id, user_id))
                raw_reminders.extend(
                    reminder
                    for reminder in (page.get("absenceReminders") or [])
                    if isinstance(reminder, dict)
                )

            for raw in raw_reminders:
                user_id = compact(raw.get("userId"))
                journal_day_id = compact(raw.get("journalDayId"))
                key = (semester, journal_id, journal_day_id, user_id)
                if not user_id or not journal_day_id or key in seen:
                    continue
                seen.add(key)
                if (journal_day_id, user_id) in existing_marks:
                    continue
                journal_day = days_by_id.get(journal_day_id)
                if journal_day is None:
                    raise NitError(
                        f"Журнал {journal['className']} — {journal['subject']}: "
                        f"для нагадування не знайдено колонку {journal_day_id}"
                    )
                result.append(
                    AbsenceReminder(
                        semester,
                        journal,
                        journal_day,
                        user_id,
                        student_names.get(user_id, ""),
                    )
                )

    result.sort(
        key=lambda item: (
            item.semester,
            item.journal.get("classNumber") or 0,
            compact(item.journal.get("classLetter")),
            parse_date_value(item.journal_day.get("date")) or date.min,
            int(item.journal_day.get("lesson") or 0),
            item.user_id,
        )
    )
    return result


def print_absence_preview(reminders: list[AbsenceReminder]) -> None:
    print("\n" + "=" * 92)
    print("ПОПЕРЕДНІЙ ПЕРЕГЛЯД ПІДТВЕРДЖЕННЯ ВІДСУТНОСТЕЙ — НІЧОГО НЕ ЗАПИСАНО")
    print("=" * 92)
    for index, reminder in enumerate(reminders, 1):
        day = parse_date_value(reminder.journal_day.get("date"))
        lesson = reminder.journal_day.get("lesson")
        print(
            f"[{index:03d}] {reminder.journal['className']} — {reminder.journal['subject']} | "
            f"{day.strftime('%d.%m.%Y') if day else '-'} | урок {lesson or '-'} | "
            f"учень …{reminder.user_id[-6:]}"
        )
    print("-" * 92)
    print(f"Буде поставлено позначок «н»: {len(reminders)}")


def normalize_class_name(value: Any) -> str:
    return re.sub(r"[\s\-–—]+", "", compact(value)).upper()


def select_absence_class(
    reminders: list[AbsenceReminder],
    requested_class: str | None = None,
) -> list[AbsenceReminder]:
    """Залишає нагадування одного класу або пропонує інтерактивний вибір."""
    counts: dict[str, int] = {}
    for reminder in reminders:
        class_name = compact(reminder.journal.get("className"))
        counts[class_name] = counts.get(class_name, 0) + 1

    class_names = sorted(
        counts,
        key=lambda name: (
            int(re.search(r"\d+", name).group()) if re.search(r"\d+", name) else 0,
            name,
        ),
    )
    if requested_class:
        normalized = normalize_class_name(requested_class)
        matches = [name for name in class_names if normalize_class_name(name) == normalized]
        if not matches:
            available = ", ".join(class_names)
            raise NitError(
                f"Серед нагадувань немає класу {requested_class!r}. "
                f"Доступні класи: {available or 'немає'}"
            )
        selected = matches[0]
    elif len(class_names) == 1:
        selected = class_names[0]
        print(f"\nКлас із непідтвердженими відсутностями: {selected}")
    else:
        options = [f"Усі класи — {len(reminders)} позначок"]
        options.extend(f"{name} — {counts[name]} позначок" for name in class_names)
        choice = choose_from_menu("Оберіть клас для підтвердження відсутностей:", options)
        if choice == 0:
            return reminders
        selected = class_names[choice - 1]

    selected_normalized = normalize_class_name(selected)
    return [
        reminder
        for reminder in reminders
        if normalize_class_name(reminder.journal.get("className")) == selected_normalized
    ]


def make_absence_backup(reminders: list[AbsenceReminder], directory: Path) -> Path:
    payload = {
        "createdAt": datetime.now().isoformat(timespec="seconds"),
        "note": "Нагадування до автоматичного підтвердження відсутностей",
        "reminders": [
            {
                "semester": reminder.semester,
                "className": reminder.journal.get("className"),
                "subject": reminder.journal.get("subject"),
                "journalId": reminder.journal.get("journalId"),
                "journalDayId": reminder.journal_day.get("id"),
                "date": reminder.journal_day.get("date"),
                "lesson": reminder.journal_day.get("lesson"),
                "userId": reminder.user_id,
            }
            for reminder in reminders
        ],
    }
    return write_secure_backup(payload, directory, "nit_absence_backup")


def extract_jwt_hint(token: str) -> dict[str, Any] | None:
    """Лише для локальної діагностики. Не перевіряє підпис JWT."""
    token = token.strip()
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    parts = token.split(".")
    if len(parts) != 3:
        return None
    try:
        payload = parts[1] + "=" * (-len(parts[1]) % 4)
        return json.loads(base64.urlsafe_b64decode(payload.encode()).decode("utf-8"))
    except Exception:
        return None


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Імпорт тем і домашніх завдань з Excel у журнал НІТ",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent(
            """
            Приклади:
              python3 nit_homework_v2.py homework.xlsx
              python3 nit_homework_v2.py homework.xlsx --mode homework
              python3 nit_homework_v2.py homework.xlsx --mode sync --only-changed
              python3 nit_homework_v2.py --lesson-report
              python3 nit_homework_v2.py --lesson-report --report-output lessons.xlsx
              python3 nit_homework_v2.py --confirm-absences
              python3 nit_homework_v2.py --confirm-absences --class-name "10-А"

            Режими:
              homework  — змінити лише домашні завдання
              fill      — заповнити лише порожні тему/ДЗ у НІТ
              sync      — синхронізувати тему та ДЗ з Excel
            """
        ),
    )
    p.add_argument("excel", nargs="?", help="Excel-файл .xlsx")
    p.add_argument("--sheet", help="Назва аркуша Excel")
    p.add_argument("--mode", choices=["homework", "fill", "sync"], help="Режим імпорту")
    p.add_argument("--semester", type=int, choices=[1, 2], help="Семестр")
    p.add_argument("--school-id", default=os.getenv("NIT_SCHOOL_ID", DEFAULT_SCHOOL_ID))
    p.add_argument("--user-id", default=os.getenv("NIT_USER_ID", DEFAULT_USER_ID))
    p.add_argument("--client-version", default=os.getenv("NIT_CLIENT_VERSION", DEFAULT_CLIENT_VERSION))
    p.add_argument("--token", default=os.getenv("NIT_TOKEN"), help="Bearer token; краще не передавати в командному рядку")
    p.add_argument("--only-changed", action="store_true", help="У preview показувати лише зміни")
    p.add_argument("--allow-clear", action="store_true", help="Дозволити порожнім клітинкам Excel очищати значення НІТ")
    p.add_argument("--yes", action="store_true", help="Після preview не питати фінальне підтвердження (не рекомендовано)")
    p.add_argument(
        "--lesson-report",
        action="store_true",
        help="Показати кількість і дати уроків у всіх моїх журналах; Excel для імпорту не потрібен",
    )
    p.add_argument(
        "--report-output",
        type=Path,
        help="Зберегти звіт про уроки у файл .xlsx або .json (разом із --lesson-report)",
    )
    p.add_argument(
        "--confirm-absences",
        action="store_true",
        help="Знайти червоні трикутники та після preview підтвердити їх позначкою «н»",
    )
    p.add_argument(
        "--class-name",
        help="Клас для --confirm-absences, наприклад 10-А; без параметра буде меню вибору",
    )
    return p.parse_args()


def main() -> int:
    args = parse_args()

    print("NIT² by FIZYK")
    if len(sys.argv) == 1 and not configure_interactive_action(args):
        return 0

    if args.report_output and not args.lesson_report:
        raise NitError("Параметр --report-output працює лише разом із --lesson-report")
    if args.lesson_report and args.confirm_absences:
        raise NitError("Оберіть лише один режим: --lesson-report або --confirm-absences")
    if args.class_name and not args.confirm_absences:
        raise NitError("Параметр --class-name працює лише разом із --confirm-absences")

    if args.lesson_report:
        print("Режим звіту: дані з НІТ лише читаються, жодних змін не буде.\n")
    elif args.confirm_absences:
        print("Режим підтвердження відсутностей: спочатку буде показано preview.\n")
    else:
        print("Безпечний режим: спочатку preview, запис — лише після підтвердження.\n")
        excel_path = Path(args.excel).expanduser() if args.excel else maybe_pick_excel_file()
        sheet_title, excel_rows, has_date = load_excel_rows(excel_path, args.sheet)
        print(f"Excel: {excel_path}")
        print(f"Аркуш: {sheet_title}")
        print(f"Рядків для імпорту: {len(excel_rows)}")
        print(f"Колонка «Дата»: {'так' if has_date else 'ні — буде обережне зіставлення за порядком'}")

    token = get_auth_token(args.token)
    if not token:
        print("Авторизацію скасовано.")
        return 0

    client = NitClient(
        token=token,
        school_id=str(args.school_id),
        user_id=str(args.user_id),
        client_version=str(args.client_version),
    )

    print("→ Перевіряю авторизацію та з’єднання з НІТ...", flush=True)
    active = client.get_active_year()
    print("✓ Авторизація працює, НІТ відповів.")
    y1, y2 = active.get("yearStart"), active.get("yearEnd")
    print(f"\nАктивний навчальний рік НІТ: {y1}/{y2}")

    sem_map: dict[int, str] = {}
    for key, num in (("firstSemester", 1), ("secondSemester", 2)):
        value = active.get(key)
        if isinstance(value, dict) and value.get("id"):
            sem_map[num] = str(value["id"])

    if args.lesson_report:
        selected_schedules = sem_map
        if args.semester is not None:
            if args.semester not in sem_map:
                raise NitError(f"У відповіді НІТ немає scheduleId для {args.semester} семестру")
            selected_schedules = {args.semester: sem_map[args.semester]}
        if not selected_schedules:
            raise NitError("У відповіді НІТ немає жодного доступного семестру")

        reports = collect_lesson_reports(client, selected_schedules)
        if not reports:
            raise NitError("Не знайдено жодного журналу з isMyJournal=true")
        print_lesson_reports(reports)
        if args.report_output:
            saved_path = save_lesson_reports(reports, args.report_output, active)
            print(f"\nЗвіт збережено: {saved_path}")
        return 0

    if args.confirm_absences:
        selected_schedules = sem_map
        if args.semester is not None:
            if args.semester not in sem_map:
                raise NitError(f"У відповіді НІТ немає scheduleId для {args.semester} семестру")
            selected_schedules = {args.semester: sem_map[args.semester]}
        if not selected_schedules:
            raise NitError("У відповіді НІТ немає жодного доступного семестру")

        reminders = collect_absence_reminders(client, selected_schedules)
        if not reminders:
            print("\nНепідтверджених нагадувань про відсутність немає.")
            return 0
        reminders = select_absence_class(reminders, args.class_name)
        print_absence_preview(reminders)

        print(
            "\nУвага: ці позначки будуть записані в офіційний журнал. "
            "Параметр --yes у цьому режимі навмисно не вимикає підтвердження."
        )
        confirm = input("Для запису введи слово ПІДТВЕРДИТИ: ").strip().upper()
        if confirm != "ПІДТВЕРДИТИ":
            print("Скасовано. У НІТ нічого не змінено.")
            return 0

        from nit_core import apply_absence_reminders

        result = apply_absence_reminders(client, reminders, Path.cwd() / "backups")
        print(
            f"\nГотово. Підтверджено: {result.updated}; "
            f"пропущено безпечно: {result.skipped}. Backup: {result.backup_path}"
        )
        for detail in result.details:
            print(f"  • {detail}")
        return 0

    semester = args.semester
    if semester is None:
        available = [x for x in (1, 2) if x in sem_map]
        idx = choose_from_menu("Оберіть семестр:", [f"{x} семестр" for x in available])
        semester = available[idx]
    if semester not in sem_map:
        raise NitError(f"У відповіді НІТ немає scheduleId для {semester} семестру")
    schedule_id = sem_map[semester]

    journals = client.get_journals(schedule_id)
    if not journals:
        raise NitError("Не знайдено жодного журналу з isMyJournal=true")

    journal_idx = choose_from_menu(
        "Оберіть журнал:",
        [f"{j['className']} — {j['subject']}" for j in journals],
    )
    journal = journals[journal_idx]
    print(f"\nОбрано: {journal['className']} — {journal['subject']}")

    journal_id = str(journal["journalId"])
    lessons = merge_journal_lessons_with_plans(
        client.get_journal_lessons(journal_id),
        client.get_training_plans(journal_id),
    )
    for lesson in lessons:
        lesson.setdefault("journalId", journal_id)
    print(f"Уроків у НІТ: {len(lessons)}")

    matches = match_rows_to_lessons(excel_rows, lessons, has_date)

    mode = args.mode
    if mode is None:
        mode_options = [
            ("homework", "Тільки домашні — тему не змінювати"),
            ("fill", "Заповнити пропуски — змінювати тільки порожні тему/ДЗ"),
            ("sync", "Синхронізувати тему + ДЗ з Excel"),
        ]
        idx = choose_from_menu("Режим імпорту:", [x[1] for x in mode_options])
        mode = mode_options[idx][0]

    changes = plan_changes(matches, mode=mode, allow_clear=args.allow_clear)
    print_preview(changes, has_date=has_date, only_changed=args.only_changed)

    changed = [x for x in changes if x.changed]
    if not changed:
        print("\nЗмін немає. Нічого записувати не потрібно.")
        return 0

    if not has_date:
        print(
            "\n⚠ У цьому Excel немає дати. Для поточного файлу це може працювати, якщо "
            "порядок 1:1 збігається з журналом. Уважно перевір preview вище."
        )

    if not args.yes:
        print(f"\nБуде виконано {len(changed)} перевірених операцій запису до НІТ.")
        confirm = input("Для запису введи слово ЗАПИСАТИ: ").strip().upper()
        if confirm != "ЗАПИСАТИ":
            print("Скасовано. У НІТ нічого не змінено.")
            return 0

    from nit_core import ImportPreview, apply_import_preview

    preview = ImportPreview(journal, semester, sheet_title, has_date, len(excel_rows), lessons, changes)
    result = apply_import_preview(client, active, preview, Path.cwd() / "backups")
    print(
        f"\nГотово. Записано і перевірено: {result.verified}/{len(changed)} уроків. "
        f"Backup: {result.backup_path}"
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nСкасовано користувачем.")
        raise SystemExit(130)
    except NitError as e:
        print(f"\nПОМИЛКА: {e}", file=sys.stderr)
        raise SystemExit(1)
