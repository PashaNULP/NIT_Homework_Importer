#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NIT Homework Importer

Імпорт тем уроків і домашніх завдань з Excel у електронний журнал НІТ.

Безпечна поведінка:
- спочатку лише показує preview;
- нічого не записує без явного підтвердження;
- перед записом створює локальний JSON-backup поточного стану журналу;
- порожні клітинки Excel за замовчуванням НЕ стирають дані в НІТ.

API було відновлено з мережевих запитів вебзастосунку НІТ:
  GET  /api/v1/academic-years/active
  GET  /api/v1/journals/classes
  GET  /api/v1/training-plans
  PUT  /api/v1/training-plans/{id}
"""

from __future__ import annotations

import argparse
import base64
import getpass
import json
import os
import re
import sys
import textwrap
from dataclasses import dataclass
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable

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

# Персональні ID зберігаються лише в ігнорованому локальному файлі.
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
        timeout: int = 12,
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
        print("→ Перевіряю з’єднання з НІТ...", flush=True)
        try:
            r = self.session.get(
                f"{API_BASE}/academic-years/active",
                headers={"Alias": "academic-year-active"},
                timeout=self.timeout,
            )
        except requests.exceptions.Timeout as e:
            raise NitError(f"НІТ не відповів за {self.timeout} с (timeout). Перевір інтернет/VPN або TLS Python.") from e
        except requests.exceptions.SSLError as e:
            raise NitError(f"Помилка TLS/SSL при підключенні до НІТ: {e}") from e
        except requests.exceptions.RequestException as e:
            raise NitError(f"Не вдалося підключитися до НІТ: {e}") from e
        print(f"✓ НІТ відповів: HTTP {r.status_code}", flush=True)
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


def match_rows_to_lessons(rows: list[ExcelRow], lessons: list[dict[str, Any]], has_date: bool) -> list[Match]:
    if has_date:
        by_day: dict[date, list[dict[str, Any]]] = {}
        for lesson in lessons:
            by_day.setdefault(iso_day_from_lesson(lesson), []).append(lesson)

        matches: list[Match] = []
        used_ids: set[str] = set()
        for row in rows:
            if row.day is None:
                raise NitError(
                    f"Excel, рядок {row.excel_row}: колонка «Дата» є, але дата порожня. "
                    "Заповни дату або прибери цей рядок."
                )
            candidates = [x for x in by_day.get(row.day, []) if compact(x.get("id")) not in used_ids]
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
            used_ids.add(compact(lesson.get("id")))
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
              python3 nit_homework.py homework.xlsx
              python3 nit_homework.py homework.xlsx --mode homework
              python3 nit_homework.py homework.xlsx --mode sync --only-changed

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
    return p.parse_args()


def main() -> int:
    args = parse_args()

    print("NIT Homework Importer")
    print("Безпечний режим: спочатку preview, запис — лише після підтвердження.\n")

    excel_path = Path(args.excel).expanduser() if args.excel else maybe_pick_excel_file()
    sheet_title, excel_rows, has_date = load_excel_rows(excel_path, args.sheet)
    print(f"Excel: {excel_path}")
    print(f"Аркуш: {sheet_title}")
    print(f"Рядків для імпорту: {len(excel_rows)}")
    print(f"Колонка «Дата»: {'так' if has_date else 'ні — буде обережне зіставлення за порядком'}")

    token = args.token
    if not token:
        print("\nВстав Bearer token НІТ і натисни Enter.")
        print("Токен не буде виведений назад програмою, але під час вставки його може бути видно в Terminal.")
        token = input("Token: ").strip()
    if not token:
        raise NitError("Bearer token не введено.")
    clean_token = token[7:].strip() if token.lower().startswith("bearer ") else token.strip()
    print(f"✓ Токен отримано ({len(clean_token)} символів). Сам токен не показую.", flush=True)

    client = NitClient(
        token=token,
        school_id=str(args.school_id),
        user_id=str(args.user_id),
        client_version=str(args.client_version),
    )

    active = client.get_active_year()
    y1, y2 = active.get("yearStart"), active.get("yearEnd")
    print(f"\nАктивний навчальний рік НІТ: {y1}/{y2}")

    sem_map: dict[int, str] = {}
    for key, num in (("firstSemester", 1), ("secondSemester", 2)):
        value = active.get(key)
        if isinstance(value, dict) and value.get("id"):
            sem_map[num] = str(value["id"])

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

    lessons = client.get_training_plans(str(journal["journalId"]))
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
        print(f"\nБуде виконано {len(changed)} PUT-запитів до НІТ.")
        confirm = input("Для запису введи слово ЗАПИСАТИ: ").strip().upper()
        if confirm != "ЗАПИСАТИ":
            print("Скасовано. У НІТ нічого не змінено.")
            return 0

    backup_dir = Path.cwd() / "backups"
    backup_path = make_backup(lessons, journal, active, semester, backup_dir)
    print(f"\nBackup створено: {backup_path}")

    ok = 0
    for i, ch in enumerate(changed, 1):
        day = iso_day_from_lesson(ch.lesson).strftime("%d.%m.%Y")
        print(f"[{i}/{len(changed)}] {day} ... ", end="", flush=True)
        try:
            client.update_training_plan(ch.lesson, ch.new_topic, ch.new_homework)
        except Exception:
            print("ПОМИЛКА")
            print(
                "\nЗапис зупинено. Попередні успішні зміни могли вже зберегтися в НІТ. "
                f"Backup: {backup_path}"
            )
            raise
        else:
            ok += 1
            print("OK")

    print(f"\nГотово. Успішно оновлено: {ok}/{len(changed)} уроків.")
    print(f"Backup до змін: {backup_path}")
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
        raise SystemExit(1)
