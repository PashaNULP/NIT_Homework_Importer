"""Спільне ядро для консольного та локального вебінтерфейсу НІТ.

Поки перевірені API-функції залишаються в ``nit_homework_v2``. Цей модуль
дає інтерфейсам стабільну точку входу та поступово прийматиме бізнес-логіку під
час подальшого розділення великого CLI-файлу.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from threading import RLock
from typing import Any

from openpyxl import Workbook

from nit_auth import FunctionSessionStore, NitUserConfig, SessionStore

from nit_homework_v2 import (
    DEFAULT_CLIENT_VERSION,
    DEFAULT_SCHOOL_ID,
    DEFAULT_USER_ID,
    AbsenceReminder,
    ExcelRow,
    JournalLessonReport,
    NitClient,
    NitError,
    PlannedChange,
    collect_absence_reminders,
    delete_keychain_session,
    get_my_permissions,
    load_keychain_session,
    load_excel_rows,
    login_session_with_password,
    iso_day_from_lesson,
    make_backup,
    make_absence_backup,
    match_rows_to_lessons,
    merge_journal_lessons_with_plans,
    parse_date_value,
    plan_changes,
    save_keychain_session,
    save_lesson_reports,
    valid_or_refreshed_session,
)


@dataclass
class AuthContext:
    client: NitClient
    active_year: dict[str, Any]
    semester_schedules: dict[int, str]
    user_config: NitUserConfig | None = None
    session_saved: bool = True
    session_warning: str | None = None
    session: dict[str, Any] | None = None
    session_store: SessionStore | None = None
    shared_session_cache: bool = False


@dataclass
class ImportPreview:
    journal: dict[str, Any]
    semester: int
    sheet_title: str
    has_date: bool
    source_rows: int
    lessons: list[dict[str, Any]]
    changes: list[PlannedChange]

    @property
    def changed(self) -> list[PlannedChange]:
        return [item for item in self.changes if item.changed]


@dataclass
class ImportResult:
    updated: int
    backup_path: Path
    created: int = 0
    verified: int = 0


@dataclass
class AbsenceResult:
    updated: int
    backup_path: Path
    skipped: int = 0
    details: tuple[str, ...] = ()


class BatchOperationError(NitError):
    """Помилка пакетного запису з явною інформацією про частковий результат."""

    def __init__(self, message: str, completed: int, total: int, backup_path: Path):
        super().__init__(message)
        self.completed = completed
        self.total = total
        self.backup_path = backup_path


@dataclass
class JournalMatrix:
    journal: dict[str, Any]
    lessons: list[dict[str, Any]]
    students: list[dict[str, Any]]
    marks: dict[tuple[str, str], dict[str, Any]]


_SESSION_LOCK = RLock()
_SESSION_CACHE: dict[tuple[str, str, str], dict[str, Any]] = {}


def _session_key(config: NitUserConfig) -> tuple[str, str, str]:
    return config.profile, config.school_id, config.user_id


def semester_schedule_map(active_year: dict[str, Any]) -> dict[int, str]:
    result: dict[int, str] = {}
    for key, semester in (("firstSemester", 1), ("secondSemester", 2)):
        value = active_year.get(key)
        if isinstance(value, dict) and value.get("id"):
            result[semester] = str(value["id"])
    return result


def _default_session_store(config: NitUserConfig) -> SessionStore:
    account = "nit.school" if config.profile == "default" else f"nit.school.{config.profile}"
    return FunctionSessionStore(
        lambda: load_keychain_session(account),
        lambda session: save_keychain_session(session, account),
        lambda: delete_keychain_session(account),
    )


def _context_from_session(
    session: dict[str, Any],
    user_config: NitUserConfig | None = None,
    session_store: SessionStore | None = None,
    shared_session_cache: bool = False,
) -> AuthContext:
    token = str(session.get("access_token") or "").strip()
    if not token:
        raise NitError("Не вдалося отримати тимчасовий access token")
    config = user_config or NitUserConfig.from_environment(DEFAULT_SCHOOL_ID, DEFAULT_USER_ID)
    client = NitClient(
        token=token,
        school_id=config.school_id,
        user_id=config.user_id,
        client_version=DEFAULT_CLIENT_VERSION,
    )
    active_year = client.get_active_year()
    schedules = semester_schedule_map(active_year)
    if not schedules:
        raise NitError("У відповіді НІТ немає доступних семестрів")
    return AuthContext(
        client,
        active_year,
        schedules,
        config,
        session=dict(session),
        session_store=session_store,
        shared_session_cache=shared_session_cache,
    )


def connect_with_saved_session(
    session_store: SessionStore | None = None,
    user_config: NitUserConfig | None = None,
) -> AuthContext | None:
    config = user_config or NitUserConfig.from_environment(DEFAULT_SCHOOL_ID, DEFAULT_USER_ID)
    use_cache = session_store is None
    store = session_store or _default_session_store(config)
    with _SESSION_LOCK:
        saved = (_SESSION_CACHE.get(_session_key(config)) if use_cache else None) or store.load()
        if not saved:
            return None
        session = valid_or_refreshed_session(saved, save_session=store.save)
        if not session:
            return None
        if use_cache:
            _SESSION_CACHE[_session_key(config)] = dict(session)
    return _context_from_session(session, config, store, use_cache)


def saved_session_available(user_config: NitUserConfig | None = None) -> bool:
    config = user_config or NitUserConfig.from_environment(DEFAULT_SCHOOL_ID, DEFAULT_USER_ID)
    with _SESSION_LOCK:
        if _SESSION_CACHE.get(_session_key(config)):
            return True
    return _default_session_store(config).load() is not None


def connect_with_password(
    username: str,
    password: str,
    session_store: SessionStore | None = None,
    user_config: NitUserConfig | None = None,
) -> AuthContext:
    config = user_config or NitUserConfig.from_environment(DEFAULT_SCHOOL_ID, DEFAULT_USER_ID)
    use_cache = session_store is None
    store = session_store or _default_session_store(config)
    session = login_session_with_password(username, password)
    warning = None
    try:
        store.save(session)
    except (NitError, OSError) as exc:
        # Відмова Keychain не повинна скасовувати вже успішний OAuth-вхід.
        # Пароль у контекст не потрапляє, а сесія працює лише до закриття процесу.
        warning = str(exc)
    context = _context_from_session(session, config, shared_session_cache=use_cache)
    if use_cache:
        with _SESSION_LOCK:
            _SESSION_CACHE[_session_key(config)] = dict(session)
    context.session_store = store
    context.session_saved = warning is None
    context.session_warning = warning
    return context


def discover_user_config(
    session: dict[str, Any],
    preferred_school_id: str | None = None,
) -> NitUserConfig:
    """Визначає userId і школу так само, як офіційний вебклієнт НІТ."""
    token = str(session.get("access_token") or "").strip()
    profile = get_my_permissions(token)
    user_id = str(profile.get("userId") or "").strip()
    permissions = profile.get("permissions")
    if not user_id or not isinstance(permissions, list):
        raise NitError("НІТ не повернув userId або список доступних шкіл.")

    school_ids: list[str] = []
    for item in permissions:
        if not isinstance(item, dict) or not item.get("active"):
            continue
        if str(item.get("forUserId") or "").strip() != user_id:
            continue
        school_id = str(item.get("schoolId") or "").strip()
        if school_id and school_id != "0" and school_id not in school_ids:
            school_ids.append(school_id)

    preferred = str(preferred_school_id or "").strip()
    if preferred:
        if preferred not in school_ids:
            raise NitError(
                "Цей обліковий запис не має активного доступу до школи, "
                "для якої налаштовано NIT²."
            )
        selected_school = preferred
    elif len(school_ids) == 1:
        selected_school = school_ids[0]
    elif not school_ids:
        raise NitError("Для цього облікового запису НІТ не знайдено активної школи.")
    else:
        raise NitError(
            "Обліковий запис має доступ до кількох шкіл. "
            "Адміністратор NIT² має вказати NIT_SCHOOL_ID."
        )
    return NitUserConfig(selected_school, user_id, profile="web")


def connect_user_with_password(
    username: str,
    password: str,
    session_store: SessionStore,
    preferred_school_id: str | None = None,
) -> AuthContext:
    """Окремий вхід вебкористувача без Keychain і спільного кешу токенів."""
    session = login_session_with_password(username, password)
    config = discover_user_config(session, preferred_school_id)
    session_store.save(session)
    context = _context_from_session(session, config, session_store, False)
    context.session_saved = True
    return context


def refresh_auth_context(context: AuthContext, min_validity_seconds: int = 300) -> bool:
    """Заздалегідь поновлює сесію і підміняє token в чинному HTTP-клієнті."""
    config = context.user_config
    store = context.session_store
    if config is None or store is None or context.session is None:
        return False
    key = _session_key(config)
    with _SESSION_LOCK:
        candidate = (
            _SESSION_CACHE.get(key) if context.shared_session_cache else None
        ) or context.session
        refreshed = valid_or_refreshed_session(
            candidate,
            save_session=store.save,
            min_validity_seconds=min_validity_seconds,
        )
        if not refreshed:
            return False
        if context.shared_session_cache:
            _SESSION_CACHE[key] = dict(refreshed)
        context.session = dict(refreshed)
        context.client.set_access_token(str(refreshed.get("access_token") or ""))
    return True


def build_all_teachers_schedule(
    client: NitClient,
    week_start: date,
) -> dict[str, Any]:
    """Формує безпечний JSON-об'єкт тижневого розкладу всіх учителів.

    Довідник НІТ також містить телефони, email і дати народження. Вони навмисно
    не копіюються до експорту: залишаються лише ім'я та дані уроків.
    """
    start_day = week_start - timedelta(days=week_start.weekday())
    end_day = start_day + timedelta(days=6)
    raw_teachers = client.get_school_teachers(include_inactive=False)
    teachers: dict[str, dict[str, Any]] = {}
    for raw in raw_teachers:
        teacher_id = str(raw.get("id") or "").strip()
        if not teacher_id or not raw.get("isActive", True):
            continue
        teachers[teacher_id] = {
            "id": teacher_id,
            "name": str(raw.get("fullName") or "").strip(),
            "lessons": [],
        }

    rows = client.get_school_journal_days(start_day, end_day, list(teachers))
    seen: set[str] = set()
    for raw in rows:
        nested = raw.get("journalTeacher")
        nested = nested if isinstance(nested, dict) else {}
        teacher_id = str(
            raw.get("journalTeacherId") or nested.get("id") or ""
        ).strip()
        if not teacher_id:
            continue
        if teacher_id not in teachers:
            teachers[teacher_id] = {
                "id": teacher_id,
                "name": str(
                    nested.get("fullName")
                    or raw.get("journalTeacherShortName")
                    or ""
                ).strip(),
                "lessons": [],
            }
        row_id = str(raw.get("id") or "").strip()
        unique_key = row_id or "|".join(
            str(raw.get(field) or "")
            for field in ("date", "lesson", "journalId", "classId", "subjectId")
        )
        if unique_key in seen:
            continue
        seen.add(unique_key)
        lesson_value = raw.get("lesson")
        try:
            lesson_number: int | float | str = float(lesson_value)
            if lesson_number.is_integer():
                lesson_number = int(lesson_number)
        except (TypeError, ValueError):
            lesson_number = str(lesson_value or "")
        teachers[teacher_id]["lessons"].append(
            {
                "id": row_id or None,
                "date": str(raw.get("date") or "")[:10],
                "lesson": lesson_number,
                "classId": raw.get("classId"),
                "className": raw.get("className"),
                "subjectId": raw.get("subjectId"),
                "subject": raw.get("subject"),
                "classroom": raw.get("classroom"),
                "journalId": raw.get("journalId"),
                "isReplacement": bool(raw.get("replace")),
            }
        )

    teacher_rows = sorted(
        teachers.values(),
        key=lambda item: (str(item.get("name") or "").casefold(), item["id"]),
    )
    for teacher in teacher_rows:
        teacher["lessons"].sort(
            key=lambda item: (
                item["date"],
                item["lesson"] if isinstance(item["lesson"], (int, float)) else 999,
                str(item["className"] or ""),
            )
        )
    return {
        "schema": "nit2.teacher-schedules.v1",
        "generatedAt": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "weekStart": start_day.isoformat(),
        "weekEnd": end_day.isoformat(),
        "summary": {
            "teachers": len(teacher_rows),
            "teachersWithLessons": sum(bool(item["lessons"]) for item in teacher_rows),
            "lessons": sum(len(item["lessons"]) for item in teacher_rows),
        },
        "teachers": teacher_rows,
    }


def load_journal_report(
    client: NitClient,
    semester: int,
    journal: dict[str, Any],
) -> JournalLessonReport:
    journal_id = str(journal.get("journalId") or "").strip()
    if not journal_id:
        raise NitError("У вибраному журналі немає journalId")
    journal_lessons = client.get_journal_lessons(journal_id)
    training_plans = client.get_training_plans(journal_id)
    lessons = merge_journal_lessons_with_plans(journal_lessons, training_plans)
    return JournalLessonReport(semester=semester, journal=journal, lessons=lessons)


def load_calendar_reports(
    client: NitClient,
    semester: int,
    journals: list[dict[str, Any]],
    cached_report: JournalLessonReport | None = None,
) -> list[JournalLessonReport]:
    """Завантажує уроки всіх журналів для спільного календаря."""
    cached_id = ""
    if cached_report is not None:
        cached_id = str(cached_report.journal.get("journalId") or "")

    reports: list[JournalLessonReport] = []
    for journal in journals:
        journal_id = str(journal.get("journalId") or "")
        # НІТ іноді повертає «мій» предмет, для якого журнал ще не створений.
        if not journal_id:
            continue
        if cached_report is not None and journal_id and journal_id == cached_id:
            reports.append(cached_report)
        else:
            reports.append(load_journal_report(client, semester, journal))
    return reports


def load_journal_matrix(
    client: NitClient,
    journal: dict[str, Any],
) -> JournalMatrix:
    """Збирає всі сторінки журналу в одну матрицю без пагінації."""
    journal_id = str(journal.get("journalId") or "").strip()
    if not journal_id:
        raise NitError("У вибраному журналі немає journalId")

    pages = client.get_journal_pages(journal_id)
    students = client.get_journal_users(journal_id)
    lessons_by_id: dict[str, dict[str, Any]] = {}
    marks: dict[tuple[str, str], dict[str, Any]] = {}

    for page in pages:
        for column in page.get("journalDay") or []:
            if not isinstance(column, dict):
                continue
            if column.get("type") not in (0, "0", "Day", "day"):
                continue
            column_id = str(column.get("id") or "").strip()
            if column_id:
                lessons_by_id[column_id] = column

        for mark in page.get("marks") or []:
            if not isinstance(mark, dict):
                continue
            user_id = str(mark.get("userId") or "").strip()
            column_id = str(mark.get("columnId") or mark.get("journalDayId") or "").strip()
            if user_id and column_id:
                marks[(user_id, column_id)] = mark

    def lesson_key(column: dict[str, Any]) -> tuple[date, int, int]:
        try:
            day = parse_date_value(column.get("date")) or date.max
        except (TypeError, ValueError):
            day = date.max
        try:
            lesson = int(column.get("lesson") or 0)
        except (TypeError, ValueError):
            lesson = 0
        try:
            index = int(column.get("index") or 0)
        except (TypeError, ValueError):
            index = 0
        return day, lesson, index

    lessons = sorted(lessons_by_id.values(), key=lesson_key)
    return JournalMatrix(journal=journal, lessons=lessons, students=students, marks=marks)


def read_excel_upload(path: Path) -> tuple[str, list[ExcelRow], bool]:
    return load_excel_rows(path)


def save_import_template(report: JournalLessonReport, path: Path) -> Path:
    """Створює редагований Excel лише з уроками, які мають training-plan ID."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Імпорт"
    sheet.append(["Дата", "Урок", "Тема уроку", "Домашнє завдання"])
    for lesson in report.lessons:
        sheet.append(
            [
                iso_day_from_lesson(lesson).strftime("%d.%m.%Y"),
                lesson.get("lesson"),
                lesson.get("lessonContent") or "",
                lesson.get("homework") or "",
            ]
        )
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column, width in {"A": 14, "B": 10, "C": 55, "D": 55}.items():
        sheet.column_dimensions[column].width = width
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    return path


def prepare_import_preview(
    client: NitClient,
    semester: int,
    journal: dict[str, Any],
    sheet_title: str,
    rows: list[ExcelRow],
    has_date: bool,
    mode: str,
    allow_clear: bool = False,
) -> ImportPreview:
    journal_id = str(journal.get("journalId") or "").strip()
    if not journal_id:
        raise NitError("У вибраному журналі немає journalId")
    journal_lessons = client.get_journal_lessons(journal_id)
    training_plans = client.get_training_plans(journal_id)
    lessons = merge_journal_lessons_with_plans(journal_lessons, training_plans)
    for lesson in lessons:
        lesson.setdefault("journalId", journal_id)
    matches = match_rows_to_lessons(rows, lessons, has_date)
    changes = plan_changes(matches, mode=mode, allow_clear=allow_clear)
    return ImportPreview(
        journal=journal,
        semester=semester,
        sheet_title=sheet_title,
        has_date=has_date,
        source_rows=len(rows),
        lessons=lessons,
        changes=changes,
    )


def apply_import_preview(
    client: NitClient,
    active_year: dict[str, Any],
    preview: ImportPreview,
    backup_directory: Path,
) -> ImportResult:
    changed = preview.changed
    if not changed:
        raise NitError("У preview немає змін для запису")

    changed.sort(
        key=lambda change: (
            iso_day_from_lesson(change.lesson),
            int(change.lesson.get("lesson") or 0),
        )
    )

    # POST /training-plans у НІТ завжди створює наступний вільний план.
    # Тому не можна перескочити через попередній порожній урок.
    empty_lessons = [lesson for lesson in preview.lessons if not lesson.get("id")]
    empty_lessons.sort(
        key=lambda lesson: (iso_day_from_lesson(lesson), int(lesson.get("lesson") or 0))
    )

    def identity(lesson: dict[str, Any]) -> str:
        return str(lesson.get("journalDayId") or "").strip()

    new_changes = [change for change in changed if not change.lesson.get("id")]
    if new_changes:
        positions = {identity(lesson): index for index, lesson in enumerate(empty_lessons)}
        changes_by_column = {identity(change.lesson): change for change in new_changes}
        last_position = max(positions[identity(change.lesson)] for change in new_changes)
        for lesson in empty_lessons[: last_position + 1]:
            change = changes_by_column.get(identity(lesson))
            day = iso_day_from_lesson(lesson).strftime("%d.%m.%Y")
            number = lesson.get("lesson") or "—"
            if change is None:
                raise NitError(
                    f"НІТ створює плани послідовно. Додайте в Excel попередній "
                    f"порожній урок {day} (урок {number})."
                )
            if not change.new_topic.strip():
                raise NitError(
                    f"Для нового уроку {day} (урок {number}) заповніть «Тема уроку»: "
                    "НІТ не дозволяє створити план без теми."
                )

    backup_directory.mkdir(parents=True, exist_ok=True)
    backup_path = make_backup(
        preview.lessons,
        preview.journal,
        active_year,
        preview.semester,
        backup_directory,
    )

    updated = 0
    created = 0
    verified = 0
    journal_id = str(preview.journal.get("journalId") or "")

    def current_lessons() -> list[dict[str, Any]]:
        calendar = client.get_journal_lessons(journal_id)
        plans = client.get_training_plans(journal_id)
        fresh = merge_journal_lessons_with_plans(calendar, plans)
        for item in fresh:
            item.setdefault("journalId", journal_id)
        return fresh

    def same_column(candidate: dict[str, Any], baseline: dict[str, Any]) -> bool:
        column_id = str(baseline.get("journalDayId") or "")
        if column_id:
            return str(candidate.get("journalDayId") or "") == column_id
        lesson_id = str(baseline.get("id") or "")
        return bool(lesson_id) and str(candidate.get("id") or "") == lesson_id

    def find_current(baseline: dict[str, Any]) -> dict[str, Any]:
        match = next((item for item in current_lessons() if same_column(item, baseline)), None)
        if match is None:
            raise NitError("Відповідна календарна колонка більше не існує в НІТ")
        return match

    def values(item: dict[str, Any]) -> tuple[str, str]:
        return str(item.get("lessonContent") or "").strip(), str(item.get("homework") or "").strip()

    for change in changed:
        try:
            fresh = find_current(change.lesson)
            baseline_values = values(change.lesson)
            fresh_values = values(fresh)
            if baseline_values != fresh_values or bool(change.lesson.get("id")) != bool(fresh.get("id")):
                raise NitError(
                    "Конфлікт: урок змінився після формування Preview. "
                    "Оновіть дані та побудуйте новий Preview."
                )

            if fresh.get("id"):
                client.update_training_plan(
                    fresh,
                    change.new_topic,
                    change.new_homework,
                )
            else:
                response = client.create_training_plan(
                    fresh,
                    journal_id,
                    change.new_topic,
                    change.new_homework,
                )
                if response and (response.get("day") or response.get("lesson") is not None):
                    expected_day = iso_day_from_lesson(change.lesson)
                    actual_day = parse_date_value(response.get("day"))
                    expected_lesson = int(change.lesson.get("lesson") or 0)
                    actual_lesson = int(response.get("lesson") or 0)
                    if actual_day != expected_day or actual_lesson != expected_lesson:
                        raise NitError(
                            "НІТ створив план не для очікуваної календарної колонки; "
                            "подальший імпорт зупинено."
                        )
                created += 1
        except Exception as exc:
            # Після мережевої помилки не повторюємо запис. Спочатку читаємо
            # фактичний стан: запит міг бути прийнятий сервером.
            try:
                actual = find_current(change.lesson)
                if values(actual) == (change.new_topic.strip(), change.new_homework.strip()):
                    updated += 1
                    verified += 1
                    if not change.lesson.get("id"):
                        created += 1
                    continue
            except Exception:
                pass
            raise BatchOperationError(
                f"Запис зупинено після {updated}/{len(changed)} змін. "
                f"Backup до змін: {backup_path}. Причина: {exc}",
                updated,
                len(changed),
                backup_path,
            ) from exc

        actual = find_current(change.lesson)
        if values(actual) != (change.new_topic.strip(), change.new_homework.strip()):
            raise BatchOperationError(
                f"НІТ не підтвердив результат після {updated}/{len(changed)} змін. "
                f"Backup: {backup_path}",
                updated,
                len(changed),
                backup_path,
            )
        updated += 1
        verified += 1
    return ImportResult(
        updated=updated,
        backup_path=backup_path,
        created=created,
        verified=verified,
    )


def load_absence_reminders(
    client: NitClient,
    semester: int,
    schedule_id: str,
) -> list[AbsenceReminder]:
    return collect_absence_reminders(client, {semester: schedule_id})


def apply_absence_reminders(
    client: NitClient,
    reminders: list[AbsenceReminder],
    backup_directory: Path,
) -> AbsenceResult:
    if not reminders:
        raise NitError("Не вибрано жодної відсутності")
    backup_directory.mkdir(parents=True, exist_ok=True)
    backup_path = make_absence_backup(reminders, backup_directory)
    updated = 0
    skipped = 0
    details: list[str] = []

    def state(reminder: AbsenceReminder) -> str:
        journal_id = str(reminder.journal.get("journalId") or "")
        column_id = str(reminder.journal_day.get("id") or "")
        user_id = str(reminder.user_id or "")
        if not journal_id or not column_id or not user_id:
            return "invalid"
        pages = client.get_journal_pages(journal_id)
        marks = [
            mark
            for page in pages
            for mark in (page.get("marks") or [])
            if isinstance(mark, dict)
            and str(mark.get("columnId") or mark.get("journalDayId") or "") == column_id
            and str(mark.get("userId") or "") == user_id
        ]
        if len(marks) > 1:
            return "ambiguous"
        if marks:
            mark = marks[0]
            other = any(
                mark.get(field) is not None
                for field in ("rating", "customMark", "leveledMark", "letterMark", "isVerbalFormed")
            )
            if mark.get("absense") == 0 and not other:
                return "confirmed"
            return "occupied"
        pending = any(
            isinstance(raw, dict)
            and str(raw.get("journalDayId") or "") == column_id
            and str(raw.get("userId") or "") == user_id
            for page in pages
            for raw in (page.get("absenceReminders") or [])
        )
        return "pending" if pending else "unknown"

    for reminder in reminders:
        current = state(reminder)
        label = (
            f"{reminder.journal.get('className') or '—'} | "
            f"{reminder.journal_day.get('date') or '—'} | "
            f"урок {reminder.journal_day.get('lesson') or '—'} | "
            f"учень …{str(reminder.user_id or '')[-6:]}"
        )
        if current == "confirmed":
            skipped += 1
            details.append(f"{label}: позначка вже підтверджена")
            continue
        if current != "pending":
            skipped += 1
            details.append(f"{label}: пропущено через стан клітинки {current}")
            continue
        try:
            client.create_absence_mark(reminder)
        except Exception as exc:
            # Не повторюємо POST автоматично: спершу з’ясовуємо результат.
            if state(reminder) == "confirmed":
                updated += 1
                continue
            raise BatchOperationError(
                f"Підтвердження зупинено після {updated}/{len(reminders)} позначок. "
                f"Backup: {backup_path}. Причина: {exc}",
                updated,
                len(reminders),
                backup_path,
            ) from exc
        if state(reminder) != "confirmed":
            raise BatchOperationError(
                f"НІТ не підтвердив позначку після {updated}/{len(reminders)} операцій. "
                f"Backup: {backup_path}",
                updated,
                len(reminders),
                backup_path,
            )
        updated += 1
    return AbsenceResult(updated=updated, backup_path=backup_path, skipped=skipped, details=tuple(details))


__all__ = [
    "AuthContext",
    "BatchOperationError",
    "AbsenceReminder",
    "AbsenceResult",
    "ExcelRow",
    "ImportPreview",
    "ImportResult",
    "JournalLessonReport",
    "JournalMatrix",
    "NitError",
    "connect_with_password",
    "connect_with_saved_session",
    "connect_user_with_password",
    "discover_user_config",
    "apply_absence_reminders",
    "apply_import_preview",
    "build_all_teachers_schedule",
    "load_absence_reminders",
    "load_calendar_reports",
    "load_journal_report",
    "load_journal_matrix",
    "prepare_import_preview",
    "read_excel_upload",
    "refresh_auth_context",
    "save_import_template",
    "saved_session_available",
    "save_lesson_reports",
]
