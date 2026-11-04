"""Персональний розклад NIT²: нормалізація teacher-diary, кеш і локальні дзвінки.

Цей модуль не виконує API-записів і не зберігає токени.
"""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
from dataclasses import asdict, dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo


KYIV = ZoneInfo("Europe/Kyiv")


@dataclass(frozen=True)
class ScheduleLesson:
    day: date
    weekday: int
    number: int
    subject: str
    class_name: str
    is_fds: bool = False
    room: str | None = None


@dataclass
class WeekSchedule:
    week_start: date
    lessons: list[ScheduleLesson]
    fetched_at: datetime
    source: str = "nit"
    warning: str | None = None


@dataclass
class BellPeriod:
    number: int
    start: str
    end: str


@dataclass
class BellSettings:
    active_profile: str
    profiles: dict[str, list[BellPeriod]]


@dataclass
class OfficialBells:
    year_start: int
    year_end: int
    periods: list[BellPeriod]
    fetched_at: datetime
    source: str = "nit"


def monday_of(day: date) -> date:
    return day - timedelta(days=day.weekday())


def normalize_teacher_diary(payload: Any, requested_week: date) -> WeekSchedule:
    if not isinstance(payload, dict) or not isinstance(payload.get("data"), list):
        raise ValueError("НІТ повернув невідомий формат teacher-diary")
    lessons: list[ScheduleLesson] = []
    for raw_day in payload["data"]:
        if not isinstance(raw_day, dict):
            continue
        try:
            lesson_day = date.fromisoformat(str(raw_day.get("date") or "")[:10])
        except ValueError:
            continue
        for raw in raw_day.get("subjects") or []:
            if not isinstance(raw, dict):
                continue
            try:
                number = int(raw.get("number"))
            except (TypeError, ValueError):
                continue
            if number <= 0:
                continue
            lessons.append(
                ScheduleLesson(
                    day=lesson_day,
                    weekday=lesson_day.weekday(),
                    number=number,
                    subject=str(raw.get("subject") or "").strip(),
                    class_name=str(raw.get("className") or "").strip(),
                    is_fds=bool(raw.get("isFds")),
                    room=(str(raw.get("room")).strip() if raw.get("room") else None),
                )
            )
    lessons.sort(key=lambda item: (item.day, item.number, item.class_name, item.subject))
    return WeekSchedule(monday_of(requested_week), lessons, datetime.now(KYIV))


def normalize_academic_year_bells(payload: Any, active_year: dict[str, Any]) -> OfficialBells:
    years = payload.get("data") if isinstance(payload, dict) else None
    if not isinstance(years, list):
        raise ValueError("НІТ повернув невідомий формат academic-years")
    start, end = active_year.get("yearStart"), active_year.get("yearEnd")
    year = next(
        (item for item in years if item.get("yearStart") == start and item.get("yearEnd") == end),
        None,
    )
    if not isinstance(year, dict):
        raise ValueError("Активний навчальний рік не знайдено у повному списку")
    periods: list[BellPeriod] = []
    for item in year.get("lessons") or []:
        if not isinstance(item, dict):
            continue
        try:
            number = int(item.get("number"))
        except (TypeError, ValueError):
            continue
        start_time = str(item.get("fromTime") or "").strip()[:5]
        end_time = str(item.get("toTime") or "").strip()[:5]
        if parse_clock(start_time) is None or parse_clock(end_time) is None:
            continue
        periods.append(BellPeriod(number, start_time, end_time))
    periods.sort(key=lambda item: item.number)
    if not periods:
        raise ValueError("Для активного року НІТ не повернув часу дзвінків")
    return OfficialBells(int(start), int(end), periods, datetime.now(KYIV))


def profile_storage_key(profile: str, school_id: str, user_id: str) -> str:
    raw = f"{profile}\0{school_id}\0{user_id}".encode("utf-8")
    return hashlib.sha256(raw).hexdigest()[:20]


def _atomic_json(path: Path, document: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    try:
        os.chmod(path.parent, 0o700)
    except OSError:
        pass
    fd, temporary = tempfile.mkstemp(prefix=".nit2_schedule_", suffix=".tmp", dir=path.parent)
    temporary_path = Path(temporary)
    try:
        if hasattr(os, "fchmod"):
            os.fchmod(fd, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump(document, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary_path, path)
        try:
            os.chmod(path, 0o600)
        except OSError:
            pass
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise


class ScheduleStorage:
    def __init__(self, root: Path, identity: str):
        self.root = root
        self.identity = identity
        self.root.mkdir(parents=True, exist_ok=True, mode=0o700)
        for directory in (self.root, self.root.parent):
            try:
                os.chmod(directory, 0o700)
            except OSError:
                pass

    def _week_path(self, week_start: date) -> Path:
        return self.root / f"week_{self.identity}_{monday_of(week_start).isoformat()}.json"

    @property
    def bells_path(self) -> Path:
        return self.root / f"bells_{self.identity}.json"

    @property
    def official_bells_path(self) -> Path:
        return self.root / f"nit_bells_{self.identity}.json"

    def save_week(self, schedule: WeekSchedule) -> None:
        _atomic_json(
            self._week_path(schedule.week_start),
            {
                "format": "NIT2-SCHEDULE-1",
                "weekStart": schedule.week_start.isoformat(),
                "fetchedAt": schedule.fetched_at.isoformat(),
                "lessons": [
                    {**asdict(lesson), "day": lesson.day.isoformat()} for lesson in schedule.lessons
                ],
            },
        )

    def load_week(self, week_start: date) -> WeekSchedule | None:
        path = self._week_path(week_start)
        if not path.exists():
            return None
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
            if raw.get("format") != "NIT2-SCHEDULE-1":
                return None
            lessons = [
                ScheduleLesson(
                    day=date.fromisoformat(item["day"]),
                    weekday=int(item["weekday"]),
                    number=int(item["number"]),
                    subject=str(item.get("subject") or ""),
                    class_name=str(item.get("class_name") or ""),
                    is_fds=bool(item.get("is_fds")),
                    room=item.get("room"),
                )
                for item in raw.get("lessons") or []
            ]
            return WeekSchedule(
                week_start=date.fromisoformat(raw["weekStart"]),
                lessons=lessons,
                fetched_at=datetime.fromisoformat(raw["fetchedAt"]),
                source="cache",
            )
        except (KeyError, TypeError, ValueError, OSError, json.JSONDecodeError):
            return None

    def load_bells(self) -> BellSettings:
        default = BellSettings("Звичайний", {"Звичайний": [], "Скорочений": []})
        if not self.bells_path.exists():
            return default
        try:
            raw = json.loads(self.bells_path.read_text(encoding="utf-8"))
            profiles = {
                str(name): [
                    BellPeriod(int(item["number"]), str(item["start"]), str(item["end"]))
                    for item in periods
                ]
                for name, periods in (raw.get("profiles") or {}).items()
            }
            if not profiles:
                return default
            active = str(raw.get("activeProfile") or "")
            return BellSettings(active if active in profiles else next(iter(profiles)), profiles)
        except (KeyError, TypeError, ValueError, OSError, json.JSONDecodeError):
            return default

    def save_bells(self, settings: BellSettings) -> None:
        _atomic_json(
            self.bells_path,
            {
                "format": "NIT2-BELLS-1",
                "activeProfile": settings.active_profile,
                "profiles": {
                    name: [asdict(period) for period in periods]
                    for name, periods in settings.profiles.items()
                },
            },
        )

    def save_official_bells(self, bells: OfficialBells) -> None:
        _atomic_json(
            self.official_bells_path,
            {
                "format": "NIT2-OFFICIAL-BELLS-1",
                "yearStart": bells.year_start,
                "yearEnd": bells.year_end,
                "fetchedAt": bells.fetched_at.isoformat(),
                "periods": [asdict(period) for period in bells.periods],
            },
        )

    def load_official_bells(self) -> OfficialBells | None:
        if not self.official_bells_path.exists():
            return None
        try:
            raw = json.loads(self.official_bells_path.read_text(encoding="utf-8"))
            if raw.get("format") != "NIT2-OFFICIAL-BELLS-1":
                return None
            return OfficialBells(
                int(raw["yearStart"]),
                int(raw["yearEnd"]),
                [BellPeriod(int(item["number"]), str(item["start"]), str(item["end"])) for item in raw.get("periods") or []],
                datetime.fromisoformat(raw["fetchedAt"]),
                source="cache",
            )
        except (KeyError, TypeError, ValueError, OSError, json.JSONDecodeError):
            return None


def parse_clock(value: str) -> time | None:
    try:
        return time.fromisoformat(value.strip())
    except (TypeError, ValueError):
        return None


def break_minutes(current: BellPeriod, following: BellPeriod | None) -> int | None:
    end = parse_clock(current.end)
    start = parse_clock(following.start) if following else None
    if end is None or start is None:
        return None
    minutes = int(
        (datetime.combine(date.today(), start) - datetime.combine(date.today(), end)).total_seconds()
        // 60
    )
    return minutes if minutes >= 0 else None


def current_and_next(
    schedule: WeekSchedule | None,
    bells: BellSettings,
    now: datetime | None = None,
) -> tuple[tuple[ScheduleLesson, BellPeriod] | None, tuple[ScheduleLesson, BellPeriod] | None]:
    moment = now.astimezone(KYIV) if now else datetime.now(KYIV)
    if schedule is None:
        return None, None
    periods = {item.number: item for item in bells.profiles.get(bells.active_profile, [])}
    timed: list[tuple[datetime, datetime, ScheduleLesson, BellPeriod]] = []
    for lesson in schedule.lessons:
        if lesson.day != moment.date() or lesson.number not in periods:
            continue
        period = periods[lesson.number]
        start_clock, end_clock = parse_clock(period.start), parse_clock(period.end)
        if start_clock is None or end_clock is None:
            continue
        start = datetime.combine(lesson.day, start_clock, KYIV)
        end = datetime.combine(lesson.day, end_clock, KYIV)
        if end <= start:
            continue
        timed.append((start, end, lesson, period))
    timed.sort(key=lambda item: item[0])
    current = next(((lesson, period) for start, end, lesson, period in timed if start <= moment < end), None)
    following = next(((lesson, period) for start, _end, lesson, period in timed if start > moment), None)
    return current, following
