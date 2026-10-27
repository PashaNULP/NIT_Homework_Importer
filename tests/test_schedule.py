from __future__ import annotations

import os
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from unittest.mock import Mock

from nit_homework_v2 import NitClient

from nit_schedule import (
    KYIV,
    BellPeriod,
    BellSettings,
    ScheduleStorage,
    WeekSchedule,
    break_minutes,
    current_and_next,
    monday_of,
    normalize_teacher_diary,
    normalize_academic_year_bells,
    profile_storage_key,
)


def payload(days):
    return {"data": days, "total": len(days)}


class ScheduleNormalizationTests(unittest.TestCase):
    def test_teacher_diary_uses_confirmed_english_date_format(self):
        client = NitClient("synthetic", "school", "user", "test")
        response = object()
        client.session.get = Mock(return_value=response)
        client._handle = Mock(return_value={"data": [], "total": 0})
        client.get_teacher_diary(date(2026, 11, 2))
        url = client.session.get.call_args.args[0]
        self.assertIn("date=Mon%20Nov%2002%202026", url)

    def test_academic_years_uses_confirmed_bells_endpoint(self):
        client = NitClient("synthetic", "school", "user", "test")
        response = object()
        client.session.get = Mock(return_value=response)
        client._handle = Mock(return_value={"data": [], "total": 0})
        client.get_academic_years()
        self.assertTrue(client.session.get.call_args.args[0].endswith("/academic-years?take=10&skip=0"))

    def test_regular_week_and_missing_room(self):
        result = normalize_teacher_diary(
            payload([
                {"date": "2026-11-02T00:00:00Z", "dayOfWeek": 1, "subjects": [
                    {"number": 1, "subject": "Physics", "className": "8-A", "isFds": False},
                    {"number": 3, "subject": "Science", "className": "7-A", "isFds": True},
                ]},
            ]),
            date(2026, 11, 2),
        )
        self.assertEqual([item.number for item in result.lessons], [1, 3])
        self.assertIsNone(result.lessons[0].room)

    def test_multiple_lessons_of_same_class_are_preserved(self):
        result = normalize_teacher_diary(
            payload([{"date": "2026-11-03T00:00:00Z", "subjects": [
                {"number": 2, "subject": "Physics", "className": "8-A"},
                {"number": 5, "subject": "Physics", "className": "8-A"},
            ]}]),
            date(2026, 11, 2),
        )
        self.assertEqual(len(result.lessons), 2)

    def test_holiday_week_is_valid_and_empty(self):
        days = [
            {"date": f"2026-10-{26 + offset:02d}T00:00:00Z", "dayOfWeek": offset + 1, "subjects": []}
            for offset in range(6)
        ]
        self.assertEqual(normalize_teacher_diary(payload(days), date(2026, 10, 26)).lessons, [])

    def test_adjacent_weeks_remain_distinct(self):
        first = normalize_teacher_diary(payload([]), date(2026, 10, 26))
        second = normalize_teacher_diary(payload([]), date(2026, 11, 2))
        self.assertNotEqual(first.week_start, second.week_start)

    def test_monday_normalization_from_weekend(self):
        self.assertEqual(monday_of(date(2026, 11, 8)), date(2026, 11, 2))


class ScheduleStorageTests(unittest.TestCase):
    def test_official_bells_are_read_from_active_academic_year(self):
        result = normalize_academic_year_bells(
            {"data": [
                {"yearStart": 2025, "yearEnd": 2026, "lessons": []},
                {"yearStart": 2026, "yearEnd": 2027, "lessons": [
                    {"number": 0, "fromTime": "08:10:00", "toTime": "08:55:00", "type": 0},
                    {"number": 1, "fromTime": "09:00:00", "toTime": "09:45:00", "type": 0},
                ]},
            ]},
            {"yearStart": 2026, "yearEnd": 2027},
        )
        self.assertEqual([(item.number, item.start, item.end) for item in result.periods], [
            (0, "08:10", "08:55"), (1, "09:00", "09:45")
        ])

    def test_cache_round_trip_and_permissions(self):
        with tempfile.TemporaryDirectory() as directory:
            storage = ScheduleStorage(Path(directory), "profile-a")
            schedule = normalize_teacher_diary(
                payload([{"date": "2026-11-02T00:00:00Z", "subjects": [
                    {"number": 1, "subject": "Physics", "className": "8-A"}
                ]}]),
                date(2026, 11, 2),
            )
            storage.save_week(schedule)
            loaded = storage.load_week(date(2026, 11, 2))
            self.assertIsNotNone(loaded)
            self.assertEqual(loaded.source, "cache")
            path = next(Path(directory).glob("week_*.json"))
            self.assertEqual(os.stat(path).st_mode & 0o777, 0o600)

    def test_profiles_are_isolated_without_exposing_identity(self):
        first = profile_storage_key("one", "school", "user-a")
        second = profile_storage_key("two", "school", "user-b")
        self.assertNotEqual(first, second)
        self.assertNotIn("user-a", first)

    def test_bells_profiles_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            storage = ScheduleStorage(Path(directory), "profile")
            settings = BellSettings("Short", {"Regular": [], "Short": [BellPeriod(1, "09:00", "09:30")]})
            storage.save_bells(settings)
            loaded = storage.load_bells()
            self.assertEqual((loaded.active_profile, loaded.profiles["Short"][0].end), ("Short", "09:30"))

    def test_official_bells_cache_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            storage = ScheduleStorage(Path(directory), "profile")
            bells = normalize_academic_year_bells(
                {"data": [{"yearStart": 2026, "yearEnd": 2027, "lessons": [
                    {"number": 1, "fromTime": "09:00:00", "toTime": "09:45:00"}
                ]}]},
                {"yearStart": 2026, "yearEnd": 2027},
            )
            storage.save_official_bells(bells)
            loaded = storage.load_official_bells()
            self.assertEqual((loaded.source, loaded.periods[0].start), ("cache", "09:00"))


class CurrentLessonTests(unittest.TestCase):
    def setUp(self):
        self.schedule = normalize_teacher_diary(
            payload([{"date": "2026-11-02T00:00:00Z", "subjects": [
                {"number": 1, "subject": "Physics", "className": "8-A"},
                {"number": 2, "subject": "Science", "className": "7-A"},
            ]}]),
            date(2026, 11, 2),
        )

    def test_no_api_times_means_no_invented_current_lesson(self):
        settings = BellSettings("Regular", {"Regular": []})
        self.assertEqual(current_and_next(self.schedule, settings, datetime(2026, 11, 2, 9, 0, tzinfo=KYIV)), (None, None))

    def test_current_and_next_with_local_bells(self):
        settings = BellSettings("Regular", {"Regular": [
            BellPeriod(1, "08:30", "09:15"), BellPeriod(2, "09:25", "10:10")
        ]})
        current, following = current_and_next(
            self.schedule, settings, datetime(2026, 11, 2, 8, 45, tzinfo=KYIV)
        )
        self.assertEqual(current[0].number, 1)
        self.assertEqual(following[0].number, 2)
        self.assertEqual(break_minutes(settings.profiles["Regular"][0], settings.profiles["Regular"][1]), 10)

    def test_shortened_profile_changes_time(self):
        settings = BellSettings("Short", {
            "Regular": [BellPeriod(1, "08:30", "09:15")],
            "Short": [BellPeriod(1, "08:30", "09:00")],
        })
        current, _ = current_and_next(
            self.schedule, settings, datetime(2026, 11, 2, 9, 5, tzinfo=KYIV)
        )
        self.assertIsNone(current)


if __name__ == "__main__":
    unittest.main()
