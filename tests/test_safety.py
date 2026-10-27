from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from nit_core import (
    AbsenceReminder,
    BatchOperationError,
    ImportPreview,
    apply_absence_reminders,
    apply_import_preview,
)
from nit_homework_v2 import ExcelRow, PlannedChange


def calendar_lesson(column: str, day: str, lesson: int) -> dict:
    return {"journalDayId": column, "day": day, "lesson": lesson, "index": lesson}


def plan(plan_id: str, journal: str, day: str, lesson: int, topic: str, homework: str) -> dict:
    return {
        "id": plan_id,
        "journalId": journal,
        "day": day,
        "lesson": lesson,
        "lessonContent": topic,
        "homework": homework,
        "homeworkAttachments": [{"id": "attachment-1"}],
    }


class ImportClient:
    def __init__(self) -> None:
        self.calendar = [
            calendar_lesson("c1", "2026-09-03", 3),
            calendar_lesson("c2", "2026-09-07", 7),
        ]
        self.plans = [plan("p1", "j1", "2026-09-03", 3, "old", "old hw")]
        self.writes = 0
        self.fail_on_write: int | None = None

    def get_journal_lessons(self, _journal_id: str):
        return [dict(item) for item in self.calendar]

    def get_training_plans(self, _journal_id: str):
        return [dict(item) for item in self.plans]

    def update_training_plan(self, lesson, topic, homework):
        self.writes += 1
        if self.fail_on_write == self.writes:
            raise OSError("network")
        current = next(item for item in self.plans if item["id"] == lesson["id"])
        current["lessonContent"] = topic
        current["homework"] = homework
        self.last_attachment = lesson["homeworkAttachments"][0]["id"]

    def create_training_plan(self, lesson, journal_id, topic, homework):
        self.writes += 1
        if self.fail_on_write == self.writes:
            raise OSError("network")
        new = plan("p2", journal_id, lesson["day"], lesson["lesson"], topic, homework)
        self.plans.append(new)
        return dict(new)


class ImportSafetyTests(unittest.TestCase):
    def preview(self, client: ImportClient, two: bool = False) -> ImportPreview:
        first = dict(client.calendar[0])
        first.update(client.plans[0])
        first["journalDayId"] = "c1"
        changes = [
            PlannedChange(first, ExcelRow(2, "new", "new hw", date(2026, 9, 3), 3), "new", "new hw", True, True)
        ]
        lessons = [first, dict(client.calendar[1])]
        if two:
            empty = lessons[1]
            empty["journalId"] = "j1"
            changes.append(
                PlannedChange(empty, ExcelRow(3, "topic 2", "hw 2", date(2026, 9, 7), 7), "topic 2", "hw 2", True, True)
            )
        return ImportPreview({"journalId": "j1", "className": "8-А", "subject": "Фізика"}, 1, "x", True, len(changes), lessons, changes)

    def test_conflict_stops_stale_preview_without_write(self):
        client = ImportClient()
        preview = self.preview(client)
        client.plans[0]["lessonContent"] = "edited elsewhere"
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(BatchOperationError) as error:
                apply_import_preview(client, {}, preview, Path(directory))
        self.assertEqual(client.writes, 0)
        self.assertIn("Конфлікт", str(error.exception))

    def test_update_preserves_attachments_and_verifies(self):
        client = ImportClient()
        with tempfile.TemporaryDirectory() as directory:
            result = apply_import_preview(client, {}, self.preview(client), Path(directory))
        self.assertEqual((result.updated, result.verified), (1, 1))
        self.assertEqual(client.last_attachment, "attachment-1")

    def test_partial_failure_reports_completed_count(self):
        client = ImportClient()
        client.fail_on_write = 2
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(BatchOperationError) as error:
                apply_import_preview(client, {}, self.preview(client, two=True), Path(directory))
        self.assertEqual(error.exception.completed, 1)

    def test_backup_failure_prevents_every_api_write(self):
        client = ImportClient()
        with tempfile.TemporaryDirectory() as directory:
            with patch("nit_core.make_backup", side_effect=OSError("disk full")):
                with self.assertRaises(OSError):
                    apply_import_preview(client, {}, self.preview(client), Path(directory))
        self.assertEqual(client.writes, 0)


class AbsenceClient:
    def __init__(self, network_after_commit: bool = False) -> None:
        self.mark = None
        self.posts = 0
        self.network_after_commit = network_after_commit

    def get_journal_pages(self, _journal_id):
        marks = [self.mark] if self.mark else []
        reminders = [] if self.mark else [{"journalDayId": "c1", "userId": "u1"}]
        return [{"marks": marks, "absenceReminders": reminders}]

    def create_absence_mark(self, _reminder):
        self.posts += 1
        self.mark = {"columnId": "c1", "userId": "u1", "absense": 0, "rating": None,
                     "customMark": None, "leveledMark": None, "letterMark": None,
                     "isVerbalFormed": None}
        if self.network_after_commit:
            raise OSError("connection lost")


class AbsenceSafetyTests(unittest.TestCase):
    def reminder(self):
        return AbsenceReminder(1, {"journalId": "j1", "className": "8-А", "subject": "Фізика"},
                               {"id": "c1", "date": "2026-09-03", "lesson": 3}, "u1")

    def test_repeated_run_does_not_duplicate(self):
        client = AbsenceClient()
        with tempfile.TemporaryDirectory() as directory:
            first = apply_absence_reminders(client, [self.reminder()], Path(directory))
            second = apply_absence_reminders(client, [self.reminder()], Path(directory))
        self.assertEqual((first.updated, second.skipped, client.posts), (1, 1, 1))

    def test_network_error_after_commit_is_verified_without_retry(self):
        client = AbsenceClient(network_after_commit=True)
        with tempfile.TemporaryDirectory() as directory:
            result = apply_absence_reminders(client, [self.reminder()], Path(directory))
        self.assertEqual((result.updated, client.posts), (1, 1))

    def test_grade_is_never_overwritten(self):
        client = AbsenceClient()
        client.mark = {"columnId": "c1", "userId": "u1", "absense": None, "rating": 10}
        with tempfile.TemporaryDirectory() as directory:
            result = apply_absence_reminders(client, [self.reminder()], Path(directory))
        self.assertEqual((result.updated, result.skipped, client.posts), (0, 1, 0))

    def test_incomplete_identifiers_are_skipped(self):
        client = AbsenceClient()
        incomplete = AbsenceReminder(
            1,
            {"journalId": "j1", "className": "8-А", "subject": "Фізика"},
            {"id": "", "date": "2026-09-03", "lesson": 3},
            "u1",
        )
        with tempfile.TemporaryDirectory() as directory:
            result = apply_absence_reminders(client, [incomplete], Path(directory))
        self.assertEqual((result.updated, result.skipped, client.posts), (0, 1, 0))


if __name__ == "__main__":
    unittest.main()
