from __future__ import annotations

import json
import os
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import nit_homework_v2
from nit_auth import FunctionSessionStore, MemorySessionStore, NitUserConfig
from nit_backup import verify_backup, write_secure_backup
from nit_core import (
    AuthContext,
    connect_with_password,
    connect_with_saved_session,
    discover_user_config,
    refresh_auth_context,
)
from nit_homework_v2 import ExcelRow, NitError, load_excel_rows, match_rows_to_lessons, valid_or_refreshed_session


class BackupTests(unittest.TestCase):
    def test_unique_atomic_backup_and_integrity(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            first = write_secure_backup({"value": 1}, root, "test")
            second = write_secure_backup({"value": 1}, root, "test")
            self.assertNotEqual(first, second)
            self.assertEqual(verify_backup(first)["value"], 1)
            self.assertEqual(os.stat(first).st_mode & 0o777, 0o600)

    def test_tampering_is_detected(self):
        with tempfile.TemporaryDirectory() as directory:
            path = write_secure_backup({"value": 1}, Path(directory), "test")
            data = json.loads(path.read_text())
            data["value"] = 2
            path.write_text(json.dumps(data))
            with self.assertRaises(ValueError):
                verify_backup(path)


class AuthTests(unittest.TestCase):
    def test_memory_session_stores_are_isolated_and_return_copies(self):
        first = MemorySessionStore()
        second = MemorySessionStore()
        source = {"access_token": "first", "nested": {"value": 1}}
        first.save(source)
        source["nested"]["value"] = 2
        loaded = first.load()
        self.assertEqual(loaded, {"access_token": "first", "nested": {"value": 1}})
        loaded["nested"]["value"] = 3
        self.assertEqual(first.load()["nested"]["value"], 1)
        self.assertIsNone(second.load())
        self.assertTrue(first.delete())
        self.assertFalse(first.delete())

    def test_user_config_is_discovered_for_configured_school(self):
        response = {
            "userId": "teacher-2",
            "permissions": [
                {
                    "active": True,
                    "forUserId": "teacher-2",
                    "schoolId": "school-a",
                },
                {
                    "active": True,
                    "forUserId": "someone-else",
                    "schoolId": "school-b",
                },
            ],
        }
        with patch("nit_core.get_my_permissions", return_value=response):
            config = discover_user_config({"access_token": "token"}, "school-a")
        self.assertEqual((config.school_id, config.user_id), ("school-a", "teacher-2"))

    def test_user_without_configured_school_is_rejected(self):
        response = {
            "userId": "teacher-2",
            "permissions": [
                {
                    "active": True,
                    "forUserId": "teacher-2",
                    "schoolId": "another-school",
                }
            ],
        }
        with (
            patch("nit_core.get_my_permissions", return_value=response),
            self.assertRaises(NitError),
        ):
            discover_user_config({"access_token": "token"}, "school-a")

    def test_valid_session_is_not_refreshed(self):
        session = {"access_token": "token", "refresh_token": "refresh", "expires_at": 99999999999}
        with patch("nit_homework_v2.refresh_oauth_session") as refresh:
            self.assertIs(valid_or_refreshed_session(session), session)
            refresh.assert_not_called()

    def test_expired_session_refresh_error_returns_none(self):
        with patch("nit_homework_v2.refresh_oauth_session", side_effect=NitError("failed")):
            self.assertIsNone(valid_or_refreshed_session({"refresh_token": "r", "expires_at": 0}))

    def test_store_and_profile_are_isolated_objects(self):
        saved = {}
        store = FunctionSessionStore(lambda: saved.get("a"), lambda value: saved.__setitem__("a", value), lambda: True)
        store.save({"access_token": "secret"})
        self.assertEqual(store.load()["access_token"], "secret")
        with patch.dict(os.environ, {"NIT_PROFILE": "teacher-2", "NIT_SCHOOL_ID": "2", "NIT_USER_ID": "u2"}):
            config = NitUserConfig.from_environment("1", "u1")
        self.assertEqual((config.profile, config.school_id, config.user_id), ("teacher-2", "2", "u2"))

    def test_keychain_manifest_uses_profile_account_not_last_chunk_account(self):
        writes: list[tuple[str, str]] = []
        with (
            patch.object(nit_homework_v2.sys, "platform", "darwin"),
            patch("nit_homework_v2._keychain_read_value", return_value=None),
            patch("nit_homework_v2._keychain_write_value", side_effect=lambda account, value: writes.append((account, value))),
            patch("nit_homework_v2.secrets.token_hex", return_value="abc123abc123"),
        ):
            nit_homework_v2.save_keychain_session(
                {"refresh_token": "refresh-token"},
                account="nit.school.teacher-2",
            )
        self.assertEqual(writes[-1][0], "nit.school.teacher-2")
        self.assertTrue(writes[-1][1].startswith("NIT2:R:"))

    def test_missing_session_store_returns_none(self):
        store = FunctionSessionStore(lambda: None, lambda _value: None, lambda: False)
        self.assertIsNone(connect_with_saved_session(session_store=store))

    def test_keychain_failure_keeps_successful_login_in_memory(self):
        def fail_save(_value):
            raise OSError("Keychain unavailable")

        store = FunctionSessionStore(lambda: None, fail_save, lambda: False)
        base_context = AuthContext(client=object(), active_year={}, semester_schedules={})
        with (
            patch("nit_core.login_session_with_password", return_value={"access_token": "a", "refresh_token": "r"}),
            patch("nit_core._context_from_session", return_value=base_context),
        ):
            context = connect_with_password("teacher@example.invalid", "synthetic-password", session_store=store)
        self.assertFalse(context.session_saved)
        self.assertIn("Keychain unavailable", context.session_warning or "")

    def test_open_context_refreshes_token_and_updates_http_client(self):
        saved: list[dict] = []

        class Client:
            token = "old"

            def set_access_token(self, token):
                self.token = token

        store = FunctionSessionStore(lambda: None, lambda value: saved.append(value), lambda: False)
        context = AuthContext(
            client=Client(),
            active_year={},
            semester_schedules={},
            user_config=NitUserConfig("97", "u1"),
            session={"access_token": "old", "refresh_token": "refresh", "expires_at": 0},
            session_store=store,
        )
        refreshed = {"access_token": "new", "refresh_token": "refresh-2", "expires_at": 99999999999}
        with patch("nit_homework_v2.refresh_oauth_session", return_value=refreshed):
            self.assertTrue(refresh_auth_context(context))
        self.assertEqual(context.client.token, "new")
        self.assertEqual(saved, [refreshed])


class ExcelTests(unittest.TestCase):
    def test_two_lessons_same_day_match_by_number(self):
        lessons = [
            {"id": "1", "day": "2026-09-03", "lesson": 1},
            {"id": "2", "day": "2026-09-03", "lesson": 2},
        ]
        rows = [ExcelRow(2, "a", "", date(2026, 9, 3), 2)]
        self.assertEqual(match_rows_to_lessons(rows, lessons, True)[0].lesson["id"], "2")

    def test_empty_cells_do_not_remove_row_with_date(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "x.xlsx"
            book = Workbook(); sheet = book.active
            sheet.append(["Дата", "Урок", "Тема уроку", "Домашнє завдання"])
            sheet.append(["03.09.2026", 3, "", ""]); book.save(path)
            _, rows, has_date = load_excel_rows(path)
        self.assertTrue(has_date)
        self.assertEqual((len(rows), rows[0].day), (1, date(2026, 9, 3)))

    def test_invalid_date_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "x.xlsx"
            book = Workbook(); sheet = book.active
            sheet.append(["Дата", "Тема уроку"]); sheet.append(["not-a-date", "topic"]); book.save(path)
            with self.assertRaises(NitError):
                load_excel_rows(path)


class RegressionTests(unittest.TestCase):
    def test_modules_import(self):
        __import__("nit_core"); __import__("nit_web"); __import__("nit_homework_v2"); __import__("nit_schedule")

    def test_web_health_payload(self):
        import nit_web

        self.assertEqual(nit_web.health(), {"status": "ok", "app": "NIT2"})

    def test_web_tls_options_require_both_files(self):
        import nit_web

        with patch.dict(os.environ, {"NIT_SSL_CERTFILE": "cert.pem"}, clear=True):
            with self.assertRaisesRegex(SystemExit, "NIT_SSL_CERTFILE і NIT_SSL_KEYFILE"):
                nit_web._tls_run_options()

    def test_web_tls_options_accept_existing_files(self):
        import nit_web

        with tempfile.TemporaryDirectory() as directory:
            cert = Path(directory) / "cert.pem"
            key = Path(directory) / "key.pem"
            cert.touch()
            key.touch()
            with patch.dict(
                os.environ,
                {"NIT_SSL_CERTFILE": str(cert), "NIT_SSL_KEYFILE": str(key)},
                clear=True,
            ):
                options, enabled = nit_web._tls_run_options()
        self.assertTrue(enabled)
        self.assertEqual(
            options,
            {"ssl_certfile": str(cert.resolve()), "ssl_keyfile": str(key.resolve())},
        )


if __name__ == "__main__":
    unittest.main()
