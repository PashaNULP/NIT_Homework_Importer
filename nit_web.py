#!/usr/bin/env python3
"""Локальний вебінтерфейс NIT² by FIZYK."""

from __future__ import annotations

import calendar
import json
import os
import secrets
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

try:
    from nicegui import app, run, ui
except ImportError:
    print(
        "Не встановлено NiceGUI. Виконайте:\n"
        "python3 -m pip install -r requirements_web.txt"
    )
    raise SystemExit(2)

from nit_core import (
    AbsenceReminder,
    AuthContext,
    ExcelRow,
    ImportPreview,
    JournalMatrix,
    JournalLessonReport,
    NitError,
    apply_absence_reminders,
    apply_import_preview,
    build_all_teachers_schedule,
    connect_with_saved_session,
    connect_user_with_password,
    load_calendar_reports,
    load_journal_report,
    load_journal_matrix,
    load_absence_reminders,
    prepare_import_preview,
    read_excel_upload,
    refresh_auth_context,
    save_import_template,
    save_lesson_reports,
)
from nit_auth import MemorySessionStore, NitUserConfig
from nit_homework_v2 import DEFAULT_SCHOOL_ID, DEFAULT_USER_ID, compact, parse_date_value
from nit_schedule import (
    KYIV,
    BellPeriod,
    BellSettings,
    OfficialBells,
    ScheduleStorage,
    WeekSchedule,
    break_minutes,
    current_and_next,
    monday_of,
    normalize_teacher_diary,
    normalize_academic_year_bells,
    parse_clock,
    profile_storage_key,
)


UK_MONTHS = (
    "",
    "Січень",
    "Лютий",
    "Березень",
    "Квітень",
    "Травень",
    "Червень",
    "Липень",
    "Серпень",
    "Вересень",
    "Жовтень",
    "Листопад",
    "Грудень",
)
WEEKDAYS = ("Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Нд")


_WEB_SESSION_STORES: dict[str, MemorySessionStore] = {}
_WEB_SESSION_CONFIGS: dict[str, NitUserConfig] = {}
_STORAGE_SECRET = os.getenv("NIT2_STORAGE_SECRET") or secrets.token_urlsafe(32)


def lesson_day(lesson: dict[str, Any]) -> date | None:
    try:
        return parse_date_value(lesson.get("day") or lesson.get("date"))
    except (TypeError, ValueError):
        return None


class NitWebApp:
    def __init__(self) -> None:
        self.client = ui.context.client
        self.server_config = NitUserConfig.from_environment(DEFAULT_SCHOOL_ID, DEFAULT_USER_ID)
        user_storage = app.storage.user
        self.session_handle = str(user_storage.get("nit2_session_handle") or "")
        if not self.session_handle:
            self.session_handle = secrets.token_urlsafe(24)
            user_storage["nit2_session_handle"] = self.session_handle
        self.session_store = _WEB_SESSION_STORES.setdefault(
            self.session_handle, MemorySessionStore()
        )
        self.auth: AuthContext | None = None
        self.semester: int | None = None
        self.journals: list[dict[str, Any]] = []
        self.selected_journal_id: str | None = None
        self.report: JournalLessonReport | None = None
        self.calendar_reports: list[JournalLessonReport] = []
        self.calendar_loaded = False
        self.calendar_loading = False
        self.calendar_load_attempted = False
        self.journal_matrix: JournalMatrix | None = None
        self.journal_loading = False
        self.journal_load_attempted = False
        self.calendar_month = date.today().replace(day=1)
        self.selected_day: date | None = None
        self.busy = False
        self.excel_name: str | None = None
        self.excel_sheet: str | None = None
        self.excel_rows: list[ExcelRow] = []
        self.excel_has_date = False
        self.import_mode = "homework"
        self.import_allow_clear = False
        self.import_preview: ImportPreview | None = None
        self.import_preview_token: str | None = None
        self.import_in_progress = False
        self.absence_reminders: list[AbsenceReminder] = []
        self.absence_loaded = False
        self.absence_loading = False
        self.absence_load_attempted = False
        self.absence_class: str | None = None
        self.absence_action_token: str | None = None
        self.absence_in_progress = False
        self.absence_row_map: dict[str, AbsenceReminder] = {}

        self.status_label = None
        self.dark_mode = None
        self.theme_button = None
        self.theme_tooltip = None
        self.logout_button = None
        self.header_controls = None
        self.header_tabs = None
        self.content = None
        self.calendar_box = None
        self.journal_box = None
        self.import_box = None
        self.absence_box = None
        self.absence_table = None
        self.overview_absence_title = None
        self.overview_absence_meta = None
        self.today_date_label = None
        self.today_focus_title = None
        self.today_focus_meta = None
        self.today_countdown_label = None
        self.today_total_label = None
        self.schedule_box = None
        self.schedule_week = monday_of(datetime.now(KYIV).date())
        self.schedule_data: WeekSchedule | None = None
        self.schedule_storage: ScheduleStorage | None = None
        self.bell_settings = BellSettings("Звичайний", {"Звичайний": [], "Скорочений": []})
        self.official_bells: OfficialBells | None = None
        self.bell_source = "local"
        self.schedule_loading = False
        self.schedule_error: str | None = None
        self.schedule_view = "lessons"
        self.offline_schedule_mode = False
        self.main_tabs = None
        self.schedule_tab = None
        self.journal_tab = None
        self.import_tab = None
        self.absence_tab = None
        self.session_refresh_failed = False
        self.session_timer = None
        self._build_shell()

    def _build_shell(self) -> None:
        ui.colors(primary="#4f46e5", secondary="#64748b", accent="#06b6d4")
        self.dark_mode = ui.dark_mode(True)
        ui.add_head_html(
            """
            <style>
              :root { color-scheme: light dark; }
              body {
                background:
                  radial-gradient(circle at 8% 8%, rgba(79,70,229,.10), transparent 28rem),
                  radial-gradient(circle at 92% 15%, rgba(6,182,212,.10), transparent 25rem),
                  #f8fafc;
                color: #172033;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
              }
              .nit-header {
                background: linear-gradient(115deg, #111827 0%, #25205f 55%, #164e63 100%);
                color: white;
                border-bottom: 1px solid rgba(255,255,255,.13);
                box-shadow: 0 12px 32px rgba(15,23,42,.18);
              }
              .brand-orbit {
                width: 46px; height: 46px; border-radius: 15px;
                display: flex; align-items: center; justify-content: center;
                background: linear-gradient(145deg, #818cf8, #22d3ee);
                box-shadow: 0 8px 22px rgba(34,211,238,.28), inset 0 1px 0 rgba(255,255,255,.45);
              }
              .brand-logo { width: 40px; height: 40px; object-fit: contain; }
              .brand-name { letter-spacing: -.04em; }
              .status-pill {
                padding: 8px 13px; border-radius: 999px;
                background: rgba(255,255,255,.10);
                border: 1px solid rgba(255,255,255,.16);
                backdrop-filter: blur(12px);
              }
              .header-controls {
                min-width: 0; max-width: 920px;
                padding: 4px 8px;
              }
              .header-controls .q-field__control {
                min-height: 42px; height: 42px;
                background: rgba(15,23,42,.34);
              }
              .header-controls .q-field__native,
              .header-controls .q-field__input,
              .header-controls .q-field__label,
              .header-controls .q-field__marginal { color: #e2e8f0 !important; }
              .header-controls .q-field--outlined .q-field__control:before {
                border-color: rgba(203,213,225,.42);
              }
              .header-semesters {
                padding: 3px; border-radius: 13px;
                background: rgba(15,23,42,.40);
                border: 1px solid rgba(203,213,225,.20);
              }
              .header-semesters .q-btn { min-height: 36px; }
              .header-tabs-wrap {
                border-top: 1px solid rgba(255,255,255,.10);
              }
              .header-tabs {
                width: min(920px, 100%); padding: 4px;
                border-radius: 14px;
                background: rgba(15,23,42,.28);
              }
              .header-tabs .q-tab {
                flex: 1 1 0 !important; min-width: 0; max-width: none;
                border-radius: 11px; min-height: 46px;
                color: #cbd5e1;
              }
              .header-tabs .q-tab__content { min-width: 0; }
              .header-tabs .q-tab__icon { font-size: 22px; }
              .header-tabs .q-tab__label {
                font-size: 13px; font-weight: 600; letter-spacing: .01em;
              }
              .header-tabs .q-tab--active {
                color: white; background: rgba(79,70,229,.72);
              }
              .nit-card {
                border: 1px solid rgba(148,163,184,.24);
                border-radius: 20px;
                background: rgba(255,255,255,.92);
                box-shadow: 0 12px 34px rgba(30,41,59,.075);
                backdrop-filter: blur(12px);
              }
              .metric { min-width: 170px; }
              .overview-dashboard-grid {
                display: grid; grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 12px; width: 100%;
              }
              .overview-panel {
                display: flex; flex-direction: column; gap: 12px !important;
              }
              .overview-widget { min-height: 170px; }
              .overview-status-row {
                display: flex; align-items: center; gap: 10px;
                width: 100%;
                min-height: 42px; padding: 8px 10px; border-radius: 12px;
                background: #f8fafc; border: 1px solid #e2e8f0;
              }
              .overview-status-icon {
                width: 34px; height: 34px; flex: 0 0 34px;
                display: flex; align-items: center; justify-content: center;
                border-radius: 10px; background: #eef2ff; color: #4f46e5;
              }
              .lesson-countdown {
                margin-left: auto; padding: 7px 11px; border-radius: 10px;
                color: #3730a3; background: #e0e7ff;
                border: 1px solid #c7d2fe;
                font-size: 13px; font-weight: 750;
                font-variant-numeric: tabular-nums; white-space: nowrap;
              }
              .journal-summary-bar {
                display: grid; grid-template-columns: minmax(360px, 720px) auto;
                align-items: center; justify-content: space-between;
                gap: 18px; width: 100%;
              }
              .journal-summary-actions { flex-wrap: nowrap; }
              .calendar-cell {
                min-height: 125px; cursor: pointer; transition: .18s ease;
                border-radius: 16px; box-shadow: none;
              }
              .calendar-cell:hover {
                transform: translateY(-2px); border-color: #818cf8;
                box-shadow: 0 8px 20px rgba(79,70,229,.13);
              }
              .calendar-selected { outline: 2px solid #4f46e5; background: #eef2ff; }
              .month-grid { min-width: 900px; table-layout: fixed; }
              .month-grid th, .month-grid td { min-width: 125px; }
              .month-grid th { text-align: center; padding: 9px; }
              .month-day {
                height: 108px; padding: 8px !important; cursor: pointer;
                transition: background .15s ease, box-shadow .15s ease;
              }
              .month-day:hover { background: #f8fafc !important; }
              .month-date {
                width: 28px; height: 28px; display: inline-flex;
                align-items: center; justify-content: center;
                border-radius: 9px; font-weight: 700; color: #334155;
              }
              .month-class-list { display: flex; flex-wrap: wrap; gap: 4px; margin-top: 5px; }
              .month-class-chip {
                display: inline-flex; align-items: center; gap: 4px;
                padding: 3px 7px; border-radius: 8px;
                color: #3730a3; background: #eef2ff;
                border: 1px solid #c7d2fe; font-size: 12px; font-weight: 600;
                line-height: 1.25;
              }
              .month-selected { background: #e0e7ff !important; box-shadow: inset 0 0 0 2px #6366f1; }
              .month-today .month-date { color: white; background: #4f46e5; }
              .month-outside { background: #f8fafc !important; cursor: default; }
              .bells-grid {
                display: grid; grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 10px; width: 100%;
              }
              @media (max-width: 700px) { .bells-grid { grid-template-columns: 1fr; } }
              .bell-card {
                display: flex; align-items: center; gap: 12px;
                min-height: 78px; padding: 12px 14px; border-radius: 14px;
                color: #172033; background: #f8fafc; border: 1px solid #e2e8f0;
              }
              .bell-number {
                width: 46px; height: 46px; flex: 0 0 46px;
                display: flex; align-items: center; justify-content: center;
                border-radius: 13px; color: white; background: #4f46e5;
                font-size: 18px; font-weight: 800;
              }
              .bell-time { font-size: 18px; font-weight: 750; letter-spacing: .01em; }
              .bell-meta { font-size: 12px; color: #64748b; }
              .bell-break {
                margin-left: auto; padding: 4px 8px; border-radius: 8px;
                color: #3730a3; background: #e0e7ff; font-size: 12px; font-weight: 700;
                white-space: nowrap;
              }
              .q-btn { border-radius: 12px; font-weight: 600; }
              .q-field--outlined .q-field__control { border-radius: 12px; }
              .q-table thead tr { background: #f8fafc; }
              .q-table th { color: #475569; font-weight: 700; }
              .nit-data-table {
                overflow: hidden; border-radius: 14px;
                border: 1px solid #dbe3ee !important;
                box-shadow: none;
              }
              .nit-data-table .q-table {
                font-family: inherit; font-size: 14px;
                font-variant-numeric: tabular-nums;
              }
              .nit-data-table .q-table thead tr { height: 46px; }
              .nit-data-table .q-table th {
                padding: 9px 12px;
                color: #475569; background: #f1f5f9;
                border-color: #dbe3ee;
                font-size: 12px; font-weight: 750;
                letter-spacing: .025em; line-height: 1.25;
                white-space: nowrap;
              }
              .nit-data-table .q-table tbody tr { transition: background .14s ease; }
              .nit-data-table .q-table tbody td {
                height: 50px; padding: 10px 12px;
                color: #334155; background: #ffffff;
                border-color: #e2e8f0;
                line-height: 1.45; vertical-align: middle;
              }
              .nit-data-table .q-table tbody tr:nth-child(even) td {
                background: #f8fafc;
              }
              .nit-data-table .q-table tbody tr:hover td {
                background: #eef2ff;
              }
              .nit-data-table .q-table__bottom {
                min-height: 48px; padding: 7px 12px;
                color: #64748b; background: #f8fafc;
                border-color: #e2e8f0;
                font-size: 13px;
              }
              .nit-data-table .q-checkbox__inner { font-size: 36px; }
              .overview-table .q-table__middle { overflow-x: hidden; }
              .overview-table table { width: 100%; table-layout: fixed; }
              .overview-table th:nth-child(1), .overview-table td:nth-child(1) { width: 9%; }
              .overview-table th:nth-child(2), .overview-table td:nth-child(2) { width: 6%; }
              .overview-table th:nth-child(3), .overview-table td:nth-child(3) { width: 28%; }
              .overview-table th:nth-child(4), .overview-table td:nth-child(4) { width: 47%; }
              .overview-table th:nth-child(5), .overview-table td:nth-child(5) { width: 10%; }
              .overview-table td:nth-child(3),
              .overview-table td:nth-child(4) {
                white-space: normal; overflow-wrap: anywhere; word-break: break-word;
              }
              .calendar-detail-table { max-width: 1320px; margin: 0 auto; }
              .calendar-detail-table .q-table__middle { overflow-x: auto; }
              .calendar-detail-table table {
                width: 100%; min-width: 960px; table-layout: fixed;
              }
              .calendar-detail-table th:nth-child(1),
              .calendar-detail-table td:nth-child(1) { width: 10%; }
              .calendar-detail-table th:nth-child(2),
              .calendar-detail-table td:nth-child(2) { width: 14%; }
              .calendar-detail-table th:nth-child(3),
              .calendar-detail-table td:nth-child(3) { width: 7%; }
              .calendar-detail-table th:nth-child(4),
              .calendar-detail-table td:nth-child(4) { width: 27%; }
              .calendar-detail-table th:nth-child(5),
              .calendar-detail-table td:nth-child(5) { width: 30%; }
              .calendar-detail-table th:nth-child(6),
              .calendar-detail-table td:nth-child(6) { width: 12%; }
              .calendar-detail-table td:nth-child(4),
              .calendar-detail-table td:nth-child(5) {
                white-space: normal; overflow-wrap: anywhere; word-break: break-word;
              }
              .schedule-scroll {
                width: 100%; max-width: 100%; overflow-x: auto;
                border: 1px solid #e2e8f0; border-radius: 16px;
                overscroll-behavior-x: contain; overscroll-behavior-y: auto;
                background: white;
              }
              .schedule-grid {
                width: 100%; min-width: 840px; border-collapse: separate; border-spacing: 0;
                font-family: inherit; font-size: 14px; font-variant-numeric: tabular-nums;
              }
              .schedule-grid th, .schedule-grid td {
                min-width: 145px; padding: 10px 12px; vertical-align: top;
                border-right: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0;
                background: white;
              }
              .schedule-grid th {
                height: 46px; background: #eef2ff; color: #334155; text-align: left;
                font-size: 12px; font-weight: 750; letter-spacing: .025em;
              }
              .schedule-grid .schedule-number {
                min-width: 58px; width: 58px; text-align: center; font-weight: 700;
                background: #f8fafc;
              }
              .schedule-lesson {
                padding: 8px; border-radius: 11px; background: #eef2ff;
                border: 1px solid #c7d2fe;
              }
              .schedule-today { background: #ecfeff !important; box-shadow: inset 0 3px #06b6d4; }
              .journal-scroll {
                width: 100%; max-width: 100%; min-width: 0;
                max-height: 65vh; overflow: auto;
                overscroll-behavior: contain; scrollbar-gutter: stable;
                border: 1px solid #e2e8f0; border-radius: 16px;
                background: white;
              }
              .journal-grid {
                border-collapse: separate; border-spacing: 0;
                min-width: 100%; width: max-content;
                font-family: inherit; font-size: 14px; font-variant-numeric: tabular-nums;
              }
              .journal-grid th, .journal-grid td {
                min-width: 70px; height: 48px; padding: 7px 9px;
                border-right: 1px solid #e2e8f0;
                border-bottom: 1px solid #e2e8f0;
                text-align: center; background: white;
              }
              .journal-grid thead th {
                position: sticky; top: 0; z-index: 3;
                height: 42px;
                color: #334155; background: #eef2ff;
                font-size: 12px; font-weight: 750; letter-spacing: .02em;
              }
              .journal-grid .journal-student {
                position: sticky; left: 0; z-index: 2;
                min-width: 320px; width: 320px;
                text-align: left; overflow: hidden; background: #f8fafc;
                box-shadow: 3px 0 7px rgba(15,23,42,.06);
              }
              .journal-grid thead .journal-student {
                top: 0; left: 0; z-index: 5; background: #e0e7ff;
              }
              .journal-grid .journal-student .row { min-width: 0; flex-wrap: nowrap; }
              .journal-grid tbody tr:nth-child(even) td { background: #f8fafc; }
              .journal-grid tbody tr:nth-child(even) .journal-student { background: #f1f5f9; }
              .journal-grid tbody tr:hover td { background: #f8fafc; }
              .journal-grid tbody tr:hover .journal-student { background: #eef2ff; }
              .journal-mark { font-weight: 700; color: #3730a3; }
              .journal-absence { color: #dc2626; background: #fef2f2 !important; }
              .journal-empty { color: #cbd5e1; }
              body.body--dark {
                background:
                  radial-gradient(circle at 8% 8%, rgba(99,102,241,.16), transparent 28rem),
                  radial-gradient(circle at 92% 15%, rgba(6,182,212,.12), transparent 25rem),
                  #070b14;
                color: #e5e7eb;
              }
              body.body--dark .nit-card {
                background: rgba(15,23,42,.94);
                border-color: rgba(100,116,139,.38);
                box-shadow: 0 14px 38px rgba(0,0,0,.24);
              }
              body.body--dark .overview-status-row {
                background: #111c30; border-color: #334155;
              }
              body.body--dark .overview-status-icon {
                color: #a5b4fc; background: #25275b;
              }
              body.body--dark .lesson-countdown {
                color: #e0e7ff; background: #28276a; border-color: #4f46e5;
              }
              body.body--dark .overview-table .q-table__middle,
              body.body--dark .overview-table tbody td {
                color: #e2e8f0; background: #0f172a; border-color: #334155;
              }
              body.body--dark .overview-table tbody tr:hover td {
                background: #111c30;
              }
              body.body--dark .calendar-selected { background: #26245f; outline-color: #818cf8; }
              body.body--dark .q-table thead tr { background: #172033; }
              body.body--dark .q-table th { color: #cbd5e1; }
              body.body--dark .nit-data-table {
                border-color: #334155 !important;
              }
              body.body--dark .nit-data-table .q-table__middle {
                background: #0f172a;
              }
              body.body--dark .nit-data-table .q-table th {
                color: #dbeafe; background: #1e293b; border-color: #3b4a61;
              }
              body.body--dark .nit-data-table .q-table tbody td {
                color: #e2e8f0; background: #0f172a; border-color: #27364b;
              }
              body.body--dark .nit-data-table .q-table tbody tr:nth-child(even) td {
                background: #111c30;
              }
              body.body--dark .nit-data-table .q-table tbody tr:hover td {
                background: #172554;
              }
              body.body--dark .nit-data-table .q-table__bottom {
                color: #94a3b8; background: #111c30; border-color: #334155;
              }
              body.body--dark .journal-scroll {
                background: #0f172a; border-color: #334155;
              }
              body.body--dark .journal-grid th,
              body.body--dark .journal-grid td {
                color: #e2e8f0; background: #0f172a;
                border-color: #334155;
              }
              body.body--dark .journal-grid thead th { color: #e0e7ff; background: #25275b; }
              body.body--dark .journal-grid .journal-student { background: #111c30; }
              body.body--dark .journal-grid thead .journal-student { background: #312e81; }
              body.body--dark .journal-grid tbody tr:nth-child(even) td { background: #111c30; }
              body.body--dark .journal-grid tbody tr:nth-child(even) .journal-student { background: #142037; }
              body.body--dark .journal-grid tbody tr:hover td { background: #172554; }
              body.body--dark .journal-grid tbody tr:hover .journal-student { background: #1e293b; }
              body.body--dark .journal-mark { color: #a5b4fc; }
              body.body--dark .journal-absence { color: #fca5a5; background: #451a1a !important; }
              body.body--dark .journal-empty { color: #475569; }
              body.body--dark .schedule-scroll { background: #0f172a; border-color: #334155; }
              body.body--dark .schedule-grid th,
              body.body--dark .schedule-grid td { color: #e2e8f0; background: #0f172a; border-color: #334155; }
              body.body--dark .schedule-grid th { background: #25275b; }
              body.body--dark .schedule-grid .schedule-number { background: #111c30; }
              body.body--dark .schedule-lesson { background: #1e1b4b; border-color: #4338ca; }
              body.body--dark .schedule-today { background: #153143 !important; }
              body.body--dark .month-day:hover { background: #172554 !important; }
              body.body--dark .month-date { color: #e2e8f0; }
              body.body--dark .month-class-chip {
                color: #c7d2fe; background: #24245a; border-color: #4f46e5;
              }
              body.body--dark .month-selected { background: #28276a !important; box-shadow: inset 0 0 0 2px #818cf8; }
              body.body--dark .month-today .month-date { color: white; background: #6366f1; }
              body.body--dark .month-outside { background: #0b1220 !important; }
              body.body--dark .bell-card {
                color: #e2e8f0; background: #111c30; border-color: #334155;
              }
              body.body--dark .bell-number { background: #4f46e5; }
              body.body--dark .bell-meta { color: #94a3b8; }
              body.body--dark .bell-break { color: #c7d2fe; background: #28276a; }
              body.body--dark .text-slate-400 { color: #94a3b8 !important; }
              body.body--dark .text-slate-500 { color: #94a3b8 !important; }
              body.body--dark .text-slate-600 { color: #cbd5e1 !important; }
              body.body--dark .text-slate-700 { color: #e2e8f0 !important; }
              body.body--dark .bg-indigo-50 { background: #1e1b4b !important; }
              body.body--dark .text-indigo-900 { color: #e0e7ff !important; }
              body.body--dark .border-indigo-100 { border-color: #3730a3 !important; }

              /* Responsive shell: keep the page itself fixed while data grids scroll. */
              html, body, #q-app { max-width: 100%; overflow-x: hidden; }
              .nicegui-content { width: 100%; max-width: 100%; padding: 0 !important; }
              .nit-content, .q-tab-panels, .q-tab-panel { min-width: 0; max-width: 100%; }
              .schedule-scroll, .journal-scroll, .q-table__middle {
                -webkit-overflow-scrolling: touch;
              }

              @media (max-width: 900px) {
                .nit-header { padding: 8px 10px !important; }
                .header-main {
                  display: grid !important;
                  grid-template-columns: minmax(0, 1fr) auto;
                  gap: 8px 10px !important;
                }
                .header-brand { min-width: 0; }
                .header-controls {
                  grid-column: 1 / -1;
                  grid-row: 2;
                  width: 100%; max-width: none; padding: 0;
                }
                .header-actions { grid-column: 2; grid-row: 1; }
                .header-tabs { width: 100%; }
                .nit-content { padding: 4px 8px 12px !important; }
                .overview-table .q-table__middle { overflow-x: auto; }
                .overview-table table { min-width: 760px; }
                .calendar-detail-table { max-width: 100%; }
                .overview-dashboard-grid { grid-template-columns: 1fr; }
                .journal-summary-bar { grid-template-columns: 1fr; }
                .journal-summary-actions { flex-wrap: wrap; }
              }

              @media (max-width: 600px) {
                .brand-orbit {
                  width: 38px; height: 38px; flex: 0 0 38px; border-radius: 12px;
                }
                .brand-logo { width: 34px; height: 34px; }
                .brand-name { font-size: 20px !important; }
                .brand-tagline { display: none; }
                .header-brand { gap: 10px !important; }
                .header-actions { gap: 2px !important; }
                .status-pill {
                  max-width: 156px; padding: 6px 9px; gap: 5px !important;
                }
                .status-pill .q-icon { display: none; }
                .status-pill .q-label {
                  overflow: hidden; text-overflow: ellipsis; font-size: 11px !important;
                }
                .header-controls { gap: 6px !important; }
                .header-controls .q-field { min-width: 0 !important; }
                .header-controls .q-field__control { min-height: 40px; height: 40px; }
                .header-semesters { flex: 0 0 auto; }
                .header-semesters .q-btn {
                  min-height: 34px; padding: 0 9px; font-size: 12px;
                }
                .header-tabs-wrap { padding-top: 4px !important; }
                .header-tabs { padding: 3px; border-radius: 12px; }
                .header-tabs .q-tab { min-height: 42px; padding: 0 4px; }
                .header-tabs .q-tab__icon { font-size: 23px; margin: 0; }
                .header-tabs .q-tab__label { display: none; }

                .nit-content { padding: 3px 6px 10px !important; gap: 6px !important; }
                .q-tab-panel { padding: 0 !important; }
                .nit-card {
                  min-width: 0; max-width: 100%; padding: 12px !important; border-radius: 15px;
                  backdrop-filter: none;
                }
                .nit-content > .row.mt-8 { margin-top: 12px !important; }
                .nit-card > .row, .nit-card > .column > .row {
                  flex-wrap: wrap;
                }
                .nit-card .q-btn__content {
                  min-width: 0; white-space: normal; line-height: 1.2;
                }
                .lesson-countdown {
                  flex: 1 1 100%; margin-left: 44px; width: max-content;
                  max-width: calc(100% - 44px); text-align: center;
                }
                .nit-card .q-field { max-width: 100%; }
                .nit-card .min-w-72, .nit-card .min-w-48, .nit-card .min-w-52 {
                  min-width: 0 !important;
                }
                .nit-card h1, .nit-card .text-xl { font-size: 18px !important; }
                .metric {
                  flex: 1 1 calc(50% - 6px) !important;
                  min-width: 0; padding: 11px !important;
                }
                .metric .row { flex-wrap: nowrap !important; }
                .metric .text-3xl { font-size: 24px !important; }
                .metric .q-icon { font-size: 26px !important; }

                .schedule-scroll, .journal-scroll {
                  border-radius: 12px; max-width: 100%; touch-action: pan-x pan-y;
                }
                .schedule-grid { min-width: 740px; }
                .schedule-grid th, .schedule-grid td {
                  min-width: 126px; padding: 8px;
                }
                .schedule-grid .schedule-number { min-width: 54px; width: 54px; }
                .month-grid { min-width: 760px; }
                .month-grid th, .month-grid td { min-width: 106px; }
                .month-day { height: 94px; padding: 6px !important; }
                .month-class-chip { padding: 2px 5px; font-size: 11px; }
                .journal-scroll { max-height: 62vh; }
                .journal-grid .journal-student {
                  min-width: 170px; width: 170px; max-width: 170px;
                }
                .journal-grid th, .journal-grid td {
                  min-width: 56px; height: 48px; padding: 5px 6px;
                }
                .journal-grid .journal-student .row {
                  align-items: flex-start; gap: 5px !important;
                }
                .journal-grid .journal-student .font-medium {
                  display: -webkit-box;
                  overflow: hidden;
                  white-space: normal !important;
                  text-overflow: ellipsis;
                  overflow-wrap: anywhere;
                  -webkit-box-orient: vertical;
                  -webkit-line-clamp: 2;
                  line-height: 1.25;
                  font-size: 12px;
                }
                .journal-grid .journal-student .text-xs {
                  width: 16px !important; flex: 0 0 16px;
                }
                .journal-grid thead .journal-student .font-bold {
                  font-size: 13px;
                }

                .bells-grid { grid-template-columns: 1fr; }
                .bell-card { padding: 10px; gap: 9px; }
                .bell-number { width: 40px; height: 40px; flex-basis: 40px; }
                .bell-time { font-size: 16px; }
                .bell-break { padding: 3px 6px; font-size: 11px; }
                .bell-editor-row { align-items: flex-end !important; }
                .bell-editor-row .q-field { flex: 1 1 calc(50% - 8px); width: auto !important; }
                .bell-editor-row .q-field:first-child { flex: 0 0 68px; }

                .q-table__container { max-width: 100%; }
                .q-table__middle { max-width: 100%; overflow-x: auto !important; }
                .overview-table table { min-width: 700px; }
                .calendar-detail-table table { min-width: 780px; }
                .responsive-controls { align-items: stretch !important; }
                .responsive-controls > .q-field { flex: 1 1 100%; width: 100%; }
                .responsive-controls > .q-checkbox { flex: 1 1 100%; }
                .responsive-controls > .q-btn { margin-left: auto; }
                .q-dialog__inner { padding: 12px !important; }
                .q-dialog__inner > .q-card {
                  width: 100% !important; max-width: calc(100vw - 24px) !important;
                  padding: 16px !important;
                }
                .q-notification { max-width: calc(100vw - 24px) !important; }
              }

              @media (max-width: 420px) {
                .status-pill { display: none !important; }
                .header-actions { min-width: 42px; }
                .nit-card .q-btn { font-size: 12px; }
              }
            </style>
            """
        )

        with ui.header().classes("nit-header px-4 md:px-7 py-2 gap-1").style(
            "flex-direction: column; align-items: stretch;"
        ):
            with ui.row().classes("header-main w-full items-center gap-4 flex-nowrap"):
                with ui.row().classes("header-brand items-center gap-4 flex-nowrap shrink-0"):
                    with ui.element("div").classes("brand-orbit"):
                        ui.image(Path(__file__).with_name("image.png")).classes("brand-logo")
                    with ui.column().classes("gap-0"):
                        ui.label("NIT²").classes("brand-name text-2xl font-bold leading-none")
                        ui.label("by FIZYK · краще ніж було").classes(
                            "brand-tagline text-xs text-cyan-100 mt-1"
                        )

                self.header_controls = ui.row().classes(
                    "header-controls grow items-center justify-center gap-2 flex-nowrap"
                )

                with ui.row().classes("header-actions items-center gap-2 flex-nowrap shrink-0"):
                    with ui.button(
                        icon="light_mode",
                        on_click=self._toggle_theme,
                    ).props("flat round color=white") as self.theme_button:
                        self.theme_tooltip = ui.tooltip("Увімкнути світлу тему")
                    with ui.row().classes("status-pill items-center gap-2 flex-nowrap"):
                        ui.icon("radio_button_checked", size="14px").classes("text-cyan-300")
                        self.status_label = ui.label("Не авторизовано").classes(
                            "text-sm text-white whitespace-nowrap"
                        )
                    with ui.button(
                        icon="logout",
                        on_click=self.logout,
                    ).props("flat round color=white") as self.logout_button:
                        ui.tooltip("Вийти з облікового запису")
                    self.logout_button.set_visibility(False)

            self.header_tabs = ui.row().classes(
                "header-tabs-wrap w-full items-center justify-center pt-1"
            )

        self.content = ui.column().classes(
            "nit-content w-full max-w-none px-2 md:px-3 pt-1 pb-3 gap-2"
        )
        self.render()
        self.session_timer = ui.timer(60.0, self._keep_session_alive)
        self.today_clock_timer = ui.timer(1.0, self._update_today_lesson_widget)

        if (
            self.session_store.load() is not None
            and self.session_handle in _WEB_SESSION_CONFIGS
        ):
            ui.timer(0.2, self.connect_saved, once=True)

    def _toggle_theme(self) -> None:
        if self.dark_mode is None:
            return
        dark = not bool(self.dark_mode.value)
        self.dark_mode.set_value(dark)
        if self.theme_button is not None:
            self.theme_button.set_icon("light_mode" if dark else "dark_mode")
        if self.theme_tooltip is not None:
            self.theme_tooltip.set_text(
                "Увімкнути світлу тему" if dark else "Увімкнути темну тему"
            )

    def render(self) -> None:
        assert self.content is not None
        if self.logout_button is not None:
            self.logout_button.set_visibility(self.auth is not None)
        self.content.clear()
        with self.content:
            if self.auth is None:
                self._clear_header_navigation()
                if self.offline_schedule_mode:
                    self._render_offline_schedule()
                else:
                    self._render_login()
            else:
                self._render_dashboard()

    def _clear_header_navigation(self) -> None:
        if self.header_controls is not None:
            self.header_controls.clear()
            self.header_controls.set_visibility(False)
        if self.header_tabs is not None:
            self.header_tabs.clear()
            self.header_tabs.set_visibility(False)

    def _render_login(self) -> None:
        with ui.row().classes("w-full justify-center mt-8"):
            with ui.card().classes("nit-card w-full max-w-lg p-7 gap-5"):
                with ui.row().classes("items-center gap-3"):
                    ui.icon("science", size="34px").classes("text-primary")
                    with ui.column().classes("gap-0"):
                        ui.label("Вітаємо у NIT²").classes("text-2xl font-bold")
                        ui.label("Інструменти вчителя від FIZYK").classes("text-sm text-slate-500")
                ui.label(
                    "Кожен браузер має окрему сесію. Токени зберігаються лише в пам’яті "
                    "сервера до виходу або його перезапуску."
                ).classes("text-sm text-slate-600")

                email = ui.input("Email НІТ").props("outlined autocomplete=username").classes("w-full")
                password = ui.input(
                    "Пароль НІТ",
                    password=True,
                    password_toggle_button=True,
                ).props("outlined autocomplete=current-password").classes("w-full")

                async def login() -> None:
                    raw_password = password.value or ""
                    password.value = ""
                    password.update()
                    await self.connect_password(email.value or "", raw_password)

                ui.button("Увійти", icon="login", on_click=login).classes("w-full")
                ui.label("Пароль використовується лише для OAuth-входу й не зберігається.").classes(
                    "text-xs text-slate-500"
                )

    def _open_offline_schedule(self) -> None:
        self.offline_schedule_mode = True
        if self.schedule_data is not None:
            self.schedule_data.source = "cache"
        self.render()

    def _close_offline_schedule(self) -> None:
        self.offline_schedule_mode = False
        self.render()

    def _render_offline_schedule(self) -> None:
        with ui.row().classes("w-full items-center justify-between"):
            ui.label("Офлайн-перегляд збереженого розкладу").classes("text-lg font-semibold")
            ui.button("До входу", icon="login", on_click=self._close_offline_schedule).props(
                "outline no-caps"
            )
        self.schedule_box = ui.column().classes("w-full gap-3")
        self._refresh_schedule_panel()

    async def connect_saved(self) -> None:
        if self.busy:
            return
        self.busy = True
        self._set_status("Поновлення збереженої сесії…")
        try:
            config = _WEB_SESSION_CONFIGS.get(self.session_handle)
            if config is None:
                raise NitError("Сесію цього браузера не знайдено")
            context = await run.io_bound(
                connect_with_saved_session,
                self.session_store,
                config,
            )
            if context is None:
                raise NitError("Збережена сесія не знайдена або не поновилася")
            await self._finish_connect(context)
        except Exception as exc:
            self.session_store.delete()
            _WEB_SESSION_CONFIGS.pop(self.session_handle, None)
            self.auth = None
            self.schedule_storage = None
            self.schedule_data = None
            self._show_error(exc)
            self._set_status("Не авторизовано")
            self.render()
        finally:
            self.busy = False

    async def connect_password(self, email: str, password: str) -> None:
        if self.busy:
            return
        if not email.strip() or not password:
            self._notify("Введіть email і пароль", type="warning")
            return
        self.busy = True
        self._set_status("Авторизація…")
        try:
            preferred_school = self.server_config.school_id or None
            context = await run.io_bound(
                connect_user_with_password,
                email.strip(),
                password,
                self.session_store,
                preferred_school,
            )
            if context.user_config is None:
                raise NitError("НІТ не повернув профіль користувача")
            _WEB_SESSION_CONFIGS[self.session_handle] = context.user_config
            await self._finish_connect(context)
        except Exception as exc:
            self.session_store.delete()
            _WEB_SESSION_CONFIGS.pop(self.session_handle, None)
            self.auth = None
            self.schedule_storage = None
            self.schedule_data = None
            self._show_error(exc)
            self._set_status("Не авторизовано")
            self.render()
        finally:
            self.busy = False

    async def _finish_connect(self, context: AuthContext) -> None:
        self.auth = context
        self.session_refresh_failed = False
        config = context.user_config
        if config is not None:
            identity = profile_storage_key(config.profile, config.school_id, config.user_id)
            self.schedule_storage = ScheduleStorage(Path.cwd() / ".nit2_data" / "schedule", identity)
            self.bell_settings = self.schedule_storage.load_bells()
            self.official_bells = self.schedule_storage.load_official_bells()
            if self.official_bells is not None:
                self.bell_source = "nit"
            self.schedule_data = self.schedule_storage.load_week(self.schedule_week)
        self.semester = sorted(context.semester_schedules)[0]
        await self._load_journals(load_first_report=True)
        await self._load_official_bells()
        await self._load_schedule_week(render_after=False, notify=False)
        year = context.active_year
        self._set_status(f"{year.get('yearStart')}/{year.get('yearEnd')} · підключено")
        self.render()

    async def logout(self) -> None:
        """Завершує лише сесію поточного браузера та очищає його дані з UI."""
        if self.busy:
            self._notify("Зачекайте завершення поточної операції", type="warning")
            return
        self.session_store.delete()
        _WEB_SESSION_CONFIGS.pop(self.session_handle, None)
        self.session_store = MemorySessionStore()
        _WEB_SESSION_STORES[self.session_handle] = self.session_store

        self.auth = None
        self.semester = None
        self.journals = []
        self.selected_journal_id = None
        self.report = None
        self.calendar_reports = []
        self.calendar_loaded = False
        self.calendar_loading = False
        self.calendar_load_attempted = False
        self.journal_matrix = None
        self.journal_loading = False
        self.journal_load_attempted = False
        self.excel_name = None
        self.excel_sheet = None
        self.excel_rows = []
        self.excel_has_date = False
        self.import_preview = None
        self.import_preview_token = None
        self.import_in_progress = False
        self.absence_reminders = []
        self.absence_loaded = False
        self.absence_loading = False
        self.absence_load_attempted = False
        self.absence_class = None
        self.absence_action_token = None
        self.absence_in_progress = False
        self.absence_row_map = {}
        self.schedule_data = None
        self.schedule_storage = None
        self.bell_settings = BellSettings(
            "Звичайний", {"Звичайний": [], "Скорочений": []}
        )
        self.official_bells = None
        self.bell_source = "local"
        self.schedule_error = None
        self.offline_schedule_mode = False
        self.session_refresh_failed = False
        self._set_status("Не авторизовано")
        self.render()

    async def _ensure_session(self) -> None:
        if self.auth is None:
            raise NitError("Сесія НІТ не активна")
        if not await run.io_bound(refresh_auth_context, self.auth, 300):
            self.session_refresh_failed = True
            raise NitError(
                "Не вдалося поновити сесію НІТ. Увійдіть знову; "
                "Preview і жодні зміни не були записані."
            )
        self.session_refresh_failed = False

    async def _keep_session_alive(self) -> None:
        if self.client.is_deleted:
            if self.session_timer is not None:
                self.session_timer.deactivate()
            return
        if not self.client.has_socket_connection:
            return
        if self.auth is None or self.busy:
            return
        already_failed = self.session_refresh_failed
        try:
            await self._ensure_session()
            if already_failed:
                self._set_connected_status()
        except Exception as exc:
            self._set_status("Сесія потребує повторного входу")
            if not already_failed:
                self._show_error(exc)

    async def _load_journals(self, load_first_report: bool = True) -> None:
        if not self.auth or self.semester is None:
            return
        await self._ensure_session()
        schedule_id = self.auth.semester_schedules[self.semester]
        self.journals = await run.io_bound(self.auth.client.get_journals, schedule_id)
        self.selected_journal_id = (
            str(self.journals[0].get("journalId")) if self.journals else None
        )
        self.report = None
        self.calendar_reports = []
        self.calendar_loaded = False
        self.calendar_loading = False
        self.calendar_load_attempted = False
        self.journal_matrix = None
        self.journal_loading = False
        self.journal_load_attempted = False
        if load_first_report and self.selected_journal_id:
            await self._load_selected_report(render_after=False)

    def _render_dashboard(self) -> None:
        assert self.auth is not None
        tabs, overview_tab, schedule_tab, journal_tab, import_tab, absence_tab = (
            self._render_header_navigation()
        )
        self.main_tabs = tabs
        self.schedule_tab = schedule_tab
        self.journal_tab = journal_tab
        self.import_tab = import_tab
        self.absence_tab = absence_tab

        with ui.tab_panels(tabs, value=overview_tab).classes("w-full bg-transparent"):
            with ui.tab_panel(overview_tab).classes("overview-panel p-0"):
                self._render_overview()
            with ui.tab_panel(schedule_tab).classes("p-0"):
                self.schedule_box = ui.column().classes("w-full gap-3")
                self._refresh_schedule_panel()
            with ui.tab_panel(journal_tab).classes("p-0"):
                self.journal_box = ui.column().classes("w-full gap-4")
                self._refresh_journal_panel()
            with ui.tab_panel(import_tab).classes("p-0"):
                self.import_box = ui.column().classes("w-full gap-4")
                self._refresh_import_panel()
            with ui.tab_panel(absence_tab).classes("p-0"):
                self.absence_box = ui.column().classes("w-full gap-4")
                self._refresh_absence_panel()

        if not self.absence_loaded and not self.absence_loading:
            ui.timer(0.5, self._auto_load_absences, once=True)

    def _render_header_navigation(self) -> tuple[Any, Any, Any, Any, Any, Any]:
        assert self.auth is not None
        assert self.header_controls is not None
        assert self.header_tabs is not None

        available = {
            semester: f"{semester} семестр"
            for semester in sorted(self.auth.semester_schedules)
        }
        journal_options = {
            str(item.get("journalId")): f"{item.get('className')} — {item.get('subject')}"
            for item in self.journals
            if item.get("journalId")
        }

        self.header_controls.clear()
        self.header_controls.set_visibility(True)
        with self.header_controls:
            ui.toggle(
                available,
                value=self.semester,
                on_change=self._semester_changed,
            ).props("no-caps dense unelevated").classes("header-semesters shrink-0")
            ui.select(
                journal_options,
                value=self.selected_journal_id,
                label="Клас і предмет",
                on_change=self._journal_changed,
            ).props("outlined dense options-dense dark").classes("grow min-w-52")
            with ui.button(
                icon="refresh",
                on_click=self.refresh_current,
            ).props("flat round color=white"):
                ui.tooltip("Оновити вибраний журнал")

        self.header_tabs.clear()
        self.header_tabs.set_visibility(True)
        with self.header_tabs:
            with ui.tabs(on_change=self._tab_changed).classes("header-tabs") as tabs:
                overview_tab = ui.tab("Огляд", icon="assessment").props("no-caps")
                schedule_tab = ui.tab("Розклад", icon="calendar_month").props("no-caps")
                journal_tab = ui.tab("Журнал", icon="assignment").props("no-caps")
                import_tab = ui.tab("Імпорт Excel", icon="upload_file").props("no-caps")
                absence_tab = ui.tab("Відсутності", icon="person_off").props("no-caps")

        return tabs, overview_tab, schedule_tab, journal_tab, import_tab, absence_tab

    def _render_overview(self) -> None:
        report = self.report
        lessons = report.lessons if report else []
        filled = sum(1 for item in lessons if item.get("isFilled"))
        unfilled = len(lessons) - filled

        with ui.element("div").classes("overview-dashboard-grid"):
            self._render_today_schedule_widget()
            self._render_attention_widget(unfilled)

        journal = report.journal if report else self._current_journal() or {}
        class_name = compact(journal.get("className")) or "Клас не вибрано"
        subject = compact(journal.get("subject")) or "Предмет"
        progress = filled / len(lessons) if lessons else 0.0
        today = datetime.now(KYIV).date()
        upcoming = sorted(
            (day for item in lessons if (day := lesson_day(item)) and day >= today)
        )
        with ui.card().classes("nit-card w-full p-4 gap-3"):
            with ui.element("div").classes("journal-summary-bar"):
                with ui.column().classes("gap-2 min-w-0"):
                    with ui.row().classes("items-center gap-2 flex-wrap"):
                        ui.icon("menu_book", size="24px").classes("text-primary")
                        ui.label(f"{class_name} · {subject}").classes(
                            "text-lg font-semibold"
                        )
                    ui.label(
                        f"{len(lessons)} уроків · {filled} заповнено · "
                        f"{unfilled} потребують уваги"
                    ).classes("text-sm text-slate-500")
                    ui.linear_progress(
                        value=progress,
                        show_value=False,
                        size="8px",
                    ).props(
                        "rounded color=primary track-color=blue-grey-8"
                    ).classes("w-full")
                    if upcoming:
                        ui.label(f"Найближчий урок: {upcoming[0]:%d.%m.%Y}").classes(
                            "text-xs text-slate-500"
                        )
                with ui.row().classes("journal-summary-actions items-center gap-2"):
                    ui.button(
                        "Відкрити журнал",
                        icon="assignment",
                        on_click=self._open_journal_tab,
                    ).props("outline no-caps")
                    ui.button(
                        "Імпорт Excel",
                        icon="upload_file",
                        on_click=self._open_import_tab,
                    ).props("no-caps")

        with ui.card().classes("nit-card w-full p-4"):
            with ui.row().classes("w-full justify-between items-center"):
                ui.label("Уроки вибраного журналу").classes("text-lg font-semibold")
                with ui.row().classes("gap-2"):
                    ui.button("Excel", icon="download", on_click=lambda: self.export_report("xlsx")).props(
                        "outline no-caps"
                    )
                    ui.button("JSON", icon="download", on_click=lambda: self.export_report("json")).props(
                        "outline no-caps"
                    )
            ui.table(
                columns=self._lesson_columns(),
                rows=self._lesson_rows(lessons),
                row_key="key",
                pagination={"rowsPerPage": 15},
            ).props("flat bordered").classes("nit-data-table overview-table w-full")

    async def _tab_changed(self, event: Any) -> None:
        selected = getattr(event.value, "label", event.value)
        selected = str(selected or "")
        if selected == "Розклад" and self.schedule_data is None and not self.schedule_loading:
            await self._load_schedule_week()
        elif selected == "Журнал" and self.journal_matrix is None and not self.journal_loading:
            await self.load_journal_matrix()
        elif selected == "Відсутності" and not self.absence_loaded and not self.absence_loading:
            await self.load_absences()

    def _open_schedule_tab(self) -> None:
        if self.main_tabs is not None and self.schedule_tab is not None:
            self.main_tabs.set_value(self.schedule_tab)
            if self.schedule_data is None and not self.schedule_loading:
                ui.timer(0.05, self._load_schedule_week, once=True)

    def _open_journal_tab(self) -> None:
        if self.main_tabs is not None and self.journal_tab is not None:
            self.main_tabs.set_value(self.journal_tab)
            if self.journal_matrix is None and not self.journal_loading:
                ui.timer(0.05, self.load_journal_matrix, once=True)

    def _open_import_tab(self) -> None:
        if self.main_tabs is not None and self.import_tab is not None:
            self.main_tabs.set_value(self.import_tab)

    def _open_absence_tab(self) -> None:
        if self.main_tabs is not None and self.absence_tab is not None:
            self.main_tabs.set_value(self.absence_tab)
            if not self.absence_loaded and not self.absence_loading:
                ui.timer(0.05, self.load_absences, once=True)

    def _render_attention_widget(self, unfilled: int) -> None:
        with ui.card().classes("nit-card overview-widget w-full p-4 gap-3"):
            with ui.row().classes("w-full items-center justify-between"):
                with ui.column().classes("gap-0"):
                    ui.label("Потребує уваги").classes("text-lg font-semibold")
                    ui.label("Те, що варто перевірити").classes("text-sm text-slate-500")
                ui.icon("notifications_active", size="26px").classes(
                    "text-orange-500" if unfilled else "text-green-500"
                )

            with ui.element("div").classes("overview-status-row"):
                ui.icon(
                    "check_circle" if unfilled == 0 else "edit_note",
                    size="20px",
                ).classes("overview-status-icon")
                with ui.column().classes("gap-0 min-w-0"):
                    ui.label(
                        "Вибраний журнал заповнений"
                        if unfilled == 0
                        else f"Не заповнено уроків: {unfilled}"
                    ).classes("font-medium")
                    ui.label("За вибраним класом і предметом").classes(
                        "text-xs text-slate-500"
                    )

            with ui.element("div").classes("overview-status-row"):
                ui.icon("person_off", size="20px").classes("overview-status-icon")
                with ui.column().classes("gap-0 min-w-0 grow"):
                    if self.absence_loaded:
                        self.overview_absence_title = ui.label(
                            f"Непідтверджених відсутностей: {len(self.absence_reminders)}"
                        ).classes("font-medium")
                        self.overview_absence_meta = ui.label(
                            "У всіх класах цього семестру"
                        ).classes(
                            "text-xs text-slate-500"
                        )
                    elif self.absence_loading:
                        self.overview_absence_title = ui.label(
                            "Перевіряю відсутності…"
                        ).classes("font-medium")
                        self.overview_absence_meta = ui.label(
                            "Завантажую червоні трикутники з НІТ"
                        ).classes("text-xs text-slate-500")
                    else:
                        self.overview_absence_title = ui.label(
                            "Відсутності ще не перевірено"
                        ).classes("font-medium")
                        self.overview_absence_meta = ui.label(
                            "Очікую на фонову перевірку"
                        ).classes(
                            "text-xs text-slate-500"
                        )
                ui.button(icon="arrow_forward", on_click=self._open_absence_tab).props(
                    "flat round"
                )

    async def _auto_load_absences(self) -> None:
        if self.auth is None or self.absence_loaded or self.absence_loading:
            return
        if self.busy:
            ui.timer(0.5, self._auto_load_absences, once=True)
            return
        await self.load_absences()

    def _update_overview_absence_widget(self) -> None:
        if self.overview_absence_title is None or self.overview_absence_meta is None:
            return
        if self.absence_loaded:
            self.overview_absence_title.set_text(
                f"Непідтверджених відсутностей: {len(self.absence_reminders)}"
            )
            self.overview_absence_meta.set_text("У всіх класах цього семестру")
        elif self.absence_loading:
            self.overview_absence_title.set_text("Перевіряю відсутності…")
            self.overview_absence_meta.set_text("Завантажую червоні трикутники з НІТ")
        elif self.absence_load_attempted:
            self.overview_absence_title.set_text("Не вдалося перевірити відсутності")
            self.overview_absence_meta.set_text("Відкрийте вкладку та спробуйте ще раз")
        else:
            self.overview_absence_title.set_text("Відсутності ще не перевірено")
            self.overview_absence_meta.set_text("Очікую на фонову перевірку")

    def _render_today_schedule_widget(self) -> None:
        today = datetime.now(KYIV).date()
        with ui.card().classes("nit-card overview-widget w-full p-4 gap-3"):
            with ui.row().classes("w-full items-center justify-between"):
                with ui.column().classes("gap-0"):
                    ui.label("Сьогодні").classes("text-lg font-semibold")
                    self.today_date_label = ui.label(today.strftime("%d.%m.%Y")).classes(
                        "text-sm text-slate-500"
                    )
                ui.icon("today", size="26px").classes("text-primary")

            with ui.element("div").classes("overview-status-row"):
                ui.icon("schedule", size="20px").classes("overview-status-icon")
                with ui.column().classes("gap-0 min-w-0 grow"):
                    self.today_focus_title = ui.label("").classes("font-medium")
                    self.today_focus_meta = ui.label("").classes(
                        "text-xs text-slate-500"
                    )
                self.today_countdown_label = ui.label("").classes("lesson-countdown")

            with ui.row().classes("w-full items-center justify-between gap-2"):
                self.today_total_label = ui.label("").classes(
                    "text-sm text-slate-500"
                )
                ui.button(
                    "Розклад",
                    icon="calendar_month",
                    on_click=self._open_schedule_tab,
                ).props("outline no-caps")

        self._update_today_lesson_widget()

    def _today_schedule(self, today: date) -> WeekSchedule | None:
        current_week = monday_of(today)
        if self.schedule_data and self.schedule_data.week_start == current_week:
            return self.schedule_data
        if self.schedule_storage is not None:
            return self.schedule_storage.load_week(current_week)
        return None

    @staticmethod
    def _countdown_text(prefix: str, target: datetime, moment: datetime) -> str:
        seconds = max(0, int((target - moment).total_seconds() + 0.999))
        hours, remainder = divmod(seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        clock = f"{hours}:{minutes:02d}:{seconds:02d}" if hours else f"{minutes:02d}:{seconds:02d}"
        return f"{prefix} {clock}"

    def _update_today_lesson_widget(self) -> None:
        labels = (
            self.today_date_label,
            self.today_focus_title,
            self.today_focus_meta,
            self.today_countdown_label,
            self.today_total_label,
        )
        if self.auth is None or any(label is None for label in labels):
            return

        moment = datetime.now(KYIV)
        today = moment.date()
        schedule = self._today_schedule(today)
        lessons = [item for item in schedule.lessons if item.day == today] if schedule else []
        bells = self._effective_bells()
        current, following = current_and_next(schedule, bells, moment)

        self.today_date_label.set_text(today.strftime("%d.%m.%Y"))
        self.today_total_label.set_text(f"Усього уроків: {len(lessons)}")

        countdown_text = ""
        if current:
            lesson, period = current
            self.today_focus_title.set_text(f"Зараз: {lesson.number} · {lesson.class_name}")
            self.today_focus_meta.set_text(
                f"{lesson.subject} · {period.start}–{period.end}"
            )
            end_clock = parse_clock(period.end)
            if end_clock is not None:
                target = datetime.combine(today, end_clock, KYIV)
                countdown_text = self._countdown_text("До кінця уроку", target, moment)
        elif following:
            lesson, period = following
            self.today_focus_title.set_text(
                f"Наступний урок: {lesson.number} · {lesson.class_name}"
            )
            self.today_focus_meta.set_text(
                f"{lesson.subject} · {period.start}–{period.end}"
            )
            start_clock = parse_clock(period.start)
            if start_clock is not None:
                target = datetime.combine(today, start_clock, KYIV)
                countdown_text = self._countdown_text("До початку", target, moment)
        elif lessons:
            periods = {
                item.number: item for item in bells.profiles.get(bells.active_profile, [])
            }
            last_lesson = max(lessons, key=lambda item: item.number)
            last_period = periods.get(last_lesson.number)
            self.today_focus_title.set_text("Навчальний день завершено")
            self.today_focus_meta.set_text(
                f"Останній урок закінчився о {last_period.end}"
                if last_period else
                f"Останнім був {last_lesson.number}-й урок"
            )
        else:
            self.today_focus_title.set_text("Уроків на сьогодні немає")
            self.today_focus_meta.set_text("Можна планувати інші справи")

        self.today_countdown_label.set_text(countdown_text)
        self.today_countdown_label.set_visibility(bool(countdown_text))

    def _refresh_schedule_panel(self) -> None:
        if self.schedule_box is None:
            return
        self.schedule_box.clear()
        with self.schedule_box:
            self._render_schedule_panel()

    def _effective_bells(self) -> BellSettings:
        if self.bell_source == "nit" and self.official_bells is not None:
            return BellSettings("НІТ", {"НІТ": self.official_bells.periods})
        return self.bell_settings

    async def _load_official_bells(self) -> None:
        if self.auth is None or self.schedule_storage is None:
            return
        try:
            await self._ensure_session()
            payload = await run.io_bound(self.auth.client.get_academic_years)
            bells = normalize_academic_year_bells(payload, self.auth.active_year)
            await run.io_bound(self.schedule_storage.save_official_bells, bells)
            self.official_bells = bells
            self.bell_source = "nit"
        except Exception:
            cached = await run.io_bound(self.schedule_storage.load_official_bells)
            if cached is not None:
                self.official_bells = cached
                self.bell_source = "nit"

    async def _load_schedule_week(self, render_after: bool = True, notify: bool = True) -> None:
        if self.auth is None or self.schedule_storage is None or self.schedule_loading:
            return
        requested_week = self.schedule_week
        self.schedule_loading = True
        self.schedule_error = None
        try:
            await self._ensure_session()
            payload = await run.io_bound(self.auth.client.get_teacher_diary, requested_week)
            schedule = normalize_teacher_diary(payload, requested_week)
            await run.io_bound(self.schedule_storage.save_week, schedule)
            self.schedule_data = schedule
        except Exception as exc:
            cached = await run.io_bound(self.schedule_storage.load_week, requested_week)
            self.schedule_error = str(exc)
            if cached is not None:
                cached.warning = self.schedule_error
                self.schedule_data = cached
                if notify:
                    self._notify("НІТ недоступний: показано збережену копію", type="warning")
            else:
                self.schedule_data = None
                if notify:
                    self._show_error(exc)
        finally:
            self.schedule_loading = False
            if render_after:
                self._refresh_schedule_panel()

    async def _move_schedule_week(self, weeks: int) -> None:
        if self.schedule_loading:
            return
        self.schedule_week = monday_of(self.schedule_week + timedelta(days=7 * weeks))
        if self.auth is None:
            self.schedule_data = self.schedule_storage.load_week(self.schedule_week) if self.schedule_storage else None
            self._refresh_schedule_panel()
        else:
            await self._load_schedule_week()

    async def _schedule_today(self) -> None:
        if self.schedule_loading:
            return
        self.schedule_week = monday_of(datetime.now(KYIV).date())
        if self.auth is None:
            self.schedule_data = self.schedule_storage.load_week(self.schedule_week) if self.schedule_storage else None
            self._refresh_schedule_panel()
        else:
            await self._load_schedule_week()

    async def _schedule_view_changed(self, event: Any) -> None:
        self.schedule_view = str(event.value or "lessons")
        self._refresh_schedule_panel()
        if self.schedule_view == "month" and not self.calendar_loaded and not self.calendar_loading:
            await self.load_calendar()

    def _render_schedule_panel(self) -> None:
        with ui.card().classes("nit-card w-full p-4 gap-3"):
            with ui.row().classes("w-full items-center justify-between gap-3"):
                with ui.column().classes("gap-0"):
                    ui.label("Персональний розклад").classes("text-xl font-semibold")
                    if self.schedule_view == "lessons":
                        end = self.schedule_week + timedelta(days=6)
                        ui.label(f"{self.schedule_week:%d.%m.%Y} — {end:%d.%m.%Y}").classes("text-sm text-slate-500")
                    elif self.schedule_view == "month":
                        ui.label("Уроки всіх класів за датами").classes("text-sm text-slate-500")
                    else:
                        ui.label("Офіційний час НІТ або локальні профілі").classes("text-sm text-slate-500")
                if self.schedule_view == "lessons":
                    with ui.row().classes("items-center gap-1"):
                        ui.button(icon="chevron_left", on_click=lambda: self._move_schedule_week(-1)).props("flat round")
                        ui.button("Сьогодні", on_click=self._schedule_today).props("outline no-caps")
                        ui.button(icon="chevron_right", on_click=lambda: self._move_schedule_week(1)).props("flat round")
                        if self.auth is not None:
                            ui.button(icon="refresh", on_click=self._load_schedule_week).props("flat round")
                            with ui.button(
                                "Усі вчителі · JSON",
                                icon="download",
                                on_click=self.export_all_teachers_schedule,
                            ).props("outline no-caps"):
                                ui.tooltip("Завантажити розклад усіх учителів за цей тиждень")
            with ui.row().classes("w-full items-center justify-between gap-3"):
                view_options = {"lessons": "Тиждень", "bells": "Дзвінки"}
                if self.auth is not None:
                    view_options = {"lessons": "Тиждень", "month": "Місяць", "bells": "Дзвінки"}
                ui.toggle(
                    view_options,
                    value=self.schedule_view,
                    on_change=self._schedule_view_changed,
                ).props("no-caps")
                if self.schedule_view == "lessons" and self.schedule_data:
                    local_time = self.schedule_data.fetched_at.astimezone(KYIV).strftime("%d.%m.%Y %H:%M")
                    source = "збережена копія" if self.schedule_data.source == "cache" else "НІТ"
                    ui.label(f"Оновлено: {local_time} · {source}").classes("text-xs text-slate-500")
            if self.schedule_loading:
                with ui.row().classes("items-center gap-2"):
                    ui.spinner(size="24px")
                    ui.label("Отримую розклад із НІТ…")
            elif self.schedule_error and self.schedule_data is None:
                ui.label(self.schedule_error).classes("text-red-600")

        if self.schedule_view == "month":
            self._render_calendar()
        elif self.schedule_view == "bells":
            self._render_bells_editor()
        else:
            self._render_week_grid()

    async def export_all_teachers_schedule(self) -> None:
        if self.auth is None:
            self._notify("Спочатку увійдіть у НІТ", type="warning")
            return
        if self.busy:
            return
        self.busy = True
        self._set_status("Формую JSON розкладу всіх учителів…")
        try:
            await self._ensure_session()
            document = await run.io_bound(
                build_all_teachers_schedule,
                self.auth.client,
                self.schedule_week,
            )
            filename = (
                f"розклад_усіх_вчителів_{document['weekStart']}_{document['weekEnd']}.json"
            )
            path = Path(tempfile.gettempdir()) / f"nit2_{secrets.token_hex(4)}_{filename}"
            content = json.dumps(document, ensure_ascii=False, indent=2) + "\n"
            await run.io_bound(path.write_text, content, "utf-8")
            ui.download(str(path), filename=filename)
        except Exception as exc:
            self._show_error(exc)
        finally:
            self.busy = False
            self._set_connected_status()

    def _render_current_lesson(self) -> None:
        current, following = current_and_next(self.schedule_data, self._effective_bells())
        with ui.row().classes("w-full items-center gap-2 rounded-xl px-3 py-2 bg-indigo-50"):
            ui.icon("schedule", size="18px").classes("text-primary shrink-0")
            if current:
                lesson, period = current
                ui.label(
                    f"Зараз: {lesson.number} урок · {lesson.class_name} · "
                    f"{lesson.subject} · до {period.end}"
                ).classes("text-sm font-medium")
            if following:
                lesson, period = following
                prefix = "• Далі" if current else "Далі"
                ui.label(
                    f"{prefix}: {lesson.number} урок · {lesson.class_name} · "
                    f"{lesson.subject} · о {period.start}"
                ).classes("text-sm text-slate-600")
            if not current and not following:
                effective = self._effective_bells()
                has_bells = bool(effective.profiles.get(effective.active_profile))
                ui.label(
                    "На сьогодні уроків більше немає."
                    if has_bells else
                    "Немає даних про час уроків."
                ).classes("text-sm text-slate-500")

    def _render_week_grid(self) -> None:
        schedule = self.schedule_data
        with ui.card().classes("nit-card w-full p-4"):
            if self.schedule_week == monday_of(datetime.now(KYIV).date()):
                self._render_current_lesson()
            if schedule is None:
                ui.label("Даних цього тижня ще немає.").classes("text-slate-500")
                return
            if schedule.source == "cache":
                ui.label("Офлайн-режим: відображається остання збережена копія.").classes(
                    "text-orange-600 font-medium"
                )
            numbers = sorted({item.number for item in schedule.lessons})
            if not numbers:
                ui.label("На цьому тижні уроків немає.").classes("text-slate-500")
                return
            by_cell: dict[tuple[int, int], list[Any]] = {}
            for lesson in schedule.lessons:
                by_cell.setdefault((lesson.weekday, lesson.number), []).append(lesson)
            today = datetime.now(KYIV).date()
            names = ("Понеділок", "Вівторок", "Середа", "Четвер", "П’ятниця")
            effective = self._effective_bells()
            periods = {item.number: item for item in effective.profiles.get(effective.active_profile, [])}
            with ui.element("div").classes("schedule-scroll"):
                with ui.element("table").classes("schedule-grid"):
                    with ui.element("thead"), ui.element("tr"):
                        with ui.element("th").classes("schedule-number"):
                            ui.label("№")
                        for weekday, name in enumerate(names):
                            day = self.schedule_week + timedelta(days=weekday)
                            with ui.element("th").classes("schedule-today" if day == today else ""):
                                ui.label(name).classes("font-bold")
                                ui.label(day.strftime("%d.%m")).classes("text-xs text-slate-500")
                    with ui.element("tbody"):
                        for number in numbers:
                            with ui.element("tr"):
                                with ui.element("td").classes("schedule-number"):
                                    ui.label(str(number))
                                    period = periods.get(number)
                                    if period:
                                        ui.label(f"{period.start}–{period.end}").classes("text-xs text-slate-500 whitespace-nowrap")
                                for weekday in range(5):
                                    day = self.schedule_week + timedelta(days=weekday)
                                    with ui.element("td").classes("schedule-today" if day == today else ""):
                                        items = by_cell.get((weekday, number), [])
                                        if not items:
                                            ui.label("—").classes("text-slate-400")
                                        for lesson in items:
                                            with ui.column().classes("schedule-lesson gap-0"):
                                                ui.label(lesson.class_name or "Без класу").classes("font-bold")
                                                ui.label(lesson.subject or "Без назви").classes("text-sm")
                                                if lesson.room:
                                                    ui.label(f"Кабінет: {lesson.room}").classes("text-xs text-slate-500")

    def _save_bells(self) -> None:
        if self.schedule_storage is None:
            return
        try:
            self.schedule_storage.save_bells(self.bell_settings)
        except Exception as exc:
            self._show_error(exc)

    def _bell_profile_changed(self, event: Any) -> None:
        value = str(event.value or "")
        if value in self.bell_settings.profiles:
            self.bell_settings.active_profile = value
            self._save_bells()
            self._refresh_schedule_panel()

    def _bell_source_changed(self, event: Any) -> None:
        value = str(event.value or "local")
        if value == "nit" and self.official_bells is None:
            return
        self.bell_source = value
        self._refresh_schedule_panel()

    def _update_bell(self, index: int, field: str, value: Any) -> None:
        periods = self.bell_settings.profiles.get(self.bell_settings.active_profile, [])
        if not 0 <= index < len(periods):
            return
        period = periods[index]
        if field == "number":
            try:
                period.number = max(1, int(value))
            except (TypeError, ValueError):
                return
        elif field == "start":
            period.start = str(value or "")
        elif field == "end":
            period.end = str(value or "")
        self._save_bells()

    def _add_bell(self) -> None:
        periods = self.bell_settings.profiles.setdefault(self.bell_settings.active_profile, [])
        periods.append(BellPeriod(max((item.number for item in periods), default=0) + 1, "", ""))
        self._save_bells()
        self._refresh_schedule_panel()

    def _delete_bell(self, index: int) -> None:
        periods = self.bell_settings.profiles.get(self.bell_settings.active_profile, [])
        if 0 <= index < len(periods):
            periods.pop(index)
            self._save_bells()
            self._refresh_schedule_panel()

    def _new_bell_profile(self) -> None:
        with ui.dialog() as dialog, ui.card().classes("w-full max-w-sm p-5 gap-4"):
            ui.label("Новий профіль дзвінків").classes("text-lg font-semibold")
            name = ui.input("Назва").props("outlined autofocus").classes("w-full")

            def create() -> None:
                value = str(name.value or "").strip()
                if not value:
                    self._notify("Введіть назву профілю", type="warning")
                    return
                self.bell_settings.profiles.setdefault(value, [])
                self.bell_settings.active_profile = value
                self._save_bells()
                dialog.close()
                self._refresh_schedule_panel()

            with ui.row().classes("w-full justify-end gap-2"):
                ui.button("Скасувати", on_click=dialog.close).props("flat no-caps")
                ui.button("Створити", on_click=create).props("no-caps")
        dialog.open()

    def _render_bells_editor(self) -> None:
        if self.official_bells is not None:
            with ui.card().classes("nit-card w-full p-4"):
                with ui.row().classes("w-full items-center justify-between gap-3"):
                    with ui.column().classes("gap-0"):
                        ui.label("Джерело часу").classes("font-semibold")
                        ui.label("Офіційні дзвінки НІТ доступні лише для читання.").classes("text-sm text-slate-500")
                    ui.toggle(
                        {"nit": "НІТ · офіційні", "local": "Локальні"},
                        value=self.bell_source,
                        on_change=self._bell_source_changed,
                    ).props("no-caps")
            if self.bell_source == "nit":
                periods = self.official_bells.periods
                with ui.card().classes("nit-card w-full p-4 gap-3"):
                    with ui.row().classes("w-full items-center justify-between"):
                        ui.label("Офіційний розклад дзвінків НІТ").classes("text-lg font-semibold")
                        source = "збережена копія" if self.official_bells.source == "cache" else "актуальні дані НІТ"
                        with ui.row().classes("items-center gap-1"):
                            ui.icon("verified", size="15px").classes("text-cyan-500")
                            ui.label(source).classes("text-xs text-slate-500")
                    with ui.element("div").classes("bells-grid"):
                        for index, period in enumerate(periods):
                            following = periods[index + 1] if index + 1 < len(periods) else None
                            pause = break_minutes(period, following)
                            start_clock, end_clock = parse_clock(period.start), parse_clock(period.end)
                            duration = None
                            if start_clock is not None and end_clock is not None:
                                duration = int(
                                    (
                                        datetime.combine(date.today(), end_clock)
                                        - datetime.combine(date.today(), start_clock)
                                    ).total_seconds()
                                    // 60
                                )
                            with ui.element("div").classes("bell-card"):
                                ui.label(str(period.number)).classes("bell-number")
                                with ui.column().classes("gap-0 min-w-0"):
                                    ui.label(f"{period.start} — {period.end}").classes("bell-time whitespace-nowrap")
                                    ui.label(
                                        f"{duration} хв · {period.number}-й урок"
                                        if duration is not None else f"{period.number}-й урок"
                                    ).classes("bell-meta")
                                ui.label(
                                    f"перерва {pause} хв" if pause is not None else "останній"
                                ).classes("bell-break")
                return
        periods = self.bell_settings.profiles.get(self.bell_settings.active_profile, [])
        with ui.card().classes("nit-card w-full p-4 gap-4"):
            with ui.row().classes("w-full items-center justify-between gap-3"):
                with ui.column().classes("gap-0"):
                    ui.label("Локальний розклад дзвінків").classes("text-lg font-semibold")
                    ui.label("Час не надходить із НІТ і зберігається лише на цьому Mac.").classes(
                        "text-sm text-slate-500"
                    )
                with ui.row().classes("items-center gap-2"):
                    ui.select(
                        list(self.bell_settings.profiles),
                        value=self.bell_settings.active_profile,
                        label="Профіль",
                        on_change=self._bell_profile_changed,
                    ).props("outlined dense options-dense").classes("min-w-48")
                    ui.button(icon="add", on_click=self._new_bell_profile).props("outline round")
            if not periods:
                ui.label("У цьому профілі ще немає дзвінків.").classes("text-slate-500")
            for index, period in enumerate(periods):
                following = periods[index + 1] if index + 1 < len(periods) else None
                pause = break_minutes(period, following)
                with ui.row().classes("bell-editor-row w-full items-center gap-3 flex-nowrap"):
                    ui.number("№", value=period.number, min=1, format="%.0f", on_change=lambda e, i=index: self._update_bell(i, "number", e.value)).props(
                        "outlined dense"
                    ).classes("w-24")
                    ui.input("Початок", value=period.start, on_change=lambda e, i=index: self._update_bell(i, "start", e.value)).props(
                        "outlined dense type=time"
                    ).classes("w-40")
                    ui.input("Кінець", value=period.end, on_change=lambda e, i=index: self._update_bell(i, "end", e.value)).props(
                        "outlined dense type=time"
                    ).classes("w-40")
                    ui.label(f"Перерва: {pause} хв" if pause is not None else "Перерва: —").classes(
                        "text-sm text-slate-500 grow"
                    )
                    ui.button(icon="delete", on_click=lambda _=None, i=index: self._delete_bell(i)).props("flat round color=negative")
            ui.button("Додати урок", icon="add", on_click=self._add_bell).props("outline no-caps")

    def _refresh_journal_panel(self) -> None:
        if self.journal_box is None:
            return
        self.journal_box.clear()
        with self.journal_box:
            self._render_journal_panel()

    def _render_journal_panel(self) -> None:
        matrix = self.journal_matrix
        with ui.card().classes("nit-card w-full p-5 gap-4"):
            with ui.row().classes("w-full items-center justify-between gap-3"):
                with ui.column().classes("gap-0"):
                    ui.label("Електронний журнал").classes("text-xl font-semibold")
                if matrix is not None:
                    ui.button(
                        "Оновити журнал",
                        icon="refresh",
                        on_click=self.load_journal_matrix,
                    ).props("no-caps")
                elif self.journal_load_attempted and not self.journal_loading:
                    ui.button(
                        "Спробувати ще раз",
                        icon="refresh",
                        on_click=self.load_journal_matrix,
                    ).props("no-caps")

            if matrix is None:
                with ui.element("div").classes(
                    "w-full rounded-2xl bg-indigo-50 border border-indigo-100 p-6"
                ):
                    with ui.row().classes("items-center gap-3"):
                        if self.journal_loading or not self.journal_load_attempted:
                            ui.spinner("dots", size="28px")
                        ui.label(
                            "Журнал завантажується автоматично…"
                            if self.journal_loading or not self.journal_load_attempted
                            else "Не вдалося завантажити журнал. Спробуйте ще раз."
                        ).classes("text-indigo-900")
                return

            with ui.row().classes("items-center gap-2"):
                ui.badge(f"{len(matrix.students)} учнів", color="primary")
                ui.badge(f"{len(matrix.lessons)} уроків", color="cyan")
                ui.label("Прокручуйте таблицю вбік →").classes("text-sm text-slate-500 ml-2")

            if not matrix.students:
                ui.label("НІТ не повернув учнів для цього журналу.").classes("text-slate-500")
                return
            if not matrix.lessons:
                ui.label("У цьому журналі ще немає уроків.").classes("text-slate-500")
                return

            with ui.element("div").classes("journal-scroll"):
                with ui.element("table").classes("journal-grid"):
                    with ui.element("thead"):
                        with ui.element("tr"):
                            with ui.element("th").classes("journal-student"):
                                ui.label("Учень").classes("font-bold")
                            for lesson in matrix.lessons:
                                day = lesson_day(lesson)
                                with ui.element("th"):
                                    ui.label(day.strftime("%d.%m") if day else "—").classes(
                                        "font-bold whitespace-nowrap"
                                    )
                    with ui.element("tbody"):
                        for index, student in enumerate(matrix.students, start=1):
                            user_id = str(student.get("userId") or "")
                            full_name = self._student_name(student)
                            with ui.element("tr"):
                                with ui.element("td").classes("journal-student"):
                                    with ui.row().classes("items-center gap-2 flex-nowrap w-full min-w-0"):
                                        ui.label(str(index)).classes(
                                            "text-xs text-slate-400 w-5 text-right"
                                        )
                                        ui.label(full_name).classes(
                                            "font-medium whitespace-nowrap overflow-hidden text-ellipsis min-w-0"
                                        )
                                for lesson in matrix.lessons:
                                    column_id = str(lesson.get("id") or "")
                                    mark = matrix.marks.get((user_id, column_id))
                                    text, css_class = self._journal_mark(mark)
                                    with ui.element("td").classes(css_class):
                                        ui.label(text or "·").classes("journal-mark" if text else "journal-empty")

    async def load_journal_matrix(self) -> None:
        if self.busy or not self.auth or not self.selected_journal_id:
            return
        journal = next(
            (item for item in self.journals if str(item.get("journalId")) == self.selected_journal_id),
            None,
        )
        if journal is None:
            return
        self.busy = True
        self.journal_loading = True
        self.journal_load_attempted = True
        self._set_status(f"Завантаження журналу: {journal.get('className')}…")
        self._refresh_journal_panel()
        try:
            await self._ensure_session()
            self.journal_matrix = await run.io_bound(
                load_journal_matrix,
                self.auth.client,
                journal,
            )
            self._refresh_journal_panel()
            self._set_connected_status()
        except Exception as exc:
            self._show_error(exc)
        finally:
            self.journal_loading = False
            self.busy = False
            self._refresh_journal_panel()
            if self.journal_matrix is not None:
                await ui.run_javascript(
                    "requestAnimationFrame(() => {"
                    "const table = document.querySelector('.journal-scroll');"
                    "if (table) { table.scrollTop = 0; table.scrollLeft = 0; }"
                    "});"
                )

    @staticmethod
    def _student_name(student: dict[str, Any]) -> str:
        full_name = compact(student.get("fullName"))
        if full_name:
            return full_name
        parts = [
            compact(student.get("lastName")),
            compact(student.get("firstName")),
            compact(student.get("middleName")),
        ]
        return " ".join(part for part in parts if part) or "Учень без імені"

    @staticmethod
    def _lesson_number(value: Any) -> str:
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return compact(value) or "—"

    @staticmethod
    def _journal_mark(mark: dict[str, Any] | None) -> tuple[str, str]:
        if not mark:
            return "", ""
        if mark.get("absense") is not None:
            return "н", "journal-absence"
        rating = mark.get("rating")
        if rating is not None:
            return compact(rating), ""
        custom = mark.get("customMark")
        if custom is not None:
            if isinstance(custom, dict):
                custom = custom.get("name") or custom.get("value")
            return compact(custom), ""
        letter = mark.get("letterMark")
        if letter is not None:
            return compact(letter), ""
        leveled = mark.get("leveledMark")
        if leveled is not None:
            level_names = {0: "П", 1: "С", 2: "Д", 3: "В", "0": "П", "1": "С", "2": "Д", "3": "В"}
            return level_names.get(leveled, compact(leveled)), ""
        if mark.get("isVerbalFormed") is True:
            return "Сформ.", ""
        return "", ""

    def _refresh_absence_panel(self) -> None:
        if self.absence_box is None:
            return
        self.absence_box.clear()
        with self.absence_box:
            self._render_absence_panel()

    def _render_absence_panel(self) -> None:
        with ui.card().classes("nit-card w-full p-5 gap-4"):
            with ui.row().classes("w-full items-center justify-between"):
                with ui.column().classes("gap-0"):
                    ui.label("Підтвердження відсутностей").classes("text-xl font-semibold")
                    ui.label(
                        "Завантажує червоні трикутники. Усі рядки можна перевірити та зняти зайві перед записом."
                    ).classes("text-sm text-slate-600")
                if self.absence_loaded:
                    ui.button(
                        "Оновити",
                        icon="refresh",
                        on_click=self.load_absences,
                    ).props("no-caps").set_enabled(not self.absence_in_progress)
                elif self.absence_load_attempted and not self.absence_loading:
                    ui.button(
                        "Спробувати ще раз",
                        icon="refresh",
                        on_click=self.load_absences,
                    ).props("no-caps")

            if not self.absence_loaded:
                with ui.row().classes("items-center gap-3"):
                    if self.absence_loading or not self.absence_load_attempted:
                        ui.spinner("dots", size="28px")
                    ui.label(
                        "Перевіряю червоні трикутники автоматично…"
                        if self.absence_loading or not self.absence_load_attempted
                        else "Не вдалося перевірити відсутності. Спробуйте ще раз."
                    ).classes("text-slate-500")
                return

            if not self.absence_reminders:
                ui.label("Непідтверджених відсутностей немає.").classes("text-slate-500")
                return

            counts: dict[str, int] = {}
            for reminder in self.absence_reminders:
                name = compact(reminder.journal.get("className"))
                counts[name] = counts.get(name, 0) + 1
            class_names = sorted(counts, key=self._class_sort_key)
            if self.absence_class not in counts:
                self.absence_class = class_names[0]

            with ui.row().classes("responsive-controls w-full items-end gap-4"):
                ui.select(
                    {name: f"{name} — {counts[name]} позначок" for name in class_names},
                    value=self.absence_class,
                    label="Клас",
                    on_change=self._absence_class_changed,
                ).props("outlined dense options-dense").classes("grow max-w-md")
                ui.button(
                    "Підтвердити вибрані",
                    icon="person_off",
                    on_click=self.open_absence_confirmation,
                ).props("color=negative no-caps").set_enabled(not self.absence_in_progress)

        selected_class = self.absence_class
        visible = [
            item
            for item in self.absence_reminders
            if compact(item.journal.get("className")) == selected_class
        ]
        rows = self._absence_rows(visible)
        self.absence_row_map = {
            row["key"]: reminder for row, reminder in zip(rows, visible)
        }
        with ui.card().classes("nit-card w-full p-5 gap-3"):
            ui.label(
                f"{selected_class}: знайдено {len(rows)} непідтверджених позначок. "
                "За замовчуванням вибрано всі."
            ).classes("font-medium")
            self.absence_table = ui.table(
                columns=self._absence_columns(),
                rows=rows,
                row_key="key",
                selection="multiple",
                pagination={"rowsPerPage": 15},
            ).props("flat bordered").classes("nit-data-table absence-data-table w-full")
            self.absence_table.selected = list(rows)
            self.absence_table.update()

    @staticmethod
    def _class_sort_key(name: str) -> tuple[int, str]:
        import re

        match = re.search(r"\d+", name)
        return (int(match.group()) if match else 0, name)

    async def load_absences(self) -> None:
        if self.busy or self.absence_in_progress or not self.auth or self.semester is None:
            return
        self.busy = True
        self.absence_loading = True
        self.absence_load_attempted = True
        self._update_overview_absence_widget()
        self._set_status("Перевірка червоних трикутників…")
        self._refresh_absence_panel()
        try:
            await self._ensure_session()
            schedule_id = self.auth.semester_schedules[self.semester]
            self.absence_reminders = await run.io_bound(
                load_absence_reminders,
                self.auth.client,
                self.semester,
                schedule_id,
            )
            self.absence_loaded = True
            self.absence_action_token = secrets.token_hex(16)
            available_classes = {
                compact(item.journal.get("className")) for item in self.absence_reminders
            }
            if self.absence_class not in available_classes:
                self.absence_class = None
            self._update_overview_absence_widget()
            self._refresh_absence_panel()
        except Exception as exc:
            self.absence_reminders = []
            self.absence_loaded = False
            self.absence_action_token = None
            self._update_overview_absence_widget()
            self._refresh_absence_panel()
            self._show_error(exc)
        finally:
            self.absence_loading = False
            self.busy = False
            self._set_connected_status()
            self._update_overview_absence_widget()
            self._refresh_absence_panel()

    def _absence_class_changed(self, event: Any) -> None:
        selected = str(event.value or "")
        if not selected or selected == self.absence_class:
            return
        self.absence_class = selected
        self.absence_action_token = secrets.token_hex(16)
        self._refresh_absence_panel()

    def open_absence_confirmation(self) -> None:
        if self.absence_table is None or not self.absence_action_token:
            return
        selected_rows = list(self.absence_table.selected or [])
        reminders = [
            self.absence_row_map[row.get("key")]
            for row in selected_rows
            if row.get("key") in self.absence_row_map
        ]
        if not reminders:
            self._notify("Виберіть хоча б один рядок", type="warning")
            return
        action_token = self.absence_action_token
        with ui.dialog() as dialog, ui.card().classes("w-full max-w-lg p-5 gap-4"):
            ui.label("Підтвердження позначок «н»").classes("text-xl font-semibold")
            ui.label(
                f"Буде записано {len(reminders)} позначок у класі {self.absence_class}. "
                "Перед першим запитом буде створено backup."
            ).classes("text-slate-700")
            ui.label(
                "Перед кожною позначкою стан клітинки буде перевірено ще раз."
            ).classes("text-sm text-slate-500")

            async def confirm() -> None:
                dialog.close()
                await self._apply_absences(action_token, reminders)

            with ui.row().classes("w-full justify-end gap-2"):
                ui.button("Скасувати", on_click=dialog.close).props("flat no-caps")
                ui.button("Поставити «н»", icon="person_off", on_click=confirm).props(
                    "color=negative no-caps"
                )
        dialog.open()

    async def _apply_absences(
        self,
        action_token: str,
        reminders: list[AbsenceReminder],
    ) -> None:
        if self.busy or self.absence_in_progress or not self.auth or self.semester is None:
            return
        if action_token != self.absence_action_token:
            self._notify("Список застарів. Перевірте нагадування повторно.", type="warning")
            return
        self.absence_action_token = None
        self.absence_in_progress = True
        self.busy = True
        self._set_status("Запис позначок «н»…")
        self._refresh_absence_panel()
        try:
            await self._ensure_session()
            result = await run.io_bound(
                apply_absence_reminders,
                self.auth.client,
                reminders,
                Path.cwd() / "backups",
            )
            schedule_id = self.auth.semester_schedules[self.semester]
            self.absence_reminders = await run.io_bound(
                load_absence_reminders,
                self.auth.client,
                self.semester,
                schedule_id,
            )
            self.absence_loaded = True
            self._update_overview_absence_widget()
            # Після запису «н» матриця журналу має бути перечитана.
            self.journal_matrix = None
            self.journal_load_attempted = False
            self.absence_action_token = secrets.token_hex(16)
            self._notify(
                f"Поставлено позначок: {result.updated}; пропущено безпечно: "
                f"{result.skipped}. Backup: {result.backup_path.name}",
                type="positive",
                timeout=10000,
            )
            if result.details:
                self._notify("\n".join(result.details[:5]), type="warning", timeout=15000)
        except Exception as exc:
            # Після часткового успіху список обов'язково треба перечитати.
            self.absence_reminders = []
            self.absence_loaded = False
            self.absence_action_token = None
            self._show_error(exc)
        finally:
            self.absence_in_progress = False
            self.busy = False
            self._set_connected_status()
            self._refresh_absence_panel()

    @staticmethod
    def _absence_columns() -> list[dict[str, Any]]:
        return [
            {"name": "date", "label": "Дата", "field": "date", "sortable": True, "align": "center"},
            {"name": "lesson", "label": "Урок", "field": "lesson", "align": "center"},
            {"name": "subject", "label": "Предмет", "field": "subject", "align": "left"},
            {"name": "student", "label": "Учень", "field": "student", "align": "left"},
            {"name": "status", "label": "Дія", "field": "status", "align": "center"},
        ]

    @staticmethod
    def _absence_rows(reminders: list[AbsenceReminder]) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        for index, reminder in enumerate(reminders):
            day = parse_date_value(reminder.journal_day.get("date"))
            key = "|".join(
                (
                    str(reminder.semester),
                    compact(reminder.journal.get("journalId")),
                    compact(reminder.journal_day.get("id")),
                    reminder.user_id,
                )
            )
            rows.append(
                {
                    "key": key or f"absence-{index}",
                    "date": day.strftime("%d.%m.%Y") if day else "—",
                    "lesson": reminder.journal_day.get("lesson") or "—",
                    "subject": compact(reminder.journal.get("subject")) or "—",
                    "student": reminder.student_name or f"…{reminder.user_id[-6:]}",
                    "status": "Поставити «н»",
                }
            )
        return rows

    def _refresh_import_panel(self) -> None:
        if self.import_box is None:
            return
        self.import_box.clear()
        with self.import_box:
            self._render_import_panel()

    def _render_import_panel(self) -> None:
        with ui.card().classes("nit-card w-full p-5 gap-4"):
            with ui.row().classes("w-full items-center justify-between"):
                ui.label("Імпорт із Excel").classes("text-xl font-semibold")
                ui.button(
                    "Завантажити шаблон",
                    icon="download",
                    on_click=self.download_import_template,
                ).props("outline no-caps").set_enabled(self.report is not None)
            ui.label(
                "Спочатку файл лише читається та зіставляється з вибраним журналом. "
                "До окремого підтвердження НІТ не змінюється."
            ).classes("text-sm text-slate-600")

            ui.upload(
                label="Оберіть файл .xlsx",
                on_upload=self._excel_uploaded,
                on_rejected=lambda: self._notify("Потрібен файл .xlsx до 10 МБ", type="warning"),
                auto_upload=True,
                max_file_size=10 * 1024 * 1024,
            ).props("accept=.xlsx flat bordered").classes("w-full")

            if self.excel_name:
                with ui.row().classes("items-center gap-2"):
                    ui.icon("description").classes("text-green-600")
                    ui.label(self.excel_name).classes("font-medium")
                    ui.badge(f"аркуш: {self.excel_sheet}", color="secondary")
                    ui.badge(f"рядків: {len(self.excel_rows)}", color="secondary")
                    ui.badge("є дата" if self.excel_has_date else "без дати", color="green" if self.excel_has_date else "orange")

            with ui.row().classes("responsive-controls w-full items-end gap-4"):
                ui.select(
                    {
                        "homework": "Тільки домашні завдання",
                        "fill": "Заповнити лише порожні поля",
                        "sync": "Синхронізувати тему та ДЗ",
                    },
                    value=self.import_mode,
                    label="Режим імпорту",
                    on_change=self._import_mode_changed,
                ).props("outlined dense options-dense").classes("grow min-w-72")
                ui.checkbox(
                    "Дозволити очищення полів",
                    value=self.import_allow_clear,
                    on_change=self._import_clear_changed,
                )
                ui.button(
                    "Побудувати preview",
                    icon="preview",
                    on_click=self.build_import_preview,
                ).props("no-caps").set_enabled(bool(self.excel_rows) and not self.import_in_progress)

            if self.import_allow_clear:
                with ui.element("div").classes("rounded-lg bg-red-50 border border-red-200 px-4 py-3 w-full"):
                    ui.label(
                        "Увага: порожні клітинки Excel зможуть очистити відповідні поля журналу."
                    ).classes("text-sm font-medium text-red-700")

        preview = self.import_preview
        if preview is None:
            with ui.card().classes("nit-card w-full p-6 items-center"):
                ui.icon("visibility", size="42px").classes("text-slate-400")
                ui.label("Завантажте Excel і натисніть «Побудувати preview».").classes("text-slate-500")
            return

        changed = preview.changed
        warning_count = sum(1 for item in preview.changes if item.warning)
        with ui.card().classes("nit-card w-full p-5 gap-4"):
            with ui.row().classes("w-full items-center justify-between"):
                with ui.column().classes("gap-0"):
                    ui.label("Попередній перегляд — ще нічого не записано").classes("text-lg font-semibold")
                    ui.label(
                        f"Зіставлено: {len(preview.changes)} · буде змінено: {len(changed)} · "
                        f"попереджень: {warning_count}"
                    ).classes("text-sm text-slate-600")
                ui.button(
                    f"Записати {len(changed)} змін",
                    icon="save",
                    on_click=self.open_import_confirmation,
                ).props("no-caps color=negative").set_enabled(bool(changed) and not self.import_in_progress)

            if not preview.has_date:
                with ui.element("div").classes("rounded-lg bg-orange-50 border border-orange-200 px-4 py-3 w-full"):
                    ui.label(
                        "У Excel немає колонки «Дата»: рядки зіставлено з уроками НІТ строго за порядком."
                    ).classes("text-sm font-medium text-orange-800")

            ui.table(
                columns=self._import_columns(),
                rows=self._import_rows(preview),
                row_key="key",
                pagination={"rowsPerPage": 15},
            ).props("flat bordered wrap-cells").classes(
                "nit-data-table import-data-table w-full"
            )

    async def _excel_uploaded(self, event: Any) -> None:
        filename = str(event.file.name or "")
        if not filename.lower().endswith(".xlsx"):
            self._notify("Підтримуються лише файли .xlsx", type="warning")
            return
        upload_directory = Path(tempfile.mkdtemp(prefix="nit_excel_"))
        upload_path = upload_directory / "upload.xlsx"
        try:
            await event.file.save(upload_path)
            sheet, rows, has_date = await run.io_bound(read_excel_upload, upload_path)
        except Exception as exc:
            self._show_error(exc)
            return
        finally:
            upload_path.unlink(missing_ok=True)
            try:
                upload_directory.rmdir()
            except OSError:
                pass

        self.excel_name = Path(filename).name
        self.excel_sheet = sheet
        self.excel_rows = rows
        self.excel_has_date = has_date
        self._invalidate_import_preview()
        self._refresh_import_panel()

    async def download_import_template(self) -> None:
        if self.report is None:
            self._notify("Спочатку виберіть журнал", type="warning")
            return
        journal = self.report.journal
        safe_class = str(journal.get("className") or "class").replace("/", "-")
        safe_subject = str(journal.get("subject") or "subject").replace("/", "-")
        filename = f"шаблон_{safe_class}_{safe_subject}.xlsx"
        path = Path(tempfile.gettempdir()) / f"nit_{datetime.now():%Y%m%d_%H%M%S}_{filename}"
        try:
            await run.io_bound(save_import_template, self.report, path)
            ui.download(str(path), filename=filename)
        except Exception as exc:
            self._show_error(exc)

    def _import_mode_changed(self, event: Any) -> None:
        self.import_mode = str(event.value or "homework")
        self._invalidate_import_preview()
        self._refresh_import_panel()

    def _import_clear_changed(self, event: Any) -> None:
        self.import_allow_clear = bool(event.value)
        self._invalidate_import_preview()
        self._refresh_import_panel()

    def _invalidate_import_preview(self) -> None:
        self.import_preview = None
        self.import_preview_token = None

    def _current_journal(self) -> dict[str, Any] | None:
        return next(
            (item for item in self.journals if str(item.get("journalId")) == self.selected_journal_id),
            None,
        )

    async def build_import_preview(self) -> None:
        if self.busy or self.import_in_progress:
            return
        if not self.auth or self.semester is None:
            self._notify("Спочатку авторизуйтеся", type="warning")
            return
        journal = self._current_journal()
        if journal is None or not self.excel_rows or not self.excel_sheet:
            self._notify("Оберіть журнал і завантажте Excel", type="warning")
            return
        self.busy = True
        self._set_status("Підготовка preview…")
        try:
            await self._ensure_session()
            self.import_preview = await run.io_bound(
                prepare_import_preview,
                self.auth.client,
                self.semester,
                journal,
                self.excel_sheet,
                self.excel_rows,
                self.excel_has_date,
                self.import_mode,
                self.import_allow_clear,
            )
            self.import_preview_token = secrets.token_hex(16)
            self._refresh_import_panel()
        except Exception as exc:
            self._invalidate_import_preview()
            self._refresh_import_panel()
            self._show_error(exc)
        finally:
            self.busy = False
            self._set_connected_status()

    def open_import_confirmation(self) -> None:
        preview = self.import_preview
        token = self.import_preview_token
        if preview is None or not token or not preview.changed:
            return
        with ui.dialog() as dialog, ui.card().classes("w-full max-w-lg p-5 gap-4"):
            ui.label("Підтвердження запису в офіційний журнал").classes("text-xl font-semibold")
            ui.label(
                f"Буде змінено {len(preview.changed)} уроків. Перед першим запитом програма створить backup."
            ).classes("text-slate-700")
            ui.label(
                "Перед кожним записом урок буде повторно звірено з актуальними даними НІТ."
            ).classes("text-sm text-slate-500")

            async def confirm() -> None:
                dialog.close()
                await self._apply_import(token)

            with ui.row().classes("w-full justify-end gap-2"):
                ui.button("Скасувати", on_click=dialog.close).props("flat no-caps")
                ui.button("Записати", icon="save", on_click=confirm).props("color=negative no-caps")
        dialog.open()

    async def _apply_import(self, preview_token: str) -> None:
        if self.import_in_progress or self.busy:
            return
        if (
            not self.auth
            or self.import_preview is None
            or self.import_preview_token != preview_token
        ):
            self._notify("Preview застарів. Побудуйте його повторно.", type="warning")
            return
        preview = self.import_preview
        self.import_preview_token = None
        self.import_in_progress = True
        self.busy = True
        self.calendar_reports = []
        self.calendar_loaded = False
        self.calendar_loading = False
        self.calendar_load_attempted = False
        self._set_status("Запис змін у НІТ…")
        self._refresh_import_panel()
        try:
            await self._ensure_session()
            result = await run.io_bound(
                apply_import_preview,
                self.auth.client,
                self.auth.active_year,
                preview,
                Path.cwd() / "backups",
            )
            self._invalidate_import_preview()
            await self._load_selected_report(render_after=True)
            created_note = f" (створено нових: {result.created})" if result.created else ""
            self._notify(
                f"Записано уроків: {result.updated}{created_note}. Backup: {result.backup_path.name}",
                type="positive",
                timeout=10000,
            )
        except Exception as exc:
            # Навіть якщо частина PUT-запитів уже пройшла, старий preview більше
            # не можна використовувати повторно. Новий preview перечитає НІТ.
            self._invalidate_import_preview()
            self._show_error(exc)
        finally:
            self.import_in_progress = False
            self.busy = False
            self._set_connected_status()
            if self.import_box is not None:
                self._refresh_import_panel()

    @staticmethod
    def _import_columns() -> list[dict[str, Any]]:
        return [
            {"name": "date", "label": "Дата", "field": "date", "sortable": True, "align": "center"},
            {"name": "lesson", "label": "Урок", "field": "lesson", "align": "center"},
            {"name": "old_topic", "label": "Тема НІТ", "field": "old_topic", "align": "left"},
            {"name": "new_topic", "label": "Тема Excel", "field": "new_topic", "align": "left"},
            {"name": "old_homework", "label": "ДЗ НІТ", "field": "old_homework", "align": "left"},
            {"name": "new_homework", "label": "ДЗ Excel", "field": "new_homework", "align": "left"},
            {"name": "status", "label": "Статус", "field": "status", "align": "center"},
            {"name": "warning", "label": "Попередження", "field": "warning", "align": "left"},
        ]

    @staticmethod
    def _import_rows(preview: ImportPreview) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        for index, change in enumerate(preview.changes):
            day = lesson_day(change.lesson)
            old_topic = compact(change.lesson.get("lessonContent"))
            old_homework = compact(change.lesson.get("homework"))
            result.append(
                {
                    "key": str(change.lesson.get("id") or f"row-{index}"),
                    "date": day.strftime("%d.%m.%Y") if day else "—",
                    "lesson": change.lesson.get("lesson") or "—",
                    "old_topic": old_topic or "—",
                    "new_topic": change.new_topic or "—",
                    "old_homework": old_homework or "—",
                    "new_homework": change.new_homework or "—",
                    "status": (
                        "Створення"
                        if change.changed and not change.lesson.get("id")
                        else "Зміна" if change.changed else "Без змін"
                    ),
                    "warning": change.warning or "—",
                }
            )
        return result

    async def _semester_changed(self, event: Any) -> None:
        try:
            semester = int(event.value)
        except (TypeError, ValueError):
            return
        if semester == self.semester or self.busy:
            return
        self.busy = True
        self.semester = semester
        self.journal_matrix = None
        self.journal_loading = False
        self.journal_load_attempted = False
        self.calendar_reports = []
        self.calendar_loaded = False
        self.calendar_loading = False
        self.calendar_load_attempted = False
        self._invalidate_import_preview()
        self.absence_reminders = []
        self.absence_loaded = False
        self.absence_loading = False
        self.absence_load_attempted = False
        self.absence_class = None
        self.absence_action_token = None
        self._set_status(f"Завантаження {semester} семестру…")
        try:
            await self._load_journals(load_first_report=True)
            self.render()
            self._set_connected_status()
        except Exception as exc:
            self._show_error(exc)
        finally:
            self.busy = False
        if self.schedule_view == "month" and not self.calendar_loaded:
            await self.load_calendar()

    async def _journal_changed(self, event: Any) -> None:
        journal_id = str(event.value or "")
        if not journal_id or journal_id == self.selected_journal_id or self.busy:
            return
        self.selected_journal_id = journal_id
        self.journal_matrix = None
        self.journal_loading = False
        self.journal_load_attempted = False
        self._invalidate_import_preview()
        await self._load_selected_report(render_after=True)

    async def _load_selected_report(self, render_after: bool = True) -> None:
        if not self.auth or self.semester is None or not self.selected_journal_id:
            return
        journal = next(
            (item for item in self.journals if str(item.get("journalId")) == self.selected_journal_id),
            None,
        )
        if journal is None:
            return
        was_busy = self.busy
        self.busy = True
        self._set_status(f"Завантаження: {journal.get('className')}…")
        try:
            await self._ensure_session()
            self.report = await run.io_bound(
                load_journal_report,
                self.auth.client,
                self.semester,
                journal,
            )
            days = [day for day in (lesson_day(item) for item in self.report.lessons) if day]
            if days:
                today = date.today()
                selected = today if any((d.year, d.month) == (today.year, today.month) for d in days) else min(days)
                self.calendar_month = selected.replace(day=1)
                self.selected_day = selected
            if render_after:
                self.render()
            self._set_connected_status()
        except Exception as exc:
            self._show_error(exc)
        finally:
            self.busy = was_busy

    async def refresh_current(self) -> None:
        if self.busy:
            return
        self.journal_matrix = None
        self.journal_loading = False
        self.journal_load_attempted = False
        self.calendar_reports = []
        self.calendar_loaded = False
        self.calendar_loading = False
        self.calendar_load_attempted = False
        self._invalidate_import_preview()
        await self._load_selected_report(render_after=True)

    def _refresh_calendar(self) -> None:
        if self.calendar_box is not None:
            self.calendar_box.clear()
            with self.calendar_box:
                self._render_calendar()
        elif self.schedule_view == "month":
            self._refresh_schedule_panel()

    def _render_calendar(self) -> None:
        if not self.calendar_loaded:
            with ui.card().classes("nit-card w-full p-5 gap-4"):
                with ui.row().classes("w-full items-center justify-between"):
                    with ui.column().classes("gap-0"):
                        ui.label("Календар усіх класів").classes("text-xl font-semibold")
                        ui.label(
                            "Уроки з усіх журналів поточного семестру в одному календарі."
                        ).classes("text-sm text-slate-500")
                    if self.calendar_load_attempted and not self.calendar_loading:
                        ui.button(
                            "Спробувати ще раз",
                            icon="refresh",
                            on_click=self.load_calendar,
                        ).props("no-caps")
                with ui.row().classes("items-center gap-3"):
                    if self.calendar_loading or not self.calendar_load_attempted:
                        ui.spinner("dots", size="28px")
                    ui.label(
                        "Збираю уроки з усіх класів…"
                        if self.calendar_loading or not self.calendar_load_attempted
                        else "Не вдалося завантажити календар. Спробуйте ще раз."
                    ).classes("text-slate-500")
            return

        by_day: dict[date, list[tuple[JournalLessonReport, dict[str, Any]]]] = {}
        for report in self.calendar_reports:
            for lesson in report.lessons:
                day = lesson_day(lesson)
                if day:
                    by_day.setdefault(day, []).append((report, lesson))

        with ui.card().classes("nit-card w-full p-4"):
            with ui.row().classes("w-full items-center justify-between"):
                ui.button(icon="chevron_left", on_click=lambda: self._move_month(-1)).props("flat round")
                ui.label(f"{UK_MONTHS[self.calendar_month.month]} {self.calendar_month.year}").classes(
                    "text-xl font-semibold"
                )
                with ui.row().classes("items-center gap-1"):
                    with ui.button(icon="refresh", on_click=self.load_calendar).props("flat round"):
                        ui.tooltip("Оновити календар усіх класів")
                    ui.button(icon="chevron_right", on_click=lambda: self._move_month(1)).props("flat round")

            today = datetime.now(KYIV).date()
            weeks = calendar.Calendar(firstweekday=0).monthdatescalendar(
                self.calendar_month.year, self.calendar_month.month
            )
            with ui.element("div").classes("schedule-scroll mt-3"):
                with ui.element("table").classes("schedule-grid month-grid"):
                    with ui.element("thead"), ui.element("tr"):
                        for weekday in WEEKDAYS:
                            with ui.element("th"):
                                ui.label(weekday).classes("text-xs font-bold")
                    with ui.element("tbody"):
                        for week in weeks:
                            with ui.element("tr"):
                                for day in week:
                                    if day.month != self.calendar_month.month:
                                        with ui.element("td").classes("month-day month-outside"):
                                            pass
                                        continue
                                    day_entries = by_day.get(day, [])
                                    css = ["month-day"]
                                    if day == self.selected_day:
                                        css.append("month-selected")
                                    if day == today:
                                        css.append("month-today")
                                    with ui.element("td").classes(" ".join(css)).on(
                                        "click", lambda _, selected=day: self._select_day(selected)
                                    ):
                                        with ui.row().classes("w-full items-center justify-between gap-1"):
                                            ui.label(str(day.day)).classes("month-date")
                                            if day_entries:
                                                ui.label(f"{len(day_entries)} ур.").classes(
                                                    "text-xs text-slate-500"
                                                )
                                        if day_entries:
                                            class_counts: dict[str, int] = {}
                                            for report, _ in day_entries:
                                                class_name = compact(report.journal.get("className")) or "Без класу"
                                                class_counts[class_name] = class_counts.get(class_name, 0) + 1
                                            with ui.element("div").classes("month-class-list"):
                                                for class_name in sorted(class_counts, key=self._class_sort_key):
                                                    count = class_counts[class_name]
                                                    text = class_name if count == 1 else f"{class_name} ×{count}"
                                                    ui.label(text).classes("month-class-chip")

        selected_entries = by_day.get(self.selected_day, []) if self.selected_day else []
        with ui.card().classes("nit-card w-full p-4"):
            title = self.selected_day.strftime("%d.%m.%Y") if self.selected_day else "Оберіть дату"
            ui.label(f"Уроки за {title}").classes("text-lg font-semibold")
            if selected_entries:
                ui.table(
                    columns=self._calendar_columns(),
                    rows=self._calendar_rows(selected_entries),
                    row_key="key",
                    pagination={"rowsPerPage": 0},
                ).props("flat bordered hide-pagination wrap-cells").classes(
                    "nit-data-table calendar-detail-table w-full"
                )
            else:
                ui.label("На цю дату уроків немає.").classes("text-slate-500")

    async def load_calendar(self) -> None:
        if self.busy or not self.auth or self.semester is None:
            return
        self.busy = True
        self.calendar_loading = True
        self.calendar_load_attempted = True
        self._set_status("Завантаження календаря всіх класів…")
        self._refresh_calendar()
        try:
            await self._ensure_session()
            self.calendar_reports = await run.io_bound(
                load_calendar_reports,
                self.auth.client,
                self.semester,
                self.journals,
                self.report,
            )
            self.calendar_loaded = True
            days = [
                day
                for report in self.calendar_reports
                for day in (lesson_day(item) for item in report.lessons)
                if day
            ]
            if days:
                today = date.today()
                same_month = [day for day in days if (day.year, day.month) == (today.year, today.month)]
                if today in days:
                    selected = today
                elif same_month:
                    selected = min(same_month, key=lambda day: abs((day - today).days))
                else:
                    selected = min(days)
                self.calendar_month = selected.replace(day=1)
                self.selected_day = selected
        except Exception as exc:
            self.calendar_reports = []
            self.calendar_loaded = False
            self._show_error(exc)
        finally:
            self.calendar_loading = False
            self.busy = False
            self._set_connected_status()
            self._refresh_calendar()

    def _move_month(self, delta: int) -> None:
        year = self.calendar_month.year
        month = self.calendar_month.month + delta
        if month < 1:
            year -= 1
            month = 12
        elif month > 12:
            year += 1
            month = 1
        self.calendar_month = date(year, month, 1)
        self.selected_day = None
        self._refresh_calendar()

    def _select_day(self, selected: date) -> None:
        self.selected_day = selected
        self._refresh_calendar()

    @staticmethod
    def _calendar_columns() -> list[dict[str, Any]]:
        return [
            {"name": "class", "label": "Клас", "field": "class", "sortable": True, "align": "center"},
            {"name": "subject", "label": "Предмет", "field": "subject", "align": "left"},
            {"name": "lesson", "label": "Урок", "field": "lesson", "align": "center"},
            {"name": "topic", "label": "Тема", "field": "topic", "align": "left"},
            {"name": "homework", "label": "Домашнє завдання", "field": "homework", "align": "left"},
            {"name": "status", "label": "Статус", "field": "status", "align": "center"},
        ]

    @classmethod
    def _calendar_rows(
        cls,
        entries: list[tuple[JournalLessonReport, dict[str, Any]]],
    ) -> list[dict[str, Any]]:
        sorted_entries = sorted(
            entries,
            key=lambda entry: (
                cls._class_sort_key(compact(entry[0].journal.get("className"))),
                compact(entry[0].journal.get("subject")),
                int(entry[1].get("lesson") or 0),
            ),
        )
        rows: list[dict[str, Any]] = []
        for index, (report, lesson) in enumerate(sorted_entries):
            journal_id = compact(report.journal.get("journalId"))
            lesson_id = compact(lesson.get("journalDayId") or lesson.get("id"))
            rows.append(
                {
                    "key": f"{journal_id}:{lesson_id or index}",
                    "class": compact(report.journal.get("className")) or "—",
                    "subject": compact(report.journal.get("subject")) or "—",
                    "lesson": lesson.get("lesson") or "—",
                    "topic": lesson.get("lessonContent") or "—",
                    "homework": lesson.get("homework") or "—",
                    "status": "Заповнено" if lesson.get("isFilled") else "Не заповнено",
                }
            )
        return rows

    @staticmethod
    def _lesson_columns() -> list[dict[str, Any]]:
        return [
            {"name": "date", "label": "Дата", "field": "date", "sortable": True, "align": "center"},
            {"name": "lesson", "label": "Урок", "field": "lesson", "align": "center"},
            {"name": "topic", "label": "Тема", "field": "topic", "align": "left"},
            {"name": "homework", "label": "Домашнє завдання", "field": "homework", "align": "left"},
            {"name": "status", "label": "Статус", "field": "status", "align": "center"},
        ]

    @staticmethod
    def _lesson_rows(lessons: list[dict[str, Any]]) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        for index, item in enumerate(lessons):
            day = lesson_day(item)
            rows.append(
                {
                    "key": str(item.get("journalDayId") or item.get("id") or index),
                    "date": day.strftime("%d.%m.%Y") if day else "—",
                    "lesson": item.get("lesson") or "—",
                    "topic": item.get("lessonContent") or "—",
                    "homework": item.get("homework") or "—",
                    "status": "Заповнено" if item.get("isFilled") else "Не заповнено",
                }
            )
        return rows

    async def export_report(self, extension: str) -> None:
        if not self.report or not self.auth:
            self._notify("Спочатку виберіть журнал", type="warning")
            return
        journal = self.report.journal
        safe_class = str(journal.get("className") or "class").replace("/", "-")
        safe_subject = str(journal.get("subject") or "subject").replace("/", "-")
        filename = f"уроки_{safe_class}_{safe_subject}.{extension}"
        path = Path(tempfile.gettempdir()) / f"nit_{datetime.now():%Y%m%d_%H%M%S}_{filename}"
        try:
            await run.io_bound(
                save_lesson_reports,
                [self.report],
                path,
                self.auth.active_year,
            )
            ui.download(str(path), filename=filename)
        except Exception as exc:
            self._show_error(exc)

    def _set_connected_status(self) -> None:
        if not self.auth:
            return
        year = self.auth.active_year
        self._set_status(f"{year.get('yearStart')}/{year.get('yearEnd')} · підключено")

    def _set_status(self, text: str) -> None:
        if self.status_label is not None:
            self.status_label.text = text
            self.status_label.update()

    def _notify(self, message: str, **kwargs: Any) -> None:
        # Async-обробник може перемалювати й видалити контейнер, з якого був
        # запущений. Контекст клієнта залишається стабільним і дозволяє безпечно
        # показувати повідомлення після refresh/render.
        with self.client:
            ui.notify(message, **kwargs)

    def _show_error(self, exc: Exception) -> None:
        message = str(exc) if isinstance(exc, NitError) else f"Неочікувана помилка: {exc}"
        self._notify(message, type="negative", timeout=8000, close_button=True)


@app.get("/health")
def health() -> dict[str, str]:
    """Легка локальна перевірка для Windows-лаунчера та моніторингу."""
    return {"status": "ok", "app": "NIT2"}


@ui.page("/")
def index() -> None:
    NitWebApp()


if __name__ in {"__main__", "__mp_main__"}:
    host = os.getenv("NIT_HOST", "127.0.0.1").strip() or "127.0.0.1"
    try:
        port = int(os.getenv("NIT_PORT", "8080"))
    except ValueError:
        raise SystemExit("NIT_PORT має бути числом від 1 до 65535")
    if not 1 <= port <= 65535:
        raise SystemExit("NIT_PORT має бути числом від 1 до 65535")
    show_browser = os.getenv("NIT_SHOW_BROWSER", "1").strip().lower() not in {
        "0", "false", "no", "off",
    }
    ui.run(
        host=host,
        port=port,
        title="NIT² by FIZYK",
        favicon=Path(__file__).with_name("image.png"),
        dark=True,
        show=show_browser,
        reload=False,
        storage_secret=_STORAGE_SECRET,
    )
