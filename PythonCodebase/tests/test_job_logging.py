"""Tests for the repository-local job logger."""

import importlib.util
import csv
from datetime import datetime
from pathlib import Path
import tempfile
import unittest


LOGGER_PATH = (
    Path(__file__).resolve().parents[2]
    / ".github"
    / "skills"
    / "job-logging"
    / "scripts"
    / "job_log.py"
)
SPEC = importlib.util.spec_from_file_location("job_log", LOGGER_PATH)
job_log = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(job_log)


class JobLoggingTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.log_directory = Path(self.temp_dir.name)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_timer_can_stop_on_a_different_day_and_records_elapsed(self):
        started_at = datetime(2026, 10, 13, 23, 30, 0)
        stopped_at = datetime(2026, 10, 14, 1, 0, 0)
        start_path, _ = job_log.start_job(
            "Testing staged workbook", self.log_directory, started_at
        )

        status = job_log.show_status(self.log_directory, stopped_at)
        description, duration, stop_path = job_log.stop_job(self.log_directory, stopped_at)

        self.assertEqual(description, "Testing staged workbook")
        self.assertEqual(duration, "01:30:00")
        self.assertEqual(start_path.name, "JobLog_20261013.csv")
        self.assertEqual(stop_path, start_path)
        self.assertEqual(status[2], "01:30:00")
        with start_path.open("r", encoding="utf-8-sig", newline="") as file:
            rows = list(csv.DictReader(file))
        self.assertEqual(rows[0]["End"], "2026-10-14T01:00:00")
        self.assertEqual(rows[0]["Duration"], "01:30:00")
        self.assertEqual(rows[0]["DurationSeconds"], "5400")
        self.assertEqual(rows[0]["DurationHours"], "1.500000")
        self.assertIsNone(job_log.show_status(self.log_directory, stopped_at))

    def test_start_rejects_a_second_active_timer(self):
        job_log.start_job("First job", self.log_directory, datetime(2026, 10, 13, 9))

        with self.assertRaisesRegex(ValueError, "already running"):
            job_log.start_job("Second job", self.log_directory, datetime(2026, 10, 13, 10))

    def test_manual_duration_accepts_units_and_appends_to_daily_log(self):
        timestamp = datetime(2026, 10, 13, 11, 15, 0)

        duration, path = job_log.log_manual(
            "Updating documentation", "1h30m", self.log_directory, timestamp
        )

        self.assertEqual(duration, "01:30:00")
        self.assertEqual(path.name, "JobLog_20261013.csv")
        with path.open("r", encoding="utf-8-sig", newline="") as file:
            rows = list(csv.DictReader(file))
        self.assertEqual(rows[0]["Type"], "Manual")
        self.assertEqual(rows[0]["DurationSeconds"], "5400")
        self.assertEqual(rows[0]["Description"], "Updating documentation")

    def test_migrates_legacy_text_and_preserves_the_original(self):
        legacy_path = self.log_directory / "JobLog_20261013.txt"
        legacy_text = (
            "START | timer123 | 2026-10-13T23:16:01 | testing job logging\n"
            "STOP | timer123 | 2026-10-13T23:16:53 | elapsed=00:00:52 | testing job logging\n"
        )
        legacy_path.write_text(legacy_text, encoding="utf-8")

        status = job_log.show_status(
            self.log_directory, datetime(2026, 10, 13, 23, 20, 0)
        )
        csv_path = self.log_directory / "JobLog_20261013.csv"

        self.assertIsNone(status)
        self.assertEqual(legacy_path.read_text(encoding="utf-8"), legacy_text)
        with csv_path.open("r", encoding="utf-8-sig", newline="") as file:
            rows = list(csv.DictReader(file))
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["DurationSeconds"], "52")
        self.assertEqual(rows[0]["Description"], "testing job logging")

    def test_duration_requires_positive_value_and_units(self):
        for invalid_duration in ("", "1.5", "0m", "abc"):
            with self.subTest(invalid_duration=invalid_duration):
                with self.assertRaises(ValueError):
                    job_log.parse_duration(invalid_duration)


if __name__ == "__main__":
    unittest.main()
