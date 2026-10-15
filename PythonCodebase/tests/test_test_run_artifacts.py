import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from test_run_artifacts import (
    archive_legacy_results,
    archive_test_run,
    create_run_directory,
    rebuild_indexes,
)


class TestRunArtifactsTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name) / ".test_runs"
        self.root.mkdir()
        self.timestamp = datetime(2026, 10, 15, 21, 30, 0)
        self.results = {
            "target": "R1 Inbound and Sampling.xlsm",
            "target_path": "ExcelFiles/R1 Inbound and Sampling.xlsm",
            "target_sha256": "workbook-hash",
            "timestamp": self.timestamp.isoformat(),
            "scenario": "R1 PathResolver startup loads",
            "scenario_path": "TestScenarios/R1 Inbound and Sampling.json",
            "scenario_sha256": "scenario-hash",
            "passed": True,
            "duration_seconds": 12.5,
            "errors": [],
        }

    def tearDown(self):
        self.temp_dir.cleanup()

    @staticmethod
    def render_markdown(results):
        return f"# {results['target']}\n\nRun: {results.get('run_id', '')}\n"

    def test_archives_json_markdown_and_date_and_workbook_indexes(self):
        run_dir = create_run_directory(
            self.root,
            self.results["target_path"],
            self.results["scenario"],
            self.timestamp,
        )

        results_path, report_path = archive_test_run(
            self.root,
            self.results.copy(),
            self.render_markdown,
            run_dir,
        )

        self.assertEqual(results_path.name, "results.json")
        self.assertEqual(report_path.name, "report.md")
        self.assertTrue(results_path.is_file())
        self.assertTrue(report_path.is_file())
        self.assertEqual(
            json.loads(results_path.read_text(encoding="utf-8"))["run_id"],
            run_dir.relative_to(self.root).as_posix(),
        )
        self.assertTrue((self.root / "index.md").is_file())
        self.assertTrue((self.root / "2026-10-15" / "index.md").is_file())
        self.assertTrue((self.root / "by_workbook" / "index.md").is_file())
        workbook_indexes = [
            path for path in (self.root / "by_workbook").glob("*.md")
            if path.name != "index.md"
        ]
        self.assertEqual(len(workbook_indexes), 1)
        self.assertIn("R1 Inbound and Sampling", workbook_indexes[0].read_text(encoding="utf-8"))
        root_index = (self.root / "index.md").read_text(encoding="utf-8")
        self.assertIn("Browse by workbook", root_index)
        self.assertIn("2026-10-15/index.md", root_index)
        self.assertIn("R1 PathResolver startup loads", root_index)

    def test_migrates_legacy_json_without_changing_original(self):
        legacy_path = self.root / "R1 Inbound and Sampling.xlsm_test_results.json"
        original = json.dumps(self.results).encode("utf-8")
        legacy_path.write_bytes(original)

        archived = archive_legacy_results(self.root, self.render_markdown)
        rebuild_indexes(self.root)
        archived_results = archived[0] / "results.json"

        self.assertEqual(legacy_path.read_bytes(), original)
        self.assertEqual(archived_results.read_bytes(), original)
        self.assertTrue((archived[0] / "report.md").is_file())
        self.assertIn("R1 Inbound and Sampling", (self.root / "index.md").read_text(encoding="utf-8"))

        second_archive = archive_legacy_results(self.root, self.render_markdown)
        self.assertEqual(second_archive, archived)

    def test_archive_rejects_artifact_directory_outside_root(self):
        outside = Path(self.temp_dir.name) / "outside"

        with self.assertRaisesRegex(ValueError, "inside the archive root"):
            archive_test_run(
                self.root,
                self.results.copy(),
                self.render_markdown,
                outside,
            )


if __name__ == "__main__":
    unittest.main()
