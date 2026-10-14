import json
import tempfile
import unittest
from pathlib import Path

from write_macros import find_changed_vba_targets, sha256_vba_source
from workbook_release import (
    promote_workbook,
    sha256_file,
    state_path,
    test_results_path as get_test_results_path,
)


class PromoteWorkbookTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        root = Path(self.temp_dir.name)
        self.production_dir = root / "ExcelFiles"
        self.staging_dir = root / "StagedExcelFiles"
        self.test_runs_dir = root / ".test_runs"
        self.production_dir.mkdir()
        self.staging_dir.mkdir()
        self.test_runs_dir.mkdir()

        self.production_path = self.production_dir / "Sample.xlsm"
        self.staged_path = self.staging_dir / self.production_path.name
        self.production_path.write_bytes(b"production version")
        self.staged_path.write_bytes(b"tested staged version")
        self.scenario_path = root / "scenario.json"
        self.scenario_path.write_text('{"name": "passing test"}', encoding="utf-8")
        self.metadata = {
            "production_path": str(self.production_path.resolve()),
            "production_sha256": sha256_file(self.production_path),
            "staged_path": str(self.staged_path.resolve()),
            "staged_sha256": sha256_file(self.staged_path),
        }
        with open(state_path(self.staged_path), "w", encoding="utf-8") as file:
            json.dump(self.metadata, file)
        self.results = {
            "target_path": str(self.staged_path.resolve()),
            "target_sha256": sha256_file(self.staged_path),
            "scenario_path": str(self.scenario_path.resolve()),
            "scenario_sha256": sha256_file(self.scenario_path),
            "passed": True,
        }
        with open(get_test_results_path(self.staged_path, self.test_runs_dir), "w", encoding="utf-8") as file:
            json.dump(self.results, file)

    def tearDown(self):
        self.temp_dir.cleanup()

    def promote(self):
        promote_workbook(
            self.production_path.name,
            self.production_dir,
            self.staging_dir,
            self.test_runs_dir,
        )

    def write_results(self):
        with open(get_test_results_path(self.staged_path, self.test_runs_dir), "w", encoding="utf-8") as file:
            json.dump(self.results, file)

    def test_promotes_passing_candidate_and_preserves_backup(self):
        self.promote()

        self.assertEqual(self.production_path.read_bytes(), b"tested staged version")
        backups = list((self.production_dir / "Backups").glob("Sample_*.xlsm"))
        self.assertEqual(len(backups), 1)
        self.assertEqual(backups[0].read_bytes(), b"production version")

    def test_rejects_failed_test_evidence(self):
        self.results["passed"] = False
        self.write_results()

        with self.assertRaisesRegex(RuntimeError, "passing scenario-based test"):
            self.promote()

        self.assertEqual(self.production_path.read_bytes(), b"production version")

    def test_rejects_candidate_changed_after_testing(self):
        self.staged_path.write_bytes(b"changed after test")

        with self.assertRaisesRegex(RuntimeError, "changed after it was built"):
            self.promote()

        self.assertEqual(self.production_path.read_bytes(), b"production version")

    def test_rejects_production_changed_after_staging(self):
        self.production_path.write_bytes(b"changed after staging")

        with self.assertRaisesRegex(RuntimeError, "changed after staging"):
            self.promote()

        self.assertEqual(self.production_path.read_bytes(), b"changed after staging")

    def test_rejects_scenario_changed_after_testing(self):
        self.scenario_path.write_text('{"name": "changed test"}', encoding="utf-8")

        with self.assertRaisesRegex(RuntimeError, "scenario changed or is missing"):
            self.promote()

        self.assertEqual(self.production_path.read_bytes(), b"production version")


class FindChangedVbaTargetsTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        root = Path(self.temp_dir.name)
        self.vba_dir = root / "VBAFiles"
        self.production_dir = root / "ExcelFiles"
        self.staging_dir = root / "StagedExcelFiles"
        self.production_dir.mkdir()
        self.staging_dir.mkdir()

    def tearDown(self):
        self.temp_dir.cleanup()

    def add_workbook(self, name, source_content):
        source_dir = self.vba_dir / name
        modules_dir = source_dir / "modules"
        modules_dir.mkdir(parents=True, exist_ok=True)
        source = modules_dir / "Module.bas"
        source.write_text(source_content, encoding="utf-8")
        workbook = self.production_dir / f"{name}.xlsm"
        workbook.write_bytes(f"{name} workbook".encode("utf-8"))
        return workbook, source_dir

    def write_staging_metadata(self, workbook, source_dir, source_hash=None):
        staged_workbook = self.staging_dir / workbook.name
        metadata_path = staged_workbook.with_suffix(".release.json")
        metadata = {
            "vba_source_sha256": source_hash or sha256_vba_source(source_dir),
        }
        metadata_path.write_text(json.dumps(metadata), encoding="utf-8")

    def find_changes(self):
        return find_changed_vba_targets(
            self.vba_dir, self.production_dir, self.staging_dir
        )

    def test_finds_all_workbooks_with_sources_changed_since_staging(self):
        first_workbook, first_source_dir = self.add_workbook("First", "Sub First()\nEnd Sub\n")
        second_workbook, second_source_dir = self.add_workbook("Second", "Sub Second()\nEnd Sub\n")
        self.add_workbook("Unchanged", "Sub Unchanged()\nEnd Sub\n")
        self.write_staging_metadata(first_workbook, first_source_dir, "old-first-hash")
        self.write_staging_metadata(second_workbook, second_source_dir, "old-second-hash")
        unchanged_workbook = self.production_dir / "Unchanged.xlsm"
        self.write_staging_metadata(
            unchanged_workbook,
            self.vba_dir / "Unchanged",
        )

        changed = self.find_changes()

        self.assertEqual(
            changed,
            [
                (first_workbook, first_source_dir),
                (second_workbook, second_source_dir),
            ],
        )

    def test_ignores_unchanged_workbook_and_stages_untracked_baseline(self):
        unchanged_workbook, unchanged_source_dir = self.add_workbook(
            "Unchanged", "Sub Unchanged()\nEnd Sub\n"
        )
        new_workbook, new_source_dir = self.add_workbook("New", "Sub NewMacro()\nEnd Sub\n")
        self.write_staging_metadata(unchanged_workbook, unchanged_source_dir)

        changed = self.find_changes()

        self.assertEqual(changed, [(new_workbook, new_source_dir)])

    def test_source_hash_changes_when_a_module_is_added_or_modified(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source_dir = Path(temp_dir) / "modules"
            source_dir.mkdir()
            module = source_dir / "Module.bas"
            module.write_text("Sub First()\nEnd Sub\n", encoding="utf-8")
            original_hash = sha256_vba_source(source_dir.parent)

            module.write_text("Sub Changed()\nEnd Sub\n", encoding="utf-8")
            modified_hash = sha256_vba_source(source_dir.parent)
            (source_dir / "Second.bas").write_text("Sub Second()\nEnd Sub\n", encoding="utf-8")
            added_hash = sha256_vba_source(source_dir.parent)

            self.assertNotEqual(original_hash, modified_hash)
            self.assertNotEqual(modified_hash, added_hash)


if __name__ == "__main__":
    unittest.main()
