import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from write_macros import (
    ADO_TYPE_LIBRARY_GUID,
    ADO_TYPE_LIBRARY_MAJOR,
    ADO_TYPE_LIBRARY_MINOR,
    ensure_ado_reference,
    find_changed_vba_targets,
    sha256_vba_source,
    stage_single_workbook,
)
from test_run_artifacts import create_run_directory
from workbook_release import (
    promote_workbook,
    sha256_file,
    state_path,
    test_workbook,
    test_results_path as get_test_results_path,
)


class EnsureAdoReferenceTests(unittest.TestCase):
    class References(list):
        def AddFromGuid(self, guid, major, minor):
            self.added = (guid, major, minor)
            self.append(SimpleNamespace(Name="ADODB"))

    def test_adds_reference_for_early_bound_ado_sources_only_once(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source_dir = Path(temp_dir) / "VBAFiles" / "Database"
            modules_dir = source_dir / "modules"
            modules_dir.mkdir(parents=True)
            (modules_dir / "Database.bas").write_text(
                "Dim connection As ADODB.Connection\n", encoding="utf-8"
            )
            references = self.References()
            vb_project = SimpleNamespace(References=references)

            self.assertTrue(ensure_ado_reference(vb_project, source_dir))
            self.assertEqual(
                references.added,
                (
                    ADO_TYPE_LIBRARY_GUID,
                    ADO_TYPE_LIBRARY_MAJOR,
                    ADO_TYPE_LIBRARY_MINOR,
                ),
            )
            self.assertFalse(ensure_ado_reference(vb_project, source_dir))
            self.assertEqual(len(references), 1)

    def test_does_not_add_reference_when_source_does_not_use_ado(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source_dir = Path(temp_dir) / "VBAFiles" / "Sample"
            modules_dir = source_dir / "modules"
            modules_dir.mkdir(parents=True)
            (modules_dir / "Module.bas").write_text(
                "Sub Run()\nEnd Sub\n", encoding="utf-8"
            )
            references = self.References()
            vb_project = SimpleNamespace(References=references)

            self.assertFalse(ensure_ado_reference(vb_project, source_dir))
            self.assertEqual(len(references), 0)


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
        self.run_directory = create_run_directory(
            self.test_runs_dir,
            self.staged_path,
            "passing test",
            datetime(2026, 10, 15, 21, 30, 0),
        )
        self.run_id = self.run_directory.relative_to(self.test_runs_dir).as_posix()
        self.scenario_path = root / "scenario.json"
        self.scenario_path.write_text('{"name": "passing test"}', encoding="utf-8")
        self.metadata = {
            "production_path": str(self.production_path.resolve()),
            "production_sha256": sha256_file(self.production_path),
            "staged_path": str(self.staged_path.resolve()),
            "staged_sha256": sha256_file(self.staged_path),
            "test_run_id": self.run_id,
            "test_scenario_sha256": sha256_file(self.scenario_path),
        }
        with open(state_path(self.staged_path), "w", encoding="utf-8") as file:
            json.dump(self.metadata, file)
        self.results = {
            "target_path": str(self.staged_path.resolve()),
            "target_sha256": sha256_file(self.staged_path),
            "scenario_path": str(self.scenario_path.resolve()),
            "scenario_sha256": sha256_file(self.scenario_path),
            "passed": True,
            "run_id": self.run_id,
        }
        with open(
            get_test_results_path(self.staged_path, self.test_runs_dir, self.run_id),
            "w",
            encoding="utf-8",
        ) as file:
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
        with open(
            get_test_results_path(self.staged_path, self.test_runs_dir, self.run_id),
            "w",
            encoding="utf-8",
        ) as file:
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

    def test_rejects_test_result_with_mismatched_run_id(self):
        self.results["run_id"] = "another/run"
        self.write_results()

        with self.assertRaisesRegex(RuntimeError, "does not match the release evidence"):
            self.promote()

        self.assertEqual(self.production_path.read_bytes(), b"production version")

    def test_rejects_test_run_id_that_escapes_archive(self):
        with self.assertRaisesRegex(ValueError, "invalid test-run identifier"):
            get_test_results_path(
                self.staged_path,
                self.test_runs_dir,
                Path("..") / "outside",
            )

    def test_rejects_release_metadata_with_mismatched_scenario_hash(self):
        self.metadata["test_scenario_sha256"] = "different-scenario-hash"
        state_path(self.staged_path).write_text(
            json.dumps(self.metadata), encoding="utf-8"
        )

        with self.assertRaisesRegex(RuntimeError, "scenario hash"):
            self.promote()

        self.assertEqual(self.production_path.read_bytes(), b"production version")


class TestWorkbookArchiveEvidenceTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        root = Path(self.temp_dir.name)
        self.production_path = root / "ExcelFiles" / "Sample.xlsm"
        self.staged_path = root / "StagedExcelFiles" / "Sample.xlsm"
        self.test_runs_dir = root / ".test_runs"
        self.scenario_path = root / "scenario.json"
        self.production_path.parent.mkdir()
        self.staged_path.parent.mkdir()
        self.test_runs_dir.mkdir()
        self.production_path.write_bytes(b"production")
        self.staged_path.write_bytes(b"staged")
        self.scenario_path.write_text('{"name": "Sample scenario"}', encoding="utf-8")
        self.metadata = {
            "production_path": str(self.production_path.resolve()),
            "production_sha256": sha256_file(self.production_path),
            "staged_path": str(self.staged_path.resolve()),
            "staged_sha256": sha256_file(self.staged_path),
        }
        state_path(self.staged_path).write_text(
            json.dumps(self.metadata), encoding="utf-8"
        )

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_test_workbook_archives_result_and_records_run_pointer(self):
        def run_test(command, cwd, check):
            run_dir = Path(command[command.index("--artifact-dir") + 1])
            timestamp = command[command.index("--run-timestamp") + 1]
            run_id = run_dir.relative_to(self.test_runs_dir).as_posix()
            results = {
                "target": self.staged_path.name,
                "target_path": str(self.staged_path.resolve()),
                "target_sha256": sha256_file(self.staged_path),
                "scenario_path": str(self.scenario_path.resolve()),
                "scenario_sha256": sha256_file(self.scenario_path),
                "timestamp": timestamp,
                "scenario": "Sample scenario",
                "passed": True,
                "run_id": run_id,
            }
            (run_dir / "results.json").write_text(json.dumps(results), encoding="utf-8")
            return SimpleNamespace(returncode=0)

        with mock.patch(
            "workbook_release.resolve_staged_workbook",
            return_value=(self.production_path, self.staged_path),
        ), mock.patch(
            "workbook_release.TEST_RUNS_DIR", self.test_runs_dir
        ), mock.patch(
            "workbook_release.subprocess.run", side_effect=run_test
        ):
            test_workbook(self.staged_path.name, str(self.scenario_path))

        updated_metadata = json.loads(state_path(self.staged_path).read_text(encoding="utf-8"))
        run_id = updated_metadata["test_run_id"]
        result_path = get_test_results_path(self.staged_path, self.test_runs_dir, run_id)
        self.assertTrue(result_path.is_file())
        self.assertEqual(json.loads(result_path.read_text(encoding="utf-8"))["passed"], True)
        self.assertEqual(updated_metadata["test_scenario_sha256"], sha256_file(self.scenario_path))

    def test_failed_rerun_replaces_old_passing_evidence(self):
        old_run = create_run_directory(
            self.test_runs_dir,
            self.staged_path,
            "old scenario",
            datetime(2026, 10, 14, 21, 30, 0),
        )
        self.metadata["test_run_id"] = old_run.relative_to(self.test_runs_dir).as_posix()
        state_path(self.staged_path).write_text(json.dumps(self.metadata), encoding="utf-8")

        def run_failed_test(command, cwd, check):
            run_dir = Path(command[command.index("--artifact-dir") + 1])
            timestamp = command[command.index("--run-timestamp") + 1]
            run_id = run_dir.relative_to(self.test_runs_dir).as_posix()
            results = {
                "target": self.staged_path.name,
                "target_path": str(self.staged_path.resolve()),
                "target_sha256": sha256_file(self.staged_path),
                "scenario_path": str(self.scenario_path.resolve()),
                "scenario_sha256": sha256_file(self.scenario_path),
                "timestamp": timestamp,
                "scenario": "Sample scenario",
                "passed": False,
                "run_id": run_id,
            }
            (run_dir / "results.json").write_text(json.dumps(results), encoding="utf-8")
            return SimpleNamespace(returncode=1)

        with mock.patch(
            "workbook_release.resolve_staged_workbook",
            return_value=(self.production_path, self.staged_path),
        ), mock.patch(
            "workbook_release.TEST_RUNS_DIR", self.test_runs_dir
        ), mock.patch(
            "workbook_release.subprocess.run", side_effect=run_failed_test
        ):
            with self.assertRaises(SystemExit):
                test_workbook(self.staged_path.name, str(self.scenario_path))

        updated_metadata = json.loads(state_path(self.staged_path).read_text(encoding="utf-8"))
        self.assertNotEqual(updated_metadata["test_run_id"], self.metadata["test_run_id"])
        failed_result = get_test_results_path(
            self.staged_path,
            self.test_runs_dir,
            updated_metadata["test_run_id"],
        )
        self.assertFalse(json.loads(failed_result.read_text(encoding="utf-8"))["passed"])


class StagePreservesArchivedEvidenceTests(unittest.TestCase):
    def test_staging_keeps_previous_test_json(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            production_dir = root / "ExcelFiles"
            source_dir = root / "VBAFiles" / "Sample"
            staging_dir = root / "StagedExcelFiles"
            test_runs_dir = root / ".test_runs"
            production_dir.mkdir()
            source_dir.mkdir(parents=True)
            test_runs_dir.mkdir()
            production_path = production_dir / "Sample.xlsm"
            production_path.write_bytes(b"production")
            test_result = test_runs_dir / "Sample.xlsm_test_results.json"
            test_result.write_text('{"passed": true}', encoding="utf-8")

            with mock.patch("write_macros.PRODUCTION_DIR", production_dir), mock.patch(
                "write_macros.write_vba_to_workbook", return_value=(0, 0)
            ):
                stage_single_workbook(production_path, source_dir, staging_dir)

            self.assertEqual(test_result.read_text(encoding="utf-8"), '{"passed": true}')


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
