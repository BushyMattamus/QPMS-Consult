"""Build, test, and promote Excel workbooks through a staging directory."""

import argparse
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from write_macros import (
    PRODUCTION_DIR,
    STAGING_DIR,
    resolve_production_workbook,
    sha256_file,
    stage_single_workbook,
)
from test_run_artifacts import create_run_directory


PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEST_RUNS_DIR = PROJECT_ROOT / ".test_runs"
TEST_RUNNER = Path(__file__).resolve().parent / "excel_test_runner.py"


def resolve_staged_workbook(argument, production_dir=PRODUCTION_DIR, staging_dir=STAGING_DIR):
    """Resolve a workbook name or path to its staged workbook."""
    production_path = resolve_production_workbook(argument, production_dir)
    staged_path = Path(staging_dir) / production_path.name
    if not staged_path.is_file():
        raise FileNotFoundError(f"Staged workbook not found; run the stage command first: {staged_path}")
    return production_path, staged_path.resolve()


def state_path(staged_path):
    return staged_path.with_suffix(".release.json")


def test_results_path(staged_path, test_runs_dir=TEST_RUNS_DIR, run_id=None):
    if not run_id:
        raise ValueError("Staging metadata does not identify a test run.")

    root = Path(test_runs_dir).resolve()
    relative_run_id = Path(run_id)
    if relative_run_id.is_absolute() or ".." in relative_run_id.parts:
        raise ValueError("Staging metadata contains an invalid test-run identifier.")
    result_path = (root / relative_run_id / "results.json").resolve()
    if os.path.commonpath([str(root), str(result_path)]) != str(root):
        raise ValueError("Staging metadata points outside the test-run archive.")
    return result_path


def _write_release_metadata(metadata_file, metadata):
    temporary_path = metadata_file.with_name(
        f".{metadata_file.name}.{uuid4().hex}.tmp"
    )
    try:
        temporary_path.write_text(json.dumps(metadata, indent=2), encoding="utf-8")
        os.replace(temporary_path, metadata_file)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def stage_workbook(argument):
    production_path = resolve_production_workbook(argument)
    staged_path = stage_single_workbook(
        production_path,
        PRODUCTION_DIR.parent / "VBAFiles" / production_path.stem,
    )
    print(f"[STAGED] {staged_path}")
    print("[NEXT] Run workbook_release.py test with a scenario before promotion.")


def test_workbook(argument, scenario_argument):
    production_path, staged_path = resolve_staged_workbook(argument)
    metadata_file = state_path(staged_path)
    with open(metadata_file, "r", encoding="utf-8") as file:
        metadata = json.load(file)

    if metadata.get("production_sha256") != sha256_file(production_path):
        raise RuntimeError("Production workbook changed after staging; rebuild the staged workbook.")
    if metadata.get("staged_sha256") != sha256_file(staged_path):
        raise RuntimeError("Staged workbook changed after it was built; rebuild it before testing.")

    scenario_path = Path(scenario_argument).resolve()
    if not scenario_path.is_file():
        raise FileNotFoundError(f"Test scenario not found: {scenario_path}")

    with open(scenario_path, "r", encoding="utf-8") as file:
        scenario_data = json.load(file)
    scenario_name = scenario_data.get("name", scenario_path.stem)
    run_started_at = datetime.now()
    run_directory = create_run_directory(
        TEST_RUNS_DIR,
        staged_path,
        scenario_name,
        run_started_at,
    )
    metadata.pop("test_run_id", None)
    metadata.pop("test_scenario_sha256", None)
    _write_release_metadata(metadata_file, metadata)

    command = [
        sys.executable,
        str(TEST_RUNNER),
        str(staged_path),
        "--scenario",
        str(scenario_path),
        "--artifact-dir",
        str(run_directory),
        "--run-timestamp",
        run_started_at.isoformat(),
    ]
    completed = subprocess.run(command, cwd=PROJECT_ROOT, check=False)

    result_file = run_directory / "results.json"
    if not result_file.is_file():
        if completed.returncode != 0:
            raise SystemExit(completed.returncode)
        raise RuntimeError("Test runner completed without writing its archived results.")

    with open(result_file, "r", encoding="utf-8") as file:
        results = json.load(file)
    staged_hash = sha256_file(staged_path)
    if results.get("target_path") != str(staged_path) or results.get("target_sha256") != staged_hash:
        raise RuntimeError("Test result does not match the staged workbook.")
    if results.get("scenario_path") != str(scenario_path):
        raise RuntimeError("Test result does not match the requested scenario.")
    if results.get("scenario_sha256") != sha256_file(scenario_path):
        raise RuntimeError("Test scenario changed while it was being tested.")
    expected_run_id = run_directory.relative_to(TEST_RUNS_DIR).as_posix()
    if results.get("run_id") != expected_run_id:
        raise RuntimeError("Test result run identifier does not match its archive directory.")

    metadata["test_run_id"] = expected_run_id
    metadata["test_scenario_sha256"] = results.get("scenario_sha256")
    _write_release_metadata(metadata_file, metadata)
    print(f"[TEST ARCHIVE] {run_directory}")
    if completed.returncode != 0:
        raise SystemExit(completed.returncode)
    if results.get("passed") is not True:
        raise RuntimeError("Test runner returned success without passing test evidence.")

def promote_workbook(
    argument,
    production_dir=PRODUCTION_DIR,
    staging_dir=STAGING_DIR,
    test_runs_dir=TEST_RUNS_DIR,
):
    production_dir = Path(production_dir).resolve()
    production_path, staged_path = resolve_staged_workbook(
        argument, production_dir, staging_dir
    )
    metadata_file = state_path(staged_path)
    if not metadata_file.is_file():
        raise RuntimeError("Missing build evidence; stage this workbook before promotion.")

    with open(metadata_file, "r", encoding="utf-8") as file:
        metadata = json.load(file)
    try:
        results_file = test_results_path(
            staged_path,
            test_runs_dir,
            metadata.get("test_run_id"),
        )
    except ValueError as error:
        raise RuntimeError(
            "Missing build or test evidence; stage and successfully test this workbook first."
        ) from error
    if not results_file.is_file():
        raise RuntimeError("The archived test result is missing; rerun the staged workbook test.")

    with open(results_file, "r", encoding="utf-8") as file:
        results = json.load(file)

    if metadata.get("production_path") != str(production_path):
        raise RuntimeError("Staging metadata does not match the requested production workbook.")
    if metadata.get("staged_path") != str(staged_path):
        raise RuntimeError("Staging metadata does not match the staged workbook.")
    if metadata.get("production_sha256") != sha256_file(production_path):
        raise RuntimeError("Production workbook changed after staging; rebuild and retest before promotion.")
    staged_hash = sha256_file(staged_path)
    if metadata.get("staged_sha256") != staged_hash:
        raise RuntimeError("Staged workbook changed after it was built; rebuild and retest before promotion.")
    if results.get("target_path") != str(staged_path):
        raise RuntimeError("Passing test results are not for this staged workbook.")
    if results.get("target_sha256") != staged_hash:
        raise RuntimeError("Staged workbook changed after testing; rerun tests before promotion.")
    if results.get("run_id") != metadata.get("test_run_id"):
        raise RuntimeError("Archived test run does not match the release evidence.")
    if metadata.get("test_scenario_sha256") != results.get("scenario_sha256"):
        raise RuntimeError("Archived test scenario hash does not match the release evidence.")
    scenario_path = Path(results.get("scenario_path", "")).resolve()
    if results.get("passed") is not True or not results.get("scenario_path"):
        raise RuntimeError("Promotion requires a passing scenario-based test.")
    if not scenario_path.is_file() or results.get("scenario_sha256") != sha256_file(scenario_path):
        raise RuntimeError("The test scenario changed or is missing; rerun the test before promotion.")

    backup_dir = production_dir / "Backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    backup_path = backup_dir / f"{production_path.stem}_{timestamp}{production_path.suffix}"
    shutil.copy2(production_path, backup_path)

    temporary_path = production_path.with_name(
        f".{production_path.stem}.promote-{uuid4().hex}{production_path.suffix}"
    )
    try:
        shutil.copy2(staged_path, temporary_path)
        if metadata.get("production_sha256") != sha256_file(production_path):
            raise RuntimeError("Production workbook changed during promotion; the staged copy was not installed.")
        os.replace(temporary_path, production_path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()

    print(f"[PROMOTED] {staged_path} -> {production_path}")
    print(f"[BACKUP] {backup_path}")


def main():
    parser = argparse.ArgumentParser(description="Stage, test, and promote Excel workbook changes.")
    commands = parser.add_subparsers(dest="command", required=True)

    stage_parser = commands.add_parser("stage", help="Build a fresh staged workbook and inject VBA source.")
    stage_parser.add_argument("workbook", help="Workbook filename or path inside ExcelFiles.")

    test_parser = commands.add_parser("test", help="Run a required scenario against a staged workbook.")
    test_parser.add_argument("workbook", help="Workbook filename or path inside ExcelFiles.")
    test_parser.add_argument("--scenario", required=True, help="JSON test scenario to execute.")

    promote_parser = commands.add_parser("promote", help="Install a tested staged workbook into ExcelFiles.")
    promote_parser.add_argument("workbook", help="Workbook filename or path inside ExcelFiles.")

    args = parser.parse_args()
    try:
        if args.command == "stage":
            stage_workbook(args.workbook)
        elif args.command == "test":
            test_workbook(args.workbook, args.scenario)
        else:
            promote_workbook(args.workbook)
    except (FileNotFoundError, RuntimeError, ValueError) as error:
        print(f"[ERROR] {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
