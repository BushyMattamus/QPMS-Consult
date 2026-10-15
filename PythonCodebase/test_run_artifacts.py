"""Archive test results and generate date- and workbook-oriented indexes."""

import hashlib
import json
import os
import re
import tempfile
from datetime import datetime
from pathlib import Path
from urllib.parse import quote
from uuid import uuid4


def _slug(value):
    slug = re.sub(r"[^A-Za-z0-9]+", "-", str(value).strip()).strip("-").lower()
    return slug or "unnamed"


def _parse_timestamp(value):
    if not isinstance(value, str) or not value:
        raise ValueError("Test result is missing a valid timestamp.")
    try:
        return datetime.fromisoformat(value)
    except ValueError as error:
        raise ValueError(f"Invalid test result timestamp: {value}") from error


def _result_identity(results):
    target_path = results.get("target_path")
    target = results.get("target")
    if not target and target_path:
        target = Path(target_path).name
    if not target:
        raise ValueError("Test result is missing its target workbook name.")
    scenario = results.get("scenario") or "unspecified scenario"
    timestamp = _parse_timestamp(results.get("timestamp"))
    return Path(target).stem, scenario, timestamp


def _run_directory(root, workbook, scenario, timestamp):
    run_name = (
        f"{timestamp:%Y%m%dT%H%M%S%f}__{_slug(scenario)}"
    )
    parent = Path(root) / timestamp.strftime("%Y-%m-%d") / workbook
    return parent / run_name


def create_run_directory(test_runs_dir, target_path, scenario, timestamp=None):
    """Create and return a unique date/workbook/scenario archive directory."""
    root = Path(test_runs_dir).resolve()
    timestamp = timestamp or datetime.now()
    workbook = Path(target_path).stem
    parent = root / timestamp.strftime("%Y-%m-%d") / workbook
    parent.mkdir(parents=True, exist_ok=True)
    run_name = f"{timestamp:%Y%m%dT%H%M%S%f}__{_slug(scenario)}"
    run_dir = parent / run_name
    while True:
        try:
            run_dir.mkdir()
            return run_dir
        except FileExistsError:
            run_dir = parent / f"{run_name}__{uuid4().hex[:8]}"


def _atomic_write_bytes(path, contents):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            "wb", dir=path.parent, prefix=f".{path.name}.", suffix=".tmp", delete=False
        ) as file:
            temporary_path = Path(file.name)
            file.write(contents)
        os.replace(temporary_path, path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()


def write_json_results(path, results):
    _atomic_write_bytes(
        path,
        (json.dumps(results, indent=2, default=str) + "\n").encode("utf-8"),
    )


def write_run_artifacts(artifact_dir, results, report_markdown):
    """Write an immutable JSON/Markdown pair into an allocated run directory."""
    artifact_dir = Path(artifact_dir)
    results_path = artifact_dir / "results.json"
    report_path = artifact_dir / "report.md"
    if results_path.exists() or report_path.exists():
        raise FileExistsError(f"Test-run artifacts already exist in {artifact_dir}")

    write_json_results(results_path, results)
    _atomic_write_bytes(report_path, (report_markdown.rstrip() + "\n").encode("utf-8"))
    return results_path, report_path


def _legacy_run_directory(root, results):
    workbook, scenario, timestamp = _result_identity(results)
    return _run_directory(root, workbook, scenario, timestamp)


def archive_legacy_results(test_runs_dir, render_markdown):
    """Copy flat legacy result JSONs into the archive without altering originals."""
    root = Path(test_runs_dir).resolve()
    if not root.exists():
        return []

    archived = []
    for legacy_path in sorted(root.glob("*.json")):
        if not (
            legacy_path.name == "last_test_results.json"
            or legacy_path.name.endswith("_test_results.json")
            or legacy_path.name.endswith("_path_test.json")
        ):
            continue
        results = json.loads(legacy_path.read_text(encoding="utf-8"))
        if (
            not isinstance(results, dict)
            or "timestamp" not in results
            or "passed" not in results
            or not (results.get("target") or results.get("target_path"))
        ):
            continue
        if results.get("run_id"):
            continue
        run_dir = _legacy_run_directory(root, results)
        results_path = run_dir / "results.json"
        report_path = run_dir / "report.md"
        run_dir.mkdir(parents=True, exist_ok=True)

        if not results_path.exists():
            _atomic_write_bytes(results_path, legacy_path.read_bytes())
        if not report_path.exists():
            workbook, _, _ = _result_identity(results)
            run_id = run_dir.relative_to(root).as_posix()
            report_data = dict(results)
            report_data["run_id"] = run_id
            report_data["target"] = report_data.get("target") or workbook
            _atomic_write_bytes(
                report_path,
                (render_markdown(report_data).rstrip() + "\n").encode("utf-8"),
            )
        archived.append(run_dir)
    return archived


def archive_test_run(test_runs_dir, results, render_markdown, artifact_dir=None):
    """Save one completed run, migrate flat legacy results, and refresh indexes."""
    root = Path(test_runs_dir).resolve()
    if artifact_dir is None:
        workbook, scenario, timestamp = _result_identity(results)
        artifact_dir = create_run_directory(
            root,
            workbook,
            scenario,
            timestamp,
        )
    else:
        artifact_dir = Path(artifact_dir).resolve()
        try:
            artifact_dir.relative_to(root)
        except ValueError as error:
            raise ValueError("Test-run artifact directory must be inside the archive root.") from error
        artifact_dir.mkdir(parents=True, exist_ok=True)

    archive_legacy_results(root, render_markdown)

    run_id = artifact_dir.relative_to(root).as_posix()
    results["run_id"] = run_id
    results["artifact_paths"] = {
        "markdown": f"{run_id}/report.md",
        "json": f"{run_id}/results.json",
    }
    results_path, report_path = write_run_artifacts(
        artifact_dir,
        results,
        render_markdown(results),
    )
    rebuild_indexes(root)
    return results_path, report_path


def _markdown_cell(value):
    return str(value or "").replace("|", "&#124;").replace("\r", " ").replace("\n", " ")


def _link(from_dir, target):
    relative_path = os.path.relpath(target, start=from_dir).replace(os.sep, "/")
    return quote(relative_path, safe="/._-")


def _read_archive_entries(root):
    entries = []
    for results_path in root.glob("20??-??-??/*/*/results.json"):
        results = json.loads(results_path.read_text(encoding="utf-8"))
        workbook, scenario, timestamp = _result_identity(results)
        report_path = results_path.with_name("report.md")
        run_dir = results_path.parent
        entries.append({
            "date": timestamp.strftime("%Y-%m-%d"),
            "timestamp": timestamp,
            "workbook": workbook,
            "scenario": scenario,
            "passed": results.get("passed") is True,
            "duration": results.get("duration_seconds"),
            "run_dir": run_dir,
            "results_path": results_path,
            "report_path": report_path,
            "date_index_path": root / timestamp.strftime("%Y-%m-%d") / "index.md",
            "workbook_index_path": _workbook_index_path(root, workbook),
        })
    return sorted(entries, key=lambda entry: entry["timestamp"].timestamp(), reverse=True)


def _workbook_index_path(root, workbook):
    suffix = hashlib.sha256(workbook.encode("utf-8")).hexdigest()[:8]
    return Path(root) / "by_workbook" / f"{_slug(workbook)}-{suffix}.md"


def _render_index(title, entries, base_dir):
    lines = [
        f"# {title}",
        "",
    ]
    if Path(base_dir).name != "by_workbook" and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", Path(base_dir).name):
        lines.extend(["[Browse by workbook](by_workbook/index.md)", ""])
    lines.extend([
        "| Date | Workbook | Scenario | Result | Duration (s) | Report | JSON |",
        "|---|---|---|---|---:|---|---|",
    ])
    for entry in entries:
        status = "PASSED" if entry["passed"] else "FAILED"
        duration = entry["duration"]
        duration_text = f"{duration:g}" if isinstance(duration, (int, float)) else ""
        date_text = _markdown_cell(entry["date"])
        workbook_text = _markdown_cell(entry["workbook"])
        if Path(base_dir).resolve() != entry["date_index_path"].parent.resolve():
            date_text = f"[{date_text}]({_link(base_dir, entry['date_index_path'])})"
        if Path(base_dir).name != "by_workbook":
            workbook_text = (
                f"[{workbook_text}]({_link(base_dir, entry['workbook_index_path'])})"
            )
        lines.append(
            "| {date} | {workbook} | {scenario} | {status} | {duration} | "
            "[Markdown]({report}) | [JSON]({json}) |".format(
                date=date_text,
                workbook=workbook_text,
                scenario=_markdown_cell(entry["scenario"]),
                status=status,
                duration=duration_text,
                report=_link(base_dir, entry["report_path"]),
                json=_link(base_dir, entry["results_path"]),
            )
        )
    if not entries:
        lines.append("| - | - | No archived test runs yet | - | - | - | - |")
    return "\n".join(lines) + "\n"


def rebuild_indexes(test_runs_dir):
    """Regenerate root, per-date, and per-workbook Markdown indexes."""
    root = Path(test_runs_dir).resolve()
    entries = _read_archive_entries(root)
    _atomic_write_bytes(
        root / "index.md",
        _render_index("Test Run Archive", entries, root).encode("utf-8"),
    )

    by_date = {}
    by_workbook = {}
    for entry in entries:
        by_date.setdefault(entry["date"], []).append(entry)
        by_workbook.setdefault(entry["workbook"], []).append(entry)

    for date, date_entries in by_date.items():
        date_dir = root / date
        _atomic_write_bytes(
            date_dir / "index.md",
            _render_index(f"Test Runs for {date}", date_entries, date_dir).encode("utf-8"),
        )

    workbook_index_dir = root / "by_workbook"
    workbook_index_lines = ["# Test Runs by Workbook", "", "| Workbook | Runs | Latest |", "|---|---:|---|"]
    for workbook, workbook_entries in sorted(by_workbook.items(), key=lambda item: item[0].casefold()):
        index_path = _workbook_index_path(root, workbook)
        _atomic_write_bytes(
            index_path,
            _render_index(f"Test Runs for {workbook}", workbook_entries, workbook_index_dir).encode("utf-8"),
        )
        latest = workbook_entries[0]
        workbook_index_lines.append(
            "| [{name}]({link}) | {count} | [{date}]({date_link}) |".format(
                name=_markdown_cell(workbook),
                link=_link(workbook_index_dir, index_path),
                count=len(workbook_entries),
                date=latest["date"],
                date_link=_link(workbook_index_dir, latest["date_index_path"]),
            )
        )
    if not workbook_index_lines[3:]:
        workbook_index_lines.append("| - | 0 | - |")
    _atomic_write_bytes(
        workbook_index_dir / "index.md",
        ("\n".join(workbook_index_lines) + "\n").encode("utf-8"),
    )

    return entries
