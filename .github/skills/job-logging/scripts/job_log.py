"""Append job timer and duration entries to portable daily CSV logs."""

import argparse
import csv
from datetime import datetime
import os
import re
from pathlib import Path
import tempfile


REPOSITORY_ROOT = Path(__file__).resolve().parents[4]
LOG_DIRECTORY = REPOSITORY_ROOT / "JobLogging"
LOG_PATTERN = "JobLog_*.csv"
LEGACY_LOG_PATTERN = "JobLog_*.txt"
CSV_FIELDS = (
    "Date",
    "Description",
    "Start",
    "End",
    "Duration",
    "DurationSeconds",
    "DurationHours",
    "Type",
)
TIMER_TYPE = "Timed"
MANUAL_TYPE = "Manual"


def log_path(log_directory, timestamp):
    """Return the daily CSV path for a timestamp."""
    return Path(log_directory) / f"JobLog_{timestamp:%Y%m%d}.csv"


def _display_duration(seconds):
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def _row(timestamp, description, start=None, end=None, duration_seconds=None, entry_type=TIMER_TYPE):
    return {
        "Date": timestamp.strftime("%Y-%m-%d"),
        "Description": description,
        "Start": start.isoformat(timespec="seconds") if start else "",
        "End": end.isoformat(timespec="seconds") if end else "",
        "Duration": _display_duration(duration_seconds) if duration_seconds is not None else "",
        "DurationSeconds": str(duration_seconds) if duration_seconds is not None else "",
        "DurationHours": f"{duration_seconds / 3600:.6f}" if duration_seconds is not None else "",
        "Type": entry_type,
    }


def _read_rows(path):
    with Path(path).open("r", encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file))


def _write_rows(path, rows):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8-sig",
            newline="",
            dir=path.parent,
            prefix=f".{path.name}.",
            suffix=".tmp",
            delete=False,
        ) as file:
            temporary_path = Path(file.name)
            writer = csv.DictWriter(file, fieldnames=CSV_FIELDS)
            writer.writeheader()
            writer.writerows(rows)
        os.replace(temporary_path, path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()


def _append_row(log_directory, timestamp, row):
    path = log_path(log_directory, timestamp)
    path.parent.mkdir(parents=True, exist_ok=True)
    needs_header = not path.exists() or path.stat().st_size == 0
    encoding = "utf-8-sig" if needs_header else "utf-8"
    with path.open("a", encoding=encoding, newline="") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_FIELDS)
        if needs_header:
            writer.writeheader()
        writer.writerow(row)
    return path


def _migrate_legacy_logs(log_directory):
    """Convert legacy text logs to CSV, leaving original text files untouched."""
    directory = Path(log_directory)
    if not directory.exists():
        return

    starts = {}
    stops = {}
    manual_rows = []
    for path in sorted(directory.glob(LEGACY_LOG_PATTERN)):
        with path.open("r", encoding="utf-8-sig") as file:
            for line in file:
                fields = line.rstrip("\r\n").split(" | ", 3)
                if len(fields) < 3:
                    continue
                if fields[0] == "START" and len(fields) == 4:
                    starts[fields[1]] = (fields[2], fields[3])
                elif fields[0] == "STOP":
                    stop_fields = fields[3].split(" | ", 1) if len(fields) == 4 else []
                    if len(stop_fields) == 2 and stop_fields[0].startswith("elapsed="):
                        stops[fields[1]] = (fields[2], stop_fields[1])
                elif fields[0] == "MANUAL" and len(fields) == 4:
                    details = fields[2].split(" | ", 1)
                    if len(details) == 2 and details[0].startswith("duration="):
                        seconds_match = re.search(r"\((\d+)s\)$", details[0])
                        if seconds_match:
                            timestamp = datetime.fromisoformat(fields[1])
                            manual_rows.append(
                                _row(
                                    timestamp,
                                    details[1],
                                    duration_seconds=int(seconds_match.group(1)),
                                    entry_type=MANUAL_TYPE,
                                )
                            )

    rows_by_date = {}
    for record_id, (start_text, description) in starts.items():
        started_at = datetime.fromisoformat(start_text)
        stop = stops.get(record_id)
        if stop is None:
            row = _row(started_at, description, start=started_at)
        else:
            stopped_at = datetime.fromisoformat(stop[0])
            seconds = max(0, int((stopped_at - started_at).total_seconds()))
            row = _row(
                started_at,
                description,
                start=started_at,
                end=stopped_at,
                duration_seconds=seconds,
            )
        rows_by_date.setdefault(started_at.strftime("%Y%m%d"), []).append(row)

    for row in manual_rows:
        rows_by_date.setdefault(row["Date"].replace("-", ""), []).append(row)

    for date_key, rows in rows_by_date.items():
        target = directory / f"JobLog_{date_key}.csv"
        if not target.exists():
            _write_rows(target, rows)


def _active_timers(log_directory):
    """Find timed CSV rows that have not yet been stopped."""
    active = []
    directory = Path(log_directory)
    if not directory.exists():
        return active

    for path in sorted(directory.glob(LOG_PATTERN)):
        rows = _read_rows(path)
        active.extend(
            (path, index, row)
            for index, row in enumerate(rows)
            if row.get("Type") == TIMER_TYPE and not row.get("End")
        )
    return active


def _parse_timestamp(value, reference):
    """Parse HH:MM (on the reference date) or YYYY-MM-DD HH:MM."""
    text = value.strip()
    for pattern in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%H:%M", "%H:%M:%S"):
        try:
            parsed = datetime.strptime(text, pattern)
        except ValueError:
            continue
        if pattern.startswith("%H"):
            return reference.replace(hour=parsed.hour, minute=parsed.minute, second=parsed.second, microsecond=0)
        return parsed
    raise ValueError(f"Time '{value}' must be HH:MM or YYYY-MM-DD HH:MM.")


def start_job(description, log_directory=LOG_DIRECTORY, now=None):
    """Start the single allowed active job timer."""
    _migrate_legacy_logs(log_directory)
    description = description.strip()
    if not description:
        raise ValueError("Job description cannot be empty.")

    active = _active_timers(log_directory)
    if active:
        active_description = active[0][2]["Description"]
        raise ValueError(
            f"A timer is already running for '{active_description}'. "
            "Stop it before starting another job."
        )

    timestamp = now or datetime.now()
    row = _row(timestamp, description, start=timestamp)
    path = _append_row(log_directory, timestamp, row)
    return path, timestamp


def stop_job(log_directory=LOG_DIRECTORY, now=None):
    """Stop the active timer and complete its CSV row."""
    _migrate_legacy_logs(log_directory)
    active = _active_timers(log_directory)
    if not active:
        raise ValueError("There is no active job timer to stop.")
    if len(active) > 1:
        raise RuntimeError("Multiple active timers were found; resolve the log entries before stopping.")

    path, row_index, row = active[0]
    started_at = datetime.fromisoformat(row["Start"])
    stopped_at = now or datetime.now()
    elapsed = int((stopped_at - started_at).total_seconds())
    if elapsed < 0:
        raise ValueError("The current time is earlier than the timer start; no stop entry was written.")

    duration = _display_duration(elapsed)
    rows = _read_rows(path)
    rows[row_index]["End"] = stopped_at.isoformat(timespec="seconds")
    rows[row_index]["Duration"] = duration
    rows[row_index]["DurationSeconds"] = str(elapsed)
    rows[row_index]["DurationHours"] = f"{elapsed / 3600:.6f}"
    _write_rows(path, rows)
    return row["Description"], duration, path


def log_manual(description, start, end, log_directory=LOG_DIRECTORY, now=None):
    """Append a completed job with a required start and end time."""
    _migrate_legacy_logs(log_directory)
    description = description.strip()
    if not description:
        raise ValueError("Job description cannot be empty.")

    reference = now or datetime.now()
    started_at = _parse_timestamp(start, reference)
    ended_at = _parse_timestamp(end, reference)
    elapsed = int((ended_at - started_at).total_seconds())
    if elapsed <= 0:
        raise ValueError("The end time must be later than the start time.")

    row = _row(
        started_at,
        description,
        start=started_at,
        end=ended_at,
        duration_seconds=elapsed,
        entry_type=MANUAL_TYPE,
    )
    path = _append_row(log_directory, started_at, row)
    return _display_duration(elapsed), path


def show_status(log_directory=LOG_DIRECTORY, now=None):
    """Return active-job details, or None when no timer is running."""
    _migrate_legacy_logs(log_directory)
    active = _active_timers(log_directory)
    if not active:
        return None
    if len(active) > 1:
        raise RuntimeError("Multiple active timers were found; resolve the log entries before continuing.")

    path, _, row = active[0]
    started_at = datetime.fromisoformat(row["Start"])
    current_time = now or datetime.now()
    elapsed = int((current_time - started_at).total_seconds())
    if elapsed < 0:
        raise ValueError("The current time is earlier than the timer start.")
    return row["Description"], started_at, _display_duration(elapsed), path


def main():
    parser = argparse.ArgumentParser(description="Append job time records to daily CSV logs.")
    commands = parser.add_subparsers(dest="command", required=True)

    start_parser = commands.add_parser("start", help="Start timing a job.")
    start_parser.add_argument("description", help="Short description of the job.")
    commands.add_parser("stop", help="Stop the active job timer.")
    commands.add_parser("status", help="Show the active job timer.")

    log_parser = commands.add_parser("log", help="Record a completed job with start and end times.")
    log_parser.add_argument("--start", required=True, help="Start time: HH:MM or YYYY-MM-DD HH:MM.")
    log_parser.add_argument("--end", required=True, help="End time: HH:MM or YYYY-MM-DD HH:MM.")
    log_parser.add_argument("description", help="Short description of the completed job.")

    args = parser.parse_args()
    try:
        if args.command == "start":
            path, timestamp = start_job(args.description)
            print(f"[STARTED] {args.description.strip()} at {timestamp:%Y-%m-%d %H:%M:%S}")
            print(f"[LOG] {path}")
        elif args.command == "stop":
            description, duration, path = stop_job()
            print(f"[STOPPED] {description} - elapsed {duration}")
            print(f"[LOG] {path}")
        elif args.command == "status":
            status = show_status()
            if status is None:
                print("[STATUS] No active job timer.")
                return 0
            description, started_at, duration, path = status
            print(f"[ACTIVE] {description} - started {started_at:%Y-%m-%d %H:%M:%S}, elapsed {duration}")
            print(f"[LOG] {path}")
        else:
            duration, path = log_manual(args.description, args.start, args.end)
            print(f"[LOGGED] {args.description.strip()} - {args.start} to {args.end}, duration {duration}")
            print(f"[LOG] {path}")
    except (OSError, RuntimeError, ValueError) as error:
        print(f"[ERROR] {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
