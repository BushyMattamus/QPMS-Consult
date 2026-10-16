---
name: job-logging
description: "Track work and time in daily CSV logs for Excel import and billing. Use when the user asks to start, stop, check, or manually record a job or its duration."
argument-hint: "start <job>, stop, status, or log --start <time> --end <time> <job>"
user-invocable: true
---

# Job Logging

Track work in the repository's `JobLogging\JobLog_YYYYMMDD.csv` files. Use the bundled script to keep timestamps and durations consistent. The CSV files use a UTF-8 BOM for Excel compatibility and contain Date, Description, Start, End, Duration, DurationSeconds, DurationHours, and Type columns.

## Actions

Run commands from the repository root. Prefer `.venv\Scripts\python.exe` on Windows when it exists; otherwise use `python`. The script path is `.github\skills\job-logging\scripts\job_log.py`.

- Start a timer when the user asks to start timing a job:
  ```powershell
  .\.venv\Scripts\python.exe .github\skills\job-logging\scripts\job_log.py start "Reviewing workbook calculations"
  ```
- Stop the active timer when the user asks to stop or finish:
  ```powershell
  .\.venv\Scripts\python.exe .github\skills\job-logging\scripts\job_log.py stop
  ```
- Check the active timer when asked what job is being timed:
  ```powershell
  .\.venv\Scripts\python.exe .github\skills\job-logging\scripts\job_log.py status
  ```
- Record a completed job with its start and end times (required for every manual entry):
  ```powershell
  .\.venv\Scripts\python.exe .github\skills\job-logging\scripts\job_log.py log --start 13:26 --end 22:28 "Investigating QC number duplicate risk"
  ```
  Times are `HH:MM` on today's date, or `YYYY-MM-DD HH:MM` for another date or a span over midnight.

Use `python` instead of the virtual-environment path if that executable is unavailable. Do not start, stop, or infer time without the user's request or explicit start and end times. Every row must have a Start and End. If a manual entry is requested without both times, ask the user for them rather than writing a row with blank times or estimating. If the user asks to start a second job while a timer is running, explain the active job and ask them to stop it first; the logger allows only one active timer. If there is no active timer to stop, report that rather than inventing an entry.

`start` adds a timed row with a blank End and Duration. `stop` completes that row with the local stop time, elapsed duration, seconds, and decimal hours. The row stays in the file for the date the timer started, including when it crosses midnight. `log --start --end` adds a completed manual row to the start date's file, with Duration calculated from the two times. CSV fields are escaped by the CSV writer for safe Excel import.

On first use, existing `JobLog_YYYYMMDD.txt` logs are converted to their corresponding CSV files. The original text files are preserved unchanged; do not delete them unless the user explicitly asks.

Report the action, job description, duration when available, and log filename after the command succeeds. If the command fails, relay the error; do not claim the log was updated.
