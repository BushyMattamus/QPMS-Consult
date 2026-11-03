"""
Excel AI Test Runner & Inspector
--------------------------------
Automated test framework for Excel workbooks (.xlsm) and VBA code.
Can be triggered manually via CLI or VS Code tasks.

Features:
- Auto-detects the most recently modified workbook or VBA module.
- Creates an isolated temporary workbook copy outside the repository.
- Archives JSON and Markdown results by date, workbook, and scenario.
- Inspects worksheet structure, named ranges, formulas, and cells into Markdown/JSON.
- Injects test inputs, evaluates formulas, and executes VBA macros.
- Captures runtime errors, formula errors (#VALUE!, #REF!, etc.), and unexpected outputs.
- Can sync code modifications from `VBAFiles/` into the target workbook.
"""

import sys
import os
import json
import csv
import hashlib
import multiprocessing
import queue
import shutil
import tempfile
import time
import traceback
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from test_run_artifacts import archive_test_run, write_json_results

# Ensure UTF-8 output encoding for Windows terminal
if sys.platform == 'win32' and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

REPO_ROOT = Path(__file__).resolve().parent.parent
EXCEL_DIR = REPO_ROOT / "ExcelFiles"
VBA_DIR = REPO_ROOT / "VBAFiles"
TEST_RUNS_DIR = REPO_ROOT / ".test_runs"
SANDBOX_WORKBOOK_DIR = Path(tempfile.gettempdir()) / "QPMS-Consult-test-runs"
DATA_DIR = REPO_ROOT / "DataFiles"
DATA_ROOT_ENV_VAR = "QPMS_DATA_ROOT"
TEST_MODE_ENV_VAR = "QPMS_TEST_MODE"
TEST_LOG_ENV_VAR = "QPMS_TEST_LOG"
DIALOG_CLOSE_TIMEOUT_SECONDS = 2.0
UTF8_BOM = b"\xef\xbb\xbf"
SCENARIO_DIR = REPO_ROOT / "TestScenarios"


def validate_dialog_rules(dialog_rules):
    """Validate expected Excel message-box responses configured by a scenario."""
    if not isinstance(dialog_rules, list):
        raise ValueError("Macro 'dialogs' must be a list.")

    for rule in dialog_rules:
        if (
            not isinstance(rule, dict)
            or not isinstance(rule.get("title"), str)
            or not rule["title"].strip()
            or rule.get("button") not in {"OK", "Yes", "No", "Cancel"}
            or not isinstance(rule.get("count", 1), int)
            or isinstance(rule.get("count", 1), bool)
            or rule.get("count", 1) < 1
        ):
            raise ValueError(
                "Each macro dialog rule must include a title, an OK/Yes/No/Cancel "
                "button, and a positive integer count."
            )
    return dialog_rules


def validate_macro_expectations(macros):
    """Validate optional scalar return-value expectations on scenario macros."""
    if not isinstance(macros, list):
        raise ValueError("Scenario 'macros' must be a list.")

    for macro in macros:
        if not isinstance(macro, dict):
            raise ValueError("Each scenario macro must be an object.")
        if isinstance(macro.get("expected_result"), (list, dict)):
            raise ValueError(
                f"Macro '{macro.get('name')}' expected_result must be a JSON scalar "
                "(true, false, number, or string)."
            )
    return macros


def macro_result_matches_expected(result, expected):
    """Compare macro return values; booleans must match booleans so 1 never equals True."""
    if isinstance(result, bool) or isinstance(expected, bool):
        return isinstance(result, bool) and isinstance(expected, bool) and result == expected
    return result == expected


def validate_csv_assertions(assertions):
    """Validate scenario assertions against CSV files in the sandbox data copy."""
    if not isinstance(assertions, list):
        raise ValueError("Scenario 'csv_assertions' must be a list.")

    for assertion in assertions:
        if not isinstance(assertion, dict):
            raise ValueError("Each CSV assertion must be an object.")
        file_name = assertion.get("file")
        if not isinstance(file_name, str) or not file_name.strip():
            raise ValueError("Each CSV assertion must include a safe relative 'file'.")
        relative_path = Path(file_name)
        match = assertion.get("match")
        numeric_match = assertion.get("numeric_match", {})
        expected_count = assertion.get("count", 1)
        encoding = assertion.get("encoding", "utf-8-sig")
        if (
            relative_path.is_absolute()
            or ".." in relative_path.parts
            or not isinstance(match, dict)
            or not match
            or not all(isinstance(key, str) for key in match)
            or not isinstance(numeric_match, dict)
            or not isinstance(expected_count, int)
            or isinstance(expected_count, bool)
            or expected_count < 0
            or not isinstance(encoding, str)
            or encoding not in ("utf-8-sig", "cp1252")
        ):
            raise ValueError(
                "Each CSV assertion must include a safe relative 'file', a non-empty "
                "'match' object, a non-negative integer 'count', and a supported encoding."
            )
        for column, expected in numeric_match.items():
            if (
                not isinstance(column, str)
                or not isinstance(expected, dict)
                or not isinstance(expected.get("expected"), (int, float))
                or isinstance(expected.get("expected"), bool)
                or not isinstance(expected.get("tolerance", 0), (int, float))
                or isinstance(expected.get("tolerance", 0), bool)
                or expected.get("tolerance", 0) < 0
            ):
                raise ValueError(
                    "CSV numeric matches must map column names to numeric expected "
                    "values and non-negative tolerances."
                )
    return assertions


def evaluate_csv_assertion(data_root, assertion):
    """Count rows matching a scenario assertion in a sandbox CSV file."""
    if data_root is None:
        raise RuntimeError("CSV assertions require a sandbox copy of DataFiles.")
    data_root = Path(data_root).resolve()
    csv_path = (data_root / assertion["file"]).resolve()
    if data_root not in csv_path.parents:
        raise ValueError(f"CSV assertion path escapes the sandbox data folder: {assertion['file']}")
    if not csv_path.is_file():
        raise FileNotFoundError(f"CSV assertion file not found in sandbox: {assertion['file']}")

    encoding = assertion.get("encoding", "utf-8-sig")
    with csv_path.open("rb") as raw_file:
        if raw_file.read(len(UTF8_BOM)) == UTF8_BOM:
            # A UTF-8 BOM read as cp1252 turns the first header into "ï»¿Name",
            # so the match silently finds no rows.
            encoding = "utf-8-sig"

    with csv_path.open(
        "r",
        encoding=encoding,
        newline="",
    ) as csv_file:
        matched_count = 0
        for row in csv.DictReader(csv_file):
            if not all(
                row.get(key) == str(value)
                for key, value in assertion["match"].items()
            ):
                continue
            numeric_matches = assertion.get("numeric_match", {})
            try:
                numeric_values_match = all(
                    abs(Decimal(row[column]) - Decimal(str(expected["expected"])))
                    <= Decimal(str(expected.get("tolerance", 0)))
                    for column, expected in numeric_matches.items()
                )
            except (KeyError, InvalidOperation, TypeError) as error:
                raise ValueError(
                    f"CSV assertion for {assertion['file']} has a missing or "
                    f"non-numeric column value: {error}"
                ) from error
            if numeric_values_match:
                matched_count += 1

    expected_count = assertion.get("count", 1)
    return {
        "file": assertion["file"],
        "match": assertion["match"],
        "expected_count": expected_count,
        "matched_count": matched_count,
        "passed": matched_count == expected_count,
    }


def _dialog_is_open(hwnd):
    import win32gui

    return bool(win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd))


def _wait_for_dialog_close(hwnd, timeout):
    deadline = time.monotonic() + timeout
    while _dialog_is_open(hwnd):
        if time.monotonic() >= deadline:
            return False
        time.sleep(0.05)
    return True


def _click_dialog_buttons(hwnd, button_ids, responses):
    """Click each button's control directly when a posted command is ignored."""
    import win32gui

    bm_click = 0x00F5
    for response in responses:
        child = win32gui.GetDlgItem(hwnd, button_ids[response])
        if child:
            win32gui.SendMessage(child, bm_click, 0, 0)
        if _wait_for_dialog_close(hwnd, 1.0):
            return True
    return False


def _terminate_test_excel(excel_pid):
    """Stop only the dedicated test Excel instance so a blocked excel.Run returns."""
    import win32api
    import win32con

    handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, excel_pid)
    try:
        win32api.TerminateProcess(handle, 1)
    finally:
        win32api.CloseHandle(handle)


def watch_excel_dialogs(excel_pid, dialog_rules, stop_event, result_queue):
    """Dismiss Excel message boxes from a separate process while COM is blocked."""
    import ctypes
    import win32gui
    import win32process

    button_ids = {"OK": 1, "Cancel": 2, "Yes": 6, "No": 7, "Abort": 3}
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    post_message = user32.PostMessageW
    post_message.argtypes = [
        ctypes.c_void_p,
        ctypes.c_uint,
        ctypes.c_size_t,
        ctypes.c_ssize_t,
    ]
    post_message.restype = ctypes.c_int
    handled_counts = [0] * len(dialog_rules)
    seen_dialogs = set()

    while not stop_event.is_set():
        dialogs = []

        def collect(hwnd, result):
            if not win32gui.IsWindow(hwnd):
                return True
            if win32gui.GetClassName(hwnd) != "#32770":
                return True
            process_id = win32process.GetWindowThreadProcessId(hwnd)[1]
            if process_id == excel_pid:
                result.append((hwnd, win32gui.GetWindowText(hwnd)))
            return True

        win32gui.EnumWindows(collect, dialogs)
        active_handles = {hwnd for hwnd, _ in dialogs}
        seen_dialogs.intersection_update(active_handles)

        for hwnd, title in dialogs:
            if hwnd in seen_dialogs:
                continue

            rule_index = next(
                (
                    index for index, rule in enumerate(dialog_rules)
                    if rule["title"] == title
                    and handled_counts[index] < rule["count"]
                ),
                None,
            )
            if rule_index is not None:
                buttons = (dialog_rules[rule_index]["button"],)
                expected = True
                handled_counts[rule_index] += 1
            else:
                buttons = ("No", "Cancel", "Abort", "OK")
                expected = False
            for response in buttons:
                if not post_message(hwnd, 0x0111, button_ids[response], 0):
                    result_queue.put({
                        "type": "error",
                        "message": (
                            f"Could not post {response} to Excel dialog '{title}': "
                            f"{ctypes.get_last_error()}."
                        ),
                    })

            dismissed = _wait_for_dialog_close(hwnd, DIALOG_CLOSE_TIMEOUT_SECONDS)
            if not dismissed:
                dismissed = _click_dialog_buttons(hwnd, button_ids, buttons)
            result_queue.put({
                "type": "dialog",
                "title": title,
                "button": buttons[0] if expected else "No/Cancel/Abort/OK fallback",
                "expected": expected,
                "dismissed": dismissed,
            })
            seen_dialogs.add(hwnd)

            if not dismissed:
                result_queue.put({
                    "type": "error",
                    "message": (
                        f"Excel dialog '{title}' stayed open after clicking "
                        f"{buttons[0]}; stopped the test Excel instance to avoid a hang."
                    ),
                })
                _terminate_test_excel(excel_pid)
                return

        stop_event.wait(0.05)


class ExcelModalDialogWatcher:
    """Run a separate Windows process to handle Excel message boxes."""

    def __init__(self, excel_hwnd, dialog_rules):
        import win32process

        self.excel_pid = win32process.GetWindowThreadProcessId(excel_hwnd)[1]
        self.dialog_rules = [dict(rule, count=rule.get("count", 1)) for rule in dialog_rules]
        self.stop_event = multiprocessing.Event()
        self.result_queue = multiprocessing.Queue()
        self.process = None
        self.events = []
        self.errors = []

    def start(self):
        self.process = multiprocessing.Process(
            target=watch_excel_dialogs,
            args=(self.excel_pid, self.dialog_rules, self.stop_event, self.result_queue),
            daemon=True,
        )
        self.process.start()

    def stop(self):
        self.stop_event.set()
        if self.process:
            self.process.join(timeout=5)
            if self.process.is_alive():
                self.process.terminate()
                self.process.join(timeout=2)
                self.errors.append("Excel dialog watcher did not stop within five seconds.")

        while True:
            try:
                result = self.result_queue.get_nowait()
            except queue.Empty:
                break
            if result["type"] == "dialog":
                self.events.append({
                    "title": result["title"],
                    "button": result["button"],
                    "expected": result["expected"],
                    "dismissed": result.get("dismissed", True),
                })
            else:
                self.errors.append(result["message"])

        if self.process and self.process.exitcode not in (0, None):
            self.errors.append(
                f"Excel dialog watcher exited with code {self.process.exitcode}."
            )
        self.result_queue.close()

    def validation_errors(self):
        errors = list(self.errors)
        for rule in self.dialog_rules:
            handled_count = sum(
                1 for event in self.events
                if event["expected"]
                and event["dismissed"]
                and event["title"] == rule["title"]
                and event["button"] == rule["button"]
            )
            missing_count = rule["count"] - handled_count
            if missing_count:
                errors.append(
                    f"Expected Excel dialog '{rule['title']}' with {rule['button']} "
                    f"{missing_count} more time(s), but it did not appear."
                )
        errors.extend(
            f"Unexpected Excel dialog '{event['title']}' was dismissed with "
            f"{event['button']}."
            for event in self.events
            if not event["expected"]
        )
        return errors


def ensure_excel_com():
    """Ensure win32com is available and initialize Excel."""
    try:
        import win32com.client as win32
        return win32
    except ImportError:
        print("[ERROR] pywin32 is not installed. Please run: pip install pywin32")
        sys.exit(1)


def sha256_file(file_path):
    """Return the SHA-256 digest of a file without loading it all into memory."""
    digest = hashlib.sha256()
    with open(file_path, "rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def find_latest_target():
    """
    Finds the most recently modified .xlsm in ExcelFiles or .bas/.cls/.frm in VBAFiles.
    Returns: (workbook_path, source_type, last_mtime)
    """
    candidates = []

    if EXCEL_DIR.exists():
        for f in EXCEL_DIR.glob("*.xlsm"):
            if not f.name.startswith("~$"):
                candidates.append((f, "excel", f.stat().st_mtime))

    if VBA_DIR.exists():
        for f in VBA_DIR.rglob("*.*"):
            if f.suffix.lower() in [".bas", ".cls", ".frm"]:
                # Map to corresponding workbook
                # e.g., VBAFiles/TEST/modules/Test.bas -> ExcelFiles/TEST.xlsm
                rel_parts = f.relative_to(VBA_DIR).parts
                if rel_parts:
                    wb_name = rel_parts[0]
                    target_wb = EXCEL_DIR / f"{wb_name}.xlsm"
                    if target_wb.exists():
                        candidates.append((target_wb, f"vba:{f.name}", f.stat().st_mtime))

    if not candidates:
        return None, None, None

    candidates.sort(key=lambda x: x[2], reverse=True)
    return candidates[0]


class ExcelSession:
    """Manages an automated Excel COM session."""
    def __init__(self, workbook_path, sandbox=True, visible=False):
        self.win32 = ensure_excel_com()
        self.original_path = Path(workbook_path).resolve()
        self.sandbox = sandbox
        self.visible = visible
        self.excel = None
        self.wb = None
        self.active_path = None
        self.sandbox_data_dir = None
        self.previous_data_root = None
        self.previous_test_mode = None
        self.test_mode_set = False
        self.test_log_path = None
        self.previous_test_log = None
        self.dialog_rules = []

    def __enter__(self):
        SANDBOX_WORKBOOK_DIR.mkdir(parents=True, exist_ok=True)
        if self.sandbox:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            sandbox_wb = SANDBOX_WORKBOOK_DIR / f"{self.original_path.stem}_test_{timestamp}{self.original_path.suffix}"
            shutil.copy2(self.original_path, sandbox_wb)
            self.active_path = sandbox_wb
            # Sandbox runs skip the success MsgBox that VBA shows only when this is unset.
            self.previous_test_mode = os.environ.get(TEST_MODE_ENV_VAR)
            os.environ[TEST_MODE_ENV_VAR] = "1"
            self.test_mode_set = True
            # Test-mode VBA ShowMessage calls are appended here instead of shown as MsgBoxes.
            self.previous_test_log = os.environ.get(TEST_LOG_ENV_VAR)
            self.test_log_path = SANDBOX_WORKBOOK_DIR / f"messages_{timestamp}.log"
            os.environ[TEST_LOG_ENV_VAR] = str(self.test_log_path)
            # The sandboxed workbook lives outside the repo, so give it a private
            # copy of DataFiles and point the VBA PathResolver at it.
            if DATA_DIR.is_dir():
                self.sandbox_data_dir = SANDBOX_WORKBOOK_DIR / f"data_{timestamp}"
                shutil.copytree(DATA_DIR, self.sandbox_data_dir)
                self.previous_data_root = os.environ.get(DATA_ROOT_ENV_VAR)
                os.environ[DATA_ROOT_ENV_VAR] = str(self.sandbox_data_dir)
        else:
            self.active_path = self.original_path

        try:
            # A dedicated Excel instance inherits the environment set above
            # and never touches an Excel window the user already has open.
            self.excel = self.win32.DispatchEx('Excel.Application')
            self.excel.Visible = self.visible
            self.excel.DisplayAlerts = False
            self.excel.ScreenUpdating = True

            self.wb = self.excel.Workbooks.Open(str(self.active_path))
        except Exception:
            if self.excel:
                try:
                    self.excel.Quit()
                except Exception:
                    pass
                self.excel = None
            if self.sandbox and self.active_path:
                self.active_path.unlink(missing_ok=True)
            self._cleanup_sandbox_data()
            raise
        return self

    def _cleanup_sandbox_data(self):
        if self.test_mode_set:
            if self.previous_test_mode is None:
                os.environ.pop(TEST_MODE_ENV_VAR, None)
            else:
                os.environ[TEST_MODE_ENV_VAR] = self.previous_test_mode
            self.test_mode_set = False
        if self.sandbox_data_dir:
            shutil.rmtree(self.sandbox_data_dir, ignore_errors=True)
            self.sandbox_data_dir = None
            if self.previous_data_root is None:
                os.environ.pop(DATA_ROOT_ENV_VAR, None)
            else:
                os.environ[DATA_ROOT_ENV_VAR] = self.previous_data_root
        if self.test_log_path:
            self.test_log_path.unlink(missing_ok=True)
            self.test_log_path = None
            if self.previous_test_log is None:
                os.environ.pop(TEST_LOG_ENV_VAR, None)
            else:
                os.environ[TEST_LOG_ENV_VAR] = self.previous_test_log

    def take_test_messages(self):
        """Return and clear the messages that test-mode VBA ShowMessage calls logged."""
        if not self.test_log_path or not self.test_log_path.exists():
            return []
        # VBA Print # writes the system ANSI code page, so read it back the same way.
        content = self.test_log_path.read_text(encoding="mbcs", errors="replace")
        self.test_log_path.unlink(missing_ok=True)
        return [line for line in content.splitlines() if line.strip()]

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.wb:
            try:
                self.wb.Close(SaveChanges=False)
            except Exception:
                pass
            self.wb = None
        if self.excel:
            try:
                self.excel.Quit()
            except Exception:
                pass
            self.excel = None
        if self.sandbox and self.active_path:
            self.active_path.unlink(missing_ok=True)
        self._cleanup_sandbox_data()

    def get_sheet(self, sheet_name=None):
        if sheet_name:
            return self.wb.Sheets(sheet_name)
        return self.wb.ActiveSheet

    def inspect_sheet(self, sheet_name=None, max_rows=30, max_cols=15):
        """
        Extracts a clean tabular summary of a sheet including values, formulas, and errors.
        """
        sheet = self.get_sheet(sheet_name)
        used_range = sheet.UsedRange
        start_row = used_range.Row
        start_col = used_range.Column
        num_rows = min(used_range.Rows.Count, max_rows)
        num_cols = min(used_range.Columns.Count, max_cols)

        cells_data = []
        errors = []

        for r in range(start_row, start_row + num_rows):
            row_data = []
            for c in range(start_col, start_col + num_cols):
                cell = sheet.Cells(r, c)
                val = cell.Value
                formula = cell.Formula
                cell_text = cell.Text

                # Check for standard Excel formula errors
                if isinstance(val, int) and val in [-2146826281, -2146826259, -2146826245]:
                    errors.append({
                        "cell": f"{sheet.Name}!R{r}C{c}",
                        "text": cell_text,
                        "formula": formula
                    })
                elif str(cell_text).startswith("#") and str(cell_text).endswith(("!", "?")):
                    errors.append({
                        "cell": f"{sheet.Name}!R{r}C{c}",
                        "text": cell_text,
                        "formula": formula
                    })

                row_data.append({
                    "val": val,
                    "formula": formula if str(formula).startswith("=") else None,
                    "text": str(cell_text) if cell_text is not None else ""
                })
            cells_data.append(row_data)

        return {
            "sheet_name": sheet.Name,
            "used_range": f"R{start_row}C{start_col}:R{start_row + num_rows - 1}C{start_col + num_cols - 1}",
            "grid": cells_data,
            "errors": errors
        }

    def run_macro(self, macro_name, *args):
        """Executes a VBA macro in the open workbook."""
        # Discard messages logged before this macro so each result reports only its own.
        self.take_test_messages()
        watcher = None
        macro_result = None
        macro_error = None
        success = False
        try:
            # Always watch: unexpected message boxes are dismissed and reported
            # as failures instead of blocking the run.
            watcher = ExcelModalDialogWatcher(self.excel.Hwnd, self.dialog_rules)
            watcher.start()
            res = self.excel.Run(macro_name, *args)
            macro_result = res
            success = True
        except Exception as e:
            macro_error = str(e)
        finally:
            if watcher:
                watcher.stop()

        messages = self.take_test_messages()
        dialogs = watcher.events if watcher else []
        if watcher:
            dialog_errors = watcher.validation_errors()
            if dialog_errors:
                success = False
                macro_error = "; ".join(filter(None, [macro_error, *dialog_errors]))
        return {
            "success": success,
            "result": macro_result,
            "error": macro_error,
            "dialogs": dialogs,
            "messages": messages,
        }

    def format_range(self, sheet_name, cell_ref, bg_color=None, font_color=None, bold=None, italic=None, font_size=None):
        """
        Applies visual formatting to a cell or range.
        bg_color and font_color can be (R, G, B) tuples or hex strings like '#DAF2D0'.
        """
        sheet = self.get_sheet(sheet_name)
        rng = sheet.Range(cell_ref)

        def to_excel_rgb(color):
            if isinstance(color, str) and color.startswith("#"):
                color = color.lstrip("#")
                r = int(color[0:2], 16)
                g = int(color[2:4], 16)
                b = int(color[4:6], 16)
                return r + (g * 256) + (b * 65536)
            elif isinstance(color, (tuple, list)) and len(color) == 3:
                return color[0] + (color[1] * 256) + (color[2] * 65536)
            return color

        if bg_color is not None:
            rng.Interior.Color = to_excel_rgb(bg_color)
        if font_color is not None:
            rng.Font.Color = to_excel_rgb(font_color)
        if bold is not None:
            rng.Font.Bold = bold
        if italic is not None:
            rng.Font.Italic = italic
        if font_size is not None:
            rng.Font.Size = font_size

    def set_cell_value(self, sheet_name, cell_ref, value):
        sheet = self.get_sheet(sheet_name)
        sheet.Range(cell_ref).Value = value

    def get_cell_value(self, sheet_name, cell_ref):
        sheet = self.get_sheet(sheet_name)
        cell = sheet.Range(cell_ref)
        return {
            "value": cell.Value,
            "formula": cell.Formula,
            "text": cell.Text
        }

    def calculate(self):
        self.excel.CalculateFull()


def generate_markdown_report(report_data):
    """Formats inspection and test results into human/AI-readable markdown."""
    md = []
    md.append(f"# Excel AI Test Report: `{report_data.get('target')}`")
    md.append(f"**Timestamp:** {report_data.get('timestamp')}")
    if report_data.get("run_id"):
        md.append(f"**Run ID:** `{report_data['run_id']}`")
    if report_data.get("duration_seconds") is not None:
        md.append(f"**Duration:** {report_data['duration_seconds']} seconds")
    md.append(f"**Sandbox Copy (deleted after test):** `{report_data.get('sandbox_path')}`\n")
    md.append(f"**Result:** {'PASSED' if report_data.get('passed') else 'FAILED'}\n")

    if report_data.get("scenario"):
        md.append(f"## Test Scenario: `{report_data['scenario']}`")
        md.append(f"**Status:** {'✅ PASSED' if report_data.get('passed') else '❌ FAILED'}\n")

    # Macro Execution Results
    if "macro_results" in report_data and report_data["macro_results"]:
        md.append("### Macro Execution")
        for m in report_data["macro_results"]:
            status_icon = "✅" if m["success"] else "❌"
            md.append(f"- **{m['name']}**: {status_icon} Success={m['success']}")
            if m.get("result") is not None:
                md.append(f"  - Return Value: `{m['result']}`")
            if "expected_result" in m:
                match_icon = "✅" if m.get("result_matches_expected") else "❌"
                md.append(f"  - Expected Return Value: `{m['expected_result']}`")
                md.append(f"  - Return Value Assertion: {match_icon} Match={m.get('result_matches_expected')}")
            if m.get("error"):
                md.append(f"  - **Error:** `{m['error']}`")
            for dialog in m.get("dialogs", []):
                md.append(
                    f"  - Dialog `{dialog['title']}`: "
                    f"{'expected' if dialog['expected'] else 'unexpected'}; "
                    f"clicked **{dialog['button']}**."
                )
            for message in m.get("messages", []):
                md.append(f"  - Logged message: `{message}`")
        md.append("")

    csv_assertions = report_data.get("csv_assertions", [])
    if csv_assertions:
        md.append("### CSV Data Assertions")
        for assertion in csv_assertions:
            status_icon = "✅" if assertion["passed"] else "❌"
            md.append(
                f"- `{assertion['file']}`: {status_icon} matched "
                f"{assertion['matched_count']} row(s); expected "
                f"{assertion['expected_count']}."
            )
        md.append("")

    # Formula/Cell Errors
    errors = report_data.get("errors", [])
    if errors:
        md.append(f"### ⚠️ Errors Detected ({len(errors)})")
        for err in errors:
            md.append(f"- **{err['cell']}**: `{err['text']}` (Formula: `{err.get('formula')}`)")
        md.append("")
    else:
        md.append("### ✅ Errors Detected: 0\n")

    ignored_formula_errors = report_data.get("ignored_formula_errors", [])
    if ignored_formula_errors:
        md.append(f"### Ignored Formula Errors ({len(ignored_formula_errors)})")
        for err in ignored_formula_errors:
            md.append(f"- **{err['cell']}**: `{err['text']}` (Formula: `{err.get('formula')}`)")
        md.append("")

    # Sheet Inspection Snapshot
    if "sheets" in report_data:
        for s in report_data["sheets"]:
            md.append(f"### Sheet Snapshot: `{s['sheet_name']}` ({s['used_range']})")
            grid = s["grid"]
            if grid:
                header = "| Row | " + " | ".join([f"Col {i+1}" for i in range(len(grid[0]))]) + " |"
                sep = "|---|" + "|".join(["---" for _ in range(len(grid[0]))]) + "|"
                md.append(header)
                md.append(sep)
                for r_idx, row in enumerate(grid[:15]):  # limit display rows
                    cols_str = " | ".join([c["text"] or ("`" + c["formula"] + "`" if c["formula"] else "") for c in row])
                    md.append(f"| {r_idx+1} | {cols_str} |")
            md.append("")

    return "\n".join(md)


def is_ignored_formula_error(sheet_name, error, ignore_rules):
    return any(
        rule["sheet"].casefold() == sheet_name.casefold()
        and rule["error"] == error["text"]
        for rule in ignore_rules
    )


def run_test_scenario(
    workbook_path,
    scenario_path=None,
    json_report_path=None,
    artifact_dir=None,
    run_timestamp=None,
):
    """
    Executes a test scenario against the workbook.
    If no scenario is specified, performs a general health check and macro test.
    Completed results are archived with date and workbook indexes.
    """
    workbook_path = Path(workbook_path).resolve()
    scenario_file = Path(scenario_path).resolve() if scenario_path else None
    initial_workbook_hash = sha256_file(workbook_path)
    run_started_at = run_timestamp or datetime.now()

    scenario_data = None
    if scenario_file:
        if not scenario_file.is_file():
            raise FileNotFoundError(f"Test scenario not found: {scenario_file}")
        with open(scenario_file, "r", encoding="utf-8") as f:
            scenario_data = json.load(f)

    ignore_formula_errors = scenario_data.get("ignore_formula_errors", []) if scenario_data else []
    if not isinstance(ignore_formula_errors, list):
        raise ValueError("Scenario 'ignore_formula_errors' must be a list.")
    for rule in ignore_formula_errors:
        if (
            not isinstance(rule, dict)
            or not isinstance(rule.get("sheet"), str)
            or not isinstance(rule.get("error"), str)
        ):
            raise ValueError(
                "Each 'ignore_formula_errors' rule must include string 'sheet' and 'error' values."
            )
    csv_assertions = validate_csv_assertions(
        scenario_data.get("csv_assertions", []) if scenario_data else []
    )
    if scenario_data and "macros" in scenario_data:
        validate_macro_expectations(scenario_data["macros"])
    if scenario_data and "active_sheet" in scenario_data:
        if not isinstance(scenario_data["active_sheet"], str) or not scenario_data["active_sheet"]:
            raise ValueError("Scenario 'active_sheet' must be a non-empty string.")

    results = {
        "target": workbook_path.name,
        "target_path": str(workbook_path),
        "target_sha256": initial_workbook_hash,
        "timestamp": run_started_at.isoformat(),
        "passed": True,
        "scenario": scenario_data.get("name", "Default Inspection & Health Check") if scenario_data else "Default Inspection",
        "scenario_path": str(scenario_file) if scenario_file else None,
        "scenario_sha256": sha256_file(scenario_file) if scenario_file else None,
        "macro_results": [],
        "errors": [],
        "ignored_formula_errors": [],
        "csv_assertions": [],
        "sheets": []
    }

    run_started_perf = time.perf_counter()
    with ExcelSession(workbook_path, sandbox=True) as session:
        results["sandbox_path"] = str(session.active_path)

        if scenario_data and scenario_data.get("active_sheet"):
            session.get_sheet(scenario_data["active_sheet"]).Activate()

        # 1. Apply inputs if provided in scenario
        if scenario_data and "inputs" in scenario_data:
            for inp in scenario_data["inputs"]:
                sheet = inp.get("sheet")
                cell = inp["cell"]
                val = inp["value"]
                print(f"[TEST] Injecting input into {sheet or 'Active'}!{cell} = {val}")
                session.set_cell_value(sheet, cell, val)

        # 2. Run macros if specified
        if scenario_data and "macros" in scenario_data:
            for m in scenario_data["macros"]:
                m_name = m["name"]
                m_args = m.get("args", [])
                session.dialog_rules = validate_dialog_rules(m.get("dialogs", []))
                print(f"[TEST] Executing macro: {m_name}")
                try:
                    m_res = session.run_macro(m_name, *m_args)
                finally:
                    session.dialog_rules = []
                m_res["name"] = m_name
                results["macro_results"].append(m_res)
                if not m_res["success"]:
                    results["passed"] = False
                if "expected_result" in m:
                    m_res["expected_result"] = m["expected_result"]
                    m_res["result_matches_expected"] = macro_result_matches_expected(
                        m_res["result"], m["expected_result"]
                    )
                    if not m_res["result_matches_expected"]:
                        results["passed"] = False
                        results["errors"].append({
                            "cell": f"Macro {m_name}",
                            "text": (
                                f"Expected return value {m['expected_result']!r}, "
                                f"got {m_res['result']!r}."
                            ),
                            "formula": None,
                        })

        # Recalculate
        session.calculate()

        for assertion in csv_assertions:
            csv_result = evaluate_csv_assertion(session.sandbox_data_dir, assertion)
            results["csv_assertions"].append(csv_result)
            if not csv_result["passed"]:
                results["passed"] = False
                results["errors"].append({
                    "cell": assertion["file"],
                    "text": (
                        f"Expected {csv_result['expected_count']} matching CSV row(s), "
                        f"found {csv_result['matched_count']}."
                    ),
                    "formula": None,
                })

        # 3. Check assertions if defined
        if scenario_data and "assertions" in scenario_data:
            for assertion in scenario_data["assertions"]:
                sheet = assertion.get("sheet")
                cell = assertion["cell"]
                expected = assertion.get("expected")
                actual_info = session.get_cell_value(sheet, cell)
                actual_val = actual_info["value"]
                
                if actual_val != expected:
                    results["passed"] = False
                    results["errors"].append({
                        "cell": f"{sheet or 'Active'}!{cell}",
                        "text": f"Expected '{expected}', got '{actual_val}'",
                        "formula": actual_info["formula"]
                    })

        # 4. Inspect sheets for standard Excel calculation errors
        sheet_count = session.wb.Sheets.Count
        for i in range(1, sheet_count + 1):
            sheet_name = session.wb.Sheets(i).Name
            sheet_inspection = session.inspect_sheet(sheet_name)
            results["sheets"].append(sheet_inspection)
            for error in sheet_inspection["errors"]:
                if is_ignored_formula_error(sheet_name, error, ignore_formula_errors):
                    results["ignored_formula_errors"].append(error)
                else:
                    results["passed"] = False
                    results["errors"].append(error)

    results["duration_seconds"] = round(time.perf_counter() - run_started_perf, 3)

    sandbox_path = Path(results["sandbox_path"])
    sandbox_path.unlink(missing_ok=True)
    if sandbox_path.exists():
        raise OSError(f"Could not remove temporary test workbook: {sandbox_path}")

    final_workbook_hash = sha256_file(workbook_path)
    results["target_sha256"] = final_workbook_hash
    if final_workbook_hash != initial_workbook_hash:
        results["passed"] = False
        results["errors"].append({
            "cell": "Workbook file",
            "text": "The workbook changed while the test was running; test evidence is invalid.",
            "formula": None
        })

    result_file, report_file = archive_test_run(
        TEST_RUNS_DIR,
        results,
        generate_markdown_report,
        artifact_dir,
    )

    if json_report_path:
        export_path = Path(json_report_path).resolve()
        if export_path != result_file.resolve():
            write_json_results(export_path, results)

    print(f"\n[REPORT GENERATED] Saved to: {report_file}")
    print(f"[RESULTS SAVED] Saved to: {result_file}")
    if json_report_path and export_path != result_file.resolve():
        print(f"[JSON EXPORT] Saved to: {export_path}")
    print(report_file.read_text(encoding="utf-8"))
    sandbox_path.unlink(missing_ok=True)
    if sandbox_path.exists():
        raise OSError(f"Could not remove temporary test workbook: {sandbox_path}")
    return results


def main():
    import argparse
    parser = argparse.ArgumentParser(description="AI Excel Test Runner & Inspector")
    parser.add_argument("workbook", nargs="?", help="Path to .xlsm workbook (optional, defaults to last modified)")
    parser.add_argument("--scenario", help="Path to JSON test scenario file")
    parser.add_argument("--json-report", help="Path for machine-readable test results")
    parser.add_argument("--artifact-dir", help="Existing per-run archive directory for JSON and Markdown outputs")
    parser.add_argument("--run-timestamp", help="ISO timestamp that determines the archive date and run time")
    parser.add_argument("--detect-only", action="store_true", help="Only detect and display the latest changed target")

    args = parser.parse_args()

    wb_path = args.workbook
    if not wb_path:
        target, src_type, mtime = find_latest_target()
        if not target:
            print("[ERROR] No Excel workbook found in ExcelFiles/")
            sys.exit(1)
        wb_path = target
        mtime_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
        print(f"[DETECTED] Target: {wb_path.name} (Triggered by {src_type}, last modified {mtime_str})")

    if args.detect_only:
        return 0

    try:
        run_timestamp = (
            datetime.fromisoformat(args.run_timestamp) if args.run_timestamp else None
        )
    except ValueError as error:
        parser.error(f"Invalid --run-timestamp: {error}")

    results = run_test_scenario(
        wb_path,
        args.scenario,
        args.json_report,
        args.artifact_dir,
        run_timestamp,
    )
    return 0 if results["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
