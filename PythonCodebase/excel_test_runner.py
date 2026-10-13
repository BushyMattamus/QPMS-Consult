"""
Excel AI Test Runner & Inspector
--------------------------------
Automated test framework for Excel workbooks (.xlsm) and VBA code.
Can be triggered manually via CLI or VS Code tasks.

Features:
- Auto-detects the most recently modified workbook or VBA module.
- Creates an isolated temporary copy in a sandbox (`.test_runs/`).
- Inspects worksheet structure, named ranges, formulas, and cells into Markdown/JSON.
- Injects test inputs, evaluates formulas, and executes VBA macros.
- Captures runtime errors, formula errors (#VALUE!, #REF!, etc.), and unexpected outputs.
- Can sync code modifications from `VBAFiles/` into the target workbook.
"""

import sys
import os
import json
import shutil
import tempfile
import traceback
from datetime import datetime
from pathlib import Path

# Ensure UTF-8 output encoding for Windows terminal
if sys.platform == 'win32' and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

REPO_ROOT = Path(__file__).resolve().parent.parent
EXCEL_DIR = REPO_ROOT / "ExcelFiles"
VBA_DIR = REPO_ROOT / "VBAFiles"
SANDBOX_DIR = REPO_ROOT / ".test_runs"
SCENARIO_DIR = REPO_ROOT / "TestScenarios"


def ensure_excel_com():
    """Ensure win32com is available and initialize Excel."""
    try:
        import win32com.client as win32
        return win32
    except ImportError:
        print("[ERROR] pywin32 is not installed. Please run: pip install pywin32")
        sys.exit(1)


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

    def __enter__(self):
        SANDBOX_DIR.mkdir(parents=True, exist_ok=True)
        if self.sandbox:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            sandbox_wb = SANDBOX_DIR / f"{self.original_path.stem}_test_{timestamp}{self.original_path.suffix}"
            shutil.copy2(self.original_path, sandbox_wb)
            self.active_path = sandbox_wb
        else:
            self.active_path = self.original_path

        try:
            # Dispatch Excel Application
            self.excel = self.win32.gencache.EnsureDispatch('Excel.Application')
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
            raise
        return self

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
        try:
            res = self.excel.Run(macro_name, *args)
            return {"success": True, "result": res, "error": None}
        except Exception as e:
            return {"success": False, "result": None, "error": str(e)}

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
    md.append(f"**Sandbox Copy (deleted after test):** `{report_data.get('sandbox_path')}`\n")

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
            if m.get("error"):
                md.append(f"  - **Error:** `{m['error']}`")
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


def run_test_scenario(workbook_path, scenario_path=None):
    """
    Executes a test scenario against the workbook.
    If no scenario is specified, performs a general health check and macro test.
    """
    scenario_data = None
    if scenario_path and Path(scenario_path).exists():
        with open(scenario_path, "r", encoding="utf-8") as f:
            scenario_data = json.load(f)

    results = {
        "target": Path(workbook_path).name,
        "timestamp": datetime.now().isoformat(),
        "passed": True,
        "scenario": scenario_data.get("name", "Default Inspection & Health Check") if scenario_data else "Default Inspection",
        "macro_results": [],
        "errors": [],
        "sheets": []
    }

    with ExcelSession(workbook_path, sandbox=True) as session:
        results["sandbox_path"] = str(session.active_path)

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
                print(f"[TEST] Executing macro: {m_name}")
                m_res = session.run_macro(m_name, *m_args)
                m_res["name"] = m_name
                results["macro_results"].append(m_res)
                if not m_res["success"]:
                    results["passed"] = False

        # Recalculate
        session.calculate()

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
            if sheet_inspection["errors"]:
                results["passed"] = False
                results["errors"].extend(sheet_inspection["errors"])

    # Output report
    SANDBOX_DIR.mkdir(parents=True, exist_ok=True)
    report_file = SANDBOX_DIR / "last_test_report.md"
    report_md = generate_markdown_report(results)
    with open(report_file, "w", encoding="utf-8") as f:
        f.write(report_md)

    print(f"\n[REPORT GENERATED] Saved to: {report_file}")
    print(report_md)
    return results


def main():
    import argparse
    parser = argparse.ArgumentParser(description="AI Excel Test Runner & Inspector")
    parser.add_argument("workbook", nargs="?", help="Path to .xlsm workbook (optional, defaults to last modified)")
    parser.add_argument("--scenario", help="Path to JSON test scenario file")
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
        return

    run_test_scenario(wb_path, args.scenario)


if __name__ == "__main__":
    main()
