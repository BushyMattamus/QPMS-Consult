import tempfile
import unittest
from pathlib import Path
from unittest import mock

from excel_test_runner import (
    _wait_for_dialog_close,
    evaluate_csv_assertion,
    generate_markdown_report,
    is_ignored_formula_error,
    macro_result_matches_expected,
    validate_csv_assertions,
    validate_dialog_rules,
    validate_macro_expectations,
)


class IgnoredFormulaErrorMatchingTests(unittest.TestCase):
    def setUp(self):
        self.rules = [{"sheet": "Helper", "error": "#REF!"}]
        self.error = {"text": "#REF!"}

    def test_matches_error_on_configured_sheet(self):
        self.assertTrue(is_ignored_formula_error("Helper", self.error, self.rules))

    def test_matches_sheet_name_case_insensitively(self):
        self.assertTrue(is_ignored_formula_error("HELPER", self.error, self.rules))

    def test_does_not_match_other_sheets_or_errors(self):
        self.assertFalse(is_ignored_formula_error("QC_Release", self.error, self.rules))
        self.assertFalse(
            is_ignored_formula_error("Helper", {"text": "#VALUE!"}, self.rules)
        )


class IgnoredFormulaErrorReportTests(unittest.TestCase):
    def test_ignored_formula_errors_are_reported_without_failing(self):
        report = generate_markdown_report({
            "target": "R2.xlsm",
            "timestamp": "2026-10-14T22:30:00",
            "sandbox_path": "sandbox.xlsm",
            "scenario": "R2 PathResolver",
            "passed": True,
            "errors": [],
            "ignored_formula_errors": [
                {
                    "cell": "Helper!R2C14",
                    "text": "#REF!",
                    "formula": "=FILTER(...)",
                }
            ],
            "sheets": [],
        })

        self.assertIn("Ignored Formula Errors (1)", report)
        self.assertIn("Helper!R2C14", report)
        self.assertIn("#REF!", report)

    def test_report_includes_run_id_and_duration(self):
        report = generate_markdown_report({
            "target": "R3 Goods Issuing.xlsm",
            "timestamp": "2026-10-15T21:30:00",
            "sandbox_path": "sandbox.xlsm",
            "scenario": "R3 PathResolver startup loads",
            "run_id": "2026-10-15/R3 Goods Issuing/20261015T213000__startup-loads",
            "duration_seconds": 42.5,
            "passed": True,
            "errors": [],
            "ignored_formula_errors": [],
            "sheets": [],
        })

        self.assertIn("**Run ID:**", report)
        self.assertIn("**Duration:** 42.5 seconds", report)


class DialogCloseWaitTests(unittest.TestCase):
    def test_reports_dismissed_when_window_closes(self):
        with mock.patch(
            "excel_test_runner._dialog_is_open", side_effect=[True, False]
        ):
            self.assertTrue(_wait_for_dialog_close(1234, timeout=1.0))

    def test_reports_not_dismissed_when_window_stays_open(self):
        with mock.patch("excel_test_runner._dialog_is_open", return_value=True):
            self.assertFalse(_wait_for_dialog_close(1234, timeout=0.1))


class ScenarioDialogRuleTests(unittest.TestCase):
    def test_accepts_expected_dialog_response(self):
        rules = [
            {"title": "Adjustment Complete", "button": "OK", "count": 1}
        ]
        self.assertEqual(validate_dialog_rules(rules), rules)

    def test_rejects_unsafe_or_incomplete_dialog_rules(self):
        with self.assertRaisesRegex(ValueError, "dialog rule"):
            validate_dialog_rules([
                {"title": "Confirm", "button": "Yes", "count": 0}
            ])


class ScenarioMacroExpectationTests(unittest.TestCase):
    def test_accepts_scalar_expected_macro_results(self):
        macros = [{"name": "Database.Initialize", "expected_result": True}]

        self.assertEqual(validate_macro_expectations(macros), macros)
        self.assertTrue(macro_result_matches_expected(True, True))
        self.assertFalse(macro_result_matches_expected(False, True))
        self.assertFalse(macro_result_matches_expected(1, True))

    def test_rejects_non_scalar_expected_macro_results(self):
        with self.assertRaisesRegex(ValueError, "JSON scalar"):
            validate_macro_expectations([
                {"name": "Database.Initialize", "expected_result": []}
            ])

    def test_report_shows_expected_macro_result(self):
        report = generate_markdown_report({
            "target": "Database_Create_Connect.xlsm",
            "timestamp": "2026-11-03T21:00:00",
            "sandbox_path": "sandbox.xlsm",
            "scenario": "Database rebuild",
            "passed": False,
            "errors": [],
            "ignored_formula_errors": [],
            "macro_results": [
                {
                    "name": "Database_Initialization.ReinitializeDatabase",
                    "success": False,
                    "result": False,
                    "expected_result": True,
                    "result_matches_expected": False,
                    "dialogs": [],
                }
            ],
            "sheets": [],
        })

        self.assertIn("Expected Return Value: `True`", report)
        self.assertIn("Return Value Assertion: ❌ Match=False", report)


class CsvAssertionTests(unittest.TestCase):
    def test_counts_rows_matching_expected_transaction_fields(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            data_root = Path(temp_dir)
            csv_path = data_root / "RM_Trans.csv"
            csv_path.write_bytes(
                (
                    "Trans_Type,QC_Number,Quantity,Master_UoM,Adj_Note,QC_Balance\n"
                    "WOF,250085,-2,mL,autotest\xa0,4826.105\n"
                ).encode("cp1252")
            )

            result = evaluate_csv_assertion(
                data_root,
                {
                    "file": "RM_Trans.csv",
                    "match": {
                        "Trans_Type": "WOF",
                        "QC_Number": "250085",
                        "Quantity": "-2",
                        "Master_UoM": "mL",
                        "Adj_Note": "autotest\xa0",
                    },
                    "numeric_match": {
                        "QC_Balance": {
                            "expected": 4826.105,
                            "tolerance": 0.0005,
                        }
                    },
                    "count": 1,
                    "encoding": "cp1252",
                },
            )

        self.assertTrue(result["passed"])
        self.assertEqual(result["matched_count"], 1)

    def test_utf8_bom_header_matches_even_with_cp1252_override(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            data_root = Path(temp_dir)
            (data_root / "QC_Number.csv").write_bytes(
                b"\xef\xbb\xbf" + (
                    "QC_Number,QC_BAL\r\n250085,4826.105\r\n"
                ).encode("ascii")
            )

            result = evaluate_csv_assertion(
                data_root,
                {
                    "file": "QC_Number.csv",
                    "match": {"QC_Number": "250085"},
                    "numeric_match": {
                        "QC_BAL": {"expected": 4826.105, "tolerance": 0.0005}
                    },
                    "count": 1,
                    "encoding": "cp1252",
                },
            )

        self.assertTrue(result["passed"])
        self.assertEqual(result["matched_count"], 1)

    def test_rejects_csv_assertions_outside_data_folder(self):
        with self.assertRaisesRegex(ValueError, "safe relative"):
            validate_csv_assertions([
                {"file": "..\\outside.csv", "match": {"QC_Number": "250085"}}
            ])

    def test_report_shows_dialog_and_csv_assertion_results(self):
        report = generate_markdown_report({
            "target": "R4.xlsm",
            "timestamp": "2026-10-14T23:00:00",
            "sandbox_path": "sandbox.xlsm",
            "scenario": "R4 write-off",
            "passed": True,
            "errors": [],
            "ignored_formula_errors": [],
            "macro_results": [
                {
                    "name": "ButtonUtils.ButtSaveStockAdjustment",
                    "success": True,
                    "dialogs": [
                        {
                            "title": "Adjustment Complete",
                            "button": "OK",
                            "expected": True,
                        }
                    ],
                }
            ],
            "csv_assertions": [
                {
                    "file": "Material Management/Tables/RM_Trans.csv",
                    "passed": True,
                    "matched_count": 1,
                    "expected_count": 1,
                }
            ],
            "sheets": [],
        })

        self.assertIn("Adjustment Complete", report)
        self.assertIn("clicked **OK**", report)
        self.assertIn("matched 1 row(s); expected 1", report)


if __name__ == "__main__":
    unittest.main()
