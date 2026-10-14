import tempfile
import unittest
from pathlib import Path

from excel_test_runner import (
    evaluate_csv_assertion,
    generate_markdown_report,
    is_ignored_formula_error,
    validate_csv_assertions,
    validate_dialog_rules,
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
