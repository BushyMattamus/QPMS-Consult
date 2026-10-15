# Test Run Archive

Excel test results are stored under the repository's ignored `.test_runs/` directory. Each completed run has a canonical, immutable JSON and Markdown pair:

```text
.test_runs/
  index.md
  by_workbook/
    r1-inbound-and-sampling-<hash>.md
  YYYY-MM-DD/
    <Workbook Name>/
      YYYYMMDDTHHMMSSffffff__<scenario>/
        results.json
        report.md
```

Use the root `index.md` for recent runs, a date folder's `index.md` to browse that day, or the matching page under `by_workbook/` to browse a workbook's history. Indexes link to the canonical run files rather than duplicating them.

`excel_test_runner.py` writes the JSON and human-readable Markdown report from the same result data. Its optional `--json-report` argument also writes a separate JSON export; the dated archive remains authoritative. The isolated temporary workbook is still deleted after each test.

`workbook_release.py test` records the archived run identifier in the staged workbook's release metadata. Promotion checks that exact run's workbook hash, scenario hash, and pass status. Staging invalidates that pointer but keeps previous archived test results.

On the next completed test after upgrading, valid flat JSON test results in the `.test_runs/` root are copied into the dated archive and rendered as Markdown. Original files are left untouched. `.test_runs/` remains excluded from Git.
