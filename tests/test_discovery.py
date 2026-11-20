import contextlib
import io
from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from holiday_stipend.cli import main
from holiday_stipend.discovery import inspect_file


class DiscoveryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)

    def test_csv_counts_variable_rows_and_samples_bom(self):
        path = self.directory / "employees.csv"
        path.write_text("\ufeffName,Department\n\"Doe, Jane\",School\nBlank\n", encoding="utf-8")
        table, = inspect_file(path, 2)
        self.assertEqual((table.row_count, table.column_count), (3, 2))
        self.assertEqual(table.samples, [("Name", "Department"), ("Doe, Jane", "School")])

    def test_empty_csv(self):
        path = self.directory / "empty.csv"
        path.write_text("", encoding="utf-8")
        table, = inspect_file(path)
        self.assertEqual((table.row_count, table.column_count, table.samples), (0, 0, []))

    def test_workbook_multiple_sheets_hidden_sheet_and_formula(self):
        path = self.directory / "employees.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Staff"
        sheet.append(["Name", "Formula"])
        sheet.append(["Jane", "=1+1"])
        hidden = workbook.create_sheet("Archive")
        hidden.sheet_state = "hidden"
        hidden.append(["Old employee"])
        workbook.save(path)
        workbook.close()
        staff, archive = inspect_file(path, 2)
        self.assertEqual((staff.name, staff.visibility, staff.row_count, staff.column_count), ("Staff", "visible", 2, 2))
        self.assertEqual(staff.samples[1], ("Jane", "=1+1"))
        self.assertEqual((archive.name, archive.visibility, archive.row_count), ("Archive", "hidden", 1))

    def test_cli_samples_are_opt_in_and_errors_do_not_stop_other_files(self):
        path = self.directory / "employees.csv"
        path.write_text("Name\nPrivate Person\n", encoding="utf-8")
        stdout, stderr = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
            result = main(["discover", str(self.directory / "missing.csv"), str(path)])
        self.assertEqual(result, 1)
        self.assertIn("missing.csv", stderr.getvalue())
        self.assertIn("Rows: 2; columns: 1", stdout.getvalue())
        self.assertNotIn("Private Person", stdout.getvalue())

        stdout = io.StringIO()
        with contextlib.redirect_stdout(stdout):
            result = main(["discover", "--sample-rows", "2", str(path)])
        self.assertEqual(result, 0)
        self.assertIn("Private Person", stdout.getvalue())

    def test_unsupported_and_invalid_csv_report_errors(self):
        bad_csv = self.directory / "bad.csv"
        bad_csv.write_text('"unterminated', encoding="utf-8")
        stderr = io.StringIO()
        with contextlib.redirect_stderr(stderr):
            result = main(["discover", str(self.directory / "old.xls"), str(bad_csv)])
        self.assertEqual(result, 1)
        self.assertIn("unsupported file type", stderr.getvalue())
        self.assertIn("unexpected end of data", stderr.getvalue())

    def test_negative_sample_count_is_rejected(self):
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as exc:
            main(["discover", "--sample-rows", "-1", "anything.csv"])
        self.assertEqual(exc.exception.code, 2)


if __name__ == "__main__":
    unittest.main()
