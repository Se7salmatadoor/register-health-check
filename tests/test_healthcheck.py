import datetime as dt
import hashlib
import os
import sys
import tempfile
import unittest
import zipfile

from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import healthcheck as hc  # noqa: E402

D = dt.date
ASOF = D(2026, 11, 25)


def md5(path):
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


def make_book(path, header, rows, title="Reg"):
    wb = Workbook()
    ws = wb.active
    ws.title = title
    ws.append(header)
    for r in rows:
        ws.append(r)
    wb.save(path)


def issues_of(path, **kw):
    wb = load_workbook(path, data_only=True)
    cfg = hc.Config(ASOF, **kw)
    out = []
    for ws in wb.worksheets:
        out += hc.Checker(ws, cfg).run()
    return out


def checks(issues):
    return [i["Check"] for i in issues]


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = self.tmp.name

    def tearDown(self):
        self.tmp.cleanup()

    def p(self, name):
        return os.path.join(self.dir, name)


class TestChecks(Base):
    H = ["No", "Employee Name", "Job Title", "ID Card", "Issue Date", "Qty"]

    def run_rows(self, rows, **kw):
        f = self.p("a.xlsx")
        make_book(f, self.H, rows)
        return issues_of(f, **kw)

    def test_clean_file_has_no_high_issues(self):
        rows = [[1, "Ali A", "Driver", "111", D(2026, 10, 1), 1],
                [2, "Sara B", "Mason", "222", D(2026, 10, 2), 1]]
        self.assertFalse([i for i in self.run_rows(rows) if i["Severity"] == "High"])

    def test_exact_duplicate_row(self):
        rows = [[1, "Ali A", "Driver", "1", D(2026, 10, 1), 1],
                [2, "Ali A", "Driver", "1", D(2026, 10, 1), 1]]
        iss = [i for i in self.run_rows(rows) if i["Check"] == "Duplicate row"]
        self.assertEqual(len(iss), 1)
        self.assertEqual(iss[0]["Row"], 3)

    def test_row_number_column_does_not_hide_duplicates(self):
        rows = [[7, "Ali A", "Driver", "1", D(2026, 10, 1), 1],
                [8, "Ali A", "Driver", "1", D(2026, 10, 1), 1]]
        self.assertIn("Duplicate row", checks(self.run_rows(rows)))

    def test_text_date_and_unreadable_date(self):
        rows = [[1, "Ali A", "Driver", "1", "07-10-2026", 1],
                [2, "Sara B", "Mason", "2", "16-3-20226", 1]]
        c = checks(self.run_rows(rows))
        self.assertIn("Date stored as text", c)
        self.assertIn("Unreadable date", c)

    def test_month_first_option(self):
        rows = [[1, "Ali A", "Driver", "1", "13-10-2026", 1]]
        self.assertNotIn("Unreadable date", checks(self.run_rows(rows)))
        rows = [[1, "Ali A", "Driver", "1", "10-13-2026", 1]]
        self.assertIn("Unreadable date", checks(self.run_rows(rows)))
        self.assertNotIn("Unreadable date", checks(self.run_rows(rows, month_first=True)))

    def test_impossible_dates_and_config(self):
        rows = [[1, "Ali A", "Driver", "1", D(2028, 8, 1), 1],
                [2, "Sara B", "Mason", "2", D(1999, 1, 1), 1]]
        self.assertEqual(checks(self.run_rows(rows)).count("Impossible date"), 2)
        iss = self.run_rows(rows, max_future_days=1000, min_year=1990)
        self.assertNotIn("Impossible date", checks(iss))

    def test_missing_key_values(self):
        rows = [[1, None, "Driver", "1", D(2026, 10, 1), 1],
                [2, "Sara B", "Mason", "2", None, 1]]
        self.assertEqual(checks(self.run_rows(rows)).count("Missing key value"), 2)

    def test_empty_column(self):
        rows = [[1, "Ali A", "Driver", None, D(2026, 10, 1), 1]]
        self.assertIn("Empty column", checks(self.run_rows(rows)))

    def test_name_spaces_and_job_conflict(self):
        rows = [[1, "Omar D ", "Carpenter", "1", D(2026, 10, 1), 1],
                [2, "Omar D", "Foreman", "1", D(2026, 10, 2), 1]]
        c = checks(self.run_rows(rows))
        self.assertIn("Extra spaces in name", c)
        self.assertIn("Same name, different job", c)

    def test_text_in_number_column(self):
        rows = [[1, "Ali A", "Driver", "1", D(2026, 10, 1), "2 + tape"]]
        self.assertIn("Text in number column", checks(self.run_rows(rows)))

    def test_id_word_does_not_match_valid(self):
        self.assertTrue(hc.has_word("ID Card", hc.ID_WORDS))
        self.assertFalse(hc.has_word("Valid Until", hc.ID_WORDS))

    def test_arabic_headers(self):
        f = self.p("ar.xlsx")
        make_book(f, ["م", "الاسم", "المهنة", "تاريخ التسليم", "العدد"],
                  [[1, "أحمد", "نجار", D(2026, 10, 1), 1],
                   [2, "أحمد", "سائق", D(2026, 10, 1), "٢ شريط"]])
        c = checks(issues_of(f))
        self.assertIn("Same name, different job", c)
        self.assertIn("Text in number column", c)


class TestExpiry(Base):
    def test_expired_and_soon(self):
        f = self.p("e.xlsx")
        make_book(f, ["Item", "Code", "Expiry Date"],
                  [["A", "1", D(2026, 10, 4)], ["B", "2", D(2026, 11, 30)],
                   ["C", "3", D(2027, 6, 1)]])
        c = checks(issues_of(f))
        self.assertEqual(c.count("Expired"), 1)
        self.assertEqual(c.count("Expires soon"), 1)

    def test_warn_days_config(self):
        f = self.p("e.xlsx")
        make_book(f, ["Item", "Code", "Expiry Date"], [["B", "2", D(2026, 11, 30)]])
        self.assertNotIn("Expires soon", checks(issues_of(f, warn_days=3)))

    def test_as_of_changes_result(self):
        f = self.p("e.xlsx")
        make_book(f, ["Item", "Code", "Expiry Date"], [["B", "2", D(2026, 12, 15)]])
        self.assertEqual(checks(issues_of(f)).count("Expired"), 0)
        wb = load_workbook(f, data_only=True)
        late = hc.Checker(wb.active, hc.Config(D(2027, 1, 1))).run()
        self.assertEqual(checks(late).count("Expired"), 1)


class TestCli(Base):
    def make_input(self, rows=None):
        f = self.p("in.xlsx")
        make_book(f, ["No", "Employee Name", "Issue Date"],
                  rows or [[1, "Ali A", D(2026, 10, 1)], [2, "Ali A", D(2026, 10, 1)]])
        return f

    def test_input_is_never_changed(self):
        f = self.make_input()
        before = md5(f)
        hc.main([f, "--as-of", "2026-11-25"])
        self.assertEqual(md5(f), before)
        self.assertTrue(os.path.exists(self.p("in_health_report.xlsx")))

    def test_report_has_watermark(self):
        from openpyxl import load_workbook
        f = self.make_input()
        out = self.p("w.xlsx")
        hc.main([f, "--out", out, "--as-of", "2026-11-25"])
        wb = load_workbook(out)
        self.assertEqual(wb["Summary"]["A5"].value, "Hussein Abdallah")
        self.assertEqual(wb["Issues"]["J1"].value, "Hussein Abdallah")
        for sh in wb.worksheets:
            self.assertEqual(sh.oddFooter.center.text, "Hussein Abdallah")

    def test_refuses_to_overwrite_input(self):
        f = self.make_input()
        before = md5(f)
        self.assertEqual(hc.main([f, "--out", f, "--force"]), 2)
        self.assertEqual(md5(f), before)

    def test_refuses_existing_report_without_force(self):
        f = self.make_input()
        out = self.p("r.xlsx")
        self.assertIn(hc.main([f, "--out", out]), (0, 1))
        stamp = md5(out)
        self.assertEqual(hc.main([f, "--out", out]), 2)
        self.assertEqual(md5(out), stamp)
        self.assertIn(hc.main([f, "--out", out, "--force"]), (0, 1))

    def test_bad_inputs_return_2(self):
        self.assertEqual(hc.main([self.p("missing.xlsx")]), 2)
        t = self.p("a.csv")
        with open(t, "w") as fh:
            fh.write("a,b\n1,2\n")
        self.assertEqual(hc.main([t]), 2)
        f = self.make_input()
        self.assertEqual(hc.main([f, "--as-of", "not-a-date"]), 2)
        self.assertEqual(hc.main([f, "--sheet", "Nope"]), 2)
        bad = self.p("bad.xlsx")
        with open(bad, "w") as fh:
            fh.write("not a workbook")
        self.assertEqual(hc.main([bad]), 2)

    def test_exit_code_reflects_high_issues(self):
        clean = self.make_input([[1, "Ali A", D(2026, 10, 1)], [2, "Sara B", D(2026, 10, 2)]])
        self.assertEqual(hc.main([clean, "--as-of", "2026-11-25", "--out", self.p("o1.xlsx")]), 0)
        dirty = self.make_input()
        self.assertEqual(hc.main([dirty, "--as-of", "2026-11-25", "--out", self.p("o2.xlsx")]), 1)

    def test_formula_text_in_input_stays_text_in_report(self):
        f = self.p("f.xlsx")
        make_book(f, ["No", "Employee Name", "Job Title", "Qty", "Issue Date"],
                  [[1, "Ali A", "Driver", "=SUM(A1:A2)", D(2026, 10, 1)],
                   [2, "@cmd", "Driver", "+1", D(2026, 10, 1)]])
        # force real text cells in the input (not formulas)
        wb = load_workbook(f)
        for c in ("D2", "D3"):
            wb.active[c].data_type = "s"
        wb.save(f)
        out = self.p("fr.xlsx")
        hc.main([f, "--as-of", "2026-11-25", "--out", out])
        with zipfile.ZipFile(out) as z:
            for n in z.namelist():
                if n.startswith("xl/worksheets/"):
                    self.assertNotIn(b"<f>", z.read(n))
        ws = load_workbook(out)["Issues"]
        vals = [c.value for row in ws.iter_rows(min_row=2) for c in row]
        self.assertIn("=SUM(A1:A2)", vals)
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value, str) and c.value[:1] in "=+-@":
                    self.assertEqual(c.data_type, "s")

    def test_uncalculated_formulas_are_reported(self):
        f = self.p("u.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.append(["Employee Name", "Issue Date", "Qty"])
        ws.append(["Ali A", D(2026, 10, 1), "=1+1"])
        wb.save(f)
        out = self.p("ur.xlsx")
        hc.main([f, "--out", out])
        wbf = load_workbook(f, data_only=False)
        wbv = load_workbook(f, data_only=True)
        self.assertEqual(hc.count_uncalculated(wbf.active, wbv.active), 1)
        text = [c.value for row in load_workbook(out)["Summary"].iter_rows() for c in row if c.value]
        self.assertTrue(any("no saved result" in str(t) for t in text))

    def test_header_row_after_title_rows(self):
        f = self.p("t.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.append(["Dummy register title"])
        ws.append([])
        ws.append(["No", "Employee Name", "Issue Date"])
        ws.append([1, "Ali A", D(2026, 10, 1)])
        ws.append([2, "Ali A", D(2026, 10, 1)])
        wb.save(f)
        iss = [i for i in issues_of(f) if i["Check"] == "Duplicate row"]
        self.assertEqual(iss[0]["Row"], 5)

    def test_empty_sheet_does_not_crash(self):
        f = self.p("e.xlsx")
        wb = Workbook()
        wb.active.title = "Data"
        wb.active.append(["Employee Name", "Issue Date"])
        wb.active.append(["Ali A", D(2026, 10, 1)])
        wb.create_sheet("Empty")
        wb.save(f)
        self.assertIn(hc.main([f, "--out", self.p("er.xlsx")]), (0, 1))


if __name__ == "__main__":
    unittest.main()
