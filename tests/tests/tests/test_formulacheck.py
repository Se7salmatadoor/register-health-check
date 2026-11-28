import hashlib
import os
import sys
import tempfile
import unittest

from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import formulacheck as fc  # noqa: E402
import healthcheck as hc  # noqa: E402


def pair(build, values=None):
    """Formula sheet from build(ws); values sheet with optional saved results {coord: value}."""
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    build(ws)
    wv = Workbook()
    v = wv.active
    v.title = "S"
    for k, val in (values or {}).items():
        v[k] = val
    return ws, v, ["S"]


def run(build, values=None):
    ws, v, names = pair(build, values)
    return fc.audit_sheet(ws, v, names)


def checks(iss):
    return [i["Check"] for i in iss]


def column(ws, n=6, col="B"):
    for r in range(1, n + 1):
        ws["A%d" % r] = r
        ws["%s%d" % (col, r)] = "=A%d*2" % r


class TestFormulaAudit(unittest.TestCase):
    def test_clean_sheet(self):
        self.assertEqual(run(column), [])

    def test_error_value(self):
        iss = run(lambda ws: ws.__setitem__("A1", "=1+1"), {"A1": "#DIV/0!"})
        self.assertIn("Formula error value", checks(iss))

    def test_broken_reference(self):
        self.assertIn("Broken reference", checks(run(lambda ws: ws.__setitem__("A1", "=#REF!+1"))))

    def test_division_by_zero(self):
        self.assertIn("Division by zero", checks(run(lambda ws: ws.__setitem__("A1", "=5/0"))))
        self.assertNotIn("Division by zero", checks(run(lambda ws: ws.__setitem__("A1", "=5/0.5"))))

    def test_missing_sheet(self):
        self.assertIn("Missing sheet", checks(run(lambda ws: ws.__setitem__("A1", "=Nope!A1"))))
        self.assertNotIn("Missing sheet", checks(run(lambda ws: ws.__setitem__("A1", "=S!B1"))))

    def test_external_link(self):
        self.assertIn("External link", checks(run(lambda ws: ws.__setitem__("A1", "='[1]Sheet1'!A1"))))

    def test_volatile(self):
        self.assertIn("Volatile function", checks(run(lambda ws: ws.__setitem__("A1", "=TODAY()"))))

    def test_circular(self):
        self.assertIn("Circular reference", checks(run(lambda ws: ws.__setitem__("A3", "=SUM(A1:A3)"))))

    def test_inconsistent(self):
        def b(ws):
            column(ws)
            ws["B3"] = "=A2*2"
        iss = run(b)
        self.assertEqual([i["Column"] for i in iss if i["Check"] == "Inconsistent formula"], ["B3"])

    def test_totals_row_not_flagged(self):
        def b(ws):
            column(ws)
            ws["B7"] = "=SUM(B1:B6)"
        self.assertEqual(run(b), [])

    def test_hardcoded(self):
        def b(ws):
            column(ws)
            ws["B3"] = 999
        iss = run(b)
        self.assertEqual([i["Column"] for i in iss if i["Check"] == "Hardcoded number in formula column"], ["B3"])

    def test_short_range(self):
        def b(ws):
            for r in range(1, 6):
                ws["A%d" % r] = r
            ws["C1"] = "=SUM(A1:A4)"
        self.assertIn("Range misses rows", checks(run(b)))

    def test_range_ok(self):
        def b(ws):
            for r in range(1, 5):
                ws["A%d" % r] = r
            ws["A5"] = "=SUM(A1:A4)"
        self.assertEqual(run(b), [])

    def test_number_text_in_range(self):
        def b(ws):
            ws["A1"], ws["A2"], ws["C1"] = "5", 3, "=SUM(A1:A2)"
        self.assertIn("Numbers stored as text", checks(run(b)))

    def test_empty_reference(self):
        self.assertIn("Reference to empty cell", checks(run(lambda ws: ws.__setitem__("A1", "=Z9+1"))))

    def test_text_in_formula_string_ignored(self):
        self.assertEqual(run(lambda ws: ws.__setitem__("A1", '="#REF! and /0 and Z9"')), [])


class TestCli(unittest.TestCase):
    def test_original_untouched_and_report(self):
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "f.xlsx")
            wb = Workbook()
            ws = wb.active
            ws.title = "S"
            column(ws)
            ws["B3"] = 999
            wb.save(src)
            before = hashlib.md5(open(src, "rb").read()).hexdigest()
            out = os.path.join(d, "r.xlsx")
            code = hc.main([src, "--formulas", "--out", out])
            self.assertEqual(code, 1)
            self.assertEqual(before, hashlib.md5(open(src, "rb").read()).hexdigest())
            rows = list(load_workbook(out)["Issues"].iter_rows(min_row=2, values_only=True))
            self.assertIn("Hardcoded number in formula column", [r[1] for r in rows])

    def test_off_by_default(self):
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "f.xlsx")
            wb = Workbook()
            ws = wb.active
            column(ws)
            ws["B3"] = 999
            wb.save(src)
            out = os.path.join(d, "r.xlsx")
            hc.main([src, "--out", out])
            rows = list(load_workbook(out)["Issues"].iter_rows(min_row=2, values_only=True))
            self.assertNotIn("Hardcoded number in formula column", [r[1] for r in rows])


if __name__ == "__main__":
    unittest.main()
