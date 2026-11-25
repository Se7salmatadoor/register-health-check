"""Tests for the manual header row / column roles and the local screen (prototype)."""
import datetime as dt
import http.client
import io
import os
import sys
import tempfile
import threading
import unittest
import uuid
from contextlib import redirect_stderr, redirect_stdout

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
import healthcheck as hc  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402

D = dt.datetime


def odd_register(path):
    """Title rows, headers on row 4, header words the auto guess does not know."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Cards"
    ws["A1"] = "Site safety cards (dummy)"
    ws.append([])
    ws.append([])
    ws.append(["Holder", "Issued on", "Good until", "Trade"])
    ws.append(["Ali Dummy", D(2026, 1, 5), D(2026, 9, 1), "Welder"])    # expired
    ws.append(["Sara Dummy", D(2026, 2, 5), D(2027, 9, 1), "Welder"])
    ws.append(["Ali Dummy", D(2026, 1, 5), D(2026, 9, 1), "Welder"])    # duplicate
    wb.save(path)


def run(argv):
    o, e = io.StringIO(), io.StringIO()
    with redirect_stdout(o), redirect_stderr(e):
        code = hc.main(argv)
    return code, o.getvalue(), e.getvalue()


class TestManualRoles(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.f = os.path.join(self.tmp.name, "odd.xlsx")
        odd_register(self.f)

    def tearDown(self):
        self.tmp.cleanup()

    def issues(self, out):
        ws = load_workbook(out)["Issues"]
        return [[c.value for c in r] for r in ws.iter_rows(min_row=2)]

    def test_auto_guess_misses_the_dates_on_odd_headers(self):
        out = os.path.join(self.tmp.name, "a.xlsx")
        run([self.f, "--out", out, "--as-of", "2026-11-25"])
        self.assertNotIn("Expired", {r[1] for r in self.issues(out)})

    def test_manual_header_and_roles_find_the_problems(self):
        out = os.path.join(self.tmp.name, "m.xlsx")
        code, _, err = run([self.f, "--sheet", "Cards", "--header-row", "4",
                            "--role", "A=name", "--role", "B=date", "--role", "C=expiry",
                            "--out", out, "--as-of", "2026-11-25"])
        self.assertEqual(code, 1, err)
        checks = {r[1] for r in self.issues(out)}
        self.assertIn("Expired", checks)
        self.assertIn("Duplicate row", checks)

    def test_ignore_role_removes_a_column(self):
        out = os.path.join(self.tmp.name, "i.xlsx")
        run([self.f, "--sheet", "Cards", "--header-row", "4", "--role", "A=name",
             "--role", "B=date", "--role", "C=ignore", "--out", out, "--as-of", "2026-11-25"])
        self.assertNotIn("Expired", {r[1] for r in self.issues(out)})

    def test_wrong_header_row_does_not_crash(self):
        out = os.path.join(self.tmp.name, "w.xlsx")
        code, _, err = run([self.f, "--sheet", "Cards", "--header-row", "2", "--out", out])
        self.assertEqual(code, 2)
        self.assertIn("no headers", err)
        code, _, err = run([self.f, "--sheet", "Cards", "--header-row", "99", "--out", out])
        self.assertEqual(code, 2)
        code, _, err = run([self.f, "--sheet", "Cards", "--header-row", "0", "--out", out])
        self.assertEqual(code, 2)

    def test_bad_role_and_missing_sheet_are_refused(self):
        out = os.path.join(self.tmp.name, "b.xlsx")
        self.assertEqual(run([self.f, "--sheet", "Cards", "--role", "A=boss", "--out", out])[0], 2)
        self.assertEqual(run([self.f, "--sheet", "Cards", "--role", "1=name", "--out", out])[0], 2)
        self.assertEqual(run([self.f, "--role", "A=name", "--out", out])[0], 2)

    def test_guess_roles_on_normal_headers(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["Employee Name", "Issue Date", "Expiry Date", "Job Title"])
        ws.append(["A", D(2026, 1, 1), D(2027, 1, 1), "x"])
        g = hc.Checker(ws, hc.Config(dt.date(2026, 11, 25))).guess_roles()
        self.assertEqual(g, {1: "name", 2: "date", 3: "expiry", 4: "job"})


def multipart(fields, files=None):
    b = uuid.uuid4().hex
    out = b""
    for k, v in fields:
        out += ("--%s\r\nContent-Disposition: form-data; name=\"%s\"\r\n\r\n%s\r\n" % (b, k, v)).encode()
    for k, fn, data in files or []:
        out += ("--%s\r\nContent-Disposition: form-data; name=\"%s\"; filename=\"%s\"\r\n"
                "Content-Type: application/octet-stream\r\n\r\n" % (b, k, fn)).encode() + data + b"\r\n"
    out += ("--%s--\r\n" % b).encode()
    return "multipart/form-data; boundary=" + b, out


class TestScreen(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ["HC_NO_BROWSER"] = "1"
        import healthcheck_gui as g
        cls.g = g
        cls.srv = g.ThreadingHTTPServer(("127.0.0.1", 0), g.Handler)
        cls.port = cls.srv.server_address[1]
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()
        cls.tmp = tempfile.TemporaryDirectory()
        cls.odd = os.path.join(cls.tmp.name, "odd.xlsx")
        odd_register(cls.odd)

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()
        cls.g.cleanup_uploads(force=True)
        cls.tmp.cleanup()

    def post(self, fields, files=None):
        ct, body = multipart(fields, files)
        c = http.client.HTTPConnection("127.0.0.1", self.port)
        c.request("POST", "/check", body, {"Content-Type": ct, "Content-Length": str(len(body))})
        r = c.getresponse()
        return r.status, r.read().decode("utf-8")

    def upload_for_preview(self, **extra):
        with open(self.odd, "rb") as fh:
            data = fh.read()
        st, page = self.post([("action", "preview"), ("as_of", "2026-11-25")], [("file", "odd.xlsx", data)])
        self.assertEqual(st, 200)
        import re
        return page, re.search(r"name='token' value='([^']+)'", page).group(1)

    def test_preview_shows_header_row_and_role_menus(self):
        page, tok = self.upload_for_preview()
        self.assertIn("Holder", page)
        self.assertEqual(page.count("<select name='role_"), 4)
        self.assertIn("name='header_row' value='4'", page)        # auto-detected row 4
        self.assertIn("تلقائي (مش متعرّف عليه)", page)           # odd headers not guessed

    def test_manual_run_end_to_end(self):
        page, tok = self.upload_for_preview()
        st, res = self.post([("action", "run"), ("token", tok), ("as_of", "2026-11-25"), ("warn_days", "10"),
                             ("sheet", "Cards"), ("header_row", "4"),
                             ("role_1", "name"), ("role_2", "date"), ("role_3", "expiry"), ("role_4", "auto")])
        self.assertEqual(st, 200)
        self.assertIn("Expired", res)
        self.assertIn("Duplicate row", res)
        self.assertIn("/report/", res)
        self.assertEqual(self.g.UPLOADS.get(tok), None)            # uploaded copy is gone

    def test_update_preview_changes_header_row(self):
        page, tok = self.upload_for_preview()
        st, p2 = self.post([("action", "preview"), ("token", tok), ("sheet", "Cards"), ("header_row", "1")])
        self.assertIn("name='header_row' value='1'", p2)
        self.assertEqual(p2.count("<select name='role_"), 1)       # row 1 has one header cell
        st, p3 = self.post([("action", "preview"), ("token", tok), ("sheet", "Cards"), ("header_row", "2")])
        self.assertIn("الصف ده فاضي", p3)

    def test_pages_are_well_formed(self):
        page, tok = self.upload_for_preview()
        self.assertEqual(page.count("<style>"), page.count("</style>"))
        self.assertIn("</html>", page)
        c = http.client.HTTPConnection("127.0.0.1", self.port)
        c.request("GET", "/")
        home = c.getresponse().read().decode()
        self.assertEqual(home.count("<style>"), home.count("</style>"))

    def test_expired_token_and_bad_file(self):
        st, p = self.post([("action", "run"), ("token", "nope")])
        self.assertIn("انتهت مدته", p)
        st, p = self.post([("action", "preview")], [("file", "a.csv", b"x")])
        self.assertIn("xlsx", p)
        st, p = self.post([("action", "run")], [("file", "bad.xlsx", b"junk")])
        self.assertIn("cannot open", p)

    def test_html_is_escaped_in_preview(self):
        wb = Workbook()
        ws = wb.active
        ws.append(["<script>alert(1)</script>", "Issue Date"])
        ws.append(["x", D(2026, 1, 1)])
        p = os.path.join(self.tmp.name, "x.xlsx")
        wb.save(p)
        with open(p, "rb") as fh:
            st, page = self.post([("action", "preview")], [("file", "x.xlsx", fh.read())])
        self.assertNotIn("<script>alert", page)

    def test_old_upload_is_removed(self):
        page, tok = self.upload_for_preview()
        self.g.UPLOADS[tok]["time"] -= self.g.UPLOAD_TTL + 5
        self.g.cleanup_uploads()
        self.assertNotIn(tok, self.g.UPLOADS)


if __name__ == "__main__":
    unittest.main()
