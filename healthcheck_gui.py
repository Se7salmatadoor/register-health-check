#!/usr/bin/env python3
"""Register Health-Check: simple local screen (prototype).

Run:  python healthcheck_gui.py
It opens a page in your browser at http://127.0.0.1:8765 . Choose an Excel
register, press the check button (or set the header row and the columns by hand
first), read the result, download the report.

- Everything stays on your computer. The page is served only to 127.0.0.1
  (this computer), nothing is sent to the internet.
- Your file is copied to a temporary folder, checked, and the copy is deleted
  right after the check (or after 30 minutes if you walk away). Your original file is never changed.
- The checks are the ones in healthcheck.py (keep both files in one folder).
- Needs Python 3 (tested with 3.10) and: pip install openpyxl
Stop it with Ctrl+C in the window where you started it.

© 2026 Hussein Abdallah. All rights reserved.
"""
import datetime as dt
import html
import io
import os
import re
import secrets
import shutil
import sys
import tempfile
import threading
import time
import webbrowser
from contextlib import redirect_stderr, redirect_stdout
from email.parser import BytesParser
from email.policy import HTTP as HTTP_POLICY
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import healthcheck as hc  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

HOST = "127.0.0.1"
PORT = int(os.environ.get("HC_PORT", "8765"))
MAX_UPLOAD = 50 * 1024 * 1024
MAX_ROWS_SHOWN = 60
WORK = tempfile.mkdtemp(prefix="healthcheck_")
REPORTS = {}  # token -> (path, download name)
UPLOADS = {}  # token -> dict(path, name, dir, time) waiting for the column screen
UPLOAD_TTL = 30 * 60
LOCK = threading.Lock()

CSS = """
body{font-family:Segoe UI,Tahoma,Arial,sans-serif;background:#f4f6fa;margin:0;color:#1b2430}
.wrap{max-width:900px;margin:0 auto;padding:20px}
h1{color:#1F3864;font-size:26px;margin:8px 0}
.card{background:#fff;border:1px solid #d7dde8;border-radius:10px;padding:18px;margin:14px 0}
label{display:block;margin:12px 0 4px;font-weight:600}
input[type=file],input[type=text],input[type=number],input[type=date]{width:100%;box-sizing:border-box;padding:9px;border:1px solid #b9c2d3;border-radius:6px;font-size:15px}
.row{display:flex;gap:12px;flex-wrap:wrap}.row>div{flex:1;min-width:150px}
button,.btn{background:#1F3864;color:#fff;border:0;border-radius:8px;padding:12px 22px;font-size:17px;cursor:pointer;text-decoration:none;display:inline-block}
.small{color:#5b6677;font-size:13px}
.err{background:#fde8e8;border:1px solid #e3a1a1;border-radius:8px;padding:12px}
.warn{background:#fff6dc;border:1px solid #e6cd7a;border-radius:8px;padding:10px;margin-top:10px}
table{border-collapse:collapse;width:100%;font-size:13px}
th{background:#1F3864;color:#fff;text-align:left;padding:6px}
td{border:1px solid #d7dde8;padding:5px;vertical-align:top}
.High{background:#F8CBAD}.Medium{background:#FFE699}.Low{background:#DDEBF7}.Info{background:#EDEDED}
.nums span{display:inline-block;margin:4px 8px 4px 0;padding:6px 12px;border-radius:6px;font-weight:600}
.pv{overflow-x:auto}.pv td,.pv th{white-space:nowrap}.pv th{background:#1F3864}.hdr td{background:#dbe4f5;font-weight:700}.hdrrow{outline:3px solid #1F3864}
select{padding:7px;border:1px solid #b9c2d3;border-radius:6px;font-size:13px;max-width:150px}
.btn2{background:#5b6677}
.mark{color:#a6a6a6;font-style:italic;font-size:13px;text-align:center;margin:24px 0 8px}
"""


def page(body):
    return ("<!doctype html><html lang='ar' dir='rtl'><head><meta charset='utf-8'>"
            "<meta name='viewport' content='width=device-width,initial-scale=1'>"
            "<title>Register Health-Check</title><style>%s</style></head><body><div class='wrap'>"
            "<h1>فحص صحة الريجستر (Register Health-Check)</h1>%s"
            "<div class='mark'>Hussein Abdallah</div></div></body></html>" % (CSS, body))


def form(message=""):
    today = dt.date.today().isoformat()
    return page("""%s
<div class='card'>
<form method='post' action='/check' enctype='multipart/form-data'>
<label>1) اختار ملف الريجستر (Excel: xlsx أو xlsm)</label>
<input type='file' name='file' accept='.xlsx,.xlsm' required>
<div class='row'>
<div><label>تاريخ المراجعة</label><input type='date' name='as_of' value='%s'></div>
<div><label>تنبيه قبل الانتهاء (أيام)</label><input type='number' name='warn_days' value='10' min='0' max='3650'></div>
</div>
<label>شيت واحد بس (اختياري، اكتب اسمه بالظبط، وللضبط اليدوي بتختاره في الخطوة الجاية)</label>
<input type='text' name='sheet' placeholder='سيبه فاضي لفحص كل الشيتات'>
<label><input type='checkbox' name='formulas' value='1' checked> افحص المعادلات كمان (أخطاء، مراجع مكسورة، صيغ مش متناسقة، أرقام مكتوبة مكان معادلة، نطاقات ناقصة)</label>
<label><input type='checkbox' name='month_first' value='1'> التواريخ المكتوبة كنص بالشهر الأول (03/04/2026 = 4 مارس)</label>
<p><button type='submit' name='action' value='run'>2) افحص مباشرة (تلقائي)</button> <button type='submit' name='action' value='preview' class='btn2'>أو: اضبط صف العناوين والأعمدة بنفسي</button></p>
</form>
<p class='small'>الملف بيتفحص على جهازك بس، بيتنسخ مؤقتًا وبيتمسح بعد الفحص، والأصل ما بيتغيرش.
النتيجة ملاحظات للمراجعة مش قرارات نهائية.</p>
</div>""" % (message, today))


def parse_multipart(content_type, raw):
    msg = BytesParser(policy=HTTP_POLICY).parsebytes(
        b"Content-Type: " + content_type.encode("latin-1") + b"\r\nMIME-Version: 1.0\r\n\r\n" + raw)
    fields, files = {}, {}
    if not msg.is_multipart():
        return fields, files
    for part in msg.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        fn = part.get_filename()
        data = part.get_payload(decode=True) or b""
        if fn is not None:
            files[name] = (fn, data)
        else:
            fields[name] = data.decode("utf-8", "replace").strip()
    return fields, files


def safe_name(fn):
    base = os.path.basename(fn.replace("\\", "/"))
    base = re.sub(r"[^\w.\- ()]", "_", base, flags=re.UNICODE)
    return base or "register.xlsx"


BAD_TYPE = "الملف لازم يكون Excel بامتداد xlsx أو xlsm (الملفات القديمة xls لازم تتحفظ كـ xlsx الأول)."


def cleanup_uploads(force=False):
    now = time.time()
    with LOCK:
        for tok in list(UPLOADS):
            u = UPLOADS[tok]
            if force or now - u["time"] > UPLOAD_TTL:
                shutil.rmtree(u["dir"], ignore_errors=True)
                del UPLOADS[tok]


def save_upload(fn, data):
    """Copy the chosen file into a temp job folder. Returns (token, error)."""
    cleanup_uploads()
    name = safe_name(fn)
    if not name.lower().endswith((".xlsx", ".xlsm")):
        return None, BAD_TYPE
    job = tempfile.mkdtemp(dir=WORK)
    src = os.path.join(job, name)
    with open(src, "wb") as fh:
        fh.write(data)
    token = secrets.token_urlsafe(16)
    with LOCK:
        UPLOADS[token] = {"path": src, "name": name, "dir": job, "time": time.time()}
    return token, None


def drop_upload(token):
    with LOCK:
        u = UPLOADS.pop(token, None)
    if u:
        shutil.rmtree(u["dir"], ignore_errors=True)


def base_argv(src, out, fields):
    argv = [src, "--out", out]
    if fields.get("as_of"):
        argv += ["--as-of", fields["as_of"]]
    try:
        wd = int(fields.get("warn_days") or 10)
        argv += ["--warn-days", str(max(0, min(wd, 3650)))]
    except ValueError:
        pass
    if fields.get("month_first"):
        argv += ["--month-first"]
    if fields.get("formulas"):
        argv += ["--formulas"]
    return argv


def run_check(token, fields, extra=()):
    """Run the checker on a saved upload, then delete the upload.
    Returns (ok, text, report_token, issue_rows, counts, notes)."""
    u = UPLOADS.get(token)
    if not u:
        return False, "الملف انتهت مدته أو اتمسح. اختاره تاني.", None, [], {}, []
    name = u["name"]
    stem = re.sub(r"\.xls[xm]$", "", name, flags=re.I)
    out = os.path.join(u["dir"], stem + "_health_report.xlsx")
    argv = base_argv(u["path"], out, fields) + list(extra)
    buf_o, buf_e = io.StringIO(), io.StringIO()
    try:
        with redirect_stdout(buf_o), redirect_stderr(buf_e):
            code = hc.main(argv)
    except Exception as e:  # keep the screen alive on unexpected errors
        code = 2
        buf_e.write("Unexpected error: %s" % e)
    try:
        os.remove(u["path"])  # the uploaded copy is not kept
    except OSError:
        pass
    if code == 2 or not os.path.exists(out):
        # keep nothing: a failed run needs the file chosen again
        drop_upload(token)
        return False, (buf_e.getvalue() or buf_o.getvalue()).strip(), None, [], {}, []
    rows, counts, notes = [], {}, []
    wb = load_workbook(out, data_only=True)
    for r in list(wb["Issues"].iter_rows(min_row=2, values_only=True)):
        if r[0] is None:
            continue
        counts[r[0]] = counts.get(r[0], 0) + 1
        rows.append(r[:8])
    for line in buf_o.getvalue().splitlines():
        if "Warning" in line:
            notes.append(line.strip())
    rtoken = secrets.token_urlsafe(16)
    keep = tempfile.mkdtemp(dir=WORK)
    kept = os.path.join(keep, os.path.basename(out))
    shutil.move(out, kept)
    drop_upload(token)
    with LOCK:
        REPORTS[rtoken] = (kept, os.path.basename(kept))
    return True, buf_o.getvalue(), rtoken, rows, counts, notes


def show(v):
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.date().isoformat() if v.time() == dt.time(0) else v.isoformat(sep=" ")
    return str(v)[:30]


def preview_page(token, fields, message=""):
    """Column screen: sheet, header row, a role for every header."""
    e = html.escape
    u = UPLOADS.get(token)
    if not u:
        return form("<div class='err'>الملف انتهت مدته. اختاره تاني.</div>")
    try:
        wb = load_workbook(u["path"], data_only=True, read_only=False)
    except Exception as ex:
        drop_upload(token)
        return form("<div class='err' dir='ltr'>cannot open the workbook: %s</div>" % e(str(ex)))
    sheets = wb.sheetnames
    sheet = fields.get("sheet") if fields.get("sheet") in sheets else sheets[0]
    ws = wb[sheet]
    auto_hdr = hc.find_header_row(ws)
    try:
        hdr = int(fields.get("header_row") or auto_hdr)
    except ValueError:
        hdr = auto_hdr
    hdr = max(1, min(hdr, max(ws.max_row, 1)))
    cfg = hc.Config(dt.date.today(), header_row=hdr)
    ck = hc.Checker(ws, cfg)
    guess = ck.guess_roles()
    label = {"date": "تاريخ", "expiry": "تاريخ انتهاء", "name": "اسم", "job": "مهنة", "id": "رقم/كود", "qty": "كمية", "ignore": "تجاهل"}
    opts_sheet = "".join("<option value='%s'%s>%s</option>" % (e(n, True), " selected" if n == sheet else "", e(n)) for n in sheets)
    keep = "".join("<input type='hidden' name='%s' value='%s'>" % (k, e(fields.get(k, ""), True))
                   for k in ("as_of", "warn_days", "month_first"))
    cols = sorted(ck.headers)
    thead = "".join("<th>%s</th>" % get_column_letter(c) for c in cols)
    selects = ""
    for c in cols:
        saved = fields.get("role_%d" % c)
        cur = saved if saved in hc.ROLE_NAMES else "auto"
        g = guess.get(c)
        auto_label = "تلقائي (%s)" % label[g] if g else "تلقائي (مش متعرّف عليه)"
        o = ["<option value='auto'%s>%s</option>" % (" selected" if cur == "auto" else "", e(auto_label))]
        for r in ("date", "expiry", "name", "job", "id", "qty", "ignore"):
            o.append("<option value='%s'%s>%s</option>" % (r, " selected" if cur == r else "", label[r]))
        selects += "<td><select name='role_%d'>%s</select></td>" % (c, "".join(o))
    first = max(1, hdr - 1)
    last = min(ws.max_row, hdr + 5)
    body_rows = ""
    for r in range(first, last + 1):
        cells = "".join("<td>%s</td>" % e(show(ws.cell(r, c).value)) for c in cols)
        cls = " class='hdr'" if r == hdr else ""
        body_rows += "<tr%s><td>%d</td>%s</tr>" % (cls, r, cells)
    note = ""
    if not ck.headers:
        note = "<div class='err'>الصف ده فاضي، اختار صف العناوين الصح.</div>"
    return page("""%s%s
<div class='card'>
<p>الملف: <b>%s</b>. الصف المظلل هو صف العناوين. غيّر رقمه وادوس «حدّث» لو غلط، وبعدين حدد نوع كل عمود.</p>
<form method='post' action='/check' enctype='multipart/form-data'>
<input type='hidden' name='token' value='%s'>%s
<div class='row'>
<div><label>الشيت</label><select name='sheet' style='max-width:none;width:100%%'>%s</select></div>
<div><label>رقم صف العناوين (المكتشف تلقائيًا: %d)</label><input type='number' name='header_row' value='%d' min='1'></div>
</div>
<p><button type='submit' name='action' value='preview' class='btn2'>حدّث المعاينة</button></p>
<div class='pv' dir='ltr'><table><tr><th>Row</th>%s</tr>
<tr><td></td>%s</tr>%s</table></div>
<p class='small'>«تلقائي» يعني الأداة تخمّن من اسم العمود. اختار نوع العمود لو التخمين غلط، أو «تجاهل» لو مش عايزه يتفحص.
الفحص بيتم على الشيت ده بس.</p>
<p><button type='submit' name='action' value='run'>افحص الشيت ده</button></p>
</form></div>""" % (message, note, e(u["name"]), e(token, True), keep, opts_sheet, auto_hdr, hdr, thead, selects, body_rows))


def result_page(fn, token, rows, counts, notes):
    e = html.escape
    nums = "".join("<span class='%s'>%s: %d</span>" % (k, k, counts.get(k, 0))
                   for k in ("High", "Medium", "Low", "Info"))
    body = ["<div class='card'><p>الملف: <b>%s</b> (ما اتغيرش)</p>" % e(safe_name(fn)),
            "<p>عدد الملاحظات: <b>%d</b></p><div class='nums' dir='ltr'>%s</div>" % (len(rows), nums)]
    for n in notes:
        body.append("<div class='warn' dir='ltr'>%s</div>" % e(n))
    body.append("<p><a class='btn' href='/report/%s'>نزّل التقرير (Excel)</a> "
                "<a class='btn' style='background:#5b6677' href='/'>افحص ملف تاني</a></p></div>" % token)
    if rows:
        body.append("<div class='card' dir='ltr'><table><tr><th>Severity</th><th>Check</th><th>Sheet</th>"
                    "<th>Row</th><th>Column</th><th>Value</th><th>Problem</th><th>Suggested fix</th></tr>")
        for r in rows[:MAX_ROWS_SHOWN]:
            body.append("<tr class='%s'>%s</tr>" % (
                e(str(r[0])), "".join("<td>%s</td>" % e("" if v is None else str(v)) for v in r)))
        body.append("</table>")
        if len(rows) > MAX_ROWS_SHOWN:
            body.append("<p class='small'>الظاهر %d ملاحظة من %d، الباقي في تقرير Excel.</p>" % (MAX_ROWS_SHOWN, len(rows)))
        body.append("</div>")
    else:
        body.append("<div class='card'>ما لقيتش ملاحظات. ده مش ضمان إن الريجستر سليم، الأداة بتفحص قواعد محددة بس.</div>")
    return page("".join(body))


class Handler(BaseHTTPRequestHandler):
    server_version = "HealthCheckLocal"

    def log_message(self, *a):
        pass

    def _host_ok(self):
        host = (self.headers.get("Host") or "").split(":")[0]
        origin = self.headers.get("Origin")
        if host not in ("127.0.0.1", "localhost"):
            return False
        if origin and re.sub(r"^https?://", "", origin).split(":")[0] not in ("127.0.0.1", "localhost"):
            return False
        return True

    def _send(self, code, body, ctype="text/html; charset=utf-8", extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if not self._host_ok():
            return self._send(403, "Forbidden", "text/plain")
        if self.path == "/":
            return self._send(200, form())
        m = re.match(r"^/report/([\w\-]+)$", self.path)
        if m and m.group(1) in REPORTS:
            path, name = REPORTS[m.group(1)]
            with open(path, "rb") as fh:
                data = fh.read()
            return self._send(200, data,
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              {"Content-Disposition": "attachment; filename*=UTF-8''%s" % __import__("urllib.parse").parse.quote(name)})
        self._send(404, "Not found", "text/plain")

    def do_POST(self):
        if not self._host_ok() or self.path != "/check":
            return self._send(403, "Forbidden", "text/plain")
        try:
            n = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            n = 0
        if n <= 0 or n > MAX_UPLOAD:
            return self._send(413, form("<div class='err'>الملف كبير أوي أو فاضي (الحد 50 ميجا).</div>"))
        raw = self.rfile.read(n)
        fields, files = parse_multipart(self.headers.get("Content-Type", ""), raw)
        action = fields.get("action", "run")
        token = fields.get("token", "")
        if token:                                   # second screen: file already chosen
            if token not in UPLOADS:
                return self._send(200, form("<div class='err'>الملف انتهت مدته أو اتمسح. اختاره تاني.</div>"))
            fn = UPLOADS[token]["name"]
        else:                                       # first screen: new file
            if "file" not in files or not files["file"][0]:
                return self._send(400, form("<div class='err'>اختار ملف Excel الأول.</div>"))
            fn, data = files["file"]
            token, err = save_upload(fn, data)
            if err:
                return self._send(200, form("<div class='err' dir='ltr'>%s</div>" % html.escape(err)))
            if fields.get("sheet") and action == "preview":
                pass
        if action == "preview":
            return self._send(200, preview_page(token, fields))
        extra = []
        if "header_row" in fields or any(k.startswith("role_") for k in fields):
            # manual screen: check only the chosen sheet with the chosen header row and roles
            if fields.get("sheet"):
                extra += ["--sheet", fields["sheet"]]
            if fields.get("header_row"):
                extra += ["--header-row", fields["header_row"]]
            for k, v in fields.items():
                m = re.fullmatch(r"role_(\d+)", k)
                if m and v in hc.ROLE_NAMES and v != "auto":
                    extra += ["--role", "%s=%s" % (get_column_letter(int(m.group(1))), v)]
        elif fields.get("sheet"):
            extra += ["--sheet", fields["sheet"]]
        ok, text, rtoken, rows, counts, notes = run_check(token, fields, extra)
        if not ok:
            return self._send(200, form("<div class='err' dir='ltr'>%s</div>" % html.escape(text or "Error")))
        self._send(200, result_page(fn, rtoken, rows, counts, notes))


def main():
    try:
        srv = ThreadingHTTPServer((HOST, PORT), Handler)
    except OSError as e:
        print("Cannot start on port %d (%s). Close the other copy or set HC_PORT." % (PORT, e))
        return 2
    url = "http://%s:%d/" % (HOST, PORT)
    print("Register Health-Check is running at %s" % url)
    print("Stop with Ctrl+C. Files stay on this computer.")
    if not os.environ.get("HC_NO_BROWSER"):
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        cleanup_uploads(force=True)
        shutil.rmtree(WORK, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
