#!/usr/bin/env python3
"""Register Health-Check (prototype).

Reads an Excel register, finds common data problems and writes a SEPARATE
report workbook. The input file is opened read-only and is never changed.

Usage:
    python3 healthcheck.py register.xlsx [--sheet NAME] [--as-of YYYY-MM-DD]
                           [--warn-days 10] [--out report.xlsx]

Checks:
  1. Dates stored as text, and dates that cannot be read
  2. Impossible dates (very old, or far in the future)
  3. Entry dates that go backwards compared with the row above (info)
  4. Exact duplicate rows and same person + same date + same item twice
  5. Missing values (empty columns, blank key cells)
  6. Expired or soon-to-expire certificates / inspections
  7. Names with extra spaces; same name with different job titles
  8. Text inside number columns
Columns are found from header words in English or Arabic. See README.md.
"""
import argparse
import datetime as dt
import os
import re
import sys
from collections import defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter

DATE_WORDS = ("date", "تاريخ")
EXPIRY_WORDS = ("expiry", "expire", "valid until", "valid to", "due date",
                "next inspection", "next calibration", "انتهاء", "صلاحية")
NAME_WORDS = ("name", "الاسم", "اسم")
JOB_WORDS = ("job", "title", "trade", "position", "المهنة", "الوظيفة", "مهنة")
ID_WORDS = ("id", "card", "رقم البطاقة", "بطاقة", "code", "كود")
NUM_WORDS = ("qty", "quantity", "count", "العدد", "الكمية", "عدد")
ROW_NO_HEADERS = {"no", "no.", "#", "م", "s.n", "sn", "s/n", "serial", "sr", "sr.", "ref", "ref.#"}

ROLE_NAMES = ("auto", "date", "expiry", "name", "job", "id", "qty", "ignore")

SEV_ORDER = {"High": 0, "Medium": 1, "Low": 2, "Info": 3}


def has_word(header, words):
    """Match header words. Short Latin words must match a whole word
    ("id" must not match "valid"); longer words and Arabic match as text."""
    h = str(header or "").strip().lower()
    tokens = re.findall(r"[a-z0-9]+|[\u0600-\u06FF]+", h)
    for w in words:
        if " " in w or len(w) >= 5 or re.search(r"[\u0600-\u06FF]", w):
            if w in h:
                return True
        elif w in tokens:
            return True
    return False


def parse_text_date(s, month_first=False):
    s = str(s).strip()
    m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,5})", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if month_first:
            d, mo = mo, d
        if y < 100:
            y += 2000
        try:
            return dt.date(y, mo, d)
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if m:
        try:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return None


def norm_name(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def find_header_row(ws, scan=15):
    best, best_n = 1, 0
    for r in range(1, min(ws.max_row, scan) + 1):
        vals = [c.value for c in ws[r]]
        n = sum(1 for v in vals if isinstance(v, str) and v.strip())
        if n > best_n:
            best, best_n = r, n
    return best


class Config:
    def __init__(self, as_of, warn_days=10, min_year=2000, max_future_days=365,
                 expiry_max_years=10, month_first=False, header_row=None, roles=None):
        self.header_row = header_row      # fixed header row number, or None = detect
        self.roles = roles or {}          # {column number: role}, see ROLE_NAMES
        self.as_of = as_of
        self.warn_days = warn_days
        self.min_year = min_year
        self.max_future_days = max_future_days
        self.expiry_max_years = expiry_max_years
        self.month_first = month_first


class Checker:
    def __init__(self, ws, cfg):
        self.ws = ws
        self.cfg = cfg
        self.as_of = cfg.as_of
        self.warn_days = cfg.warn_days
        self.issues = []
        self.hdr_row = cfg.header_row or find_header_row(ws)
        self.headers = {}
        for c in ws[self.hdr_row]:
            if c.value not in (None, ""):
                self.headers[c.column] = str(c.value).strip()
        self.rows = []
        for r in range(self.hdr_row + 1, ws.max_row + 1):
            vals = {col: ws.cell(r, col).value for col in self.headers}
            if any(v not in (None, "") for v in vals.values()):
                self.rows.append((r, vals))

    def ptd(self, v):
        return parse_text_date(v, self.cfg.month_first)

    def add(self, sev, check, row, col, value, msg, tip):
        self.issues.append({
            "Severity": sev, "Check": check, "Sheet": self.ws.title,
            "Row": row, "Column": self.headers.get(col, "") if col else "",
            "Value": "" if value is None else str(value)[:80],
            "Problem": msg, "Suggested fix": tip})

    def cols(self, words, exclude=()):
        out = []
        for col, h in self.headers.items():
            if has_word(h, words) and not has_word(h, exclude):
                out.append(col)
        return out

    def guess_roles(self):
        """What the header words suggest, as {column number: role}. For screens and tests."""
        out = {}
        date_cols = self.cols(DATE_WORDS)
        for c in date_cols:
            out[c] = "expiry" if has_word(self.headers[c], EXPIRY_WORDS) else "date"
        name_cols = self.cols(NAME_WORDS)
        for c in name_cols:
            out.setdefault(c, "name")
        for c in self.cols(JOB_WORDS):
            out.setdefault(c, "job")
        for c in self.cols(NUM_WORDS):
            out.setdefault(c, "qty")
        for c in self.cols(ID_WORDS):
            if c not in name_cols:
                out.setdefault(c, "id")
        return out

    def run(self):
        date_cols = self.cols(DATE_WORDS)
        expiry_cols = [c for c in date_cols if has_word(self.headers[c], EXPIRY_WORDS)]
        name_cols = self.cols(NAME_WORDS)
        job_cols = self.cols(JOB_WORDS)
        num_cols = self.cols(NUM_WORDS)
        id_cols = [c for c in self.cols(ID_WORDS) if c not in name_cols]
        # manual roles replace the guess for the columns the user set
        for col, role in self.cfg.roles.items():
            if col not in self.headers or role == "auto":
                continue
            for lst in (date_cols, expiry_cols, name_cols, job_cols, num_cols, id_cols):
                if col in lst:
                    lst.remove(col)
            if role in ("date", "expiry"):
                date_cols.append(col)
                if role == "expiry":
                    expiry_cols.append(col)
            elif role == "name":
                name_cols.append(col)
            elif role == "job":
                job_cols.append(col)
            elif role == "qty":
                num_cols.append(col)
            elif role == "id":
                id_cols.append(col)
        for lst in (date_cols, name_cols, job_cols, num_cols, id_cols):
            lst.sort()
        expiry_cols.sort()
        entry_cols = [c for c in date_cols if c not in expiry_cols]
        self.check_dates(date_cols, expiry_cols)
        self.check_order(entry_cols)
        self.check_duplicates(name_cols, entry_cols)
        self.check_missing(name_cols, entry_cols, id_cols)
        self.check_expiry(expiry_cols)
        self.check_names(name_cols, job_cols)
        self.check_numbers(num_cols)
        return self.issues

    # 1 and 2
    def check_dates(self, date_cols, expiry_cols):
        lo = dt.date(self.cfg.min_year, 1, 1)
        for col in date_cols:
            far = self.as_of + dt.timedelta(
                days=365 * self.cfg.expiry_max_years if col in expiry_cols
                else self.cfg.max_future_days)
            for r, vals in self.rows:
                v = vals[col]
                if v in (None, ""):
                    continue
                d = to_date(v)
                if d is None and isinstance(v, str):
                    d = self.ptd(v)
                    if d is None:
                        self.add("High", "Unreadable date", r, col, v,
                                 "Date cannot be read.", "Retype as a real Excel date.")
                        continue
                    self.add("Medium", "Date stored as text", r, col, v,
                             "Date is text, so sorting and formulas can fail.",
                             "Convert to a real date.")
                elif d is None:
                    continue
                if d < lo or d > far:
                    self.add("High", "Impossible date", r, col, d.isoformat(),
                             "Date is outside the allowed range (before %d or after %s)." % (self.cfg.min_year, far.isoformat()),
                             "Check for a typo in the year.")

    # 3
    def check_order(self, entry_cols):
        for col in entry_cols[:1]:
            prev = None
            for r, vals in self.rows:
                d = to_date(vals[col]) or (self.ptd(vals[col]) if isinstance(vals[col], str) else None)
                if d is None:
                    continue
                if prev and d < prev:
                    self.add("Info", "Date out of order", r, col, d.isoformat(),
                             "Date is earlier than the row above (%s)." % prev.isoformat(),
                             "Fine if rows are not in date order. Otherwise check the date.")
                prev = d

    # 4
    def check_duplicates(self, name_cols, entry_cols):
        skip = {c for c, h in self.headers.items()
                if str(h).strip().lower() in ROW_NO_HEADERS}
        seen = {}
        for r, vals in self.rows:
            key = tuple(norm_name(vals[c]) for c in sorted(vals) if c not in skip)
            if sum(1 for k in key if k) < 2:
                continue
            if key in seen:
                self.add("High", "Duplicate row", r, None, "",
                         "Same as row %d (all columns match)." % seen[key],
                         "Delete one copy if it is a double entry.")
            else:
                seen[key] = r
        if name_cols and entry_cols:
            nc, dc = name_cols[0], entry_cols[0]
            groups = defaultdict(list)
            for r, vals in self.rows:
                n = norm_name(vals[nc])
                d = to_date(vals[dc]) or (str(vals[dc]) if vals[dc] else None)
                if n and d:
                    groups[(n, str(d))].append(r)
            dup_rows = {r for r, v in self.issues_dup_rows()}
            for (n, d), rs in groups.items():
                if len(rs) > 1:
                    for r in rs[1:]:
                        if r in dup_rows:
                            continue
                        self.add("Low", "Same person, same day", r, nc, n,
                                 "Same name and date also in row %d." % rs[0],
                                 "Check if this is a second issue on purpose.")

    def issues_dup_rows(self):
        return [(i["Row"], i) for i in self.issues if i["Check"] == "Duplicate row"]

    # 5
    def check_missing(self, name_cols, entry_cols, id_cols):
        n = len(self.rows)
        if n == 0:
            return
        for col, h in self.headers.items():
            filled = sum(1 for _, v in self.rows if v[col] not in (None, ""))
            if filled == 0:
                self.add("Medium", "Empty column", None, col, "",
                         "Column has no values in any row.",
                         "Fill it, or remove it if not needed.")
            elif col in id_cols and filled / n < 0.5:
                self.add("Medium", "Mostly empty ID", None, col, "",
                         "Only %d of %d rows have a value." % (filled, n),
                         "Fill the ID so each record can be traced.")
        for col in (name_cols[:1] + entry_cols[:1]):
            for r, vals in self.rows:
                if vals[col] in (None, ""):
                    self.add("High", "Missing key value", r, col, "",
                             "Key cell is blank while the row has other data.",
                             "Fill it in.")

    # 6
    def check_expiry(self, expiry_cols):
        for col in expiry_cols:
            for r, vals in self.rows:
                v = vals[col]
                d = to_date(v) or (self.ptd(v) if isinstance(v, str) else None)
                if d is None or d.year < self.cfg.min_year or d.year > self.as_of.year + self.cfg.expiry_max_years:
                    continue
                left = (d - self.as_of).days
                if left < 0:
                    self.add("High", "Expired", r, col, d.isoformat(),
                             "Expired %d days ago (as of %s)." % (-left, self.as_of.isoformat()),
                             "Renew or stop using the item.")
                elif left <= self.warn_days:
                    self.add("Medium", "Expires soon", r, col, d.isoformat(),
                             "Expires in %d days." % left, "Plan the renewal.")

    # 7
    def check_names(self, name_cols, job_cols):
        for col in name_cols:
            for r, vals in self.rows:
                v = vals[col]
                if isinstance(v, str) and v != re.sub(r"\s+", " ", v).strip():
                    self.add("Low", "Extra spaces in name", r, col, repr(v),
                             "Leading, trailing or double spaces.",
                             "Trim the spaces so matching works.")
        if name_cols and job_cols:
            nc, jc = name_cols[0], job_cols[0]
            jobs = defaultdict(lambda: defaultdict(list))
            for r, vals in self.rows:
                n, j = norm_name(vals[nc]), norm_name(vals[jc])
                if n and j:
                    jobs[n][j].append(r)
            for n, d in jobs.items():
                if len(d) > 1:
                    first = min(min(v) for v in d.values())
                    self.add("Medium", "Same name, different job", first, jc,
                             " / ".join(sorted(d)),
                             "'%s' appears with %d different job titles (rows %s)." % (
                                 n, len(d), ", ".join(str(x) for v in d.values() for x in v[:2])),
                             "Use one standard job title per person.")

    # 8
    def check_numbers(self, num_cols):
        for col in num_cols:
            for r, vals in self.rows:
                v = vals[col]
                if isinstance(v, str) and v.strip():
                    self.add("Medium", "Text in number column", r, col, v,
                             "Number column holds text, so totals will skip it.",
                             "Keep only the number here and move notes to a Notes column.")


def put(ws, row, col, value):
    """Write a cell. Text that starts with = + - @ is stored as plain text,
    so Excel never runs data copied from the input as a formula."""
    c = ws.cell(row, col)
    c.value = value
    if isinstance(value, str):
        c.data_type = "s"
    return c


def count_uncalculated(ws_formula, ws_values):
    """Formula cells with no saved result (file never calculated in Excel)."""
    n = 0
    for row in ws_formula.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                if ws_values.cell(c.row, c.column).value is None:
                    n += 1
    return n


WATERMARK = "Hussein Abdallah"


def write_report(path, src, issues, stats, cfg):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    head = PatternFill("solid", fgColor="1F3864")
    white = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    sev_fill = {"High": "F8CBAD", "Medium": "FFE699", "Low": "DDEBF7", "Info": "EDEDED"}

    put(ws, 1, 1, "Register Health-Check report").font = Font(bold=True, size=16)
    put(ws, 2, 1, "File checked: %s (not changed)" % src)
    put(ws, 3, 1, "Checked as of: %s. Allowed dates: from %d, up to %d days ahead. "
        "Warning window: %d days." % (cfg.as_of.isoformat(), cfg.min_year,
                                      cfg.max_future_days, cfg.warn_days))
    c = put(ws, 4, 1, "Findings are leads to review, not final decisions.")
    c.font = Font(italic=True, color="7F7F7F")
    r = 6
    for i, h in enumerate(("Sheet", "Rows checked", "Issues", "High", "Medium", "Low", "Info"), 1):
        c = put(ws, r, i, h)
        c.fill, c.font, c.border = head, white, box
    for s in stats:
        r += 1
        vals = (s["sheet"], s["rows"], s["total"], s["High"], s["Medium"], s["Low"], s["Info"])
        for i, v in enumerate(vals, 1):
            put(ws, r, i, v).border = box
    r += 2
    put(ws, r, 1, "Issues by check").font = Font(bold=True)
    r += 1
    for i, h in enumerate(("Check", "Count", "Worst severity"), 1):
        c = put(ws, r, i, h)
        c.fill, c.font, c.border = head, white, box
    by = defaultdict(list)
    for it in issues:
        by[it["Check"]].append(it["Severity"])
    for chk, sevs in sorted(by.items(), key=lambda kv: (min(SEV_ORDER[s] for s in kv[1]), -len(kv[1]))):
        r += 1
        worst = min(sevs, key=lambda s: SEV_ORDER[s])
        for i, v in enumerate((chk, len(sevs), worst), 1):
            put(ws, r, i, v).border = box
        ws.cell(r, 3).fill = PatternFill("solid", fgColor=sev_fill[worst])
    notes = [s for s in stats if s.get("uncalculated")]
    if notes:
        r += 2
        put(ws, r, 1, "Warning: formulas with no saved result").font = Font(bold=True, color="C00000")
        for s in notes:
            r += 1
            put(ws, r, 1, "%s: %d formula cells show as empty. Open the file in Excel, "
                "save it, then run again." % (s["sheet"], s["uncalculated"]))
    for col, w in zip("ABCDEFG", (34, 14, 16, 8, 10, 8, 8)):
        ws.column_dimensions[col].width = w

    wi = wb.create_sheet("Issues")
    cols = ["Severity", "Check", "Sheet", "Row", "Column", "Value", "Problem", "Suggested fix"]
    for i, h in enumerate(cols, 1):
        c = put(wi, 1, i, h)
        c.fill, c.font, c.border = head, white, box
    issues = sorted(issues, key=lambda x: (SEV_ORDER[x["Severity"]], x["Check"], x["Row"] or 0))
    for r, it in enumerate(issues, 2):
        for i, h in enumerate(cols, 1):
            c = put(wi, r, i, it[h])
            c.border = box
            c.alignment = Alignment(wrap_text=True, vertical="top")
        wi.cell(r, 1).fill = PatternFill("solid", fgColor=sev_fill[it["Severity"]])
    for col, w in zip("ABCDEFGH", (10, 24, 20, 6, 22, 24, 62, 66)):
        wi.column_dimensions[col].width = w
    wi.freeze_panes = "A2"
    wi.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(cols)), max(len(issues) + 1, 2))
    mark = Font(italic=True, size=11, color="A6A6A6")
    ws.cell(5, 1, WATERMARK).font = mark
    wi.cell(1, 10, WATERMARK).font = mark
    wi.column_dimensions["J"].width = 26
    for sh in wb.worksheets:
        sh.oddFooter.center.text = WATERMARK
        sh.oddFooter.center.size = 9
        sh.page_setup.orientation = "landscape"
        sh.page_setup.fitToWidth = 1
        sh.page_setup.fitToHeight = 0
        sh.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(path)


def fail(msg):
    print("Error: " + msg, file=sys.stderr)
    return 2


def main(argv=None):
    ap = argparse.ArgumentParser(description="Register Health-Check (prototype)")
    ap.add_argument("file", help="Excel register (.xlsx or .xlsm)")
    ap.add_argument("--sheet", help="check only this sheet")
    ap.add_argument("--as-of", default=None, help="YYYY-MM-DD, default today")
    ap.add_argument("--warn-days", type=int, default=10,
                    help="warn when an expiry is within this many days (default 10)")
    ap.add_argument("--min-year", type=int, default=2000,
                    help="dates before this year are impossible (default 2000)")
    ap.add_argument("--max-future-days", type=int, default=365,
                    help="issue/inspection dates further ahead than this are impossible (default 365)")
    ap.add_argument("--expiry-max-years", type=int, default=10,
                    help="expiry dates further ahead than this are impossible (default 10)")
    ap.add_argument("--month-first", action="store_true",
                    help="read text dates like 03/04/2026 as month first (default day first)")
    ap.add_argument("--header-row", type=int, default=None,
                    help="row number of the column headers (default: detect). Needs --sheet.")
    ap.add_argument("--role", action="append", default=[], metavar="COL=ROLE",
                    help="set a column role, e.g. B=name, F=expiry (roles: %s). Needs --sheet." % ", ".join(ROLE_NAMES))
    ap.add_argument("--out", default=None, help="report path (default: next to the input)")
    ap.add_argument("--force", action="store_true", help="overwrite an existing report")
    a = ap.parse_args(argv)

    if not os.path.isfile(a.file):
        return fail("file not found: %s" % a.file)
    if not a.file.lower().endswith((".xlsx", ".xlsm")):
        return fail("only .xlsx and .xlsm files are supported")
    try:
        as_of = dt.date.fromisoformat(a.as_of) if a.as_of else dt.date.today()
    except ValueError:
        return fail("--as-of must look like 2026-11-25")
    out = a.out or re.sub(r"\.xls[xm]$", "", a.file, flags=re.I) + "_health_report.xlsx"
    if os.path.exists(out) and os.path.samefile(out, a.file):
        return fail("refusing to overwrite the input file")
    if os.path.exists(out) and not a.force:
        return fail("report already exists: %s (use --force to overwrite it)" % out)
    if not out.lower().endswith(".xlsx"):
        return fail("report path must end with .xlsx")

    if (a.header_row or a.role) and not a.sheet:
        return fail("--header-row and --role need --sheet (they apply to one sheet)")
    if a.header_row is not None and a.header_row < 1:
        return fail("--header-row must be 1 or more")
    roles = {}
    for item in a.role:
        col, _, role = item.partition("=")
        col, role = col.strip().upper(), role.strip().lower()
        if not re.fullmatch(r"[A-Z]{1,3}", col) or role not in ROLE_NAMES:
            return fail("bad --role %r (use COL=ROLE, roles: %s)" % (item, ", ".join(ROLE_NAMES)))
        roles[column_index_from_string(col)] = role
    cfg = Config(as_of, a.warn_days, a.min_year, a.max_future_days,
                 a.expiry_max_years, a.month_first, a.header_row, roles)
    try:
        wb = load_workbook(a.file, data_only=True)
        wbf = load_workbook(a.file, data_only=False)
    except Exception as e:  # unreadable or protected file
        return fail("cannot open the workbook: %s" % e)
    if a.sheet and a.sheet not in wb.sheetnames:
        return fail("sheet not found: %s (sheets: %s)" % (a.sheet, ", ".join(wb.sheetnames)))
    sheets = [wb[a.sheet]] if a.sheet else wb.worksheets
    if a.header_row and a.header_row > wb[a.sheet].max_row:
        return fail("--header-row %d is below the last row of the sheet (%d rows)" % (a.header_row, wb[a.sheet].max_row))
    if a.header_row and not any(c.value not in (None, "") for c in wb[a.sheet][a.header_row]):
        return fail("row %d has no headers (it is empty)" % a.header_row)
    allissues, stats = [], []
    for ws in sheets:
        if ws.max_row < 2:
            continue
        ck = Checker(ws, cfg)
        iss = ck.run()
        allissues += iss
        cnt = defaultdict(int)
        for i in iss:
            cnt[i["Severity"]] += 1
        stats.append({"sheet": ws.title, "rows": len(ck.rows), "total": len(iss),
                      "High": cnt["High"], "Medium": cnt["Medium"],
                      "Low": cnt["Low"], "Info": cnt["Info"],
                      "uncalculated": count_uncalculated(wbf[ws.title], ws)})
    write_report(out, a.file, allissues, stats, cfg)
    print("Checked %d sheet(s), %d issue(s). Report: %s" % (len(stats), len(allissues), out))
    for k in ("High", "Medium", "Low", "Info"):
        print("  %-6s %d" % (k, sum(1 for i in allissues if i["Severity"] == k)))
    for s in stats:
        if s["uncalculated"]:
            print("  Warning: %s has %d formula cells with no saved result." % (
                s["sheet"], s["uncalculated"]))
    return 1 if any(i["Severity"] == "High" for i in allissues) else 0


if __name__ == "__main__":
    sys.exit(main())
