"""Formula audit for Excel sheets. Works offline and never changes the input.

It reads each sheet twice: once with the formulas, once with the values Excel
saved. It cannot recalculate, so it finds errors and pattern problems; it cannot
prove that a formula gives the right number.
"""
import re

from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter, range_boundaries

ERRORS = {"#REF!", "#DIV/0!", "#N/A", "#VALUE!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!", "#CALC!"}
VOLATILE = ("INDIRECT", "OFFSET", "TODAY", "NOW", "RAND", "RANDBETWEEN")
RANGE_FUNCS = ("SUM", "AVERAGE", "COUNT", "MIN", "MAX")
MAX_RANGE_CELLS = 20000

_STR = re.compile(r'"(?:[^"]|"")*"')
_REF = re.compile(
    r"(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!)?"
    r"(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![A-Za-z0-9_(])")
_SHEETREF = re.compile(r"(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!")
_SIMPLE = re.compile(r"^=\s*(%s)\(\s*(\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)\s*\)\s*$" % "|".join(RANGE_FUNCS), re.I)


def is_formula(v):
    return isinstance(v, str) and v.startswith("=") and len(v) > 1


def refs(formula):
    """(sheet or None, 'A1' or 'A1:B2') for each reference, text in quotes removed."""
    f = _STR.sub('""', formula)
    out = []
    for m in _REF.finditer(f):
        sheet = m.group(1) or m.group(2)
        out.append((sheet.replace("''", "'") if sheet else None, m.group(3).replace("$", "")))
    return out


def looks_numeric_text(v):
    if not isinstance(v, str):
        return False
    s = v.strip().replace(",", "")
    return bool(s) and bool(re.fullmatch(r"[-+]?\d+(\.\d+)?", s))


def audit_sheet(wsf, wsv, sheetnames):
    """wsf: sheet loaded with formulas, wsv: same sheet with saved values."""
    issues = []

    def add(sev, check, cell, msg, tip, value=None):
        issues.append({
            "Severity": sev, "Check": check, "Sheet": wsf.title, "Row": cell.row,
            "Column": cell.coordinate,
            "Value": "" if value is None else str(value)[:80],
            "Problem": msg, "Suggested fix": tip})

    formulas = {}
    for row in wsf.iter_rows():
        for c in row:
            if is_formula(c.value):
                formulas[(c.row, c.column)] = c
    names = {n.lower() for n in sheetnames}

    for (r, col), c in formulas.items():
        f = c.value
        stripped = _STR.sub('""', f)
        # saved error values
        sv = wsv.cell(r, col).value
        if isinstance(sv, str) and sv in ERRORS:
            add("High", "Formula error value", c, "The cell shows %s." % sv,
                "Open the cell and follow its inputs to find the cause.", f)
        if "#REF!" in stripped:
            add("High", "Broken reference", c, "The formula contains #REF!, a reference to a deleted cell or sheet.",
                "Rebuild the reference.", f)
        if re.search(r"/\s*0(?![\d.])", stripped):
            add("High", "Division by zero", c, "The formula divides by a literal 0.",
                "Check the divisor.", f)
        # sheet names that do not exist
        for m in _SHEETREF.finditer(stripped):
            nm = (m.group(1) or m.group(2)).replace("''", "'")
            if re.fullmatch(r"\[\d+\].*", nm):
                continue
            if nm.lower() not in names:
                add("High", "Missing sheet", c, "The formula points to sheet '%s', which is not in this file." % nm,
                    "Fix the sheet name or restore the sheet.", f)
        if re.search(r"\[\d+\]", stripped) or re.search(r"'[^']*\[[^\]]+\.xls[xmb]?\][^']*'!", stripped, re.I):
            add("Medium", "External link", c, "The formula reads from another workbook.",
                "Copy the needed data into this file, or keep the other file next to it.", f)
        up = stripped.upper()
        for fn in VOLATILE:
            if re.search(r"\b%s\(" % fn, up):
                add("Info", "Volatile function", c, "%s recalculates on every change and can change results silently." % fn,
                    "Use it on purpose only.", f)
                break
        # self reference and ranges
        for sh, ref in refs(f):
            if sh is not None and sh.lower() != wsf.title.lower():
                continue
            try:
                c1, r1, c2, r2 = range_boundaries(ref)
            except ValueError:
                continue
            if c1 <= col <= c2 and r1 <= r <= r2:
                add("High", "Circular reference", c, "The formula refers to its own cell.",
                    "Remove the cell from its own range.", f)
            elif ":" not in ref and (c2 - c1 + 1) * (r2 - r1 + 1) == 1:
                tgt = (r1, c1)
                if tgt not in formulas and wsf.cell(r1, c1).value in (None, ""):
                    add("Low", "Reference to empty cell", c, "The formula uses %s, which is empty." % ref,
                        "Check that the right cell is referenced.", f)
        # ranges
        m = _SIMPLE.match(f)
        if m:
            try:
                c1, r1, c2, r2 = range_boundaries("%s:%s" % (m.group(2).replace("$", ""), m.group(3).replace("$", "")))
            except ValueError:
                c1 = None
            if c1 is not None and c1 == c2 and (r2 - r1 + 1) <= MAX_RANGE_CELLS:
                nxt = r2 + 1
                nc = wsf.cell(nxt, c1)
                if (nxt, c1) != (r, col) and nc.value not in (None, "") and \
                        (isinstance(nc.value, (int, float)) or is_formula(nc.value)) and \
                        not (c1 == col and nxt >= r):
                    add("Medium", "Range misses rows", c,
                        "%s stops at row %d but %s%d just below holds a number." % (
                            m.group(1).upper(), r2, get_column_letter(c1), nxt),
                        "Extend the range to include the row, or use a whole-column range.", f)
                bad = [rr for rr in range(r1, r2 + 1)
                       if looks_numeric_text(wsf.cell(rr, c1).value) and not is_formula(wsf.cell(rr, c1).value)]
                if bad:
                    add("Medium", "Numbers stored as text", c,
                        "%d cell(s) in the range hold numbers saved as text (first: %s%d), so they are skipped." % (
                            len(bad), get_column_letter(c1), bad[0]),
                        "Convert those cells to real numbers.", f)

    # per column: inconsistent formulas and hardcoded numbers inside a formula column
    bycol = {}
    for (r, col), c in formulas.items():
        bycol.setdefault(col, []).append(r)
    for col, rows in bycol.items():
        rows.sort()
        if len(rows) >= 3:
            norm = {}
            for r in rows:
                c = formulas[(r, col)]
                try:
                    norm[r] = Translator(c.value, origin=c.coordinate).translate_formula("%s1" % get_column_letter(col))
                except Exception:
                    norm[r] = c.value
            for i in range(1, len(rows) - 1):
                r = rows[i]
                before, after = norm[rows[i - 1]], norm[rows[i + 1]]
                if before == after and norm[r] != before:
                    c = formulas[(r, col)]
                    add("Medium", "Inconsistent formula", c,
                        "Different from the formulas just above and below, which match each other.",
                        "Compare with the neighbours. Copy the usual formula down if this was not on purpose.",
                        c.value)
        fset = set(rows)
        lo, hi = rows[0], rows[-1]
        for r in range(lo + 1, hi):
            if r in fset:
                continue
            v = wsf.cell(r, col).value
            if isinstance(v, (int, float)) and not isinstance(v, bool) and \
                    (r - 1) in fset and (r + 1) in fset:
                add("High", "Hardcoded number in formula column", wsf.cell(r, col),
                    "A typed number sits between formulas; it will not update.",
                    "Replace it with the formula used above and below.", v)
    return issues
