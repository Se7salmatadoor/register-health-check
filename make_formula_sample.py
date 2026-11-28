"""Builds sample/dummy_formulas.xlsx: a small sheet with deliberate formula problems (dummy data).
Open it once in Excel or LibreOffice and save, so the results are stored, then run:
  python healthcheck.py sample/dummy_formulas.xlsx --formulas
"""
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Costs"
ws.append(["Item", "Qty", "Price", "Total", "Share"])
for i in range(2, 9):
    ws.append(["Item %d" % (i - 1), i, 10 * i, "=B%d*C%d" % (i, i), "=D%d/D$12" % i])
ws["D5"] = 999                      # typed number over a formula
ws["D7"] = "=B7*C6"                 # wrong row, inconsistent
ws["B9"], ws["C9"], ws["D9"] = 3, 15, 45   # extra row below the range
ws["A12"], ws["D12"] = "Sum", "=SUM(D2:D8)"   # misses row 9
ws["A13"], ws["D13"] = "Average Qty", "=AVERAGE(B2:B9)/0"
ws["A14"], ws["D14"] = "Lookup", "=VLOOKUP(B2,Gone!A1:B5,2,0)"
ws["A15"], ws["D15"] = "Loop", "=SUM(D14:D15)"
ws["B10"], ws["C10"] = "7", 4        # number saved as text
ws["D10"] = "=B10*C10"
import os
os.makedirs("sample", exist_ok=True)
wb.save("sample/dummy_formulas.xlsx")
