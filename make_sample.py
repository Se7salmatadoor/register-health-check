"""Builds a dummy register with known planted problems. No real data."""
import datetime as dt
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "PPE Register"
ws.append(["Dummy example data - not real people"])
ws.append(["No", "Employee Name", "Job Title", "ID Card", "Issue Date", "Helmet Qty", "Notes"])
D = dt.date
rows = [
    ("Ahmed Sample", "Carpenter", None, D(2026, 10, 1), 1),
    ("Mona Example", "Electrician", None, D(2026, 10, 2), 1),
    ("Omar Demo", "Carpenter", None, D(2026, 10, 3), 2),
    ("Omar Demo", "Foreman", None, D(2026, 10, 9), 1),        # same name, other job
    ("Sara Test", "Surveyor", None, D(2026, 10, 4), 1),
    ("Sara Test", "Surveyor", None, D(2026, 10, 4), 1),       # exact duplicate
    ("Khaled Dummy ", "Driver", None, D(2026, 10, 5), 1),     # trailing space
    ("Lina Sample", "Mason", None, "16-3-20226", 1),
    ("Hany Demo", "Welder", None, D(2028, 8, 1), 1),          # far future
    ("Nour Example", "Helper", None, "07-10-2026", 1),        # text date
    ("Yara Test", "Helper", None, D(2026, 9, 20), 1),         # earlier than row above
    ("Adel Sample", "Rigger", None, D(2026, 10, 12), "2 + warning tape"),  # text in number
    ("Salma Demo", "Cleaner", None, None, 1),                 # missing date
    (None, "Guard", None, D(2026, 10, 13), 1),                # missing name
]
for i, (n, j, idc, d, q) in enumerate(rows, 1):
    ws.append([i, n, j, idc, d, q, None])

w2 = wb.create_sheet("Lifting Equipment")
w2.append(["Item", "Equipment Code", "Inspection Date", "Expiry Date"])
w2.append(["Crane A", "MC-01", D(2026, 4, 10), D(2026, 10, 4)])      # expired
w2.append(["Chain sling", "CS-01", D(2026, 5, 1), D(2026, 12, 1)])
w2.append(["Welding machine", "AW-01", D(2026, 5, 1), D(2026, 11, 30)])  # soon
w2.append(["Generator", "G-01", D(2026, 5, 1), D(2026, 9, 30)])      # expired
w2.append(["Light tower", "TL-01", D(2029, 1, 1), D(2030, 1, 1)])    # far future inspection
w2.append(["Compressor", "CO-1", D(2026, 6, 1), "2027-06-01"])       # text date
wb.save("sample/dummy_register.xlsx")
print("sample written")
