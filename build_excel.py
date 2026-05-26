import sys, json, base64, io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU, cm_to_EMU

# =========================
# LOAD INPUT DATA
# =========================
# (Keeping your sys.argv logic, but using defaults for demonstration)
try:
    data = json.loads(open(sys.argv[1]).read())
except:
    data = {
        'staffName': 'JOHN DOE', 
        'staffUnit': 'Operations', 
        'tasks': [], 
        'excelFile': 'Daily_Tracker.xlsx'
    }

staff_name = data['staffName']
staff_unit = data['staffUnit']
tasks      = data['tasks']
out_path   = data['excelFile']

# =========================
# EXACT IMAGE COLORS (Hex)
# =========================
GREEN_DARK   = "385623" # The "Daily Deliverables Tracker" green
GREEN_LIGHT  = "E2EFDA" # Instruction box green
GREEN_LABEL  = "548235" # Name/Unit label green
YELLOW_BOX   = "FFFF00"
WHITE        = "FFFFFF"
BLACK        = "000000"
HEADER_GREY  = "D9D9D9"

# =========================
# HELPERS
# =========================
def fill(color):
    return PatternFill("solid", fgColor=color)

def font(size=11, bold=False, color=BLACK, name="Calibri"):
    return Font(name=name, size=size, bold=bold, color=color)

thin = Side(style='thin', color=BLACK)
def get_border():
    return Border(left=thin, right=thin, top=thin, bottom=thin)

# =========================
# WORKBOOK SETUP
# =========================
wb = Workbook()
ws = wb.active
ws.title = "Daily Tracker"
ws.sheet_view.showGridLines = False # Cleaner look like the image

# Column Widths to match the proportions
widths = {"A": 6, "B": 25, "C": 45, "D": 12, "E": 18, "F": 40, "G": 40}
for col, width in widths.items():
    ws.column_dimensions[col].width = width

# =========================
# HEADER ROW 1: Instructions | Logo | Tracker Info
# =========================
ws.row_dimensions[1].height = 90

# B1: Yellow Meeting Box
ws["B1"] = "Team Stand Up Meeting:\nMorning - 9:00am\nProgress Check-in - 4:00pm"
ws["B1"].fill, ws["B1"].font, ws["B1"].border = fill(YELLOW_BOX), font(size=11, bold=True), get_border()
ws["B1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# C1: Placeholder for Logo (Image anchor is handled later)
# Instructions Box (Merged F1:G1)
ws.merge_cells("F1:G1")
ws["F1"] = ("How to Complete the Tracker:\n"
            "Morning 8:00am: Fill your Proposed Tasks + Expected Completion Time\n\n"
            "Evening (At The End Of The Day) 4:00pm: Fill in your completed task and "
            "share with your line manager/HOD and upload necessary proof to a shared drive")
ws["F1"].fill = fill(GREEN_LIGHT)
ws["F1"].font = font(size=10, bold=True)
ws["F1"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
for cell in ws["F1:G1"][0]: cell.border = get_border()

# =========================
# NAME & UNIT ROWS (2 & 3)
# =========================
ws.row_dimensions[2].height = 25
ws.row_dimensions[3].height = 25

# Labels (A2, A3)
for addr, txt in [("A2", "Name"), ("A3", "Unit")]:
    ws[addr] = txt
    ws[addr].fill, ws[addr].font, ws[addr].border = fill(GREEN_LABEL), font(size=12, bold=True, color=WHITE), get_border()
    ws[addr].alignment = Alignment(horizontal="center", vertical="center")

# Values (B2, B3)
ws["B2"], ws["B3"] = staff_name.upper(), staff_unit
for addr in ["B2", "B3"]:
    ws[addr].fill, ws[addr].font, ws[addr].border = fill(GREEN_DARK), font(size=12, bold=True, color=WHITE), get_border()
    ws[addr].alignment = Alignment(horizontal="center", vertical="center")

# Main Title (C2:F3 Merged)
ws.merge_cells("C2:F3")
ws["C2"] = "Daily Deliverables Tracker"
ws["C2"].fill = fill(GREEN_DARK)
ws["C2"].font = font(size=24, bold=True, color=WHITE, name="Calibri Light")
ws["C2"].alignment = Alignment(horizontal="center", vertical="center")
# Apply borders/fill to merged range
for r in range(2, 4):
    for c in range(3, 7):
        ws.cell(row=r, column=c).border = get_border()
        ws.cell(row=r, column=c).fill = fill(GREEN_DARK)

# Date and Time Boxes (G2, G3)
current_dt = datetime.now()
ws["G2"], ws["G3"] = f"Date: {current_dt.strftime('%d/%m/%Y')}", f"Time: {current_dt.strftime('%I:%M %p')}"

for addr in ["G2", "G3"]:
    ws[addr].fill, ws[addr].font, ws[addr].border = fill(GREEN_DARK), font(size=11, bold=True, color=WHITE), get_border()
    ws[addr].alignment = Alignment(horizontal="left", vertical="center")
# =========================
# TABLE HEADERS (Row 4)
# =========================
ws.row_dimensions[4].height = 45
headers = ["S/N", "Clients/Focus Area", "Proposed Task/Deliverable for the Day", 
           "Priority\n(H/M/L)", "Expected Time\nof Completion", 
           "Actual Deliverables for Today", "Achievement/Result for the Day"]

for i, text in enumerate(headers, start=1):
    cell = ws.cell(row=4, column=i)
    cell.value = text
    cell.fill, cell.font, cell.border = fill(HEADER_GREY), font(size=10, bold=True), get_border()
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# =========================
# TASK ROWS
# =========================
start_row = 5
if not tasks: # Add empty rows for visual consistency if no data
    tasks = [{"client":""}] * 5

for idx, task in enumerate(tasks, start=1):
    row = start_row + idx - 1
    ws.row_dimensions[row].height = 60
    
    vals = [idx, task.get("client",""), task.get("proposed",""), task.get("priority",""), 
            task.get("time",""), task.get("actual",""), task.get("result","")]
    
    for col_idx, val in enumerate(vals, start=1):
        cell = ws.cell(row=row, column=col_idx)
        cell.value = val
        cell.border = get_border()
        cell.alignment = Alignment(horizontal="center" if col_idx in [1, 4, 5] else "left", 
                                   vertical="top", wrap_text=True)


wb.save(out_path)
