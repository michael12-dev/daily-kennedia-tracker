# =========================================================
# EXACT KENNEDIA DAILY TRACKER STYLE
# =========================================================

import sys
import json
import base64
import io

from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side
)
from openpyxl.drawing.image import Image as XLImage

# =========================================================
# LOAD JSON DATA
# =========================================================

data = json.loads(open(sys.argv[1]).read())

staff_name = data['staffName']
staff_unit = data['staffUnit']
tasks      = data['tasks']
out_path   = data['excelFile']

# =========================================================
# COLORS
# =========================================================

GREEN_DARK   = "548235"
GREEN_LIGHT  = "D9EAD3"
YELLOW       = "FFFF00"
WHITE        = "FFFFFF"
BLACK        = "000000"
HEADER_GREY  = "E7E6E6"

# =========================================================
# HELPERS
# =========================================================

def fill(color):
    return PatternFill(
        fill_type='solid',
        fgColor=color
    )

def font(
    size=11,
    bold=False,
    color=BLACK,
    name='Calibri'
):
    return Font(
        name=name,
        size=size,
        bold=bold,
        color=color
    )

thin = Side(style='thin', color='000000')

def border():
    return Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

def align(
    h='left',
    v='center',
    wrap=True
):
    return Alignment(
        horizontal=h,
        vertical=v,
        wrap_text=wrap
    )

# =========================================================
# WORKBOOK
# =========================================================

wb = Workbook()
ws = wb.active
ws.title = "Daily Tracker"

# =========================================================
# COLUMN WIDTHS
# =========================================================

widths = {
    'A': 11,
    'B': 33,
    'C': 48,
    'D': 10,
    'E': 19,
    'F': 48,
    'G': 45
}

for col, width in widths.items():
    ws.column_dimensions[col].width = width

# =========================================================
# ROW HEIGHTS
# =========================================================

ws.row_dimensions[1].height = 65
ws.row_dimensions[2].height = 32
ws.row_dimensions[3].height = 30
ws.row_dimensions[4].height = 48
ws.row_dimensions[5].height = 180

# =========================================================
# TOP LEFT YELLOW BOX
# =========================================================

ws.merge_cells('B1:C1')

ws['B1'] = (
    "Team Stand Up Meeting:\n"
    "Morning - 9:00am\n"
    "Progress Check-in - 4:00pm"
)

ws['B1'].fill = fill(YELLOW)

ws['B1'].font = font(
    size=12,
    bold=True
)

ws['B1'].alignment = align(
    h='center',
    v='center'
)

ws['B1'].border = border()

# =========================================================
# LOGO AREA
# =========================================================

ws.merge_cells('D1:E1')

# =========================================================
# TOP RIGHT INSTRUCTION BOX
# =========================================================

ws.merge_cells('F1:G3')

instruction_text = (
    "How to Complete the Tracker:\n\n"
    "Morning 8:00am: Fill your Proposed Tasks + "
    "Expected Completion Time\n\n"
    "Evening (At The End Of The Day) 4:00pm:\n"
    "Fill in your completed task and share with "
    "your line manager/HOD and upload necessary "
    "proof to a shared drive"
)

ws['F1'] = instruction_text

ws['F1'].fill = fill(GREEN_LIGHT)

ws['F1'].font = font(
    size=10,
    bold=False
)

ws['F1'].alignment = align(
    h='left',
    v='top'
)

ws['F1'].border = border()

# Make title bold manually inside instruction text
ws['F1'].font = Font(
    name='Calibri',
    size=10,
    bold=False
)

# =========================================================
# NAME ROW
# =========================================================

ws['A2'] = "Name"

ws['A2'].fill = fill(GREEN_DARK)

ws['A2'].font = font(
    size=14,
    bold=True,
    color=WHITE
)

ws['A2'].alignment = align(
    h='center'
)

ws['A2'].border = border()

ws['B2'] = staff_name.upper()

ws['B2'].fill = fill(GREEN_DARK)

ws['B2'].font = font(
    size=13,
    bold=True,
    color=WHITE
)

ws['B2'].alignment = align(
    h='right'
)

ws['B2'].border = border()

# =========================================================
# TITLE AREA
# =========================================================

ws.merge_cells('C2:E3')

ws['C2'] = "Daily Deliverables Tracker"

ws['C2'].fill = fill(GREEN_DARK)

ws['C2'].font = font(
    size=20,
    bold=True,
    color=WHITE
)

ws['C2'].alignment = align(
    h='center',
    v='center'
)

ws['C2'].border = border()

# =========================================================
# UNIT ROW
# =========================================================

ws['A3'] = "Unit"

ws['A3'].fill = fill(GREEN_DARK)

ws['A3'].font = font(
    size=14,
    bold=True,
    color=WHITE
)

ws['A3'].alignment = align(
    h='center'
)

ws['A3'].border = border()

ws['B3'] = staff_unit

ws['B3'].fill = fill(GREEN_DARK)

ws['B3'].font = font(
    size=13,
    bold=True,
    color=WHITE
)

ws['B3'].alignment = align(
    h='right'
)

ws['B3'].border = border()

# =========================================================
# TABLE HEADERS
# =========================================================

headers = [
    'S/N',
    'Clients/Focus Area',
    'Proposed Task/Deliverable for the Day',
    'Priority\n(H/M/L)',
    'Expected Time\nof Completion',
    'Actual Deliverables for Today',
    'Achievement/Result for the Day'
]

for i, header in enumerate(headers):

    cell = ws.cell(row=4, column=i + 1)

    cell.value = header

    cell.fill = fill(HEADER_GREY)

    cell.font = font(
        size=11,
        bold=True
    )

    cell.alignment = align(
        h='center' if i in [0,3,4] else 'left',
        v='center'
    )

    cell.border = border()

# =========================================================
# TASK ROWS
# =========================================================

start_row = 5

for idx, task in enumerate(tasks):

    row = start_row + idx

    ws.row_dimensions[row].height = 210

    values = [
        idx + 1,
        task.get('client', ''),
        task.get('proposed', ''),
        task.get('priority', ''),
        task.get('time', ''),
        task.get('actual', ''),
        task.get('result', '')
    ]

    for col_index, value in enumerate(values):

        cell = ws.cell(
            row=row,
            column=col_index + 1
        )

        cell.value = value

        cell.font = font(
            size=10
        )

        cell.alignment = align(
            h='center' if col_index in [0,3,4] else 'left',
            v='top'
        )

        cell.border = border()

# =========================================================
# EMBED LOGO
# =========================================================

LOGO_B64 = "YOUR_BASE64_LOGO_HERE"

try:

    logo_bytes = base64.b64decode(LOGO_B64)

    img_io = io.BytesIO(logo_bytes)

    img = XLImage(img_io)

    img.width = 220
    img.height = 52

    img.anchor = 'D1'

    ws.add_image(img)

except Exception as e:
    print("Logo Error:", e)

# =========================================================
# SAVE FILE
# =========================================================

wb.save(out_path)

print(f"Saved: {out_path}")
