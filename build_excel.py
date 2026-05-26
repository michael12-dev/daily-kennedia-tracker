import sys
import json
import base64
import io
import math

from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side,
    GradientFill
)
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import FormulaRule

# =========================
# LOAD DATA
# =========================

data = json.loads(open(sys.argv[1]).read())

staff_name = data['staffName']
staff_unit = data['staffUnit']
date_str   = data['date']
tasks      = data['tasks']
out_path   = data['excelFile']

# =========================
# COLORS
# =========================

C_DARK_GREEN  = '375623'
C_MED_GREEN   = '548235'
C_ACCENT      = '70AD47'
C_LIGHT_GREEN = 'E2EFDA'
C_ALT_ROW     = 'F7FBF4'
C_BORDER      = 'D9EAD3'
C_NAME_BG     = 'C6E0B4'
C_WHITE       = 'FFFFFF'
C_BLACK       = '000000'

# =========================
# HELPERS
# =========================

def solid(hex_color):
    return PatternFill(
        fill_type='solid',
        fgColor=hex_color
    )

def gradient():
    return GradientFill(
        stop=('70AD47', '548235')
    )

def font(
    bold=False,
    size=11,
    color=C_BLACK,
    name='Segoe UI'
):
    return Font(
        bold=bold,
        size=size,
        color=color,
        name=name
    )

def border_all(color=C_BORDER):
    s = Side(style='thin', color=color)

    return Border(
        left=s,
        right=s,
        top=s,
        bottom=s
    )

def border_medium(color=C_MED_GREEN):
    s = Side(style='medium', color=color)

    return Border(
        left=s,
        right=s,
        top=s,
        bottom=s
    )

def align(h='left', v='center', wrap=False):
    return Alignment(
        horizontal=h,
        vertical=v,
        wrap_text=wrap
    )

# =========================
# WORKBOOK
# =========================

wb = Workbook()
ws = wb.active
ws.title = 'Daily Report'

# =========================
# COLUMN WIDTHS
# =========================

col_widths = {
    'A': 8,
    'B': 26,
    'C': 42,
    'D': 12,
    'E': 18,
    'F': 42,
    'G': 38
}

for col, width in col_widths.items():
    ws.column_dimensions[col].width = width

# =========================
# ROW HEIGHTS
# =========================

ws.row_dimensions[1].height = 55
ws.row_dimensions[2].height = 28
ws.row_dimensions[3].height = 28
ws.row_dimensions[4].height = 24
ws.row_dimensions[5].height = 36

# =========================
# FREEZE PANES
# =========================

ws.freeze_panes = 'A6'

# =========================
# NAME ROW
# =========================

ws['A2'] = 'Name'
ws['A2'].font = font(
    bold=True,
    size=12,
    color=C_DARK_GREEN
)
ws['A2'].fill = solid(C_NAME_BG)
ws['A2'].alignment = align('right', 'center')
ws['A2'].border = border_all()

ws['B2'] = staff_name
ws['B2'].font = font(
    bold=True,
    size=11,
    color=C_DARK_GREEN
)
ws['B2'].fill = solid(C_LIGHT_GREEN)
ws['B2'].alignment = align('left', 'center')
ws['B2'].border = border_all()

# =========================
# TITLE AREA
# =========================

ws.merge_cells('C2:G3')

ws['C2'] = 'Daily Deliverables Tracker'

ws['C2'].font = font(
    bold=True,
    size=20,
    color=C_WHITE
)

ws['C2'].fill = gradient()

ws['C2'].alignment = align(
    'center',
    'center'
)

ws['C2'].border = border_medium()

# =========================
# UNIT ROW
# =========================

ws['A3'] = 'Unit'

ws['A3'].font = font(
    bold=True,
    size=12,
    color=C_DARK_GREEN
)

ws['A3'].fill = solid(C_NAME_BG)
ws['A3'].alignment = align('right', 'center')
ws['A3'].border = border_all()

ws['B3'] = staff_unit

ws['B3'].font = font(
    bold=True,
    size=11,
    color=C_DARK_GREEN
)

ws['B3'].fill = solid(C_LIGHT_GREEN)
ws['B3'].alignment = align('left', 'center')
ws['B3'].border = border_all()

# =========================
# DATE ROW
# =========================

ws['A4'] = 'Date'

ws['A4'].font = font(
    bold=True,
    size=11,
    color=C_DARK_GREEN
)

ws['A4'].fill = solid(C_NAME_BG)
ws['A4'].alignment = align('right', 'center')
ws['A4'].border = border_all()

ws['B4'] = date_str

ws['B4'].font = font(
    size=11,
    color=C_DARK_GREEN
)

ws['B4'].fill = solid(C_LIGHT_GREEN)
ws['B4'].alignment = align('left', 'center')
ws['B4'].border = border_all()

for col in ['C', 'D', 'E', 'F', 'G']:
    ws[f'{col}4'].fill = solid(C_MED_GREEN)
    ws[f'{col}4'].border = border_medium()

# =========================
# HEADERS
# =========================

headers = [
    'S/N',
    'Clients/Focus Area',
    'Proposed Task/Deliverable for the Day',
    'Priority',
    'Expected Time',
    'Actual Deliverables for Today',
    'Achievement/Result for the Day'
]

for i, header in enumerate(headers):

    cell = ws.cell(row=5, column=i + 1)

    cell.value = header

    cell.font = font(
        bold=True,
        size=12,
        color=C_WHITE
    )

    cell.fill = solid(C_MED_GREEN)

    cell.alignment = align(
        'center',
        'center',
        wrap=True
    )

    cell.border = border_medium()

# =========================
# TASK ROWS
# =========================

for idx, task in enumerate(tasks):

    row = idx + 6

    proposed = str(task.get('proposed', ''))
    actual   = str(task.get('actual', ''))

    longest = max(
        len(proposed),
        len(actual)
    )

    lines = math.ceil(longest / 35)

    row_height = max(
        35,
        lines * 18
    )

    ws.row_dimensions[row].height = row_height

    is_alt = idx % 2 == 1

    row_fill = solid(
        C_ALT_ROW if is_alt else 'FFFFFF'
    )

    vals = [
        idx + 1,
        task.get('client', ''),
        proposed,
        task.get('priority', 'M'),
        task.get('time', ''),
        actual,
        task.get('result', ''),
    ]

    for ci, val in enumerate(vals):

        cell = ws.cell(
            row=row,
            column=ci + 1
        )

        cell.value = val

        cell.font = font(
            size=11,
            color=C_DARK_GREEN
        )

        cell.fill = row_fill

        cell.alignment = align(
            'center' if ci in [0, 3] else 'left',
            'center',
            wrap=True
        )

        cell.border = border_all()

# =========================
# CONDITIONAL FORMATTING
# =========================

high_fill = solid('F4CCCC')
medium_fill = solid('FCE5CD')
low_fill = solid('D9EAD3')

ws.conditional_formatting.add(
    f'D6:D{len(tasks)+5}',
    FormulaRule(
        formula=['D6="H"'],
        fill=high_fill
    )
)

ws.conditional_formatting.add(
    f'D6:D{len(tasks)+5}',
    FormulaRule(
        formula=['D6="M"'],
        fill=medium_fill
    )
)

ws.conditional_formatting.add(
    f'D6:D{len(tasks)+5}',
    FormulaRule(
        formula=['D6="L"'],
        fill=low_fill
    )
)

# =========================
# AUTO FILTER
# =========================

ws.auto_filter.ref = f'A5:G{len(tasks)+5}'

# =========================
# FOOTER SECTION
# =========================

footer_row = len(tasks) + 8

ws.merge_cells(f'A{footer_row}:C{footer_row}')
ws.merge_cells(f'E{footer_row}:G{footer_row}')

ws[f'A{footer_row}'] = 'Prepared By: ____________________'
ws[f'E{footer_row}'] = 'Reviewed By: ____________________'

ws[f'A{footer_row}'].font = font(
    bold=True,
    size=11,
    color=C_DARK_GREEN
)

ws[f'E{footer_row}'].font = font(
    bold=True,
    size=11,
    color=C_DARK_GREEN
)

# =========================
# LOGO
# =========================

LOGO_B64 = "YOUR_BASE64_LOGO_HERE"

try:

    logo_bytes = base64.b64decode(LOGO_B64)

    img_io = io.BytesIO(logo_bytes)

    img = XLImage(img_io)

    img.width = 145
    img.height = 42

    img.anchor = 'D1'

    ws.add_image(img)

except Exception as e:
    print(f'Logo warning: {e}')

# =========================
# SAVE
# =========================

wb.save(out_path)

print(f'Saved professionally styled report: {out_path}')
