from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

import openpyxl
import subprocess

# XLSX - contains 6 columns (3 cols=lines RECTO / 3 cols=lines verso) ~ 23 char
SRC = '/home/odoo/src/private/flashcard-pdf-jke/flashcard_sample.xlsx'
COLS = 4
ROWS = 10
DEBUG = True

DST = '.pdf'.join(SRC.rsplit('.xlsx', 1))
COLOR_GRID_RECTO = colors.Color(0.85, 0.85, 0.85)  # 1 = white, 0 = black
COLOR_GRID_VERSO = colors.Color(1, 1, 1)  # 1 = white, 0 = black


def log(*s):
    if DEBUG:
        print(*s)


def create_table(recto, verso):
    default_style = [
        ('FONTNAME', (0, 0), (-1, -1), 'Courier'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]
    tableRecto = Table(recto, colWidths=[(21 / COLS) * cm] * COLS, rowHeights=[(28.5 / ROWS) * cm] * ROWS)
    tableRecto.setStyle(TableStyle(default_style + [
        ('GRID', (0, 0), (-1, -1), 1, COLOR_GRID_RECTO),
    ]))

    tableVerso = Table(verso, colWidths=[(21 / COLS) * cm] * COLS, rowHeights=[(28.5 / ROWS) * cm] * ROWS)
    tableVerso.setStyle(TableStyle(default_style + [
        ('INNERGRID', (0, 0), (-1, -1), 1, COLOR_GRID_VERSO),
        ('LINEABOVE', (0, 0), (-1, 0), 1, COLOR_GRID_VERSO),
        ('LINEBELOW', (0, -1), (COLS - 1, -1), 1, COLOR_GRID_VERSO),
    ]))
    return tableRecto, tableVerso


doc = SimpleDocTemplate(DST, pagesize=A4, rightMargin=0 * cm, leftMargin=0 * cm, topMargin=0.5 * cm, bottomMargin=0 * cm)
wb = openpyxl.load_workbook(SRC)
sheet = wb.active

elements = []
dataRecto = []
dataVerso = []
rowRecto = []
rowVerso = []
pending_cards = False

for row_idx, words in enumerate(sheet.iter_rows(min_row=1, max_col=6, values_only=True)):
    pending_cards = True
    mo = row_idx % COLS

    if row_idx and ((row_idx) % COLS == 0):
        log('add line', rowRecto, rowVerso)
        rowVerso = rowVerso[::-1]  # inverse order
        dataRecto.append(rowRecto)
        dataVerso.append(rowVerso)
        rowRecto = []
        rowVerso = []

    rowRecto.append(f"{words[0] or ''}\n{words[1] or ''}\n{words[2] or ''}".rstrip("\n"))
    rowVerso.append(f"{words[3] or ''}\n{words[4] or ''}\n{words[5] or ''}".rstrip("\n"))
    # log(row_idx, words, COLS, mo)

    if row_idx and row_idx % (COLS * ROWS) == 0:
        log('add sheet %s cards' % len(dataVerso))
        pending_cards = False
        tableRecto, tableVerso = create_table(dataRecto, dataVerso)
        elements.append(tableRecto)
        elements.append(tableVerso)
        dataRecto = []
        dataVerso = []

# last - not completed page
if pending_cards:
    for i in range(COLS - len(rowRecto)):
        log('add empty card to complete line')
        rowRecto.append('')
        rowVerso.append('')
    log('add line', rowRecto, rowVerso[::1])
    dataRecto.append(rowRecto)
    dataVerso.append(rowVerso[::-1])

    for j in range(ROWS - len(dataRecto)):
        log('add empty ROWS to complete sheet')
        dataRecto.append([''] * COLS)
        dataVerso.append([''] * COLS)

    log('add final sheet')
    tableRecto, tableVerso = create_table(dataRecto, dataVerso)
    elements.append(tableRecto)
    elements.append(tableVerso)

# log(elements)
doc.build(elements)
print(f"Flashcard created to: {DST}")
subprocess.call(('xdg-open', DST))
