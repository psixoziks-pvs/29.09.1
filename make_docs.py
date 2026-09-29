# -*- coding: utf-8 -*-
"""
Задание: «Создайте документ, аналогичный образцу» — ПЛАН КЛАССА.
Doc1.docx и Doc2.docx перезаписываются готовыми документами Word
(таблица со схемой расстановки парт). Стиль границ задается вручную,
чтобы не зависеть от наличия стиля "Table Grid" в шаблоне.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_cell_borders(cell, sz=8):
    """Тонкие чёрные границы со всех сторон ячейки."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), '000000')
        borders.append(el)
    tcPr.append(borders)


def fill_cell(cell, color):
    """Заливка ячейки серым цветом (для стола учителя/доски)."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)


def write_cell(cell, text, bold=False, size=12):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = 'Times New Roman'


def make_plan_doc(path, title_line, rows_layout, teacher_text='СТОЛ УЧИТЕЛЯ',
                  blackboard_text='ДОСКА', note_text=None):
    """
    rows_layout: список стр; каждая строка — список подписей парт
    (сверху вниз: первый ряд ближе к доске).
    """
    doc = Document()

    # Ориентация книжная, поля небольшие
    sec = doc.sections[0]
    sec.left_margin = Cm(2)
    sec.right_margin = Cm(2)
    sec.top_margin = Cm(2)
    sec.bottom_margin = Cm(2)

    # Задание
    p = doc.add_paragraph()
    r = p.add_run(title_line)
    r.bold = True
    r.font.size = Pt(14)
    r.font.name = 'Times New Roman'

    # Заголовок
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rh = h.add_run('ПЛАН КЛАССА')
    rh.bold = True
    rh.font.size = Pt(16)
    rh.font.name = 'Times New Roman'

    cols = max(len(row) for row in rows_layout)
    # Строки: доска + ряды парт + стол учителя
    n_rows = len(rows_layout) + 2
    table = doc.add_table(rows=n_rows, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    width = Cm(min(15.0 / cols, 3.2))
    height = Cm(1.1)

    # 1) Доска — объединённая первая строка
    board = table.rows[0].cells[0]
    for c in table.rows[0].cells[1:]:
        board = board.merge(c)
    write_cell(board, blackboard_text, bold=True, size=12)
    fill_cell(board, 'D9D9D9')

    # 2) Ряды парт
    for i, row_labels in enumerate(rows_layout):
        row = table.rows[i + 1]
        for j in range(cols):
            cell = row.cells[j]
            if j < len(row_labels):
                write_cell(cell, row_labels[j], bold=True)
            else:
                write_cell(cell, '')
            set_cell_borders(cell)
            cell.width = width
        row.height = height

    # 3) Стол учителя — объединённая последняя строка
    desk = table.rows[-1].cells[0]
    for c in table.rows[-1].cells[1:]:
        desk = desk.merge(c)
    write_cell(desk, teacher_text, bold=True, size=12)
    fill_cell(desk, 'D9D9D9')

    # Границы для доски и стола тоже
    for cell in (table.rows[0].cells[0], table.rows[-1].cells[0]):
        set_cell_borders(cell)

    # Подпись под планом
    if note_text is None:
        note_text = ('Парты расположены в {} ряда(ов).'.format(len(rows_layout)))
    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rn = note.add_run(note_text)
    rn.font.size = Pt(11)
    rn.font.name = 'Times New Roman'
    rn.font.color.rgb = RGBColor(0x40, 0x40, 0x40)

    doc.save(path)
    print('saved:', path)


# ---------- Doc1: 12 парт, 3 ряда по 4 ----------
rows1 = [
    ['№1', '№2', '№3', '№4'],
    ['№5', '№6', '№7', '№8'],
    ['№9', '№10', '№11', '№12'],
]
make_plan_doc('/workspace/Doc1.docx',
              '1. Создайте документ, аналогичный образцу:',
              rows1,
              note_text='Парты расположены в три ряда по четыре парты.')

# ---------- Doc2: 12 парт, 4 ряда по 3 ----------
rows2 = [
    ['№1', '№2', '№3'],
    ['№4', '№5', '№6'],
    ['№7', '№8', '№9'],
    ['№10', '№11', '№12'],
]
make_plan_doc('/workspace/Doc2.docx',
              '2. Создайте документ, аналогичный образцу:',
              rows2,
              note_text='Парты расположены в четыре ряда по три парты.')
