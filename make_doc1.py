# -*- coding: utf-8 -*-
"""Doc1.docx - План школьной лаборатории (рисунок из фигур Word + таблицы)."""
import math
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from wps_shapes import Canvas, emu, add_title, add_caption, set_cell, table_borders, landscape

WALL = "2B3A46"
DOOR = "C9A227"
TXT = "1E2A32"
GRID = "D3DDE4"
NAVY = "1F3B63"


def shd(fill):
    e = OxmlElement("w:shd")
    e.set(qn("w:val"), "clear")
    e.set(qn("w:fill"), fill)
    return e


doc = Document()
sec = doc.sections[0]
landscape(sec)
st = doc.styles["Normal"]
st.font.name = "Calibri"
st.font.size = Pt(10)

add_title(doc, "ПЛАН ШКОЛЬНОЙ ЛАБОРАТОРИИ", size=22, space_after=2)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Кабинет химии № 12 · учебный корпус «Б», 1 этаж · масштаб 1 : 50")
r.font.size = Pt(10)
r.italic = True
r.font.color.rgb = RGBColor(0x5A, 0x6A, 0x76)
p.paragraph_format.space_after = Pt(4)

CW, CH = 25.9, 14.4          # размер сцены в см (доступно на альбомной странице)
cv = Canvas(doc, CW, CH, top_cm=0.1, left_cm=0.0)
E = emu

# ---------------- геометрия плана ----------------
OX, OY, OW, OH = 1.9, 2.1, 18.0, 10.0   # наружные стены: X 1.9..19.9, Y 2.1..12.1
X_SPLIT = 13.9                           # внутренняя стена
X_SPLIT2 = 16.9                          # стена комнаты подготовки
Y_MID = 7.1                              # перегородка реактивная / приборная


def label(x, y, w, h, text, size=9, bold=False, color=TXT, align="ctr", anchor="ctr", rot=0):
    cv.rect(E(x), E(y), E(w), E(h), fill=None, line=None, text=text, size=size,
            bold=bold, color=color, align=align, anchor=anchor, rot=rot)


def room(x, y, w, h, fill="FFFFFF"):
    cv.rect(E(x), E(y), E(w), E(h), fill=fill, line=WALL, line_w=1.2)


def wall_v(x, y0, y1, t=0.22):
    cv.rect(E(x - t / 2), E(y0), E(t), E(y1 - y0), fill=WALL, line=WALL, line_w=0.3)


def wall_h(y, x0, x1, t=0.22):
    cv.rect(E(x0), E(y - t / 2), E(x1 - x0), E(t), fill=WALL, line=WALL, line_w=0.3)


# ---- фоновая координатная сетка (1 клетка = 1 м) ----
for i in range(int(OX - 1), int(OX + OW + 2)):
    gx = float(i)
    if OX - 0.5 <= gx <= OX + OW + 0.5:
        cv.line(E(gx), E(1.2), E(gx), E(12.9), color=GRID, w=0.25)
for gy in [float(i) for i in range(2, 13)]:
    cv.line(E(1.4), E(gy), E(20.4), E(gy), color=GRID, w=0.25)

# ---- помещения ----
room(OX, OY, X_SPLIT - OX, OH, "FFFFFF")                       # 1 рабочая зона
room(X_SPLIT, OY, X_SPLIT2 - X_SPLIT, Y_MID - OY, "FBF6E9")    # 3 реактивная
room(X_SPLIT, Y_MID, X_SPLIT2 - X_SPLIT, OY + OH - Y_MID, "F1F5EE")  # 4 приборная
room(X_SPLIT2, OY, OX + OW - X_SPLIT2, OH, "F6EFF6")           # 2 комната подготовки
wall_v(X_SPLIT, OY, OY + OH)
wall_v(X_SPLIT2, OY, OY + OH)
wall_h(Y_MID, X_SPLIT, X_SPLIT2)
cv.rect(E(OX - 0.24), E(OY - 0.24), E(OW + 0.48), E(OH + 0.48),
       fill=None, line=WALL, line_w=2.6)                        # двойная наружная обводка

# ---- названия помещений ----
label(OX, OY + 0.15, X_SPLIT - OX, 0.7, "1. РАБОЧАЯ ЗОНА УЧЕНИКОВ", size=11, bold=True)
label(X_SPLIT, OY + 0.15, X_SPLIT2 - X_SPLIT, 0.6, "3. РЕАКТИВНАЯ", size=8.5, bold=True)
label(X_SPLIT, Y_MID + 0.12, X_SPLIT2 - X_SPLIT, 0.6, "4. ПРИБОРНАЯ", size=8.5, bold=True)
label(X_SPLIT2, OY + 0.15, OX + OW - X_SPLIT2, 0.9, "2. КОМНАТА\nПОДГОТОВКИ", size=8.5, bold=True)

# ---- ученические столы и стулья ----
for row in range(2):
    for col in range(4):
        x = 2.9 + col * 2.6
        y = 5.7 + row * 3.2
        cv.rect(E(x), E(y), E(2.1), E(1.1), fill="EFE6D5", line="8A7A5C", line_w=0.9)
        label(x, y, 2.1, 1.1, "Стол %d" % (row * 4 + col + 1), size=7)
        for s in (-0.5, 0.5):
            cv.rect(E(x + 1.05 + s - 0.17), E(y - 0.40), E(0.34), E(0.30),
                    fill="CBD5DC", line="7A8A96", line_w=0.4)
            cv.rect(E(x + 1.05 + s - 0.17), E(y + 1.20), E(0.34), E(0.30),
                    fill="CBD5DC", line="7A8A96", line_w=0.4)

# ---- демонстрационный стол, стол преподавателя, раковина ----
cv.rect(E(5.4), E(2.7), E(5.4), E(1.6), fill="D9E7F0", line="4B6A82", line_w=1.1)
label(5.4, 2.7, 5.4, 1.6, "ДЕМОНСТРАЦИОННЫЙ СТОЛ\nвытяжной шкаф, подвод воды", size=7.5, bold=True)
cv.rect(E(11.4), E(2.7), E(2.2), E(1.4), fill="E7DFCF", line="8A7A5C", line_w=1.0)
label(11.4, 2.7, 2.2, 1.4, "Стол\nпреподавателя", size=7.5, bold=True)
cv.rect(E(3.3), E(2.75), E(1.2), E(0.8), fill="DFF3FA", line="2E7C93", line_w=0.9)
label(3.3, 2.75, 1.2, 0.8, "Раковина", size=6)

# ---- шкафы в реактивной ----
for i in range(3):
    x = 14.2 + (i % 2) * 1.35
    y = 3.6 + (i // 2) * 1.6
    cv.rect(E(x), E(y), E(1.15), E(1.3), fill="E8DCC0", line="8A7A5C", line_w=0.7)
    label(x, y, 1.15, 1.3, "Шкаф\n%d" % (i + 1), size=6.5)

# ---- стеллажи в приборной ----
for i in range(2):
    y = 8.4 + i * 1.5
    cv.rect(E(14.25), E(y), E(2.2), E(1.0), fill="DCE6D2", line="6F8560", line_w=0.7)
    label(14.25, y, 2.2, 1.0, "Стеллаж\n%s" % ("приборы" if i == 0 else "посуда"), size=6.5)

# ---- оборудование комнаты подготовки ----
prep = [(17.2, 3.9, 2.3, 1.1, "Рабочий стол"), (17.2, 5.4, 2.3, 0.9, "Весы, посуда"),
        (17.2, 6.7, 2.3, 1.7, "Вытяжной\nшкаф"), (17.2, 8.8, 2.3, 1.0, "Холодильник\n(реагенты)"),
        (17.2, 10.2, 2.3, 1.0, "Сушильный\nшкаф")]
for (x, y, w, h, t) in prep:
    cv.rect(E(x), E(y), E(w), E(h), fill="EFE0EF", line="8A5C8A", line_w=0.8)
    label(x, y, w, h, t, size=6.5)


# ---- двери ----
def door_v(cy, x, w=1.2, right=True):
    """Дверь в вертикальной стене: проём, полотно, дуга распахивания."""
    cv.rect(E(x - 0.16), E(cy - w / 2 - 0.06), E(0.32), E(w + 0.12), fill="FFFFFF", line=None)
    a0, a1 = (90, 180) if right else (270, 360)
    bx = x - w if right else x
    cv.rect(E(bx), E(cy - 0.06), E(w), E(0.12), fill=DOOR, line="8A6D1F", line_w=0.5)
    pts = [(x, cy)]
    for i in range(13):
        ang = math.radians(a0 + (a1 - a0) * i / 12)
        pts.append((x + w * math.cos(ang), cy + w * math.sin(ang)))
    cv.poly(pts, closed=False, fill=None, line="7A8A96", w=0.6)


def door_h(cx, y, w=1.4, up=True):
    """Дверь в горизонтальной стене."""
    cv.rect(E(cx - w / 2 - 0.06), E(y - 0.16), E(w + 0.12), E(0.32), fill="FFFFFF", line=None)
    a0, a1 = (180, 270) if up else (0, 90)
    yy = y - w if up else y
    cv.rect(E(cx - 0.06), E(yy), E(0.12), E(w), fill=DOOR, line="8A6D1F", line_w=0.5)
    pts = [(cx, y)]
    for i in range(13):
        ang = math.radians(a0 + (a1 - a0) * i / 12)
        pts.append((cx + w * math.cos(ang), y + w * math.sin(ang)))
    cv.poly(pts, closed=False, fill=None, line="7A8A96", w=0.6)


door_v(5.0, X_SPLIT, 1.2, right=True)      # рабочая -> реактивная
door_v(9.0, X_SPLIT, 1.2, right=True)      # рабочая -> приборная
door_v(4.2, X_SPLIT2, 1.2, right=False)    # реактивная -> подготовка
door_v(8.2, X_SPLIT2, 1.2, right=False)    # приборная -> подготовка
door_h(3.6, OY + OH, 1.4, up=True)         # главный вход (южная стена)
door_h(11.0, OY, 1.4, up=False)            # аварийный выход (северная стена)

# ---- окна (северная стена) ----
for (x0, x1) in [(4.4, 6.4), (7.4, 9.4), (14.6, 16.2), (17.6, 19.2)]:
    cv.rect(E(x0), E(OY - 0.30), E(x1 - x0), E(0.34), fill="7EC8E3", line="2B3A46", line_w=0.6)

# ---- средства безопасности и инженерия ----
cv.rect(E(2.15), E(11.15), E(0.55), E(0.6), geom="roundRect", adj={"adj": 32000},
        fill="D94A38", line="8E2B20", line_w=0.6)
label(1.4, 10.4, 2.1, 0.5, "Огнетушитель", size=6, color="B03426")
cv.rect(E(13.0), E(2.35), E(0.6), E(0.6), fill="FFFFFF", line="D94A38", line_w=0.9)
cv.rect(E(13.06), E(2.57), E(0.48), E(0.16), fill="D94A38", line=None)
cv.rect(E(13.22), E(2.41), E(0.16), E(0.48), fill="D94A38", line=None)
label(12.3, 3.0, 2.0, 0.5, "Аптечка", size=6, color="B03426")
cv.rect(E(20.0), E(5.6), E(0.5), E(0.8), fill="EDEDED", line="555555", line_w=0.8)
label(19.6, 6.5, 1.6, 0.5, "Щит", size=6)
cv.rect(E(12.6), E(4.6), E(0.5), E(0.5), geom="ellipse", fill="F5C445", line="8A6D1F", line_w=0.7)
label(12.6, 4.6, 0.5, 0.5, "Г", size=6.5, bold=True)
cv.rect(E(7.2), E(11.2), E(1.1), E(0.55), geom="ellipse", fill="DDEAF3", line="4B6A82", line_w=0.7)
label(6.5, 11.75, 2.5, 0.45, "Вентрешётка", size=6)
cv.rect(E(12.6), E(11.2), E(0.5), E(0.5), geom="ellipse", fill="F5C445", line="8A6D1F", line_w=0.7)
label(12.6, 11.2, 0.5, 0.5, "В", size=6.5, bold=True)

# ---- стрелки эвакуации ----
cv.line(E(6.0), E(11.4), E(3.9), E(13.0), color="D94A38", w=1.5, arrow_end=True)
cv.line(E(11.6), E(11.0), E(10.6), E(13.0), color="D94A38", w=1.5, arrow_end=True)
cv.line(E(15.6), E(11.2), E(14.2), E(13.0), color="D94A38", w=1.5, arrow_end=True)
label(2.0, 13.05, 8.4, 0.6, "← путь эвакуации к главному входу", size=7.5, color="B03426", align="l")
label(9.6, 0.3, 6.4, 0.6, "аварийный выход →", size=7.5, color="B03426", align="l")


# ---- размерные линии ----
def dim_h(x0, x1, y, text):
    cv.line(E(x0), E(y), E(x1), E(y), color="5A6A76", w=0.7)
    cv.line(E(x0), E(y - 0.2), E(x0), E(y + 0.2), color="5A6A76", w=0.7)
    cv.line(E(x1), E(y - 0.2), E(x1), E(y + 0.2), color="5A6A76", w=0.7)
    label((x0 + x1) / 2 - 1.6, y - 0.78, 3.2, 0.5, text, size=7.5, color="35505F")


def dim_v(y0, y1, x, text):
    cv.line(E(x), E(y0), E(x), E(y1), color="5A6A76", w=0.7)
    cv.line(E(x - 0.2), E(y0), E(x + 0.2), E(y0), color="5A6A76", w=0.7)
    cv.line(E(x - 0.2), E(y1), E(x + 0.2), E(y1), color="5A6A76", w=0.7)
    label(x - 0.85, (y0 + y1) / 2 - 1.4, 1.7, 2.8, text, size=7.5, color="35505F", rot=-90)


dim_h(OX, X_SPLIT, 1.5, "12,0 м")
dim_h(X_SPLIT, OX + OW, 1.5, "6,0 м")
dim_v(OY, OY + OH, 1.0, "10,0 м")
for x0 in (OX, X_SPLIT, OX + OW):
    cv.line(E(x0), E(1.65), E(x0), E(OY - 0.3), color="9AA9B4", w=0.4)
cv.line(E(1.2), E(OY), E(1.66), E(OY), color="9AA9B4", w=0.4)
cv.line(E(1.2), E(OY + OH), E(1.66), E(OY + OH), color="9AA9B4", w=0.4)

# ---- роза сторон света ----
NX, NY = 23.7, 3.3
cv.rect(E(NX - 1.5), E(NY - 1.5), E(3.0), E(3.0), geom="ellipse", fill="F7FAFC",
        line="7A8A96", line_w=0.9)
cv.poly([(NX, NY - 1.2), (NX + 0.34, NY), (NX - 0.34, NY)], closed=True, fill="D94A38", line=None)
cv.poly([(NX, NY + 1.2), (NX + 0.34, NY), (NX - 0.34, NY)], closed=True, fill="6B7C88", line=None)
label(NX - 0.7, NY - 2.3, 1.4, 0.55, "С (N)", size=8.5, bold=True)
label(NX - 1.6, NY + 1.6, 3.2, 0.5, "верх рисунка — север", size=6, color="5A6A76")

# ---- линейный масштаб ----
SX, SY = 22.1, 7.6
label(SX - 0.2, SY - 0.85, 5.0, 0.55, "Масштаб 1 : 50", size=8, bold=True, align="l")
for i in range(4):
    cv.rect(E(SX + i * 1.0), E(SY), E(1.0), E(0.28),
            fill=("2B3A46" if i % 2 == 0 else "FFFFFF"), line="2B3A46", line_w=0.6)
for i, t in enumerate(["0", "1", "2", "3", "4 м"]):
    label(SX + i * 1.0 - 0.35, SY + 0.32, 0.9, 0.45, t, size=6.5)
label(SX - 0.2, SY + 0.95, 5.4, 0.5, "1 клетка сетки = 1,0 м", size=6.5, color="5A6A76", align="l")

# ---- номера помещений ----
for (nx, ny, t) in [(2.2, 2.4, "1"), (14.1, 2.4, "3"), (14.1, 7.4, "4"), (17.1, 2.4, "2")]:
    cv.rect(E(nx), E(ny), E(0.6), E(0.6), geom="ellipse", fill=NAVY, line=None)
    label(nx, ny, 0.6, 0.6, t, size=8.5, bold=True, color="FFFFFF")

# ---- технические данные справа от плана ----
label(21.0, 10.2, 4.9, 3.6,
      "Пятно застройки: 24,0 × 10,0 м\nПолезная площадь: 180 м²\nВысота потолка: 3,6 м\n"
      "Отделка: плитка + краска\nВентиляция: приточно-вытяжная\nОсвещение: LED, 2 ряда",
      size=7.5, align="l", anchor="t", color="35505F")

add_caption(doc, "Рисунок 1 — План школьной лаборатории. Все элементы построены фигурами Word: "
                 "прямоугольник, округлый прямоугольник, эллипс, произвольный многоугольник "
                 "(дуги дверей и роза ветров), линия со стрелкой; надписи выполнены текстом внутри фигур.",
             size=8.5)

# ================= Таблица 1. Экспликация помещений =================
add_title(doc, "Таблица 1 — Экспликация помещений", size=14, space_after=4, space_before=8)
t1 = doc.add_table(rows=1, cols=6)
t1.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = ["№ на плане", "Наименование помещения", "Размеры, м", "Площадь, м²", "Норма мест", "Назначение"]
data = [
    ("1", "Рабочая зона учеников (учебный зал)", "12,0 × 10,0", "120,0", "24",
     "Фронтальные опыты и практические работы за 8 столами"),
    ("2", "Комната подготовки", "6,0 × 10,0", "60,0", "2",
     "Приготовление реактивов, весовая, сушильное и холодильное оборудование"),
    ("3", "Реактивная", "3,0 × 5,0", "15,0", "—",
     "Хранение реактивов в металлических шкафах с запирающимися дверцами"),
    ("4", "Приборная", "3,0 × 5,0", "15,0", "—",
     "Измерительные приборы, модели, лабораторная посуда"),
    ("—", "Итого по лаборатории", "24,0 × 10,0", "180,0", "26",
     "Кабинет химии, 1 этаж, два эвакуационных выхода"),
]
table_borders(t1)
for i, h in enumerate(hdr):
    set_cell(t1.rows[0].cells[i], h, bold=True, size=9, align="center", fill=NAVY,
             color=(0xFF, 0xFF, 0xFF))
for row in data:
    cells = t1.add_row().cells
    for i, val in enumerate(row):
        set_cell(cells[i], val, size=9,
                 align="center" if i in (0, 2, 3, 4) else "left",
                 bold=(row[0] == "—"))
    if row[0] == "—":
        for c in cells:
            c._tc.get_or_add_tcPr().append(shd("DCE6F1"))
w1 = [Cm(1.8), Cm(6.0), Cm(2.6), Cm(2.2), Cm(2.0), Cm(7.0)]
for r_ in t1.rows:
    for i, c in enumerate(r_.cells):
        c.width = w1[i]

# ================= Таблица 2. Оборудование =================
add_title(doc, "Таблица 2 — Расстановка оборудования по зонам плана", size=14,
          space_after=4, space_before=10)
t2 = doc.add_table(rows=1, cols=5)
t2.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr2 = ["Зона плана", "Оборудование (фигура на рисунке 1)", "Кол-во", "Координаты на плане, м",
        "Требования безопасности"]
data2 = [
    ("Рабочая зона", "Ученический стол с подводом воды (бежевый прямоугольник)", "8",
     "X 2,9–13,1; Y 5,7 и 8,9", "проход между столами не менее 0,7 м"),
    ("Рабочая зона", "Стул ученический (серый квадрат)", "24", "вокруг каждого стола",
     "устойчивые, с задней спинкой"),
    ("Рабочая зона", "Демонстрационный стол с вытяжным шкафом (голубой)", "1",
     "X 5,4–10,8; Y 2,7–4,3", "приточно-вытяжная вентиляция, не менее 3 обменов воздуха в час"),
    ("Рабочая зона", "Стол преподавателя (песочный)", "1", "X 11,4–13,6; Y 2,7–4,1",
     "обзор всех рабочих мест, доступ к щиту"),
    ("Рабочая зона", "Раковина со смесителем (голубой квадрат)", "1", "X 3,3–4,5; Y 2,75–3,55",
     "вода у демонстрационного стола, слив в обвязке"),
    ("Реактивная", "Металлический шкаф для реактивов", "3", "X 14,2–16,7; Y 3,6–6,5",
     "запирается на ключ, отдельная вытяжка, ящики раздельно по группам веществ"),
    ("Приборная", "Стеллаж с приборами и посудой", "2", "X 14,25–16,45; Y 8,4–10,9",
     "крепление тяжёлых приборов, нижние полки под посуду"),
    ("Комната подготовки", "Рабочий стол, весы, вытяжной и сушильный шкафы, холодильник",
     "1 / 1 / 1 / 1 / 1", "X 17,2–19,5; Y 3,9–11,2",
     "хранение легколетучих веществ вне учебного зала"),
    ("Помещение целиком", "Огнетушитель ОП-5, аптечка, электрощит, вентрешётка, клапан газа",
     "1 / 1 / 1 / 1 / 1", "углы и стена у входа", "подступы не загромождать, знаки F01, A03, E02"),
    ("Помещение целиком", "Эвакуационные выходы (двери с дугой распахивания)", "2",
     "главный вход + аварийный выход", "двери открываются наружу по ходу выхода"),
]
table_borders(t2)
for i, h in enumerate(hdr2):
    set_cell(t2.rows[0].cells[i], h, bold=True, size=8.5, align="center", fill=NAVY,
             color=(0xFF, 0xFF, 0xFF))
for row in data2:
    cells = t2.add_row().cells
    for i, val in enumerate(row):
        set_cell(cells[i], val, size=8.5, align="center" if i == 2 else "left")
w2 = [Cm(2.8), Cm(6.2), Cm(1.6), Cm(3.6), Cm(5.4)]
for r_ in t2.rows:
    for i, c in enumerate(r_.cells):
        c.width = w2[i]

doc.add_page_break()

# ================= Таблица 3. Условные обозначения =================
add_title(doc, "Таблица 3 — Условные обозначения и их построение из фигур Word", size=14,
          space_after=4, space_before=2)
t3 = doc.add_table(rows=1, cols=3)
t3.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr3 = ["Как выглядит на рисунке 1", "Что обозначает", "Из какой фигуры Word собрана"]
data3 = [
    ("Тёмная двойная рамка по контуру здания", "Наружные несущие стены, толщина 40 см",
     "два вложенных «Прямоугольника» без заливки, разная толщина линии (2,6 пт и 1,2 пт)"),
    ("Сплошные тёмные полосы внутри", "Межкомнатные перегородки (12 см)",
     "узкие «Прямоугольники» шириной 0,22 м с заливкой цвета стены"),
    ("Голубые полоски в верхней стене", "Оконные проёмы (стена обращена на север)",
     "четыре «Прямоугольника» с голубой заливкой поверх линии стены"),
    ("Жёлтый штрих + тонкая четверть-окружность", "Двери и сектор распахивания створки",
     "«Прямоугольник» (полотно) + незамкнутый «Произвольный многоугольник» из 13 точек-дуги"),
    ("Бежевые прямоугольники с подписью", "Ученические и демонстрационный столы",
     "«Прямоугольники» с заливкой и текстом внутри самой фигуры"),
    ("Маленькие серые квадраты вокруг столов", "Стулья (по 4 на стол)",
     "«Прямоугольники» 0,34 × 0,30 м"),
    ("Красные линии с наконечником", "Пути эвакуации людей к выходам",
     "фигура «Линия» со стрелкой на конце"),
    ("Круг с красной и серой стрелками", "Роза сторон света (направление «Север»)",
     "«Эллипс» + два замкнутых «Произвольных многоугольника» (треугольные лучи)"),
    ("Чёрно-белая полоса из четырёх делений", "Линейный масштаб 1 : 50 (1 деление = 1 м)",
     "четыре «Прямоугольника» 1,0 × 0,28 см с чередующейся заливкой"),
    ("Серая сетка по всему плану", "Координатная сетка: 1 клетка = 1 м",
     "горизонтальные и вертикальные «Линии» толщиной 0,25 пт"),
    ("Красный скруглённый прямоугольник", "Переносной огнетушитель ОП-5",
     "фигура «Округлый прямоугольник» со значением скругления 32 %"),
    ("Красный крест в белой рамке", "Аптечка первой помощи",
     "два пересекающихся «Прямоугольника» поверх прямоугольной рамки"),
    ("Жёлтые круги с буквами «Г» и «В»", "Клапан подачи газа и воздухозабор",
     "фигура «Эллипс» с текстом внутри"),
    ("Серый прямоугольник «Щит»", "Электрощиток освещения и розеточных групп",
     "«Прямоугольник» с заливкой и подписью"),
    ("Размерные линии со штрихами и числами", "Габариты помещений (12,0 м; 6,0 м; 10,0 м)",
     "«Линии» + повёрнутый на −90° текстовый «Прямоугольник» без рамки"),
    ("Тёмно-синие кружки с цифрами 1–4", "Номера помещений (связь с таблицей 1)",
     "«Эллипс» с тёмно-синей заливкой и белым текстом"),
]
table_borders(t3)
for i, h in enumerate(hdr3):
    set_cell(t3.rows[0].cells[i], h, bold=True, size=9, align="center", fill=NAVY,
             color=(0xFF, 0xFF, 0xFF))
for idx, row in enumerate(data3):
    cells = t3.add_row().cells
    for i, val in enumerate(row):
        set_cell(cells[i], val, size=8.5, align="left")
    if idx % 2 == 1:
        for c in cells:
            c._tc.get_or_add_tcPr().append(shd("F2F6FA"))
w3 = [Cm(5.6), Cm(5.6), Cm(8.4)]
for r_ in t3.rows:
    for i, c in enumerate(r_.cells):
        c.width = w3[i]

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(10)
r = p.add_run("Примечание. Рисунок 1 создан средствами «Вставка → Фигуры» Microsoft Word "
              "(объекты DrawingML, сгруппированные в один плавающий холст). Растровые изображения "
              "(PNG/JPG) в документе не используются, поэтому план остаётся векторным и масштабируется без потерь.")
r.italic = True
r.font.size = Pt(9)
r.font.color.rgb = RGBColor(0x5A, 0x6A, 0x76)

doc.save("Doc1.docx")
print("saved Doc1.docx")
