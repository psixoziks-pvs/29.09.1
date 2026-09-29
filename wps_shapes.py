# -*- coding: utf-8 -*-
"""
Помощник для создания .docx c рисунками из ФИГУР Word (WordProcessingShapes / DrawingML)
и таблиц. Никаких картинок (PNG/JPG) - только векторные фигуры:
прямоугольники, эллипсы, дуги, многоугольники, линии со стрелками и текст внутри фигур.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

NSMAP = dict(nsmap)
NSMAP["wps"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
NSMAP["wpg"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
NSMAP["wp"] = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"


def _el(tag, attrs=None, parent=None):
    """Создаёт элемент в пространстве имён drawingml (префикс a:) или wordprocessingshape (wps:)."""
    if ":" in tag:
        pref, local = tag.split(":", 1)
        uri = NSMAP.get(pref)
        if uri is None:
            raise KeyError(pref)
    else:
        local = tag
        uri = "http://schemas.openxmlformats.org/drawingml/2006/main"
    from docx.oxml.parser import parse_xml
    # собираем элемент строкой, чтобы сохранить нужные префиксы пространств имён
    decls = ' xmlns:%s="%s"' % (pref if ":" in tag else "a", uri)
    e = parse_xml("<%s%s%s/>" % (tag, decls, ""))
    if parent is not None:
        parent.append(e)
    for k, v in (attrs or {}).items():
        e.set(k, str(v))
    return e


def _nsdecls():
    parts = []
    for pfx in ("w", "wp", "a", "wps", "wpg", "pic", "r"):
        uri = NSMAP.get(pfx)
        if uri:
            parts.append('xmlns:%s="%s"' % (pfx, uri))
    return " ".join(parts)


EMU_PER_CM = 360000


def emu(cm):
    return int(round(cm * EMU_PER_CM))


A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS = NSMAP["wps"]


def _child(node, tag):
    """Первый дочерний элемент по имени {uri}local."""
    if ":" in tag:
        pfx, loc = tag.split(":")
        uri = NSMAP[pfx]
    else:
        uri, loc = A, tag
    return node.find("{%s}%s" % (uri, loc))


def _qn(tag):
    if ":" in tag:
        pfx, loc = tag.split(":")
        return "{%s}%s" % (NSMAP[pfx], loc)
    return "{%s}%s" % (A, tag)


class Canvas:
    """Одна «сцена»: плавающая область (anchor), в которой фигуры позиционируются в ЭМЮ."""

    _counter = [0]

    def __init__(self, doc, width_cm, height_cm, top_cm=0.4, left_cm=0.0):
        self.width = emu(width_cm)
        self.height = emu(height_cm)
        Canvas._counter[0] += 1
        self._id = 500 + Canvas._counter[0] * 100
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        # высота «строки» под плавающий рисунок = высота сцены, чтобы текст не наезжал
        pPr = p._p.get_or_add_pPr()
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:line"), str(int(height_cm * 567)))   # строка нужной высоты (в 1/20 pt)
        sp.set(qn("w:lineRule"), "exact")
        pPr.append(sp)
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "center" if left_cm is None else "left")
        pPr.append(jc)
        ind = OxmlElement("w:ind")
        if left_cm is not None:
            ind.set(qn("w:left"), str(int(left_cm * 567)))
        pPr.append(ind)
        self._left = left_cm
        run = p.add_run()
        self.anchor = self._make_anchor(run, top_cm)
        r_el = self.anchor.getprevious()
        if r_el is not None and r_el.tag.endswith("}r") and len(r_el) == 0:
            p._p.remove(r_el)
        self.para = p

    def _make_anchor(self, run, top_cm):
        from docx.oxml.parser import parse_xml
        xml = (
            '<wp:anchor %s distT="0" distB="0" distL="0" distR="0" simplePos="0" '
            'relativeHeight="%d" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            '<wp:positionH relativeFrom="margin">%s</wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>%d</wp:posOffset></wp:positionV>'
            '<wp:extent cx="%d" cy="%d"/>'
            '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:wrapNone/>'
            '<wp:docPr id="%d" name="Рисунок из фигур %d"/>'
            '<wp:cNvGraphicFramePr/>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="%s">'
            '<wps:wsp xmlns:wps="%s"><wps:cNvSpPr/><wps:spPr bwMode="auto">'
            '<a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>'
            '</wps:spPr><wps:bodyPr/></wps:wsp>'
            '</a:graphicData></a:graphic>'
            '</wp:anchor>'
        ) % (_nsdecls(), 20 + Canvas._counter[0],
             ('<wp:align>center</wp:align>' if self._left is None
              else '<wp:posOffset>%d</wp:posOffset>' % emu(self._left)),
             int(top_cm * EMU_PER_CM), self.width, self.height,
             self._id, Canvas._counter[0],
             "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
             WPS_NS, self.width, self.height)
        anchor = parse_xml(xml)
        run._r.addnext(anchor)
        return anchor

    # ---------- низкоуровневые примитивы ----------
    def _graphic_data(self):
        graphic = None
        for ch in self.anchor.iterchildren():
            if ch.tag.endswith("}graphic"):
                graphic = ch
                break
        gd = _child(graphic, "a:graphicData")
        # подменяем URI на группу
        gd.set("uri", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup")
        for child in list(gd):
            gd.remove(child)
        grp = _el("wpg:wgp")
        gd.append(grp)
        nvgrp = _el("wpg:nvGrpSpPr", parent=grp)
        _el("wpg:cNvPr", {"id": self._id + 1, "name": "Группа фигур"}, nvgrp)
        _el("wpg:cNvGrpSpPr", None, nvgrp)
        grpPr = _el("wpg:grpSpPr", None, grp)
        xfrm = _el("a:xfrm", None, grpPr)
        _el("a:off", {"x": 0, "y": 0}, xfrm)
        _el("a:ext", {"cx": self.width, "cy": self.height}, xfrm)
        _el("a:chOff", {"x": 0, "y": 0}, xfrm)
        _el("a:chExt", {"cx": self.width, "cy": self.height}, xfrm)
        return grp

    def _add_shape_node(self, kind, name):
        grp = self._graphic_data_cache
        node = _el("wps:wsp")
        nv = _el("wps:nvSpPr", None, node)
        self._seq = getattr(self, "_seq", 0) + 1
        _el("wps:cNvPr", {"id": self._id + 10 + self._seq, "name": name}, nv)
        _el("wps:cNvSpPr", None, nv)
        _el("wps:spPr", {"bwMode": "auto"}, node)
        _el("wps:bodyPr", {"wrap": "square", "lIns": "36576", "tIns": "18288",
                           "rIns": "36576", "bIns": "18288", "anchor": "ctr",
                           "horzOverflow": "overflow"}, node)
        grp.append(node)
        return node

    def rect(self, x, y, w, h, fill=None, line=None, line_w=1.0, geom="rect", adj=None,
             text=None, size=9, bold=False, color="000000", rot=0, dash=None, align="ctr",
             anchor="ctr", name=None, alpha=None):
        """Главная функция: фигура с необязательным текстом. Координаты в ЭМЮ."""
        node = self._add_shape_node("sp", name or ("Фигура %d" % (getattr(self, "_seq", 0) + 1)))
        spPr = _child(node, "wps:spPr")
        xfrm = _el("a:xfrm", None, spPr)
        if rot:
            xfrm.set("rot", str(int(rot * 60000)))
        _el("a:off", {"x": int(x), "y": int(y)}, xfrm)
        _el("a:ext", {"cx": int(w), "cy": int(h)}, xfrm)
        pg = _el("a:prstGeom", {"prst": geom}, spPr)
        av = _el("a:avLst", None, pg)
        for key, val in (adj or {}).items():
            _el("a:gd", {"name": key, "fmla": "val %d" % int(val)}, av)
        if fill is None:
            _el("a:noFill", None, spPr)
        else:
            sf = _el("a:solidFill", None, spPr)
            c = _el("a:srgbClr", {"val": fill}, sf)
            if alpha is not None:
                _el("a:alpha", {"val": str(int(alpha * 1000))}, c)
        ln = _el("a:ln", {"w": max(1, int(line_w * 12700)), "cap": "flat"}, spPr)
        if line is None:
            _el("a:noFill", None, ln)
        else:
            sf = _el("a:solidFill", None, ln)
            _el("a:srgbClr", {"val": line}, sf)
            if dash:
                _el("a:prstDash", {"val": dash}, ln)
        _el("a:round", None, ln)
        # текст внутри фигуры
        if text is not None:
            tx = _el("wps:txbx", None, node)
            p_el = _el("w:p", None, tx)
            pPr = _el("w:pPr", None, p_el)
            algn = {"ctr": "center", "l": "left", "r": "right"}[align]
            _el("w:jc", {"val": algn}, pPr)
            _el("w:spacing", {"before": "0", "after": "0"}, pPr)
            body = _child(node, "wps:bodyPr")
            body.set("anchor", anchor)
            lines = text.split("\n")
            for i, line_txt in enumerate(lines):
                if i:
                    br_p = _el("w:p", None, tx)
                    bpPr = _el("w:pPr", None, br_p)
                    _el("w:jc", {"val": algn}, bpPr)
                    _el("w:spacing", {"before": "0", "after": "0"}, bpPr)
                    cur = br_p
                else:
                    cur = p_el
                r = _el("w:r", None, cur)
                rPr = _el("w:rPr", None, r)
                if bold:
                    _el("w:b", None, rPr)
                _el("w:sz", {"val": str(int(size * 2))}, rPr)
                _el("w:szCs", {"val": str(int(size * 2))}, rPr)
                _el("w:color", {"val": color}, rPr)
                _el("w:t", None, r).text = line_txt
        return node

    def line(self, x1, y1, x2, y2, color="000000", w=1.0, dash=None, arrow_end=False,
             arrow_beg=False):
        """Прямая линия (выносящая выноска/стрелка)."""
        import math
        dx, dy = x2 - x1, y2 - y1
        length = math.hypot(dx, dy)
        ang = math.degrees(math.atan2(dy, dx))
        node = self._add_shape_node("ln", "Линия")
        spPr = _child(node, "wps:spPr")
        bp = _child(node, "wps:bodyPr")
        for a in list(bp.attrib):
            del bp.attrib[a]
        xfrm = _el("a:xfrm", {"rot": str(int(ang * 60000))}, spPr)
        _el("a:off", {"x": int(x1), "y": int(y1)}, xfrm)
        _el("a:ext", {"cx": int(length), "cy": 0}, xfrm)
        _el("a:prstGeom", {"prst": "line"}, spPr).append(_el("a:avLst"))
        ln = _el("a:ln", {"w": max(1, int(w * 12700))}, spPr)
        sf = _el("a:solidFill", None, ln)
        _el("a:srgbClr", {"val": color}, sf)
        if dash:
            _el("a:prstDash", {"val": dash}, ln)
        if arrow_beg:
            _el("a:headEnd", {"type": "triangle", "w": "med", "len": "med"}, ln)
        if arrow_end:
            _el("a:tailEnd", {"type": "triangle", "w": "med", "len": "med"}, ln)
        return node

    def poly(self, pts, closed=True, fill=None, line="000000", w=1.0, name=None,
             text=None, size=9, color="000000", bold=False):
        """Произвольный многоугольник. pts - список (x,y) в ЭМЮ относительно (0,0) сцены."""
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        minx, miny = min(xs), min(ys)
        bw, bh = max(xs) - minx, max(ys) - miny
        bw = max(bw, 1)
        bh = max(bh, 1)
        node = self._add_shape_node("cust", name or "Многоугольник")
        spPr = _child(node, "wps:spPr")
        xfrm = _el("a:xfrm", None, spPr)
        _el("a:off", {"x": int(minx), "y": int(miny)}, xfrm)
        _el("a:ext", {"cx": int(bw), "cy": int(bh)}, xfrm)
        cust = _el("a:custGeom", None, spPr)
        _el("a:avLst", None, cust)
        _el("a:gdLst", None, cust)
        ahLst = _el("a:ahLst", None, cust)
        _el("a:cxnLst", None, cust)
        _el("a:rect", {"l": "0", "t": "0", "r": "0", "b": "0"}, cust)
        pathLst = _el("a:pathLst", None, cust)
        path = _el("a:path", {"w": int(bw), "h": int(bh)}, pathLst)
        cmds = []
        for i, (px, py) in enumerate(pts):
            lx, ly = int(px - minx), int(py - miny)
            cmds.append(("moveTo" if i == 0 else "lnTo", lx, ly))
        if closed:
            _el("a:closePath", None, path)
        for cmd, lx, ly in cmds:
            ce = _el("a:%s" % cmd, None, path)
            _el("a:pt", {"x": lx, "y": ly}, ce)
        if fill is None:
            _el("a:noFill", None, spPr)
        else:
            sf = _el("a:solidFill", None, spPr)
            _el("a:srgbClr", {"val": fill}, sf)
        ln = _el("a:ln", {"w": max(1, int(w * 12700))}, spPr)
        if line is None:
            _el("a:noFill", None, ln)
        else:
            sf = _el("a:solidFill", None, ln)
            _el("a:srgbClr", {"val": line}, sf)
        if text is not None:
            self._set_text(node, text, size=size, color=color, bold=bold, align="ctr")
        return node

    def _set_text(self, node, text, size=9, color="000000", bold=False, align="ctr"):
        tx = _el("wps:txbx", None, node)
        p_el = _el("w:p", None, tx)
        pPr = _el("w:pPr", None, p_el)
        _el("w:jc", {"val": {"ctr": "center", "l": "left", "r": "right"}[align]}, pPr)
        _el("w:spacing", {"before": "0", "after": "0"}, pPr)
        r = _el("w:r", None, p_el)
        rPr = _el("w:rPr", None, r)
        if bold:
            _el("w:b", None, rPr)
        _el("w:sz", {"val": str(int(size * 2))}, rPr)
        _el("w:color", {"val": color}, rPr)
        _el("w:t", None, r).text = text

    @property
    def _graphic_data_cache(self):
        if not hasattr(self, "_gd"):
            self._gd = self._graphic_data()
        return self._gd


# ------------------------------------------------------------------ утилиты документа
def add_title(doc, text, size=20, space_after=10, space_before=4, italic=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(size)
    r.italic = italic
    r.font.color.rgb = RGBColor(0x1F, 0x3B, 0x63)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    return p


def add_caption(doc, text, size=9):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    return p


def set_cell(cell, text, bold=False, size=9, align="left", fill=None, color=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "right": WD_ALIGN_PARAGRAPH.RIGHT}[align]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = RGBColor(*color)
    if fill:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), fill)
        cell._tc.get_or_add_tcPr().append(shd)
    return cell


def table_borders(table, color="6E7B8B", sz=6):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:%s" % edge)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(sz))
        e.set(qn("w:color"), color)
        borders.append(e)
    tblPr.append(borders)


def landscape(section):
    w, h = section.page_width, section.page_height
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = max(w, h), min(w, h)
    section.left_margin = Cm(1.6)
    section.right_margin = Cm(1.6)
    section.top_margin = Cm(1.4)
    section.bottom_margin = Cm(1.4)
