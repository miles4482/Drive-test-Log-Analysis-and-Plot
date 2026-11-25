#!/usr/bin/env python3
"""Publish Bogura Drive-Test Report v1.93 from v1.90 (Neon restyle; Excel-openable).

Gemini 4 Neon (xhigh) design: strip clutter notes, replace harsh yellow section
bars, restyle banners/tables/cards, keep every map, chart, and movable circle.
Edits the xlsx ZIP directly — never openpyxl-save.
"""

from __future__ import annotations

import copy
import re
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_DRAW = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"

ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)
ET.register_namespace("xdr", NS_DRAW)
ET.register_namespace("a", NS_A)
ET.register_namespace("c", NS_C)

M = {"m": NS_MAIN}
D = {"xdr": NS_DRAW, "a": NS_A, "c": NS_C, "r": NS_REL}

# Style index map after rebuild (see build_styles)
S = {
    "title": 1,
    "title_fill": 2,
    "subtitle": 3,
    "subtitle_fill": 4,
    "section_neutral": 5,
    "section_active": 6,
    "section_idle": 7,
    "h2": 8,
    "th": 9,
    "label_w": 10,
    "value_w": 11,
    "num_w": 12,
    "pct_w": 13,
    "time_w": 14,
    "label_g": 15,
    "value_g": 16,
    "num_g": 17,
    "pct_g": 18,
    "time_g": 19,
    "note": 20,
    "body": 21,
    "card_cap": 22,
    "card_label": 23,
    "card_body": 24,
    "param_cap": 25,
    "existing_cap": 26,
    "new_cap": 27,
    "param_sub": 28,
    "existing_sub": 29,
    "new_sub": 30,
    "param_site": 31,
    "param_exist": 32,
    "param_new": 33,
    "param_new_chg": 34,
    "param_note": 35,
    "callout_h": 36,
    "callout_b": 37,
    "badge_active": 38,
    "badge_idle": 39,
    "badge_4th": 40,
    "badge_tilt": 41,
    "badge_new": 42,
    "badge_ep": 43,
    "badge_swap": 44,
    "badge_3swap": 45,
    "badge_third": 46,
    "total": 47,
    "log_th": 48,
    "log_w": 49,
    "log_g": 50,
    "log_time_w": 51,
    "log_time_g": 52,
    "log_num_w": 53,
    "log_num_g": 54,
    "scatter": 55,
    "feeder_site": 56,
    "kpi_label_w": 57,
    "kpi_label_g": 58,
}


def q(tag: str) -> str:
    return f"{{{NS_MAIN}}}{tag}"


def col_letters(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def parse_ref(ref: str) -> tuple[str, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    if not m:
        raise ValueError(ref)
    return m.group(1), int(m.group(2))


def shift_ref(ref: str, delta: int, min_row: int) -> str:
    m = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
    if m:
        a, r1, b, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        if r1 >= min_row:
            r1 += delta
        if r2 >= min_row:
            r2 += delta
        return f"{a}{r1}:{b}{r2}"
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    if m:
        a, r = m.group(1), int(m.group(2))
        if r >= min_row:
            r += delta
        return f"{a}{r}"
    return ref


def inline_str(text: str) -> ET.Element:
    c = ET.Element(q("c"))
    c.set("t", "inlineStr")
    is_el = ET.SubElement(c, q("is"))
    t = ET.SubElement(is_el, q("t"))
    if text[:1].isspace() or text[-1:].isspace():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return c


def set_inline(cell: ET.Element, text: str, style: int | None = None) -> None:
    cell.set("t", "inlineStr")
    for child in list(cell):
        cell.remove(child)
    is_el = ET.SubElement(cell, q("is"))
    t = ET.SubElement(is_el, q("t"))
    if text[:1].isspace() or text[-1:].isspace():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    if style is not None:
        cell.set("s", str(style))


def cell_text(cell: ET.Element) -> str:
    if cell.get("t") == "inlineStr":
        return "".join(t.text or "" for t in cell.iter(q("t")))
    v = cell.find(q("v"))
    return v.text if v is not None and v.text else ""


def ensure_cell(row: ET.Element, col: str, style: int | None = None) -> ET.Element:
    r = int(row.get("r"))
    ref = f"{col}{r}"
    for c in row.findall(q("c")):
        if c.get("r") == ref:
            if style is not None:
                c.set("s", str(style))
            return c
    c = inline_str("")
    c.set("r", ref)
    if style is not None:
        c.set("s", str(style))
    # keep cells sorted by column
    inserted = False
    for i, existing in enumerate(list(row)):
        if existing.tag != q("c"):
            continue
        letters, _ = parse_ref(existing.get("r"))
        if col_index(letters) > col_index(col):
            row.insert(i, c)
            inserted = True
            break
    if not inserted:
        row.append(c)
    return c


def get_row(sheet: ET.Element, r: int) -> ET.Element | None:
    sd = sheet.find(q("sheetData"))
    for row in sd.findall(q("row")):
        if int(row.get("r")) == r:
            return row
    return None


def clear_row_cells(row: ET.Element) -> None:
    for c in list(row.findall(q("c"))):
        row.remove(c)


def parse_merge_bounds(ref: str) -> tuple[int, int, int, int] | None:
    m = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref or "")
    if not m:
        return None
    c1, r1, c2, r2 = col_index(m.group(1)), int(m.group(2)), col_index(m.group(3)), int(m.group(4))
    if r1 > r2 or c1 > c2:
        return None
    return c1, r1, c2, r2


def merges_overlap(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    return a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]


def sanitize_merges(sheet: ET.Element) -> None:
    """Drop invalid / overlapping merges. First kept ref wins (call with preferred first)."""
    cleaned: list[str] = []
    bounds: list[tuple[int, int, int, int]] = []
    for ref in collect_merges(sheet):
        b = parse_merge_bounds(ref)
        if b is None:
            continue
        if any(merges_overlap(b, kb) for kb in bounds):
            continue
        cleaned.append(ref)
        bounds.append(b)
    set_merges(sheet, cleaned)


def delete_rows(sheet: ET.Element, rows: set[int]) -> None:
    sd = sheet.find(q("sheetData"))
    for row in list(sd.findall(q("row"))):
        if int(row.get("r")) in rows:
            sd.remove(row)
    merges = sheet.find(q("mergeCells"))
    if merges is None:
        return
    for mc in list(merges.findall(q("mergeCell"))):
        b = parse_merge_bounds(mc.get("ref"))
        if b is None:
            merges.remove(mc)
            continue
        _, r1, _, r2 = b
        if any(r in rows for r in range(r1, r2 + 1)):
            merges.remove(mc)


def shift_sheet_rows(sheet: ET.Element, delta: int, min_row: int) -> None:
    sd = sheet.find(q("sheetData"))
    for row in sd.findall(q("row")):
        r = int(row.get("r"))
        if r < min_row:
            continue
        new_r = r + delta
        row.set("r", str(new_r))
        for c in row.findall(q("c")):
            letters, _ = parse_ref(c.get("r"))
            c.set("r", f"{letters}{new_r}")
    merges = sheet.find(q("mergeCells"))
    if merges is not None:
        for mc in merges.findall(q("mergeCell")):
            mc.set("ref", shift_ref(mc.get("ref"), delta, min_row))
    dim = sheet.find(q("dimension"))
    if dim is not None and dim.get("ref"):
        dim.set("ref", shift_ref(dim.get("ref"), delta, min_row))
    af = sheet.find(q("autoFilter"))
    if af is not None and af.get("ref"):
        af.set("ref", shift_ref(af.get("ref"), delta, min_row))


def set_merges(sheet: ET.Element, refs: list[str]) -> None:
    old = sheet.find(q("mergeCells"))
    parent = sheet
    idx = list(parent).index(old) if old is not None else None
    if old is not None:
        parent.remove(old)
    if not refs:
        return
    merges = ET.Element(q("mergeCells"))
    merges.set("count", str(len(refs)))
    for ref in refs:
        mc = ET.SubElement(merges, q("mergeCell"))
        mc.set("ref", ref)
    if idx is not None:
        parent.insert(idx, merges)
    else:
        # before printOptions / pageMargins / drawing
        insert_at = len(list(parent))
        for i, child in enumerate(list(parent)):
            if child.tag in (q("printOptions"), q("pageMargins"), q("pageSetup"), q("drawing"), q("autoFilter")):
                insert_at = i
                break
        parent.insert(insert_at, merges)


def collect_merges(sheet: ET.Element) -> list[str]:
    merges = sheet.find(q("mergeCells"))
    if merges is None:
        return []
    return [mc.get("ref") for mc in merges.findall(q("mergeCell"))]


def set_tab_color(sheet: ET.Element, rgb: str) -> None:
    pr = sheet.find(q("sheetPr"))
    if pr is None:
        pr = ET.Element(q("sheetPr"))
        sheet.insert(0, pr)
    tab = pr.find(q("tabColor"))
    if tab is None:
        tab = ET.SubElement(pr, q("tabColor"))
    tab.set("rgb", "FF" + rgb)


def set_banner(sheet: ET.Element, title: str, subtitle: str, last_col: str) -> None:
    r1 = get_row(sheet, 1)
    r2 = get_row(sheet, 2)
    if r1 is None or r2 is None:
        return
    r1.set("ht", "34")
    r1.set("customHeight", "1")
    r2.set("ht", "22")
    r2.set("customHeight", "1")
    clear_row_cells(r1)
    clear_row_cells(r2)
    c1 = ensure_cell(r1, "A", S["title"])
    set_inline(c1, title, S["title"])
    # paint fill across merge visually by styling A only (merge carries style of top-left)
    for col_i in range(2, col_index(last_col) + 1):
        ensure_cell(r1, col_letters(col_i), S["title_fill"])
        ensure_cell(r2, col_letters(col_i), S["subtitle_fill"])
    c2 = ensure_cell(r2, "A", S["subtitle"])
    set_inline(c2, subtitle, S["subtitle"])


def style_solution_badge(text: str) -> int:
    t = text.strip().lower()
    if t == "active":
        return S["badge_active"]
    if t == "idle":
        return S["badge_idle"]
    if "4th sector" in t:
        return S["badge_4th"]
    if "new 3-sector" in t or t.startswith("new site"):
        return S["badge_new"]
    if "on site ep" in t:
        return S["badge_ep"]
    if t == "3-sector swap":
        return S["badge_3swap"]
    if t == "swap":
        return S["badge_swap"]
    if t == "third sector":
        return S["badge_third"]
    if "tilt" in t or "azimuth" in t or "uptilt" in t:
        return S["badge_tilt"]
    return S["body"]


def build_styles() -> bytes:
    fonts = [
        # 0 Normal lock
        '<font><name val="Calibri"/><family val="2"/><color theme="1"/><sz val="11"/><scheme val="minor"/></font>',
        # 1 title white 17 bold
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="00FFFFFF"/><sz val="17"/></font>',
        # 2 subtitle silver 10
        '<font><name val="Segoe UI"/><color rgb="00CBD5E1"/><sz val="10"/></font>',
        # 3 section neutral
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="000F2942"/><sz val="11"/></font>',
        # 4 section active
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="001E3A8A"/><sz val="11"/></font>',
        # 5 section idle
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="001E293B"/><sz val="11"/></font>',
        # 6 h2
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="000F2942"/><sz val="11"/></font>',
        # 7 dark header white 9.5
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="00FFFFFF"/><sz val="9.5"/></font>',
        # 8 body
        '<font><name val="Segoe UI"/><color rgb="001E293B"/><sz val="9.5"/></font>',
        # 9 body strong
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="000F172A"/><sz val="9.5"/></font>',
        # 10 muted
        '<font><name val="Segoe UI"/><color rgb="00475569"/><sz val="9.5"/></font>',
        # 11 label 9 bold
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="000F172A"/><sz val="9"/></font>',
        # 12 teal new
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="00065F46"/><sz val="9.5"/></font>',
        # 13 callout title
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="000369A1"/><sz val="9.5"/></font>',
        # 14 badge active
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="001E40AF"/><sz val="9"/></font>',
        # 15 badge idle
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="00475569"/><sz val="9"/></font>',
        # 16 badge 4th
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="003730A3"/><sz val="9"/></font>',
        # 17 badge tilt
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="00115E59"/><sz val="9"/></font>',
        # 18 badge new
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="009F1239"/><sz val="9"/></font>',
        # 19 badge ep
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="0092400E"/><sz val="9"/></font>',
        # 20 badge swap
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="009A3412"/><sz val="9"/></font>',
        # 21 badge third
        '<font><name val="Segoe UI"/><color rgb="0064748B"/><sz val="9"/></font>',
        # 22 existing header white
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="00FFFFFF"/><sz val="9.5"/></font>',
        # 23 proposed header white
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="00FFFFFF"/><sz val="9.5"/></font>',
        # 24 existing sub
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="00334155"/><sz val="9"/></font>',
        # 25 proposed sub
        '<font><name val="Segoe UI"/><b val="1"/><color rgb="000F766E"/><sz val="9"/></font>',
    ]
    fills = [
        "<fill><patternFill/></fill>",
        '<fill><patternFill patternType="gray125"/></fill>',
        '<fill><patternFill patternType="solid"><fgColor rgb="000F2942"/></patternFill></fill>',  # 2 title
        '<fill><patternFill patternType="solid"><fgColor rgb="001B3B5A"/></patternFill></fill>',  # 3 subtitle
        '<fill><patternFill patternType="solid"><fgColor rgb="00E8EEF5"/></patternFill></fill>',  # 4 neutral
        '<fill><patternFill patternType="solid"><fgColor rgb="00E8F0FE"/></patternFill></fill>',  # 5 active
        '<fill><patternFill patternType="solid"><fgColor rgb="00EDF2F7"/></patternFill></fill>',  # 6 idle
        '<fill><patternFill patternType="solid"><fgColor rgb="00FFFFFF"/></patternFill></fill>',  # 7 white
        '<fill><patternFill patternType="solid"><fgColor rgb="00F8FAFC"/></patternFill></fill>',  # 8 zebra
        '<fill><patternFill patternType="solid"><fgColor rgb="00F1F5F9"/></patternFill></fill>',  # 9 soft
        '<fill><patternFill patternType="solid"><fgColor rgb="001E3A5F"/></patternFill></fill>',  # 10 card cap
        '<fill><patternFill patternType="solid"><fgColor rgb="00475569"/></patternFill></fill>',  # 11 existing
        '<fill><patternFill patternType="solid"><fgColor rgb="000F766E"/></patternFill></fill>',  # 12 new
        '<fill><patternFill patternType="solid"><fgColor rgb="00E2E8F0"/></patternFill></fill>',  # 13 param sub
        '<fill><patternFill patternType="solid"><fgColor rgb="00E6F4F1"/></patternFill></fill>',  # 14 new sub
        '<fill><patternFill patternType="solid"><fgColor rgb="00ECFDF5"/></patternFill></fill>',  # 15 changed
        '<fill><patternFill patternType="solid"><fgColor rgb="00E0F2FE"/></patternFill></fill>',  # 16 callout h
        '<fill><patternFill patternType="solid"><fgColor rgb="00F0F7FF"/></patternFill></fill>',  # 17 callout b
        '<fill><patternFill patternType="solid"><fgColor rgb="00DBEAFE"/></patternFill></fill>',  # 18 badge active
        '<fill><patternFill patternType="solid"><fgColor rgb="00E0E7FF"/></patternFill></fill>',  # 19 badge 4th
        '<fill><patternFill patternType="solid"><fgColor rgb="00CCFBF1"/></patternFill></fill>',  # 20 badge tilt
        '<fill><patternFill patternType="solid"><fgColor rgb="00FFE4E6"/></patternFill></fill>',  # 21 badge new
        '<fill><patternFill patternType="solid"><fgColor rgb="00FEF3C7"/></patternFill></fill>',  # 22 badge ep
        '<fill><patternFill patternType="solid"><fgColor rgb="00FFEDD5"/></patternFill></fill>',  # 23 badge swap
    ]
    borders = [
        "<border><left/><right/><top/><bottom/><diagonal/></border>",
        # 1 subtle grid
        '<border><left style="thin"><color rgb="00E2E8F0"/></left><right style="thin"><color rgb="00E2E8F0"/></right>'
        '<top style="thin"><color rgb="00E2E8F0"/></top><bottom style="thin"><color rgb="00E2E8F0"/></bottom></border>',
        # 2 bottom only
        '<border><left/><right/><top/><bottom style="thin"><color rgb="00E2E8F0"/></bottom><diagonal/></border>',
        # 3 header anchor
        '<border><left style="thin"><color rgb="001E3A5F"/></left><right style="thin"><color rgb="001E3A5F"/></right>'
        '<top style="thin"><color rgb="000F2942"/></top><bottom style="medium"><color rgb="000F2942"/></bottom></border>',
        # 4 total
        '<border><left style="thin"><color rgb="00E2E8F0"/></left><right style="thin"><color rgb="00E2E8F0"/></right>'
        '<top style="thin"><color rgb="0094A3B8"/></top><bottom style="double"><color rgb="000F2942"/></bottom></border>',
        # 5 banner cyan bottom
        '<border><left/><right/><top/><bottom style="medium"><color rgb="000EA5E9"/></bottom><diagonal/></border>',
        # 6 neutral section
        '<border><left style="medium"><color rgb="000F2942"/></left><right style="thin"><color rgb="00CBD5E1"/></right>'
        '<top style="thin"><color rgb="00CBD5E1"/></top><bottom style="medium"><color rgb="000F2942"/></bottom></border>',
        # 7 active section
        '<border><left style="medium"><color rgb="002563EB"/></left><right style="thin"><color rgb="00BFDBFE"/></right>'
        '<top style="thin"><color rgb="00BFDBFE"/></top><bottom style="medium"><color rgb="002563EB"/></bottom></border>',
        # 8 idle section
        '<border><left style="medium"><color rgb="0064748B"/></left><right style="thin"><color rgb="00CBD5E1"/></right>'
        '<top style="thin"><color rgb="00CBD5E1"/></top><bottom style="medium"><color rgb="0064748B"/></bottom></border>',
        # 9 card label
        '<border><left style="medium"><color rgb="001E3A5F"/></left><right style="thin"><color rgb="00CBD5E1"/></right>'
        '<top style="thin"><color rgb="00CBD5E1"/></top><bottom style="thin"><color rgb="00CBD5E1"/></bottom></border>',
        # 10 card body
        '<border><left style="thin"><color rgb="00E2E8F0"/></left><right style="thin"><color rgb="00CBD5E1"/></right>'
        '<top style="thin"><color rgb="00CBD5E1"/></top><bottom style="thin"><color rgb="00CBD5E1"/></bottom></border>',
        # 11 callout
        '<border><left style="medium"><color rgb="000284C7"/></left><right style="thin"><color rgb="00BAE6FD"/></right>'
        '<top style="thin"><color rgb="00BAE6FD"/></top><bottom style="thin"><color rgb="00BAE6FD"/></bottom></border>',
        # 12 feeder group
        '<border><left style="thin"><color rgb="00E2E8F0"/></left><right style="thin"><color rgb="00E2E8F0"/></right>'
        '<top style="thin"><color rgb="00E2E8F0"/></top><bottom style="medium"><color rgb="00CBD5E1"/></bottom></border>',
    ]

    def xf(nf, font, fill, border, align="", apply_num=False):
        bits = [f'numFmtId="{nf}"', f'fontId="{font}"', f'fillId="{fill}"', f'borderId="{border}"', 'xfId="0"']
        if apply_num:
            bits.append('applyNumberFormat="1"')
        if align:
            bits.append('applyAlignment="1"')
        bits.append('applyFont="1"')
        bits.append('applyFill="1"')
        bits.append('applyBorder="1"')
        open_tag = "<xf " + " ".join(bits) + ">"
        if align:
            return open_tag + f"<alignment {align}/></xf>"
        return open_tag[:-1] + "/>"

    xfs = [
        xf(0, 0, 0, 0),  # 0
        xf(0, 1, 2, 0, 'horizontal="left" vertical="center" indent="1"'),  # 1 title
        xf(0, 1, 2, 0),  # 2 title fill
        xf(0, 2, 3, 5, 'horizontal="left" vertical="center" indent="1"'),  # 3 subtitle
        xf(0, 2, 3, 5),  # 4 subtitle fill
        xf(0, 3, 4, 6, 'horizontal="left" vertical="center" indent="1"'),  # 5 section neutral
        xf(0, 4, 5, 7, 'horizontal="left" vertical="center" indent="1"'),  # 6 section active
        xf(0, 5, 6, 8, 'horizontal="left" vertical="center" indent="1"'),  # 7 section idle
        xf(0, 6, 0, 0, 'horizontal="left" vertical="center"'),  # 8 h2
        xf(0, 7, 2, 3, 'horizontal="center" vertical="center" wrapText="1"'),  # 9 th
        xf(0, 8, 7, 1, 'horizontal="left" vertical="center"'),  # 10 label w
        xf(0, 8, 7, 1, 'horizontal="left" vertical="center"'),  # 11 value w
        xf(0, 8, 7, 1, 'horizontal="center" vertical="center"'),  # 12 num w  (general / use with nf)
        xf(167, 8, 7, 1, 'horizontal="center" vertical="center"', True),  # 13 pct w
        xf(166, 8, 7, 1, 'horizontal="center" vertical="center"', True),  # 14 time w
        xf(0, 8, 8, 1, 'horizontal="left" vertical="center"'),  # 15 label g
        xf(0, 8, 8, 1, 'horizontal="left" vertical="center"'),  # 16 value g
        xf(0, 8, 8, 1, 'horizontal="center" vertical="center"'),  # 17 num g
        xf(167, 8, 8, 1, 'horizontal="center" vertical="center"', True),  # 18 pct g
        xf(166, 8, 8, 1, 'horizontal="center" vertical="center"', True),  # 19 time g
        xf(0, 10, 0, 0, 'horizontal="left" vertical="center" wrapText="1"'),  # 20 note
        xf(0, 8, 0, 0, 'horizontal="left" vertical="top" wrapText="1"'),  # 21 body
        xf(0, 7, 10, 0, 'horizontal="left" vertical="center" indent="1"'),  # 22 card cap
        xf(0, 11, 9, 9, 'horizontal="left" vertical="center" wrapText="1" indent="1"'),  # 23 card label
        xf(0, 8, 7, 10, 'horizontal="left" vertical="top" wrapText="1"'),  # 24 card body
        xf(0, 7, 10, 3, 'horizontal="center" vertical="center"'),  # 25 param cap
        xf(0, 22, 11, 3, 'horizontal="center" vertical="center"'),  # 26 existing cap
        xf(0, 23, 12, 3, 'horizontal="center" vertical="center"'),  # 27 new cap
        xf(0, 11, 13, 3, 'horizontal="center" vertical="center"'),  # 28 param sub
        xf(0, 24, 6, 3, 'horizontal="center" vertical="center"'),  # 29 existing sub
        xf(0, 25, 14, 3, 'horizontal="center" vertical="center"'),  # 30 new sub
        xf(0, 9, 7, 1, 'horizontal="left" vertical="center"'),  # 31 param site
        xf(0, 10, 7, 1, 'horizontal="center" vertical="center"'),  # 32 param exist
        xf(0, 8, 7, 1, 'horizontal="center" vertical="center"'),  # 33 param new
        xf(0, 12, 15, 1, 'horizontal="center" vertical="center"'),  # 34 param new chg
        xf(0, 8, 7, 1, 'horizontal="left" vertical="center"'),  # 35 param note
        xf(0, 13, 16, 11, 'horizontal="left" vertical="center" indent="1"'),  # 36 callout h
        xf(0, 8, 17, 11, 'horizontal="left" vertical="top" wrapText="1" indent="1"'),  # 37 callout b
        xf(0, 14, 18, 1, 'horizontal="center" vertical="center"'),  # 38 badge active
        xf(0, 15, 9, 1, 'horizontal="center" vertical="center"'),  # 39 badge idle
        xf(0, 16, 19, 1, 'horizontal="center" vertical="center"'),  # 40 badge 4th
        xf(0, 17, 20, 1, 'horizontal="center" vertical="center"'),  # 41 badge tilt
        xf(0, 18, 21, 1, 'horizontal="center" vertical="center"'),  # 42 badge new
        xf(0, 19, 22, 1, 'horizontal="center" vertical="center"'),  # 43 badge ep
        xf(0, 20, 23, 1, 'horizontal="center" vertical="center"'),  # 44 badge swap
        xf(0, 18, 21, 1, 'horizontal="center" vertical="center"'),  # 45 badge 3swap
        xf(0, 21, 9, 1, 'horizontal="center" vertical="center"'),  # 46 badge third
        xf(0, 9, 9, 4, 'horizontal="center" vertical="center"'),  # 47 total
        xf(0, 7, 2, 3, 'horizontal="center" vertical="center"'),  # 48 log th
        xf(0, 8, 7, 2, 'horizontal="center" vertical="center"'),  # 49 log w
        xf(0, 8, 8, 2, 'horizontal="center" vertical="center"'),  # 50 log g
        xf(170, 8, 7, 2, 'horizontal="center" vertical="center"', True),  # 51 log time w
        xf(170, 8, 8, 2, 'horizontal="center" vertical="center"', True),  # 52 log time g
        xf(171, 8, 7, 2, 'horizontal="center" vertical="center"', True),  # 53 log num w
        xf(171, 8, 8, 2, 'horizontal="center" vertical="center"', True),  # 54 log num g
        xf(0, 3, 4, 6, 'horizontal="left" vertical="center" indent="1"'),  # 55 scatter
        xf(0, 3, 4, 6, 'horizontal="left" vertical="center" indent="1"'),  # 56 feeder site
        xf(3, 8, 7, 1, 'horizontal="center" vertical="center"', True),  # 57 kpi int w
        xf(3, 8, 8, 1, 'horizontal="center" vertical="center"', True),  # 58 kpi int g
        xf(2, 8, 7, 1, 'horizontal="center" vertical="center"', True),  # 59 kpi float w
        xf(2, 8, 8, 1, 'horizontal="center" vertical="center"', True),  # 60 kpi float g
        xf(168, 8, 7, 1, 'horizontal="center" vertical="center"', True),  # 61 1dp w
        xf(168, 8, 8, 1, 'horizontal="center" vertical="center"', True),  # 62 1dp g
        xf(169, 8, 7, 1, 'horizontal="center" vertical="center"', True),  # 63 5dp w
        xf(169, 8, 8, 1, 'horizontal="center" vertical="center"', True),  # 64 5dp g
        xf(0, 9, 7, 1, 'horizontal="left" vertical="top" wrapText="1" indent="1"'),  # 65 detail body
        xf(0, 9, 8, 1, 'horizontal="left" vertical="top" wrapText="1" indent="1"'),  # 66 detail body zebra
        xf(0, 8, 7, 1, 'horizontal="center" vertical="center"'),  # 67 center w
        xf(0, 8, 8, 1, 'horizontal="center" vertical="center"'),  # 68 center g
        xf(0, 9, 7, 1, 'horizontal="center" vertical="center"'),  # 69 strong center w
        xf(0, 9, 8, 1, 'horizontal="center" vertical="center"'),  # 70 strong center g
        xf(0, 8, 7, 1, 'horizontal="right" vertical="center"'),  # 71 right w
        xf(0, 8, 8, 1, 'horizontal="right" vertical="center"'),  # 72 right g
        xf(0, 8, 7, 1, 'horizontal="left" vertical="top" wrapText="1"'),  # 73 wrap w
        xf(0, 8, 8, 1, 'horizontal="left" vertical="top" wrapText="1"'),  # 74 wrap g
        xf(0, 7, 10, 3, 'horizontal="left" vertical="center" indent="1"'),  # 75 card title left
        xf(0, 8, 9, 1, 'horizontal="left" vertical="center"'),  # 76 soft label
    ]
    # remap extras used above
    S["kpi_int_w"] = 57
    S["kpi_int_g"] = 58
    S["kpi_float_w"] = 59
    S["kpi_float_g"] = 60
    S["kpi_1dp_w"] = 61
    S["kpi_1dp_g"] = 62
    S["kpi_5dp_w"] = 63
    S["kpi_5dp_g"] = 64
    S["detail_w"] = 65
    S["detail_g"] = 66
    S["center_w"] = 67
    S["center_g"] = 68
    S["strong_c_w"] = 69
    S["strong_c_g"] = 70
    S["right_w"] = 71
    S["right_g"] = 72
    S["wrap_w"] = 73
    S["wrap_g"] = 74
    S["card_title"] = 75
    S["soft_label"] = 76

    xml = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<styleSheet xmlns="{NS_MAIN}">',
        '<numFmts count="8">',
        '<numFmt numFmtId="164" formatCode="yyyy-mm-dd h:mm:ss"/>',
        '<numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm"/>',
        '<numFmt numFmtId="166" formatCode="yyyy-mm-dd hh:mm:ss"/>',
        '<numFmt numFmtId="167" formatCode="0.0%"/>',
        '<numFmt numFmtId="168" formatCode="#,##0.0"/>',
        '<numFmt numFmtId="169" formatCode="0.00000"/>',
        '<numFmt numFmtId="170" formatCode="yyyy-mm-dd hh:mm:ss.000"/>',
        '<numFmt numFmtId="171" formatCode="0.000"/>',
        "</numFmts>",
        f'<fonts count="{len(fonts)}">',
        *fonts,
        "</fonts>",
        f'<fills count="{len(fills)}">',
        *fills,
        "</fills>",
        f'<borders count="{len(borders)}">',
        *borders,
        "</borders>",
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
        f'<cellXfs count="{len(xfs)}">',
        *xfs,
        "</cellXfs>",
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
        '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>',
        # Keep indexed palette from the source workbook so Excel theme/index colours resolve.
        "<colors><indexedColors>"
        '<rgbColor rgb="00000000"/><rgbColor rgb="00FFFFFF"/><rgbColor rgb="00FF0000"/>'
        '<rgbColor rgb="0000FF00"/><rgbColor rgb="000000FF"/><rgbColor rgb="00FFFF00"/>'
        '<rgbColor rgb="00FF00FF"/><rgbColor rgb="0000FFFF"/><rgbColor rgb="00000000"/>'
        '<rgbColor rgb="00FFFFFF"/><rgbColor rgb="00FF0000"/><rgbColor rgb="0000FF00"/>'
        '<rgbColor rgb="000000FF"/><rgbColor rgb="00FFFF00"/><rgbColor rgb="00FF00FF"/>'
        '<rgbColor rgb="0000FFFF"/><rgbColor rgb="00800000"/><rgbColor rgb="00008000"/>'
        '<rgbColor rgb="00000080"/><rgbColor rgb="00808000"/><rgbColor rgb="00800080"/>'
        '<rgbColor rgb="00008080"/><rgbColor rgb="00C0C0C0"/><rgbColor rgb="00808080"/>'
        '<rgbColor rgb="009999FF"/><rgbColor rgb="00993366"/><rgbColor rgb="00FFFFCC"/>'
        '<rgbColor rgb="00CCFFFF"/><rgbColor rgb="00660066"/><rgbColor rgb="00FF8080"/>'
        '<rgbColor rgb="000066CC"/><rgbColor rgb="00CCCCFF"/><rgbColor rgb="00000080"/>'
        '<rgbColor rgb="00FF00FF"/><rgbColor rgb="00FFFF00"/><rgbColor rgb="0000FFFF"/>'
        '<rgbColor rgb="00800080"/><rgbColor rgb="00800000"/><rgbColor rgb="00008080"/>'
        '<rgbColor rgb="000000FF"/><rgbColor rgb="0000CCFF"/><rgbColor rgb="00CCFFFF"/>'
        '<rgbColor rgb="00CCFFCC"/><rgbColor rgb="00FFFF99"/><rgbColor rgb="0099CCFF"/>'
        '<rgbColor rgb="00FF99CC"/><rgbColor rgb="00CC99FF"/><rgbColor rgb="00FFCC99"/>'
        '<rgbColor rgb="003366FF"/><rgbColor rgb="0033CCCC"/><rgbColor rgb="0099CC00"/>'
        '<rgbColor rgb="00FFCC00"/><rgbColor rgb="00FF9900"/><rgbColor rgb="00FF6600"/>'
        '<rgbColor rgb="00666699"/><rgbColor rgb="00969696"/><rgbColor rgb="00003366"/>'
        '<rgbColor rgb="00339966"/><rgbColor rgb="00003300"/><rgbColor rgb="00333300"/>'
        '<rgbColor rgb="00993300"/><rgbColor rgb="00993366"/><rgbColor rgb="00333399"/>'
        '<rgbColor rgb="00333333"/>'
        "</indexedColors></colors>",
        "</styleSheet>",
    ]
    return "".join(xml).encode("utf-8")


def shift_drawing_rows(drawing_xml: bytes, delta: int, min_row_0based: int = 0) -> bytes:
    text = drawing_xml.decode("utf-8")

    def repl(match: re.Match) -> str:
        row = int(match.group(1))
        if row >= min_row_0based:
            row += delta
        return f"<row>{row}</row>"

    # only from/to row tags inside anchors
    text = re.sub(r"<row>(\d+)</row>", repl, text)
    return text.encode("utf-8")


def polish_charts(parts: dict[str, bytes]) -> None:
    title_rpr = (
        '<a:rPr sz="1050" b="1">'
        '<a:solidFill><a:srgbClr val="0F2942"/></a:solidFill>'
        '<a:latin typeface="Segoe UI"/>'
        "</a:rPr>"
    )
    for name, data in list(parts.items()):
        if not name.startswith("xl/charts/chart") or not name.endswith(".xml"):
            continue
        xml = data.decode("utf-8")
        # Insert rPr inside the first title run: <a:r><a:t>… → <a:r><a:rPr…/><a:t>…
        # Do NOT wrap another <a:r> (that nested-tag bug made Excel reject v1.91).
        xml = re.sub(r"(<a:r>)(<a:t>)", rf"\1{title_rpr}\2", xml, count=1)
        # series colour upgrades for KPI vs charts
        xml = xml.replace('srgbClr val="48C9B0"', 'srgbClr val="0D9488"')
        xml = xml.replace('srgbClr val="5DADE2"', 'srgbClr val="2563EB"')
        xml = xml.replace('<a:ln w="25000">', '<a:ln w="28575">')
        # Recolor series fill only. Do not inject <dPt> after <val> — wrong CT_Ser
        # child order makes Excel report the workbook as corrupt.
        if "RSRP ranges (dBm)" in xml and "C0392B" in xml:
            xml = xml.replace(
                '<spPr><a:solidFill><a:srgbClr val="C0392B"/></a:solidFill>',
                '<spPr><a:solidFill><a:srgbClr val="15803D"/></a:solidFill>',
                1,
            )
        parts[name] = xml.encode("utf-8")


def restyle_table_block(
    sheet: ET.Element,
    start: int,
    end: int,
    label_cols: set[str],
    value_cols: set[str],
    int_cols: set[str] | None = None,
    float_cols: set[str] | None = None,
    pct_cols: set[str] | None = None,
    time_cols: set[str] | None = None,
) -> None:
    int_cols = int_cols or set()
    float_cols = float_cols or set()
    pct_cols = pct_cols or set()
    time_cols = time_cols or set()
    for r in range(start, end + 1):
        row = get_row(sheet, r)
        if row is None:
            continue
        zebra = r % 2 == 0
        for c in row.findall(q("c")):
            letters, _ = parse_ref(c.get("r"))
            if letters in label_cols:
                c.set("s", str(S["label_g"] if zebra else S["label_w"]))
            elif letters in pct_cols:
                c.set("s", str(S["pct_g"] if zebra else S["pct_w"]))
            elif letters in time_cols:
                c.set("s", str(S["time_g"] if zebra else S["time_w"]))
            elif letters in int_cols:
                c.set("s", str(S["kpi_int_g"] if zebra else S["kpi_int_w"]))
            elif letters in float_cols:
                c.set("s", str(S["kpi_float_g"] if zebra else S["kpi_float_w"]))
            elif letters in value_cols:
                # text values
                if c.get("t") in (None, "n"):
                    c.set("s", str(S["center_g"] if zebra else S["center_w"]))
                else:
                    c.set("s", str(S["value_g"] if zebra else S["value_w"]))


def style_header_row(sheet: ET.Element, r: int, cols: list[str]) -> None:
    row = get_row(sheet, r)
    if row is None:
        return
    row.set("ht", "22")
    row.set("customHeight", "1")
    for col in cols:
        cell = None
        for c in row.findall(q("c")):
            if c.get("r") == f"{col}{r}":
                cell = c
                break
        if cell is not None:
            cell.set("s", str(S["th"]))


def transform_cover(sheet: ET.Element) -> None:
    remap_legacy_styles(sheet)
    set_tab_color(sheet, "0F2942")
    set_banner(
        sheet,
        "Bogura 4G LTE Drive-Test Report",
        "Executive Summary  |  Active & Idle Mode  |  Generated 2026-11-11 07:19",
        "T",
    )
    # section bars
    for r, style, text in (
        (3, S["section_idle"], "Idle Mode"),
        (42, S["section_active"], "Active Mode"),
        (72, S["section_neutral"], "Drive-Test Route"),
        (110, S["section_neutral"], "Bad Spot Action Summary"),
    ):
        row = get_row(sheet, r)
        if row is None:
            continue
        row.set("ht", "24")
        row.set("customHeight", "1")
        clear_row_cells(row)
        cell = ensure_cell(row, "A", style)
        set_inline(cell, text, style)
        for i in range(2, 21):
            ensure_cell(row, col_letters(i), style)

    # remove clutter rows 18-24, 36 note, 50 long note, 91 note
    for r in range(18, 25):
        row = get_row(sheet, r)
        if row is not None:
            clear_row_cells(row)
            row.set("ht", "4")
            row.set("customHeight", "1")
    for r in (36, 91):
        row = get_row(sheet, r)
        if row is not None:
            clear_row_cells(row)
            row.set("ht", "4")
            row.set("customHeight", "1")
    # clear D50 long note but keep A50/B50
    row50 = get_row(sheet, 50)
    if row50 is not None:
        for c in list(row50.findall(q("c"))):
            letters, _ = parse_ref(c.get("r"))
            if letters not in ("A", "B"):
                row50.remove(c)

    # rename headings
    for r, col, text, style in (
        (5, "A", "Dataset", S["h2"]),
        (5, "D", "Radio KPIs (all samples)", S["h2"]),
        (5, "L", "Radio KPI Distributions — Idle Mode", S["h2"]),
        (25, "A", "Band mix (DL EARFCN)", S["h2"]),
        (25, "E", "RSRP ranges (dBm)", S["h2"]),
        (44, "A", "Dataset", S["h2"]),
        (44, "D", "Radio KPIs — Active Mode (Robi)", S["h2"]),
        (44, "L", "Radio KPI Distributions — Active Mode", S["h2"]),
        (74, "A", "Drive details", S["h2"]),
        (74, "L", "DT route plot", S["h2"]),
        (74, "P", "Bad Spot Summary by Mode", S["card_title"]),
    ):
        row = get_row(sheet, r)
        if row is None:
            continue
        cell = None
        for c in row.findall(q("c")):
            if c.get("r") == f"{col}{r}":
                cell = c
                break
        if cell is None and col == "P":
            cell = ensure_cell(row, col, style)
        if cell is not None:
            set_inline(cell, text, style)

    # table headers
    for r, cols in (
        (6, list("ABCDEFGHIJK")),
        (26, list("ABC") + list("EFG")),
        (45, list("ABCDEFGHIJK")),
        (75, list("AB")),
        (76, list("PQR")),
        (112, list("ABC")),
        (119, list("ABCDEFGH")),
    ):
        style_header_row(sheet, r, cols)

    # dataset / KPI zebra
    restyle_table_block(sheet, 7, 17, {"A"}, {"B"}, int_cols={"E"}, float_cols=set("FGHIJK"), time_cols={"B"})
    # fix: B values that are text should stay value style; re-apply text rows
    for r in (7, 8, 9, 15, 16, 17):
        row = get_row(sheet, r)
        if row is None:
            continue
        zebra = r % 2 == 0
        for c in row.findall(q("c")):
            if c.get("r").startswith("B") and c.get("t") == "inlineStr":
                c.set("s", str(S["value_g"] if zebra else S["value_w"]))
            if c.get("r").startswith("A"):
                c.set("s", str(S["label_g"] if zebra else S["label_w"]))
            if c.get("r")[0] in "DEFGHIJK":
                letters = c.get("r")[0]
                if letters == "D":
                    c.set("s", str(S["label_g"] if zebra else S["label_w"]))
                elif letters == "E":
                    c.set("s", str(S["kpi_int_g"] if zebra else S["kpi_int_w"]))
                else:
                    c.set("s", str(S["kpi_float_g"] if zebra else S["kpi_float_w"]))

    for r in range(27, 32):
        row = get_row(sheet, r)
        if row is None:
            continue
        zebra = r % 2 == 0
        for c in row.findall(q("c")):
            letters, _ = parse_ref(c.get("r"))
            if letters in ("A", "E"):
                c.set("s", str(S["label_g"] if zebra else S["label_w"]))
            elif letters in ("B", "F"):
                c.set("s", str(S["kpi_int_g"] if zebra else S["kpi_int_w"]))
            elif letters in ("C", "G"):
                c.set("s", str(S["pct_g"] if zebra else S["pct_w"]))

    for r in range(46, 57):
        row = get_row(sheet, r)
        if row is None:
            continue
        zebra = r % 2 == 0
        for c in row.findall(q("c")):
            letters, _ = parse_ref(c.get("r"))
            if letters == "A":
                c.set("s", str(S["label_g"] if zebra else S["label_w"]))
            elif letters == "B":
                if c.get("t") == "inlineStr":
                    c.set("s", str(S["value_g"] if zebra else S["value_w"]))
                else:
                    c.set("s", str(S["kpi_int_g"] if zebra else S["kpi_int_w"]))
            elif letters == "D":
                c.set("s", str(S["label_g"] if zebra else S["label_w"]))
            elif letters == "E":
                c.set("s", str(S["kpi_int_g"] if zebra else S["kpi_int_w"]))
            elif letters in "FGHIJK":
                c.set("s", str(S["kpi_float_g"] if zebra else S["kpi_float_w"]))

    for r in range(76, 90):
        row = get_row(sheet, r)
        if row is None:
            continue
        zebra = r % 2 == 0
        for c in row.findall(q("c")):
            letters, _ = parse_ref(c.get("r"))
            if letters == "A":
                c.set("s", str(S["label_g"] if zebra else S["label_w"]))
            elif letters == "B":
                if c.get("t") == "inlineStr":
                    c.set("s", str(S["value_g"] if zebra else S["value_w"]))
                else:
                    c.set("s", str(S["center_g"] if zebra else S["center_w"]))
            elif letters == "P":
                txt = cell_text(c)
                if r == 81:
                    c.set("s", str(S["total"]))
                else:
                    c.set("s", str(style_solution_badge(txt)))
            elif letters in ("Q", "R"):
                if r == 81:
                    c.set("s", str(S["total"]))
                else:
                    c.set("s", str(S["center_g"] if zebra else S["center_w"]))

    for r in range(113, 118):
        row = get_row(sheet, r)
        if row is None:
            continue
        for c in row.findall(q("c")):
            letters, _ = parse_ref(c.get("r"))
            if letters == "A":
                if r == 117:
                    c.set("s", str(S["total"]))
                else:
                    c.set("s", str(style_solution_badge(cell_text(c))))
            else:
                c.set("s", str(S["total"] if r == 117 else S["center_w"]))

    # detail list 120-244
    for r in range(120, 245):
        row = get_row(sheet, r)
        if row is None:
            continue
        zebra = r % 2 == 0
        for c in row.findall(q("c")):
            letters, _ = parse_ref(c.get("r"))
            txt = cell_text(c)
            if letters == "A":
                c.set("s", str(style_solution_badge(txt)))
            elif letters == "B":
                c.set("s", str(S["strong_c_g"] if zebra else S["strong_c_w"]))
            elif letters == "C":
                c.set("s", str(S["center_g"] if zebra else S["center_w"]))
            elif letters in ("D", "E"):
                c.set("s", str(S["wrap_g"] if zebra else S["wrap_w"]))
            elif letters == "F":
                c.set("s", str(S["right_g"] if zebra else S["right_w"]))
            elif letters == "G":
                c.set("s", str(style_solution_badge(txt)))
            elif letters == "H":
                c.set("s", str(S["detail_g"] if zebra else S["detail_w"]))

    # finish any leftover 900x styles on Cover
    for c in sheet.iter(q("c")):
        st = c.get("s")
        if st == "9001":
            c.set("s", str(S["card_body"]))
        elif st == "9002":
            c.set("s", str(S["param_site"]))
        elif st in ("9003", "9006", "9007"):
            c.set("s", str(S["body"]))
        elif st == "9004":
            c.set("s", str(S["h2"]))
        elif st == "9005":
            c.set("s", str(S["card_cap"]))
        elif st == "9008":
            c.set("s", str(S["center_w"]))
        elif st == "9009":
            c.set("s", str(S["center_g"]))

    # rebuild merges cleanly
    merges = [
        "A1:T1",
        "A2:T2",
        "A3:T3",
        "L5:T5",
        "A42:T42",
        "L44:T44",
        "A72:T72",
        "P74:R74",
        "A110:T110",
    ]
    for r in range(120, 245):
        merges.append(f"H{r}:T{r}")
    # keep E119:H119 as separate header cells — only merge H body rows
    set_merges(sheet, merges)


def transform_kpi_map_sheet(sheet: ET.Element, name: str, unit: str) -> None:
    set_tab_color(sheet, "2563EB")
    titles = {
        "RSRP": "RSRP Coverage Maps",
        "RSRQ": "RSRQ Coverage Maps",
        "SINR": "SINR Coverage Maps",
    }
    set_banner(
        sheet,
        titles[name],
        f"Unit: {unit}  |  Active Mode (Left) vs Idle Mode (Right)",
        "R",
    )
    bands = {
        4: ("All Bands", "All Bands"),
        47: ("L900", "L900"),
        87: ("L1800", "L1800"),
        127: ("L2100", "L2100"),
        167: ("L2600", "L2600"),
    }
    for r, (a, b) in bands.items():
        row = get_row(sheet, r)
        if row is None:
            continue
        row.set("ht", "24")
        row.set("customHeight", "1")
        # left active
        left = None
        right = None
        for c in row.findall(q("c")):
            letters, _ = parse_ref(c.get("r"))
            if letters == "A":
                left = c
            if letters == "J":
                right = c
        if left is not None:
            set_inline(left, f"{name} — Active Mode ({a})", S["section_active"])
        if right is not None:
            set_inline(right, f"{name} — Idle Mode ({b})", S["section_idle"])
    # narrow gutter col I
    cols = sheet.find(q("cols"))
    if cols is not None:
        for col in cols.findall(q("col")):
            if col.get("min") == "9" and col.get("max") == "9":
                col.set("width", "4")
    set_merges(
        sheet,
        [
            "A1:R1",
            "A2:R2",
            "A4:H4",
            "J4:R4",
            "A47:H47",
            "J47:R47",
            "A87:H87",
            "J87:R87",
            "A127:H127",
            "J127:R127",
            "A167:H167",
            "J167:R167",
        ],
    )


def transform_vs_sheet(sheet: ET.Element, mode: str) -> None:
    set_tab_color(sheet, "0D9488")
    if mode == "Idle":
        set_banner(sheet, "RSRP vs KPIs — Idle Mode", "Idle Mode Scanner  |  KPI Correlation & Trend Analysis", "M")
    else:
        set_banner(sheet, "RSRP vs KPIs — Active Mode", "Active Mode (Robi UE)  |  KPI Correlation & Trend Analysis", "M")
    row = get_row(sheet, 48)
    if row is not None:
        row.set("ht", "24")
        row.set("customHeight", "1")
        clear_row_cells(row)
        c1 = ensure_cell(row, "A", S["scatter"])
        set_inline(c1, "RSRP vs SINR Scatter & Trend", S["scatter"])
        c2 = ensure_cell(row, "H", S["scatter"])
        set_inline(c2, "RSRP vs RSRQ Scatter & Trend", S["scatter"])
        for col in list("BCDEF") + list("IJKLM"):
            ensure_cell(row, col, S["scatter"])
    set_merges(sheet, ["A1:M1", "A2:M2", "A48:F48", "H48:M48"])


def transform_sample_log(sheet: ET.Element) -> None:
    set_tab_color(sheet, "64748B")
    set_banner(sheet, "Sample Drive Log — Idle Mode", "Idle Mode  |  8,016 Sample Records", "J")
    style_header_row(sheet, 4, list("ABCDEFGHIJ"))
    sd = sheet.find(q("sheetData"))
    for row in sd.findall(q("row")):
        r = int(row.get("r"))
        if r < 5:
            continue
        zebra = r % 2 == 0
        for c in row.findall(q("c")):
            letters, _ = parse_ref(c.get("r"))
            if letters == "A":
                c.set("s", str(S["log_time_g"] if zebra else S["log_time_w"]))
            elif letters in "BCDEF":
                c.set("s", str(S["log_num_g"] if zebra else S["log_num_w"]))
            else:
                c.set("s", str(S["log_g"] if zebra else S["log_w"]))
    # autoFilter
    old = sheet.find(q("autoFilter"))
    if old is not None:
        sheet.remove(old)
    af = ET.Element(q("autoFilter"))
    af.set("ref", "A4:J8019")
    # insert after sheetData / mergeCells
    merges = sheet.find(q("mergeCells"))
    idx = list(sheet).index(merges) + 1 if merges is not None else list(sheet).index(sheet.find(q("sheetData"))) + 1
    sheet.insert(idx, af)
    set_merges(sheet, ["A1:J1", "A2:J2"])


def transform_ep(sheet: ET.Element) -> None:
    remap_legacy_styles(sheet)
    set_tab_color(sheet, "0F766E")
    # ensure rows 1-2 exist as banner
    r1 = get_row(sheet, 1)
    r2 = get_row(sheet, 2)
    if r1 is not None:
        r1.set("ht", "34")
        r1.set("customHeight", "1")
        clear_row_cells(r1)
        set_inline(ensure_cell(r1, "A", S["title"]), "Engineering Parameter Action Plan", S["title"])
        for i in range(2, 15):
            ensure_cell(r1, col_letters(i), S["title_fill"])
    if r2 is not None:
        r2.set("ht", "22")
        r2.set("customHeight", "1")
        clear_row_cells(r2)
        set_inline(
            ensure_cell(r2, "A", S["subtitle"]),
            "218 Sector Changes  |  Height (m)  |  Azimuth, ET & MT (°)",
            S["subtitle"],
        )
        for i in range(2, 15):
            ensure_cell(r2, col_letters(i), S["subtitle_fill"])
    # header rows 3-4
    r3 = get_row(sheet, 3)
    r4 = get_row(sheet, 4)
    if r3 is not None:
        r3.set("ht", "22")
        r3.set("customHeight", "1")
        for c in r3.findall(q("c")):
            letters, _ = parse_ref(c.get("r"))
            txt = cell_text(c)
            if txt == "Existing Plan":
                c.set("s", str(S["existing_cap"]))
            elif txt == "Proposed Plan":
                c.set("s", str(S["new_cap"]))
            else:
                c.set("s", str(S["th"]))
    if r4 is not None:
        r4.set("ht", "20")
        r4.set("customHeight", "1")
        for c in r4.findall(q("c")):
            letters, _ = parse_ref(c.get("r"))
            if letters in "DEFG":
                c.set("s", str(S["existing_sub"]))
            elif letters in "HIJK":
                c.set("s", str(S["new_sub"]))
            else:
                c.set("s", str(S["th"]))
    for r in range(5, 223):
        row = get_row(sheet, r)
        if row is None:
            continue
        row.set("ht", "20")
        row.set("customHeight", "1")
        zebra = r % 2 == 0
        cells = {parse_ref(c.get("r"))[0]: c for c in row.findall(q("c"))}
        for letters, c in cells.items():
            txt = cell_text(c)
            if letters == "A":
                c.set("s", str(S["param_site"] if not zebra else S["soft_label"]))
            elif letters in ("B", "C"):
                c.set("s", str(S["value_g"] if zebra else S["value_w"]))
            elif letters in "DEFG":
                c.set("s", str(S["param_exist"]))
            elif letters in "HIJK":
                # changed?
                exist_col = col_letters(col_index(letters) - 4)
                exist = cells.get(exist_col)
                exist_txt = cell_text(exist) if exist is not None else ""
                if txt and (not exist_txt or exist_txt != txt):
                    c.set("s", str(S["param_new_chg"]))
                else:
                    c.set("s", str(S["param_new"]))
            elif letters == "L":
                c.set("s", str(style_solution_badge(txt)))
            elif letters == "N":
                c.set("s", str(style_solution_badge(txt)))
            else:
                c.set("s", str(S["center_g"] if zebra else S["center_w"]))
    old = sheet.find(q("autoFilter"))
    if old is not None:
        sheet.remove(old)
    af = ET.Element(q("autoFilter"))
    af.set("ref", "A4:N222")
    merges = sheet.find(q("mergeCells"))
    idx = list(sheet).index(merges) + 1 if merges is not None else list(sheet).index(sheet.find(q("sheetData"))) + 1
    sheet.insert(idx, af)
    for c in sheet.iter(q("c")):
        st = c.get("s")
        if st in ("9001", "9003", "9006", "9007"):
            c.set("s", str(S["value_w"]))
        elif st == "9002":
            c.set("s", str(S["param_site"]))
        elif st == "9008":
            c.set("s", str(S["center_w"]))
        elif st == "9009":
            c.set("s", str(S["center_g"]))
    set_merges(
        sheet,
        [
            "A1:N1",
            "A2:N2",
            "A3:A4",
            "B3:B4",
            "C3:C4",
            "D3:G3",
            "H3:K3",
            "L3:L4",
            "M3:M4",
            "N3:N4",
        ],
    )


def transform_feeder(sheet: ET.Element) -> None:
    remap_legacy_styles(sheet)
    set_tab_color(sheet, "7C3AED")
    # delete rows 3-4 then shift
    delete_rows(sheet, {3, 4})
    shift_sheet_rows(sheet, -2, 5)
    set_banner(
        sheet,
        "Feeder Swap Analysis",
        "Idle Mode  |  6 Confirmed Sites (5 Two-Sector Swaps, 1 Three-Sector Rotation)",
        "F",
    )
    # header now at row 4 (was 6)
    style_header_row(sheet, 4, list("ABCDEF"))
    for r in range(5, 23):
        row = get_row(sheet, r)
        if row is None:
            continue
        row.set("ht", "20")
        row.set("customHeight", "1")
        site_block = (r - 5) // 3
        zebra = site_block % 2 == 1
        end_of_group = (r - 5) % 3 == 2
        for c in row.findall(q("c")):
            letters, _ = parse_ref(c.get("r"))
            txt = cell_text(c)
            if letters == "A":
                c.set("s", str(S["param_site"]))
            elif letters == "F":
                c.set("s", str(style_solution_badge(txt)))
            else:
                c.set("s", str(S["center_g"] if zebra else S["center_w"]))
            if end_of_group:
                # keep style; group separator via border style 12 would need dedicated xf — skip
                pass
    for r in (24, 66, 108, 150, 192, 234):
        row = get_row(sheet, r)
        if row is None:
            continue
        row.set("ht", "24")
        row.set("customHeight", "1")
        for c in row.findall(q("c")):
            set_inline(c, cell_text(c), S["feeder_site"])
            break
        for i in range(2, 7):
            ensure_cell(row, col_letters(i), S["feeder_site"])
    set_merges(
        sheet,
        [
            "A1:F1",
            "A2:F2",
            "A24:F24",
            "A66:F66",
            "A108:F108",
            "A150:F150",
            "A192:F192",
            "A234:F234",
        ],
    )


def _find_desc_start_col(row: ET.Element) -> str | None:
    for c in row.findall(q("c")):
        if cell_text(c) in ("Bad spot description", "Bad Spot Analysis"):
            return parse_ref(c.get("r"))[0]
    return None


def remap_legacy_styles(sheet: ET.Element) -> None:
    """Map pre-v1.91 style indices onto temporary holders before new styles apply.

    Old 37 (body) collides with new callout_b=37, so rewrite leftovers first.
    """
    for c in sheet.iter(q("c")):
        s = c.get("s")
        if s is None:
            continue
        # stash legacy body/label styles in free high ids that we immediately rewrite
        if s == "37":
            c.set("s", "9001")
        elif s == "36":
            c.set("s", "9002")
        elif s == "34":
            c.set("s", "9003")
        elif s == "33":
            c.set("s", "9004")
        elif s == "32":
            c.set("s", "9005")
        elif s == "16":
            c.set("s", "9006")
        elif s == "19":
            c.set("s", "9007")
        elif s == "20":
            c.set("s", "9008")
        elif s == "21":
            c.set("s", "9009")


def transform_bad_spot_sheet(sheet: ET.Element, mode: str, count: int) -> None:
    remap_legacy_styles(sheet)
    set_tab_color(sheet, "B45309" if mode == "Idle" else "DC2626")
    title = f"Bad Spot Analysis — RSRP ({mode} Mode)"
    subtitle = f"{mode} Mode  |  {count} Identified Bad Spots"
    set_banner(sheet, title, subtitle, "V")

    # delete clutter rows 3-7 and shift
    delete_rows(sheet, {3, 4, 5, 6, 7})
    shift_sheet_rows(sheet, -5, 8)

    # combined view now at row 4
    row4 = get_row(sheet, 4)
    if row4 is not None:
        row4.set("ht", "24")
        row4.set("customHeight", "1")
        for c in row4.findall(q("c")):
            if parse_ref(c.get("r"))[0] == "A":
                set_inline(c, f"Combined Coverage Map — All Bands ({mode} Mode)", S["section_neutral"])
            else:
                c.set("s", str(S["section_neutral"]))

    # remove trailing footnote rows
    sd = sheet.find(q("sheetData"))
    max_r = max(int(r.get("r")) for r in sd.findall(q("row")))
    # footnote starts near the end with "The combined view"
    for row in list(sd.findall(q("row"))):
        r = int(row.get("r"))
        texts = [cell_text(c) for c in row.findall(q("c"))]
        if any(t.startswith("The combined view") for t in texts):
            sd.remove(row)
            # also remove following empty/note rows in the merge
            for rr in list(sd.findall(q("row"))):
                if int(rr.get("r")) == r + 1 and not any(cell_text(c) for c in rr.findall(q("c"))):
                    sd.remove(rr)

    # restyle each Bad Spot block
    headers = []
    for row in sd.findall(q("row")):
        for c in row.findall(q("c")):
            if re.fullmatch(r"Bad Spot \d+", cell_text(c)):
                headers.append(int(row.get("r")))
                break

    for r in headers:
        row = get_row(sheet, r)
        if row is None:
            continue
        row.set("ht", "24")
        row.set("customHeight", "1")
        for c in row.findall(q("c")):
            set_inline(c, cell_text(c), S["section_neutral"]) if parse_ref(c.get("r"))[0] == "A" else c.set(
                "s", str(S["section_neutral"])
            )

    # description cards + parameter tables + idle callouts
    for row in sd.findall(q("row")):
        r = int(row.get("r"))
        for c in list(row.findall(q("c"))):
            txt = cell_text(c)
            letters, _ = parse_ref(c.get("r"))
            if txt in ("Bad spot description", "Bad Spot Analysis"):
                set_inline(c, "Bad Spot Analysis", S["card_cap"])
            elif txt in (
                "Serving Cells details",
                "Serving Cells",
                "Bad Spot Details",
                "Root Cause Analysis",
                "Suggestions",
            ):
                nice = {
                    "Serving Cells details": "Serving Cells",
                    "Serving Cells": "Serving Cells",
                    "Bad Spot Details": "Bad Spot Details",
                    "Root Cause Analysis": "Root Cause Analysis",
                    "Suggestions": "Suggestions",
                }[txt]
                set_inline(c, nice, S["card_label"])
            elif txt == "Parameter change":
                set_inline(c, "Parameter Change", S["param_cap"])
            elif txt == "Existing":
                set_inline(c, "Existing Plan", S["existing_cap"])
            elif txt == "New Plan":
                set_inline(c, "Proposed Plan", S["new_cap"])
            elif txt in ("Idle vs active", "Idle vs Active Comparison"):
                set_inline(c, "Idle vs Active Comparison", S["callout_h"])
            elif txt == "Identified RSRP bad spots":
                set_inline(c, "Identified RSRP Bad Spots", S["h2"])

    # walk parameter tables by locating Parameter Change rows
    for row in sd.findall(q("row")):
        r = int(row.get("r"))
        start_col = None
        for c in row.findall(q("c")):
            if cell_text(c) == "Parameter Change":
                start_col = parse_ref(c.get("r"))[0]
                break
        if start_col is None:
            continue
        c0 = col_index(start_col)
        # style group headers on this row
        for c in row.findall(q("c")):
            letters, _ = parse_ref(c.get("r"))
            ci = col_index(letters)
            if c0 <= ci <= c0 + 2:
                c.set("s", str(S["param_cap"]))
            elif c0 + 3 <= ci <= c0 + 6:
                c.set("s", str(S["existing_cap"]))
            elif c0 + 7 <= ci <= c0 + 10:
                c.set("s", str(S["new_cap"]))
        # subheader
        sub = get_row(sheet, r + 1)
        if sub is not None:
            for c in sub.findall(q("c")):
                letters, _ = parse_ref(c.get("r"))
                ci = col_index(letters)
                if c0 <= ci <= c0 + 2:
                    c.set("s", str(S["param_sub"]))
                elif c0 + 3 <= ci <= c0 + 6:
                    c.set("s", str(S["existing_sub"]))
                elif c0 + 7 <= ci <= c0 + 10:
                    c.set("s", str(S["new_sub"]))
        # data rows until blank / next section
        for dr in range(r + 2, r + 6):
            drow = get_row(sheet, dr)
            if drow is None:
                break
            texts = [cell_text(c) for c in drow.findall(q("c"))]
            if not any(texts):
                break
            if any(t.startswith("Idle vs") or t == "Bad Spot Analysis" or re.fullmatch(r"Bad Spot \d+", t or "") for t in texts):
                break
            cells = {parse_ref(c.get("r"))[0]: c for c in drow.findall(q("c"))}
            if len(cells) == 1 and "No azimuth" in (texts[0] or ""):
                for c in drow.findall(q("c")):
                    c.set("s", str(S["param_note"]))
                continue
            for letters, c in cells.items():
                ci = col_index(letters)
                if c0 <= ci <= c0 + 2:
                    c.set("s", str(S["param_site"] if ci == c0 else S["value_w"]))
                elif c0 + 3 <= ci <= c0 + 6:
                    c.set("s", str(S["param_exist"]))
                elif c0 + 7 <= ci <= c0 + 10:
                    exist = cells.get(col_letters(ci - 4))
                    exist_txt = cell_text(exist) if exist is not None else ""
                    txt = cell_text(c)
                    if txt and (not exist_txt or exist_txt != txt):
                        c.set("s", str(S["param_new_chg"]))
                    else:
                        c.set("s", str(S["param_new"]))

    # finish leftover legacy styles (900x) and prose bodies
    for row in sd.findall(q("row")):
        for c in row.findall(q("c")):
            st = c.get("s")
            txt = cell_text(c)
            if st == "9001":
                c.set("s", str(S["callout_b"] if ("Idle vs" in txt or "median" in txt.lower()) else S["card_body"]))
            elif st == "9002":
                if txt in ("Serving Cells", "Bad Spot Details", "Root Cause Analysis", "Suggestions"):
                    c.set("s", str(S["card_label"]))
                elif txt in ("Site", "Cell", "Sector", "Height", "Azimuth", "ET", "MT"):
                    c.set("s", str(S["param_sub"]))
                else:
                    c.set("s", str(S["param_site"]))
            elif st == "9003":
                c.set("s", str(S["callout_b"] if ("Idle vs" in txt or "median" in txt.lower()) else S["card_body"]))
            elif st == "9004":
                c.set("s", str(S["callout_h"] if "Idle" in txt else S["h2"]))
            elif st == "9005":
                c.set("s", str(S["card_cap"]))
            elif st == "9006":
                c.set("s", str(S["body"]))
            elif st == "9007":
                c.set("s", str(S["note"]))
            elif st == "9008":
                c.set("s", str(S["center_w"]))
            elif st == "9009":
                c.set("s", str(S["center_g"]))
            elif txt and len(txt) > 80 and st not in {
                str(S["card_cap"]),
                str(S["card_label"]),
                str(S["card_body"]),
                str(S["callout_h"]),
                str(S["callout_b"]),
                str(S["param_note"]),
                str(S["detail_w"]),
                str(S["detail_g"]),
            }:
                if txt.startswith("Idle vs"):
                    c.set("s", str(S["callout_b"]))
                else:
                    c.set("s", str(S["card_body"]))

    # bottom summary table — close I..O gap by shifting P..X -> I..Q
    summary_header = None
    for row in sd.findall(q("row")):
        for c in row.findall(q("c")):
            if cell_text(c) == "Circle" and c.get("r").startswith("A"):
                # confirm this is summary header by Mean KPI nearby
                texts = {parse_ref(cc.get("r"))[0]: cell_text(cc) for cc in row.findall(q("c"))}
                if texts.get("H") == "Mean KPI" and texts.get("P") == "Mean RSRP":
                    summary_header = int(row.get("r"))
                    break
        if summary_header:
            break
    if summary_header:
        mapping = {
            "P": "I",
            "Q": "J",
            "R": "K",
            "S": "L",
            "T": "M",
            "U": "N",
            "V": "O",
            "W": "P",
            "X": "Q",
        }
        for r in range(summary_header, summary_header + 200):
            row = get_row(sheet, r)
            if row is None:
                continue
            # stop if we left the table
            a = None
            for c in row.findall(q("c")):
                if c.get("r") == f"A{r}":
                    a = cell_text(c)
                    break
            if r > summary_header and (a is None or a == ""):
                # might be title row before table
                if r == summary_header:
                    continue
                # allow empty? break if past first blank after data started
                if r > summary_header + 1 and not any(cell_text(c) for c in row.findall(q("c"))):
                    break
            # rebuild cells with shifted columns
            cells = list(row.findall(q("c")))
            new_cells = []
            for c in cells:
                letters, rr = parse_ref(c.get("r"))
                if letters in mapping:
                    c.set("r", f"{mapping[letters]}{rr}")
                new_cells.append(c)
            # sort
            for c in cells:
                row.remove(c)
            for c in sorted(new_cells, key=lambda x: col_index(parse_ref(x.get("r"))[0])):
                row.append(c)
            zebra = r % 2 == 0
            if r == summary_header:
                row.set("ht", "24")
                row.set("customHeight", "1")
                for c in row.findall(q("c")):
                    c.set("s", str(S["th"]))
            else:
                for c in row.findall(q("c")):
                    letters, _ = parse_ref(c.get("r"))
                    if letters == "A":
                        c.set("s", str(S["strong_c_g"] if zebra else S["strong_c_w"]))
                    elif letters == "Q":
                        c.set("s", str(S["detail_g"] if zebra else S["detail_w"]))
                    elif letters in ("C", "D"):
                        c.set("s", str(S["kpi_5dp_g"] if zebra else S["kpi_5dp_w"]))
                    elif letters in ("E", "F", "L"):
                        c.set("s", str(S["kpi_int_g"] if zebra else S["kpi_int_w"]))
                    elif letters == "G":
                        c.set("s", str(S["pct_g"] if zebra else S["pct_w"]))
                    elif letters in ("H", "I", "J", "K", "M"):
                        c.set("s", str(S["kpi_float_g"] if zebra else S["kpi_float_w"]))
                    else:
                        c.set("s", str(S["wrap_g"] if zebra else S["wrap_w"]))
        # rebuild merges: preferred banners first, then body merges, then summary Q:V
        # v1.91/v1.92 left A1:W1 + A1:V1 (and A4:W7 / A4:W4) overlapping — Excel rejects that.
        kept: list[str] = ["A1:V1", "A2:V2", "A4:V4"]
        header_set = set(headers)
        for m in collect_merges(sheet):
            nums = [int(x) for x in re.findall(r"\d+", m)]
            if not nums:
                continue
            if min(nums) >= summary_header:
                continue
            # drop old title / methodology / pre-shift combined-header merges in rows 1–4
            if m.startswith("A") and min(nums) <= 4:
                continue
            if m in ("A1:W1", "A2:W2", "A3:W3", "A4:W7", "A9:W9", "A4:W4", "A1:V1", "A2:V2", "A4:V4"):
                continue
            if re.match(r"A\d+:[A-Z]+\d+$", m):
                top = int(re.findall(r"\d+", m)[0])
                if top in header_set:
                    continue
            kept.append(m)
        for r in headers:
            kept.append(f"A{r}:Q{r}")
        for r in range(summary_header + 1, summary_header + 1 + count):
            kept.append(f"Q{r}:V{r}")
        seen: set[str] = set()
        uniq: list[str] = []
        for m in kept:
            if m not in seen:
                seen.add(m)
                uniq.append(m)
        set_merges(sheet, uniq)
        sanitize_merges(sheet)

    # normalize Q width
    cols = sheet.find(q("cols"))
    if cols is not None:
        for col in cols.findall(q("col")):
            mn = int(col.get("min"))
            mx = int(col.get("max"))
            if mn <= 17 <= mx:
                # split if needed — simplify: set width 13 for cols 17-22
                pass
        # ensure cols 17-22 exist with width 13
        for cidx in range(17, 23):
            found = False
            for col in cols.findall(q("col")):
                if col.get("min") == str(cidx) and col.get("max") == str(cidx):
                    col.set("width", "13")
                    found = True
                    break
            if not found:
                el = ET.SubElement(cols, q("col"))
                el.set("min", str(cidx))
                el.set("max", str(cidx))
                el.set("width", "13")
                el.set("customWidth", "1")

    sanitize_merges(sheet)


def transform_drawing_vs(drawing: bytes) -> bytes:
    # shift Chart 2 from col 8 to 7; widen Chart 3
    text = drawing.decode("utf-8")
    # careful: only the second chart (RSRP vs RSRQ) is at col 8 row 3
    text = text.replace(
        "<from><col>8</col><colOff>0</colOff><row>3</row>",
        "<from><col>7</col><colOff>0</colOff><row>3</row>",
        1,
    )
    # widen dual chart at row 21
    text = re.sub(
        r"(<from><col>0</col><colOff>0</colOff><row>21</row><rowOff>0</rowOff></from><ext cx=\")\d+(\")",
        r"\g<1>11290000\2",
        text,
        count=1,
    )
    return text.encode("utf-8")


def transform_drawing_cover(drawing: bytes) -> bytes:
    text = drawing.decode("utf-8")
    # move charts at col 17 to col 16
    text = text.replace("<from><col>17</col>", "<from><col>16</col>")
    return text.encode("utf-8")


def sheet_xml_bytes(sheet: ET.Element) -> bytes:
    # Match Excel-friendly declaration (double quotes, standalone).
    body = ET.tostring(sheet, encoding="utf-8")
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body


def publish(src: Path, dest: Path) -> Path:
    import zipfile

    with zipfile.ZipFile(src) as zf:
        names = zf.namelist()
        parts = {n: zf.read(n) for n in names}

    # rebuild styles first so S map is complete
    parts["xl/styles.xml"] = build_styles()

    # parse sheets
    def load(path: str) -> ET.Element:
        return ET.fromstring(parts[path])

    def store(path: str, root: ET.Element) -> None:
        parts[path] = sheet_xml_bytes(root)

    cover = load("xl/worksheets/sheet1.xml")
    transform_cover(cover)
    store("xl/worksheets/sheet1.xml", cover)

    for path, name, unit in (
        ("xl/worksheets/sheet2.xml", "RSRP", "dBm"),
        ("xl/worksheets/sheet3.xml", "RSRQ", "dB"),
        ("xl/worksheets/sheet4.xml", "SINR", "dB"),
    ):
        sh = load(path)
        transform_kpi_map_sheet(sh, name, unit)
        store(path, sh)

    idle = load("xl/worksheets/sheet5.xml")
    transform_bad_spot_sheet(idle, "Idle", 26)
    store("xl/worksheets/sheet5.xml", idle)
    parts["xl/drawings/drawing5.xml"] = shift_drawing_rows(parts["xl/drawings/drawing5.xml"], -5, 0)

    active = load("xl/worksheets/sheet8.xml")
    transform_bad_spot_sheet(active, "Active", 99)
    store("xl/worksheets/sheet8.xml", active)
    parts["xl/drawings/drawing8.xml"] = shift_drawing_rows(parts["xl/drawings/drawing8.xml"], -5, 0)

    for path, mode in (
        ("xl/worksheets/sheet11.xml", "Idle"),
        ("xl/worksheets/sheet12.xml", "Active"),
    ):
        sh = load(path)
        transform_vs_sheet(sh, mode)
        store(path, sh)
    parts["xl/drawings/drawing11.xml"] = transform_drawing_vs(parts["xl/drawings/drawing11.xml"])
    parts["xl/drawings/drawing12.xml"] = transform_drawing_vs(parts["xl/drawings/drawing12.xml"])

    sample = load("xl/worksheets/sheet13.xml")
    transform_sample_log(sample)
    store("xl/worksheets/sheet13.xml", sample)

    ep = load("xl/worksheets/sheet15.xml")
    transform_ep(ep)
    store("xl/worksheets/sheet15.xml", ep)

    feeder = load("xl/worksheets/sheet16.xml")
    transform_feeder(feeder)
    store("xl/worksheets/sheet16.xml", feeder)
    parts["xl/drawings/drawing16.xml"] = shift_drawing_rows(parts["xl/drawings/drawing16.xml"], -2, 0)

    parts["xl/drawings/drawing1.xml"] = transform_drawing_cover(parts["xl/drawings/drawing1.xml"])
    polish_charts(parts)

    tmp = dest.with_suffix(".xlsx.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            zf.writestr(name, parts[name])
    tmp.replace(dest)
    print(f"Wrote {dest} ({dest.stat().st_size / 1e6:.1f} MB)", flush=True)
    return dest


def verify(path: Path) -> list[str]:
    import zipfile
    from xml.etree import ElementTree as XET

    problems = []
    with zipfile.ZipFile(path) as zf:
        # every XML part must parse (Excel rejects the whole file on one bad chart)
        for name in zf.namelist():
            if not name.endswith(".xml") and not name.endswith(".rels"):
                continue
            raw = zf.read(name)
            try:
                XET.fromstring(raw)
            except XET.ParseError as exc:
                problems.append(f"XML parse {name}: {exc}")
        # clutter tokens
        for name in zf.namelist():
            if not name.startswith("xl/worksheets/sheet") or not name.endswith(".xml"):
                continue
            if name.endswith("sheet14.xml"):
                continue
            text = zf.read(name).decode("utf-8", errors="ignore")
            for token in (
                "Gridlines off",
                "Freeze panes off",
                "How to read this workbook",
                "No numbering",
                "no circles",
                "Combined view has no circles",
            ):
                if token in text:
                    problems.append(f"{name}: still has {token!r}")
            if "F4D03F" in text:
                problems.append(f"{name}: still has yellow fill hex in sheet?")
        styles = zf.read("xl/styles.xml").decode("utf-8")
        if "F4D03F" in styles:
            problems.append("styles still contain #F4D03F")
        # drawing counts
        counts = {
            "xl/drawings/drawing1.xml": 8,
            "xl/drawings/drawing2.xml": 10,
            "xl/drawings/drawing3.xml": 10,
            "xl/drawings/drawing4.xml": 10,
            "xl/drawings/drawing5.xml": 53,
            "xl/drawings/drawing8.xml": 298,
            "xl/drawings/drawing11.xml": 5,
            "xl/drawings/drawing12.xml": 5,
            "xl/drawings/drawing16.xml": 6,
        }
        for name, n in counts.items():
            xml = zf.read(name).decode("utf-8")
            got = xml.count("<oneCellAnchor>")
            if got != n:
                problems.append(f"{name}: anchors {got}, expected {n}")
        # active-idle gap invariant on drawing8
        d8 = zf.read("xl/drawings/drawing8.xml").decode("utf-8")
        root = ET.fromstring(zf.read("xl/drawings/drawing8.xml"))
        pics = []
        for a in root:
            fr = a.find(f"{{{NS_DRAW}}}from")
            if fr is None:
                continue
            name_el = a.find(f".//{{{NS_DRAW}}}cNvPr")
            nm = name_el.get("name") if name_el is not None else ""
            if "<pic>" in ET.tostring(a, encoding="unicode") or a.find(f"{{{NS_DRAW}}}pic") is not None:
                row = int(fr.find(f"{{{NS_DRAW}}}row").text)
                row_off = int(fr.find(f"{{{NS_DRAW}}}rowOff").text)
                pics.append((nm, row, row_off))
        active_pics = [(n, r, o) for n, r, o in pics if n.startswith("Image")]
        idle_pics = [(n, r, o) for n, r, o in pics if n.startswith("Idle")]
        if len(idle_pics) != 99:
            problems.append(f"idle pics {len(idle_pics)}, expected 99")
        active_sorted = sorted(active_pics, key=lambda x: x[1])
        idle_sorted = sorted(idle_pics, key=lambda x: x[1])
        serial_actives = active_sorted[1:]
        if len(serial_actives) != 99:
            problems.append(f"active serial pics {len(serial_actives)}, expected 99")
        else:
            bad = 0
            for (an, ar, ao), (inn, ir, io) in zip(serial_actives, idle_sorted):
                if ir - ar != 27 or io != 22860:
                    bad += 1
            if bad:
                problems.append(f"active-idle gap mismatch on {bad} spots")
        ellipses = d8.count('prst="ellipse"')
        if ellipses != 99:
            problems.append(f"drawing8 ellipses {ellipses}, expected 99")
        d5 = zf.read("xl/drawings/drawing5.xml").decode("utf-8")
        if d5.count('prst="ellipse"') != 26:
            problems.append(f"drawing5 ellipses {d5.count('prst=\"ellipse\"')}, expected 26")
        cover = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        merges = [m.get("ref") for m in cover.findall(f".//{{{NS_MAIN}}}mergeCell")]
        if len(merges) != len(set(merges)):
            problems.append(f"cover duplicate merges {len(merges)-len(set(merges))}")
        if len(merges) > 400:
            problems.append(f"cover merges still high: {len(merges)}")
        # overlapping merges make Excel report the file as corrupt
        for name in zf.namelist():
            if not name.startswith("xl/worksheets/sheet") or not name.endswith(".xml"):
                continue
            root = ET.fromstring(zf.read(name))
            refs = [m.get("ref") for m in root.findall(f".//{{{NS_MAIN}}}mergeCell")]
            bounds = []
            for ref in refs:
                b = parse_merge_bounds(ref)
                if b is None:
                    problems.append(f"{name}: invalid merge {ref!r}")
                    continue
                if any(merges_overlap(b, kb) for _, kb in bounds):
                    problems.append(f"{name}: overlapping merge {ref}")
                    break
                bounds.append((ref, b))
        # chart title runs must not nest <a:r>
        for name in zf.namelist():
            if not name.startswith("xl/charts/chart") or not name.endswith(".xml"):
                continue
            text = zf.read(name).decode("utf-8")
            if "</a:rPr><a:r><a:t>" in text:
                problems.append(f"{name}: nested <a:r> title corruption")
            if re.search(r"</val>\s*<dPt>", text):
                problems.append(f"{name}: dPt after val (invalid ser order)")
        styles = zf.read("xl/styles.xml").decode("utf-8")
        if "<colors>" not in styles:
            problems.append("styles missing <colors>")
    return problems


if __name__ == "__main__":
    src = Path("output/Bogura_DriveTest_Report_v1.90.xlsx")
    dest = Path("output/Bogura_DriveTest_Report_v1.93.xlsx")
    publish(src, dest)
    probs = verify(dest)
    print("problems:", probs)
    if probs:
        raise SystemExit(1)
    shutil.copy2(dest, Path("output/Bogura_DriveTest_Report.xlsx"))
    print("Copied to Bogura_DriveTest_Report.xlsx")
