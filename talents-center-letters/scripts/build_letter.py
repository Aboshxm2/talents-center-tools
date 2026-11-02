#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build an official letter of مركز الموهبة والإبداع والتميز (Taibah University)
on the official letterhead (الكليشة) from a small JSON file.

Everything that makes the letter "official" — letterhead image, header, footer,
fonts, colours, spacing, the greeting/closing styles and the signature block —
comes from assets/template.docx. This script only writes the letter's words
into it, so the boilerplate is never re-typed or drifted by hand.

Standard library only (no python-docx needed).

    python3 build_letter.py letter.json -o "خطاب.docx"
    python3 build_letter.py --example        # print a complete sample letter.json

letter.json fields (only "recipient" and "body" are required):
    recipient    addressee line, e.g. "سعادة عميد شؤون الطلاب" ("\\n" = extra line)
    honorific    word printed at the far left of the addressee line: الموقر / المحترم / حفظه الله ...
    greeting     default "السلام عليكم ورحمة الله وبركاته"; "" to omit
    body         list of paragraphs; an item may also be {"bullets": [...]} or {"numbered": [...]}
                 inline: **bold**   [[missing info]] -> yellow-highlighted placeholder
    closing      default from config; e.g. "وتفضلوا بقبول خالص التحية والتقدير،،،"
    signature    {"title": "...", "name": "..."} to override the config signature
    date         "today" (Umm al-Qura Hijri, Riyadh time), an explicit date, or "" (blank)
    number       letter number (usually left blank for the registry)
    attachments  text for المشفوعات, e.g. "١" or "نموذج الترشيح"
    font_size    body text size in pt for this letter (default: config body_font_size, 12)
    layout       auto (default) | spacious | standard | compact

Numeric dates anywhere in the letter (22/5/1448, 1448/5/22, 2026-11-02) are stored day-first
in Arabic-Indic digits with right-to-left marks, which is what makes Word show them the right
way round (day on the right, year next to هـ).
"""

import argparse
import datetime
import json
import os
import re
import sys
import zipfile
from xml.dom import minidom
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
DEFAULT_TEMPLATE = os.path.join(SKILL_DIR, "assets", "template.docx")
DEFAULT_CONFIG = os.path.join(SKILL_DIR, "assets", "config.json")

# --------------------------------------------------------------------------
# Formatting taken from the template (كليشة الخطابات). Keep in sync with it.
# --------------------------------------------------------------------------
GREEN = "0A8D6E"  # addressee, greeting, closing, signature
NAVY = "090D48"   # body text
BLUE = "4056E3"   # unit lines under the letterhead

STYLES = {
    #               latin font             arabic (cs) font   colour  half-pts bold
    "body":        dict(ascii="Alexandria", cs="Tosh A", color=NAVY, size=None, bold=False),
    "recipient":   dict(ascii="Tosh A Bold", cs="Tosh A", color=GREEN, size=24, bold=True),
    "greeting":    dict(ascii="Tosh A b", cs="Tosh A Medium", color=GREEN, size=24, bold=False),
    "closing":     dict(ascii="Tosh A b", cs="Tosh A Medium", color=GREEN, size=28, bold=False,
                        eastAsia="Times New Roman"),
    "sig_title":   dict(ascii="Tosh A b", cs="Tosh A Medium", color=GREEN, size=28, bold=False,
                        eastAsia="Times New Roman"),
    "sig_name":    dict(ascii="Tosh A b", cs="Tosh A", color=GREEN, size=28, bold=False,
                        eastAsia="Times New Roman"),
    "field":       dict(ascii="Alexandria", cs="Tosh A Medium", color=NAVY, size=22, bold=False),
    "header_unit": dict(ascii="Alexandria SemiBold", cs="Tosh A Medium", color=BLUE, size=16,
                        bold=False),
}

# Body text size in points (config "body_font_size"). The template's 11pt looked small in
# practice; 12pt Tosh A matches the visual size of the center's usual 16pt Sakkal Majalla
# letters (measured on a signed letter by alif height).
DEFAULT_BODY_PT = 12


def apply_font_sizes(body_pt):
    """Body size from config; headings stay a step above it, as in the template."""
    hp = lambda pt: int(round(pt * 2))  # noqa: E731  (Word sizes are half-points)
    STYLES["body"]["size"] = hp(body_pt)
    STYLES["recipient"]["size"] = hp(body_pt + 1)
    STYLES["greeting"]["size"] = hp(body_pt + 1)
    for key in ("closing", "sig_title", "sig_name"):
        STYLES[key]["size"] = max(28, hp(body_pt + 2))  # template: 14pt

TEXT_W = 8306          # twips between the template's page margins (11906 - 2*1800)
BODY_IND = -907        # template body paragraphs reach 907 twips into both margins
ADDR_IND_START = -850  # template addressee paragraph indents
ADDR_IND_END = -794
TOP_MARGIN_MM = 25.4   # template top margin (1440 twips)

# Where the letterhead image prints its "الرقم / التاريخ / المشفوعات" labels, in EMU
# from the page's top-left corner (measured on the letterhead image as it is anchored
# in the template header). x_right = just left of the printed label.
FIELD_BOXES = {
    "number":      dict(y_center=705_500, x_left=600_000, x_right=1_715_000, cover_dots=True),
    "date":        dict(y_center=1_004_600, x_left=600_000, x_right=1_640_000, cover_dots=False),
    "attachments": dict(y_center=1_282_400, x_left=600_000, x_right=1_345_000, cover_dots=True),
}
FIELD_BOX_H = 330_000

LAYOUTS = {
    # top: mm from page top where the addressee starts. Spacing in twips.
    "spacious": dict(top=50.0, before=240, after=100, auto_after=True, line=360,
                     list_after=60, list_line=360, closing_before=480, sig_spacers=3),
    "standard": dict(top=50.0, before=0, after=160, auto_after=False, line=360,
                     list_after=80, list_line=340, closing_before=280, sig_spacers=1),
    "compact":  dict(top=47.0, before=0, after=100, auto_after=False, line=300,
                     list_after=40, list_line=288, closing_before=160, sig_spacers=0),
}

# --------------------------------------------------------------------------
# Umm al-Qura calendar, 1440/1/1 .. 1470/12/30 AH (≈ Sep 2018 .. Apr 2049).
# Month lengths (1 = 30 days, 0 = 29 days) from the official Umm al-Qura tables
# (data via the MIT-licensed hijri-converter package).
# --------------------------------------------------------------------------
UQ_FIRST_YEAR = 1440
UQ_BASE_RJD = 58373  # Reduced Julian Day of 1 Muharram 1440 (= 11 Sep 2018)
UQ_MONTH_BITS = (
    "010111010100101011011010010101011010101010101011010110010101011101001001"
    "011101100100101110101010010110110101001010110110101001010110110100101010"
    "111010010101011100101010011101010101001101011010100101011101010010011011"
    "101001001101110100100110110101010011010110101010101010101101010010110110"
    "101001010111010100100111101010010101101101001010101101010101001101101100"
    "100110101110"
)


def gregorian_to_hijri(d):
    """Return (year, month, day) in the Umm al-Qura calendar, or None if out of range."""
    rjd = d.toordinal() + 1721425 - 2400000
    start = UQ_BASE_RJD
    if rjd < start:
        return None
    for i, bit in enumerate(UQ_MONTH_BITS):
        length = 30 if bit == "1" else 29
        if rjd < start + length:
            return UQ_FIRST_YEAR + i // 12, i % 12 + 1, rjd - start + 1
        start += length
    return None


AR_DIGITS = str.maketrans("0123456789۰۱۲۳۴۵۶۷۸۹", "٠١٢٣٤٥٦٧٨٩٠١٢٣٤٥٦٧٨٩")
RLM = "‏"  # invisible right-to-left mark

# Numeric dates. Word lays out "/"-separated numbers inside Arabic text right-to-left, piece by
# piece, so a date must be STORED day-first (٢٢/٥/١٤٤٨) to show the day on the right and the
# year next to هـ — the way a hand-typed date looks. Storing it year-first shows it reversed.
# A right-to-left mark after each separator makes non-Word viewers (PDF, web) agree with Word.
_SEP_GAP = r"[‎‏]?\s*"
_DATE_DMY = re.compile(r"(?<![\d/])(\d{1,2})\s*([/\-.])" + _SEP_GAP + r"(\d{1,2})\s*\2" + _SEP_GAP
                       + r"(\d{4})(?![\d/])")
_DATE_YMD = re.compile(r"(?<![\d/])(\d{4})\s*([/\-.])" + _SEP_GAP + r"(\d{1,2})\s*\2" + _SEP_GAP
                       + r"(\d{1,2})(?![\d/])")


def _plausible_date(d, m, y):
    d, m, y = int(d), int(m), int(y)
    return 1 <= d <= 31 and 1 <= m <= 12 and (1300 <= y <= 1600 or 1900 <= y <= 2100)


def rtl_date(d, m, y, sep="/"):
    """Day-first, Arabic-Indic digits, a right-to-left mark after each separator."""
    return f"{d}{sep}{RLM}{m}{sep}{RLM}{y}".translate(AR_DIGITS)


def normalize_dates(text):
    """Rewrite numeric dates (22/5/1448, 1448/5/22, 2026-11-02 …) so Word shows them correctly."""
    def ymd(mt):
        y, sep, m, d = mt.groups()
        return rtl_date(d, m, y, sep) if _plausible_date(d, m, y) else mt.group(0)

    def dmy(mt):
        d, sep, m, y = mt.groups()
        return rtl_date(d, m, y, sep) if _plausible_date(d, m, y) else mt.group(0)

    return _DATE_DMY.sub(dmy, _DATE_YMD.sub(ymd, text))


def format_date(value):
    """'today' -> today's Umm al-Qura date (reads ٢٢/٥/١٤٤٨هـ); other dates normalised the same way."""
    value = (value or "").strip()
    if not value:
        return "", None
    if value.lower() in ("today", "اليوم"):
        riyadh_now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=3)))
        h = gregorian_to_hijri(riyadh_now.date())
        if h is None:
            return "", "today's date is outside the built-in Hijri table; date left blank"
        y, m, dd = h
        return rtl_date(dd, m, y) + "هـ", None
    for pattern, order in ((r"(\d{4})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{1,2})", "ymd"),
                           (r"(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{4})", "dmy")):
        m = re.fullmatch(pattern + r"\s*(هـ|م)?", value)
        if m:
            a, b, c, suffix = m.groups()
            y, mo, dd = (a, b, c) if order == "ymd" else (c, b, a)
            if not suffix:
                suffix = "هـ" if 1300 <= int(y) <= 1600 else "م"
            return rtl_date(dd, mo, y) + suffix, None
    return normalize_dates(value), None


# --------------------------------------------------------------------------
# Run building: split text into Arabic (rtl) and Latin/number (ltr) runs and
# handle the tiny inline markup (**bold**, [[placeholder]]).
# --------------------------------------------------------------------------
_ARABIC = ((0x0600, 0x06FF), (0x0750, 0x077F), (0x08A0, 0x08FF), (0xFB50, 0xFDFF), (0xFE70, 0xFEFF))


def _strong(ch):
    o = ord(ch)
    for a, b in _ARABIC:
        if a <= o <= b:
            return "R"
    if ch.isalnum():  # Latin letters, Western digits, other scripts
        return "L"
    return None  # neutral: spaces, punctuation, symbols


def split_bidi(text):
    """[(segment, is_rtl)] — neutrals go Latin only when both neighbours are Latin."""
    if not text:
        return []
    kinds = [_strong(c) for c in text]
    resolved = list(kinds)
    i, n = 0, len(text)
    while i < n:
        if kinds[i] is None:
            j = i
            while j < n and kinds[j] is None:
                j += 1
            before = next((kinds[k] for k in range(i - 1, -1, -1) if kinds[k]), "R")
            after = kinds[j] if j < n else "R"
            fill = "L" if before == "L" and after == "L" else "R"
            for k in range(i, j):
                resolved[k] = fill
            i = j
        else:
            i += 1
    out, cur, cur_kind = [], [], resolved[0]
    for ch, kind in zip(text, resolved):
        if kind != cur_kind:
            out.append(("".join(cur), cur_kind == "R"))
            cur, cur_kind = [], kind
        cur.append(ch)
    out.append(("".join(cur), cur_kind == "R"))
    return out


_MARKUP = re.compile(r"(\*\*.+?\*\*|\[\[.+?\]\])", re.S)


def parse_inline(text):
    """[(text, bold, placeholder)] from a string with **bold** / [[placeholder]]."""
    parts = []
    for piece in _MARKUP.split(text):
        if not piece:
            continue
        if piece.startswith("**") and piece.endswith("**") and len(piece) > 4:
            parts.append((piece[2:-2], True, False))
        elif piece.startswith("[[") and piece.endswith("]]") and len(piece) > 4:
            parts.append(("[" + piece[2:-2].strip() + "]", False, True))
        else:
            parts.append((piece, False, False))
    return parts


def rpr_xml(style, bold=False, highlight=None, rtl=True):
    s = STYLES[style]
    fonts = f'w:ascii="{s["ascii"]}" w:hAnsi="{s["ascii"]}" w:cs="{s["cs"]}"'
    if s.get("eastAsia"):
        fonts = (f'w:ascii="{s["ascii"]}" w:eastAsia="{s["eastAsia"]}" '
                 f'w:hAnsi="{s["ascii"]}" w:cs="{s["cs"]}"')
    parts = [f"<w:rFonts {fonts}/>"]
    if bold or s["bold"]:
        parts.append("<w:b/><w:bCs/>")
    parts.append(f'<w:color w:val="{s["color"]}"/>')
    if s["size"]:
        parts.append(f'<w:sz w:val="{s["size"]}"/><w:szCs w:val="{s["size"]}"/>')
    if highlight:
        parts.append(f'<w:highlight w:val="{highlight}"/>')
    if rtl:
        parts.append("<w:rtl/>")
    return "<w:rPr>" + "".join(parts) + "</w:rPr>"


def runs_xml(text, style, placeholders=None):
    out = []
    for piece, bold, is_ph in parse_inline(normalize_dates(text)):
        if is_ph and placeholders is not None:
            placeholders.append(piece)
        for seg, rtl in split_bidi(piece):
            out.append("<w:r>" + rpr_xml(style, bold=bold, highlight="yellow" if is_ph else None,
                                         rtl=rtl)
                       + f'<w:t xml:space="preserve">{escape(seg)}</w:t></w:r>')
    return "".join(out)


def para_xml(inner_runs, style, ppr_extra="", jc=None, ltr=False, keep_next=False,
             num_id=None, p_style=None):
    """A paragraph whose pPr children follow the schema order."""
    ppr = []
    if p_style:
        ppr.append(f'<w:pStyle w:val="{p_style}"/>')
    if keep_next:
        ppr.append("<w:keepNext/>")
    if num_id is not None:
        ppr.append(f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>')
    # explicit direction on every paragraph (also inside text boxes), not just via the style
    ppr.append('<w:bidi w:val="0"/>' if ltr else "<w:bidi/>")
    ppr.append(ppr_extra)  # spacing / ind / contextualSpacing, already in order
    if jc:
        ppr.append(f'<w:jc w:val="{jc}"/>')
    ppr.append(rpr_xml(style))
    return "<w:p><w:pPr>" + "".join(ppr) + "</w:pPr>" + inner_runs + "</w:p>"


# --------------------------------------------------------------------------
# Floating text boxes for the الرقم / التاريخ / المشفوعات fields
# (same structure Word uses for the template's own header text box).
# --------------------------------------------------------------------------
_shape_id = [100]


def field_textbox(key, value, first):
    box = FIELD_BOXES[key]
    _shape_id[0] += 1
    sid = _shape_id[0]
    x, cx = box["x_left"], box["x_right"] - box["x_left"]
    y = box["y_center"] - FIELD_BOX_H // 2
    cy = FIELD_BOX_H
    para = para_xml(runs_xml(value, "field"), "field",
                    ppr_extra='<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')
    if box["cover_dots"]:  # hide the printed dots behind the typed value
        fill, vml_fill = '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>', 'filled="t" fillcolor="white"'
    else:
        fill, vml_fill = "<a:noFill/>", 'filled="f"'
    emu_pt = lambda v: f"{v / 12700:.2f}pt"  # noqa: E731
    shapetype = ('<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" '
                 'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
                 '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>') if first else ""
    return (
        '<w:r><w:rPr><w:noProof/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
        f'<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{251670000 + sid}" '
        'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{x}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{y}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{sid}" name="field {key}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill}<a:ln><a:noFill/></a:ln></wps:spPr>'
        f'<wps:txbx><w:txbxContent>{para}</w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" '
        'anchor="ctr" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp>'
        '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>'
        f'<mc:Fallback><w:pict>{shapetype}'
        f'<v:shape id="field_{key}" o:spid="_x0000_s{2000 + sid}" type="#_x0000_t202" '
        f'style="position:absolute;margin-left:{emu_pt(x)};margin-top:{emu_pt(y)};width:{emu_pt(cx)};'
        f'height:{emu_pt(cy)};z-index:{251670000 + sid};mso-position-horizontal-relative:page;'
        f'mso-position-vertical-relative:page;v-text-anchor:middle" {vml_fill} stroked="f">'
        f'<v:textbox inset="0,0,0,0"><w:txbxContent>{para}</w:txbxContent></v:textbox>'
        '</v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>'
    )


# --------------------------------------------------------------------------
# Layout choice: rough page-fit estimate (Word's real fonts are not available
# here, so this is deliberately conservative).
# --------------------------------------------------------------------------
CHARS_PER_LINE = 92        # Tosh A at 11pt across the 17.85 cm body width
LIST_CHARS_PER_LINE = 84
LINE_PT_SINGLE = 16.0      # single-spaced line height of 11pt Arabic text
BODY_BOTTOM_MM = 271.0     # keep clear of the footer printed on the letterhead


def _lines(text, per_line):
    text = re.sub(r"\*\*|\[\[|\]\]", "", text)
    return max(1, -(-len(text) // max(1, int(per_line))))


def estimate_height_pt(letter, layout, body_pt=DEFAULT_BODY_PT):
    L = LAYOUTS[layout]
    k = body_pt / 11.0                          # constants above were measured at 11pt
    h = 0.0
    rec_lines = sum(_lines(x, 70 / k) for x in letter["recipient"].split("\n"))
    h += rec_lines * 22 * k + 8                 # addressee
    if letter.get("greeting"):
        h += 22 * k + 12                        # greeting
    line_pt = LINE_PT_SINGLE * k * L["line"] / 240
    list_pt = LINE_PT_SINGLE * k * L["list_line"] / 240
    after = 14 if L["auto_after"] else L["after"] / 20
    for item in letter["body"]:
        if isinstance(item, dict):
            for li in item.get("bullets", []) + item.get("numbered", []):
                h += _lines(li, LIST_CHARS_PER_LINE / k) * list_pt + L["list_after"] / 20
            h += after
        else:
            h += _lines(item, CHARS_PER_LINE / k) * line_pt + L["before"] / 20 + after
    h += L["closing_before"] / 20 + 26          # closing line
    h += L["sig_spacers"] * 26                  # empty lines before the signature
    h += 3 * 24 + 12                            # signature table (title, gap, name)
    return h


def choose_layout(letter, requested, body_pt=DEFAULT_BODY_PT):
    if requested in LAYOUTS:
        return requested
    for name in ("spacious", "standard", "compact"):
        avail_pt = (BODY_BOTTOM_MM - LAYOUTS[name]["top"]) * 72 / 25.4
        if estimate_height_pt(letter, name, body_pt) <= avail_pt * 0.95:
            return name
    return "compact"


# --------------------------------------------------------------------------
# Document parts
# --------------------------------------------------------------------------
def addressee_xml(recipient, honorific, placeholders):
    lines = [x.strip() for x in recipient.split("\n") if x.strip()]
    if not honorific:
        return "".join(
            para_xml(runs_xml(x, "recipient", placeholders), "recipient",
                     ppr_extra=f'<w:spacing w:after="80"/><w:ind w:left="{ADDR_IND_START}" '
                               f'w:right="{ADDR_IND_END}"/>', keep_next=True)
            for x in lines)
    # Borderless two-cell RTL table: addressee on the right, honorific at the far left.
    # (bidiVisual puts the first cell on the right; centring avoids any RTL indent ambiguity.)
    width = TEXT_W - ADDR_IND_START - ADDR_IND_END
    hon_w = 2200
    right = "".join(
        para_xml(runs_xml(x, "recipient", placeholders), "recipient",
                 ppr_extra='<w:spacing w:after="0" w:line="276" w:lineRule="auto"/>', keep_next=True)
        for x in lines)
    left = para_xml(runs_xml(honorific, "recipient", placeholders), "recipient",
                    ppr_extra='<w:spacing w:after="0" w:line="276" w:lineRule="auto"/>',
                    jc="left", ltr=True, keep_next=True)
    no_borders = "".join(f'<w:{b} w:val="nil"/>' for b in
                         ("top", "left", "bottom", "right", "insideH", "insideV"))
    return (
        '<w:tbl><w:tblPr><w:bidiVisual/>'
        f'<w:tblW w:w="{width}" w:type="dxa"/><w:jc w:val="center"/>'
        f'<w:tblBorders>{no_borders}</w:tblBorders><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'
        '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" '
        'w:noHBand="1" w:noVBand="1"/></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{width - hon_w}"/><w:gridCol w:w="{hon_w}"/></w:tblGrid>'
        '<w:tr><w:trPr><w:cantSplit/></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="{width - hon_w}" w:type="dxa"/></w:tcPr>{right}</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="{hon_w}" w:type="dxa"/></w:tcPr>{left}</w:tc>'
        '</w:tr></w:tbl>'
        # small gap paragraph after the table (also stops Word merging tables)
        + para_xml("", "recipient", ppr_extra='<w:spacing w:after="0" w:line="120" w:lineRule="exact"/>',
                   keep_next=True)
    )


def signature_xml(title, name):
    """Same table as the template's signature block (left column, centred)."""
    cell_p = lambda text, style: para_xml(runs_xml(text, style) if text else "", style, jc="center",  # noqa: E731
                                          keep_next=True)
    no_borders = "".join(f'<w:{b} w:val="none" w:sz="0" w:space="0" w:color="auto"/>' for b in
                         ("top", "left", "bottom", "right", "insideH", "insideV"))
    return (
        '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:bidiVisual/>'
        '<w:tblW w:w="9013" w:type="dxa"/><w:tblInd w:w="715" w:type="dxa"/>'
        f'<w:tblBorders>{no_borders}</w:tblBorders>'
        '<w:tblCellMar><w:top w:w="108" w:type="dxa"/><w:bottom w:w="108" w:type="dxa"/></w:tblCellMar>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" '
        'w:noHBand="0" w:noVBand="1"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="4755"/><w:gridCol w:w="4258"/></w:tblGrid>'
        '<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="1239"/></w:trPr>'
        '<w:tc><w:tcPr><w:tcW w:w="4755" w:type="dxa"/></w:tcPr>' + cell_p("", "sig_title") + '</w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="4258" w:type="dxa"/></w:tcPr>'
        + cell_p(title, "sig_title") + cell_p("", "sig_title") + cell_p(name, "sig_name")
        + '</w:tc></w:tr></w:tbl>'
    )


def body_xml(letter, cfg, layout, placeholders, field_values):
    L = LAYOUTS[layout]
    out = []

    # 1) Top spacer: pushes the addressee just below the letterhead and anchors the
    #    الرقم / التاريخ / المشفوعات text boxes (page 1 only).
    spacer_twips = round((L["top"] - TOP_MARGIN_MM) * 56.6929)
    boxes = "".join(field_textbox(k, v, first=(i == 0))
                    for i, (k, v) in enumerate((k, v) for k, v in field_values if v))
    out.append(para_xml(boxes, "body",
                        ppr_extra=f'<w:spacing w:before="0" w:after="0" w:line="{spacer_twips}" '
                                  'w:lineRule="exact"/>'))

    # 2) Addressee + greeting
    out.append(addressee_xml(letter["recipient"], letter.get("honorific", ""), placeholders))
    greeting = letter.get("greeting")
    if greeting:
        out.append(para_xml(runs_xml(greeting, "greeting", placeholders), "greeting",
                            ppr_extra=f'<w:spacing w:before="60" w:after="120"/>'
                                      f'<w:ind w:left="{ADDR_IND_START}" w:right="{ADDR_IND_END}"/>',
                            keep_next=True))

    # 3) Body
    after_attr = (f'w:after="{L["after"]}" w:afterAutospacing="1"' if L["auto_after"]
                  else f'w:after="{L["after"]}"')
    body_ppr = (f'<w:spacing w:before="{L["before"]}" {after_attr} w:line="{L["line"]}" w:lineRule="auto"/>'
                f'<w:ind w:left="{BODY_IND}" w:right="{BODY_IND}"/>')
    items = letter["body"]
    for idx, item in enumerate(items):
        last = idx == len(items) - 1
        if isinstance(item, str):
            out.append(para_xml(runs_xml(item, "body", placeholders), "body", ppr_extra=body_ppr,
                                jc="both", keep_next=last))
            continue
        kind = "bullets" if "bullets" in item else "numbered" if "numbered" in item else None
        if kind is None:
            raise SystemExit(f"Unknown body item: {item!r} (use a string, bullets or numbered)")
        num_id = 1 if kind == "bullets" else 90
        entries = item[kind]
        for j, entry in enumerate(entries):
            is_last_entry = j == len(entries) - 1
            after = (L["after"] if is_last_entry else L["list_after"])
            # w:left/w:right are start/end indents; in these RTL paragraphs start = right side.
            lst_ppr = (f'<w:spacing w:before="0" w:after="{after}" w:line="{L["list_line"]}" '
                       f'w:lineRule="auto"/><w:ind w:left="{BODY_IND + 720}" w:right="{BODY_IND}" '
                       'w:hanging="360"/>')
            out.append(para_xml(runs_xml(entry, "body", placeholders), "body", ppr_extra=lst_ppr,
                                jc="both", num_id=num_id, p_style="ListParagraph",
                                keep_next=last and is_last_entry))

    # 4) Closing + signature block
    closing = letter.get("closing") or cfg["default_closing"]
    out.append(para_xml(runs_xml(closing, "closing", placeholders), "closing",
                        ppr_extra=f'<w:spacing w:before="{L["closing_before"]}" w:after="120"/>',
                        jc="center", keep_next=True))
    for _ in range(L["sig_spacers"]):
        out.append(para_xml("", "closing", jc="center", keep_next=True))
    sig = letter.get("signature") or {}
    out.append(signature_xml(sig.get("title") or cfg["signature_title"],
                             sig.get("name") or cfg["signature_name"]))
    # template's closing paragraph after the table
    out.append('<w:p><w:pPr><w:spacing w:after="120"/><w:ind w:right="-510"/><w:rPr>'
               '<w:rFonts w:ascii="Alexandria" w:eastAsia="Times New Roman" w:hAnsi="Alexandria" '
               f'w:cs="Alexandria"/><w:color w:val="{GREEN}"/><w:sz w:val="24"/><w:szCs w:val="24"/>'
               '<w:rtl/></w:rPr></w:pPr></w:p>')
    return "".join(out)


def patch_header(xml, unit_line):
    """Write the unit name (e.g. مركز الموهبة والإبداع والتميز) under «عمادة شؤون الطلاب»
    in the header text box, and give the box room for the second line."""
    if not unit_line:
        return xml
    anchor = "<w:t>عمادة شؤون الطلاب</w:t></w:r></w:p>"
    if anchor not in xml:
        print("warning: header text box not found; unit line not added", file=sys.stderr)
        return xml
    new_p = para_xml(runs_xml(unit_line, "header_unit"), "header_unit",
                     ppr_extra='<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', jc="center")
    # the empty paragraph right after the deanship line (in both the DrawingML and VML copies)
    xml = re.sub(re.escape(anchor) + r"<w:p [^>]*>.*?</w:p>", lambda m: anchor + new_p, xml, flags=re.S)
    # tighten the gap between the two lines
    xml = xml.replace('<w:spacing w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>',
                      '<w:spacing w:after="40" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>')
    # taller box (template box is one line high and clips anything more)
    xml = xml.replace('cy="282575"', 'cy="460000"').replace("height:22.25pt", "height:36.2pt")
    return xml


NUMBERED_ABSTRACT = (
    '<w:abstractNum w:abstractNumId="90"><w:multiLevelType w:val="singleLevel"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1-"/>'
    '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
    '<w:rPr><w:rFonts w:ascii="Alexandria" w:hAnsi="Alexandria" w:cs="Tosh A"/></w:rPr></w:lvl>'
    '</w:abstractNum>'
)


def patch_numbering(xml):
    if 'w:abstractNumId="90"' in xml:
        return xml
    i = xml.find("<w:num ")
    if i < 0:
        i = xml.find("</w:numbering>")
    xml = xml[:i] + NUMBERED_ABSTRACT + xml[i:]
    return xml.replace("</w:numbering>",
                       '<w:num w:numId="90"><w:abstractNumId w:val="90"/></w:num></w:numbering>')


def patch_core(xml, title):
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    xml = re.sub(r"<dc:title>.*?</dc:title>|<dc:title/>", f"<dc:title>{escape(title)}</dc:title>", xml)
    xml = re.sub(r'(<dcterms:created [^>]*>).*?(</dcterms:created>)', rf"\g<1>{now}\g<2>", xml)
    xml = re.sub(r'(<dcterms:modified [^>]*>).*?(</dcterms:modified>)', rf"\g<1>{now}\g<2>", xml)
    return xml


def build(letter, out_path, template=DEFAULT_TEMPLATE, cfg=None, layout_req=None):
    cfg = cfg or {}
    if not letter.get("recipient") or not letter.get("body"):
        raise SystemExit("letter.json needs at least 'recipient' and 'body'")
    if "greeting" not in letter:
        letter["greeting"] = cfg.get("default_greeting", "السلام عليكم ورحمة الله وبركاته")
    # optional: a calligraphic font for the salam line (config "greeting_font")
    STYLES["greeting"]["cs"] = cfg.get("greeting_font") or "Tosh A Medium"
    body_pt = float(letter.get("font_size") or cfg.get("body_font_size") or DEFAULT_BODY_PT)
    apply_font_sizes(body_pt)
    layout = choose_layout(letter, layout_req or letter.get("layout", "auto"), body_pt)

    notes, field_values = [], []
    date_text, warn = format_date(letter.get("date", ""))
    if warn:
        notes.append(warn)
    for key, value in (("number", (letter.get("number") or "").strip()), ("date", date_text),
                       ("attachments", (letter.get("attachments") or "").strip())):
        field_values.append((key, value))

    placeholders = []
    with zipfile.ZipFile(template) as zin:
        doc = zin.read("word/document.xml").decode("utf-8")
        body_start = doc.index("<w:body>") + len("<w:body>")
        sect = re.search(r"<w:sectPr\b.*?</w:sectPr>", doc, re.S).group(0)
        new_doc = (doc[:body_start] + body_xml(letter, cfg, layout, placeholders, field_values)
                   + sect + "</w:body></w:document>")
        replaced = {"word/document.xml": new_doc}
        replaced["word/header1.xml"] = patch_header(zin.read("word/header1.xml").decode("utf-8"),
                                                    cfg.get("header_unit_line", ""))
        replaced["word/numbering.xml"] = patch_numbering(zin.read("word/numbering.xml").decode("utf-8"))
        title = os.path.splitext(os.path.basename(out_path))[0]
        replaced["docProps/core.xml"] = patch_core(zin.read("docProps/core.xml").decode("utf-8"), title)

        for name, xml in replaced.items():  # fail loudly rather than write a broken file
            try:
                minidom.parseString(xml.encode("utf-8"))
            except Exception as e:  # pragma: no cover
                raise SystemExit(f"internal error: {name} is not well-formed XML ({e})")

        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = replaced[info.filename].encode("utf-8") if info.filename in replaced \
                    else zin.read(info.filename)
                zout.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
    return dict(output=os.path.abspath(out_path), layout=layout, date=date_text,
                placeholders=placeholders, notes=notes)


EXAMPLE = {
    "recipient": "سعادة عميد شؤون الطلاب",
    "honorific": "الموقر",
    "body": [
        "انطلاقًا من دور الجامعة في رعاية طلبتها وتنمية قدراتهم، وتهيئتهم للمشاركة الفاعلة في المسابقات "
        "والهاكاثونات المحلية والدولية، تعتزم عمادة شؤون الطلاب ممثلةً في مركز الموهبة والإبداع والتميز "
        "تنفيذ «برنامج الجاهزية التنافسية».",
        "ويهدف البرنامج إلى رفع جاهزية طلبة الجامعة للمشاركة في المنافسات، وتطوير أفكارهم ومشاريعهم، "
        "وتنمية مهاراتهم في إعداد العروض، والتعريف بمعايير التقييم والتحكيم.",
        "ونظرًا لما تتمتع به مؤسسة الملك عبدالعزيز ورجاله للموهبة والإبداع «موهبة» من خبرة متخصصة في "
        "تأهيل الموهوبين، يأمل المركز الاستفادة من خبراتها ومشاركتها في تنفيذ البرنامج من خلال المجالات الآتية:",
        {"bullets": [
            "تقديم ورشة تدريبية حضورية أو عن بُعد حول الاستعداد للمسابقات ومعايير تقييم المشاريع.",
            "الاطلاع على أفكار الطلبة ونماذج مشاريعهم، وتقديم التغذية الراجعة المتخصصة.",
            "المشاركة في تحكيم مشاريع الطلبة خلال اليوم الختامي للبرنامج.",
        ]},
        "عليه.. نأمل من سعادتكم التكرم بالموافقة على مخاطبة مؤسسة «موهبة» رسميًا، ودعوتها للمشاركة في "
        "تنفيذ البرنامج وفق المجالات الموضحة أعلاه.",
        "ولمزيد من المعلومات والتنسيق، يمكن التواصل مع د/ مروة الحربي على رقم الجوال (0544650166).",
    ],
    "closing": "وتفضلوا بقبول خالص التحية والتقدير،،،",
    "date": "today",
}


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("letter", nargs="?", help="letter.json (UTF-8)")
    ap.add_argument("-o", "--output", help="output .docx path")
    ap.add_argument("--template", default=DEFAULT_TEMPLATE)
    ap.add_argument("--config", default=DEFAULT_CONFIG)
    ap.add_argument("--layout", choices=["auto", *LAYOUTS], help="override the letter's layout")
    ap.add_argument("--example", action="store_true", help="print a sample letter.json and exit")
    args = ap.parse_args(argv)

    if args.example:
        print(json.dumps(EXAMPLE, ensure_ascii=False, indent=2))
        return 0
    if not args.letter or not args.output:
        ap.error("letter.json and -o output.docx are required")
    with open(args.letter, encoding="utf-8") as f:
        letter = json.load(f)
    cfg = {}
    if args.config and os.path.exists(args.config):
        with open(args.config, encoding="utf-8") as f:
            cfg = json.load(f)
    info = build(letter, args.output, template=args.template, cfg=cfg, layout_req=args.layout)
    print(f"saved: {info['output']}")
    print(f"layout: {info['layout']}")
    print(f"date: {info['date'] or '(blank)'}")
    if info["placeholders"]:
        print("placeholders to fill (highlighted yellow): " + " | ".join(info["placeholders"]))
    for n in info["notes"]:
        print("note: " + n)
    return 0


if __name__ == "__main__":
    sys.exit(main())
