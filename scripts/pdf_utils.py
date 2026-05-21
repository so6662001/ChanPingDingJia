# -*- coding: utf-8 -*-
"""ReportLab 中文 PDF 公共工具。"""
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Table, TableStyle

CHINESE_FONT_PATH = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
CHINESE_FONT = "WQYMicroHei"

_REGISTERED = False


def register_chinese_font():
    global _REGISTERED
    if _REGISTERED:
        return
    pdfmetrics.registerFont(TTFont(CHINESE_FONT, CHINESE_FONT_PATH))
    _REGISTERED = True


def styles():
    register_chinese_font()
    return {
        "title": ParagraphStyle("title", fontName=CHINESE_FONT, fontSize=22, leading=28,
                                alignment=TA_CENTER, textColor=colors.HexColor("#1F3864"),
                                spaceAfter=8),
        "subtitle": ParagraphStyle("subtitle", fontName=CHINESE_FONT, fontSize=12, leading=18,
                                   alignment=TA_CENTER, textColor=colors.HexColor("#595959"),
                                   spaceAfter=18),
        "h1": ParagraphStyle("h1", fontName=CHINESE_FONT, fontSize=16, leading=22,
                             textColor=colors.HexColor("#1F3864"),
                             spaceBefore=14, spaceAfter=8),
        "h2": ParagraphStyle("h2", fontName=CHINESE_FONT, fontSize=13, leading=20,
                             textColor=colors.HexColor("#2E75B6"),
                             spaceBefore=10, spaceAfter=6),
        "h3": ParagraphStyle("h3", fontName=CHINESE_FONT, fontSize=11, leading=18,
                             textColor=colors.HexColor("#1F3864"),
                             spaceBefore=6, spaceAfter=4),
        "body": ParagraphStyle("body", fontName=CHINESE_FONT, fontSize=10.5, leading=17,
                               alignment=TA_LEFT),
        "note": ParagraphStyle("note", fontName=CHINESE_FONT, fontSize=9.5, leading=14,
                               textColor=colors.HexColor("#595959"),
                               alignment=TA_LEFT),
        "callout": ParagraphStyle("callout", fontName=CHINESE_FONT, fontSize=10.5, leading=17,
                                  textColor=colors.HexColor("#C00000"),
                                  alignment=TA_LEFT),
        "tbl_head": ParagraphStyle("tbl_head", fontName=CHINESE_FONT, fontSize=10, leading=14,
                                   textColor=colors.white, alignment=TA_CENTER),
        "tbl_cell": ParagraphStyle("tbl_cell", fontName=CHINESE_FONT, fontSize=9.5, leading=13,
                                   alignment=TA_LEFT),
        "tbl_cell_c": ParagraphStyle("tbl_cell_c", fontName=CHINESE_FONT, fontSize=9.5, leading=13,
                                     alignment=TA_CENTER),
    }


def make_table(data, col_widths=None, head_rows=1, zebra=True, font_size=9.5):
    """data: 第一行为表头；单元格可为 str 或 Paragraph。"""
    register_chinese_font()
    style_cmds = [
        ("FONT", (0, 0), (-1, -1), CHINESE_FONT, font_size),
        ("LEADING", (0, 0), (-1, -1), font_size + 3),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFBFBF")),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]
    if head_rows > 0:
        style_cmds += [
            ("BACKGROUND", (0, 0), (-1, head_rows - 1), colors.HexColor("#305496")),
            ("TEXTCOLOR", (0, 0), (-1, head_rows - 1), colors.white),
            ("ALIGN", (0, 0), (-1, head_rows - 1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, head_rows - 1), CHINESE_FONT),
        ]
    if zebra:
        for i in range(head_rows, len(data)):
            if (i - head_rows) % 2 == 1:
                style_cmds.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F2F2F2")))
    t = Table(data, colWidths=col_widths, repeatRows=head_rows)
    t.setStyle(TableStyle(style_cmds))
    return t


PAGE_W, PAGE_H = A4
MARGIN = 16 * mm
USABLE_W = PAGE_W - 2 * MARGIN
