# -*- coding: utf-8 -*-
"""通用 Markdown → 中文 PDF 渲染器（基于 reportlab + WQYMicroHei）。

支持：标题 # / ## / ### / ####，段落，无序/有序列表，引用块 >，
Markdown 表格（|），分隔线 ---，粗体 **，斜体 *，行内代码 `。
"""
import os
import re
import sys

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, PageBreak,
    Paragraph, Spacer, Table, TableStyle, KeepTogether,
)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pdf_utils import register_chinese_font, CHINESE_FONT, MARGIN, USABLE_W  # noqa: E402

register_chinese_font()


# -----------------------------------------------------------------
# 样式
# -----------------------------------------------------------------

def _style(name, size, leading, **kw):
    base = dict(fontName=CHINESE_FONT, fontSize=size, leading=leading)
    base.update(kw)
    return ParagraphStyle(name, **base)


STYLES = {
    "h1":      _style("h1", 18, 26, textColor=colors.HexColor("#1F3864"),
                      spaceBefore=12, spaceAfter=10),
    "h2":      _style("h2", 14, 22, textColor=colors.HexColor("#2E75B6"),
                      spaceBefore=10, spaceAfter=6),
    "h3":      _style("h3", 12, 19, textColor=colors.HexColor("#1F3864"),
                      spaceBefore=8, spaceAfter=4),
    "h4":      _style("h4", 11, 17, textColor=colors.HexColor("#305496"),
                      spaceBefore=6, spaceAfter=3),
    "body":    _style("body", 10, 16),
    "list":    _style("list", 10, 16, leftIndent=14, bulletIndent=2),
    "quote":   _style("quote", 10, 16, leftIndent=10,
                      textColor=colors.HexColor("#595959"),
                      borderColor=colors.HexColor("#2E75B6"),
                      borderPadding=(4, 8, 4, 8)),
    "tbl_h":   _style("tbl_h", 9.5, 13, textColor=colors.white),
    "tbl_c":   _style("tbl_c", 9, 13),
    "title":   _style("title", 22, 30, textColor=colors.HexColor("#1F3864"),
                      alignment=1, spaceAfter=8),
    "subtitle":_style("subtitle", 12, 18, textColor=colors.HexColor("#595959"),
                      alignment=1, spaceAfter=18),
}


# -----------------------------------------------------------------
# 行内格式
# -----------------------------------------------------------------

def _inline(text):
    """处理行内 markdown：粗体 / 斜体 / 行内代码 / 转义。"""
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = re.sub(r"`([^`]+)`",
                  r'<font face="Courier" color="#C00000">\1</font>', text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", text)
    text = re.sub(r"(?<!\*)\*([^*]+)\*(?!\*)", r"<i>\1</i>", text)
    text = text.replace("☐", "□")
    return text


# -----------------------------------------------------------------
# 表格
# -----------------------------------------------------------------

def _build_table(rows):
    n = len(rows[0])
    col_w = USABLE_W / n
    weights = [1.0] * n
    if n >= 4:
        weights[1] = 1.6
        if n >= 5:
            weights[-1] = 1.6
    total = sum(weights)
    col_widths = [USABLE_W * w / total for w in weights]

    data = []
    for ri, row in enumerate(rows):
        out = []
        for cell_text in row:
            style = STYLES["tbl_h"] if ri == 0 else STYLES["tbl_c"]
            out.append(Paragraph(_inline(cell_text), style))
        data.append(out)

    style_cmds = [
        ("FONT", (0, 0), (-1, -1), CHINESE_FONT, 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFBFBF")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#305496")),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    for i in range(1, len(data)):
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i),
                               colors.HexColor("#F2F2F2")))
    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle(style_cmds))
    return t


# -----------------------------------------------------------------
# Markdown 解析
# -----------------------------------------------------------------

def _flush_paragraph(buf, story):
    if buf:
        text = " ".join(b.strip() for b in buf if b.strip())
        if text:
            story.append(Paragraph(_inline(text), STYLES["body"]))
        buf.clear()


def _flush_list(items, story, ordered):
    for i, txt in enumerate(items, 1):
        bullet = f"{i}." if ordered else "·"
        p = Paragraph(f"{bullet} {_inline(txt)}", STYLES["list"])
        story.append(p)


def _flush_quote(lines, story):
    text = " ".join(l.strip() for l in lines if l.strip())
    if text:
        story.append(Paragraph(_inline(text), STYLES["quote"]))


def _flush_table(lines, story):
    rows = []
    for ln in lines:
        ln = ln.strip()
        if ln.startswith("|"):
            ln = ln[1:]
        if ln.endswith("|"):
            ln = ln[:-1]
        cells = [c.strip() for c in ln.split("|")]
        rows.append(cells)
    rows = [r for r in rows if not all(set(c) <= set("-: ") for c in r)]
    if rows:
        story.append(_build_table(rows))
        story.append(Spacer(1, 4))


def md_to_flowables(md_text):
    """Return reportlab flowables from Markdown text."""
    story = []
    lines = md_text.splitlines()
    i = 0
    para_buf = []
    while i < len(lines):
        line = lines[i].rstrip()
        stripped = line.strip()

        if stripped == "":
            _flush_paragraph(para_buf, story)
            i += 1
            continue

        if stripped.startswith("```"):
            _flush_paragraph(para_buf, story)
            i += 1
            code_lines = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            i += 1
            code_text = "\n".join(code_lines)
            story.append(Paragraph(
                f'<font face="Courier" size="9">{_inline(code_text).replace(chr(10), "<br/>")}</font>',
                STYLES["body"]))
            continue

        if stripped == "---":
            _flush_paragraph(para_buf, story)
            story.append(Spacer(1, 6))
            i += 1
            continue

        m = re.match(r"^(#{1,4})\s+(.*)$", stripped)
        if m:
            _flush_paragraph(para_buf, story)
            level = len(m.group(1))
            text = m.group(2)
            key = f"h{level}"
            if level == 1 and not story:
                story.append(Paragraph(_inline(text), STYLES["title"]))
            else:
                story.append(Paragraph(_inline(text), STYLES[key]))
            i += 1
            continue

        if stripped.startswith("|"):
            _flush_paragraph(para_buf, story)
            tbl_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl_lines.append(lines[i])
                i += 1
            _flush_table(tbl_lines, story)
            continue

        if stripped.startswith(">"):
            _flush_paragraph(para_buf, story)
            qlines = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                qlines.append(lines[i].strip()[1:].strip())
                i += 1
            _flush_quote(qlines, story)
            continue

        m = re.match(r"^[-*]\s+(.*)$", stripped)
        if m:
            _flush_paragraph(para_buf, story)
            items = []
            while i < len(lines):
                m2 = re.match(r"^[-*]\s+(.*)$", lines[i].strip())
                if m2:
                    items.append(m2.group(1))
                    i += 1
                elif lines[i].strip().startswith("  ") or lines[i].strip() == "":
                    if lines[i].strip() == "":
                        break
                    items[-1] += " " + lines[i].strip()
                    i += 1
                else:
                    break
            _flush_list(items, story, ordered=False)
            continue

        m = re.match(r"^\d+\.\s+(.*)$", stripped)
        if m:
            _flush_paragraph(para_buf, story)
            items = []
            while i < len(lines):
                m2 = re.match(r"^\d+\.\s+(.*)$", lines[i].strip())
                if m2:
                    items.append(m2.group(1))
                    i += 1
                else:
                    break
            _flush_list(items, story, ordered=True)
            continue

        para_buf.append(stripped)
        i += 1

    _flush_paragraph(para_buf, story)
    return story


# -----------------------------------------------------------------
# 文档构建
# -----------------------------------------------------------------

def build_pdf(md_path, pdf_path, header_left, cover_title,
              cover_subtitle="", cover_publisher=""):
    register_chinese_font()
    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    md_text = re.sub(r"^# .+?\n", "", md_text, count=1)

    def header_footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(CHINESE_FONT, 8.5)
        canvas.setFillColor(colors.HexColor("#595959"))
        canvas.drawString(MARGIN, A4[1] - 12 * mm, header_left)
        canvas.drawRightString(A4[0] - MARGIN, A4[1] - 12 * mm,
                               "营销总裁室 · 销售管理部")
        canvas.setStrokeColor(colors.HexColor("#1F3864"))
        canvas.line(MARGIN, A4[1] - 14 * mm, A4[0] - MARGIN, A4[1] - 14 * mm)
        canvas.drawCentredString(A4[0] / 2, 10 * mm, f"— 第 {doc.page} 页 —")
        canvas.restoreState()

    doc = BaseDocTemplate(pdf_path, pagesize=A4,
                          leftMargin=MARGIN, rightMargin=MARGIN,
                          topMargin=18 * mm, bottomMargin=14 * mm,
                          title=cover_title)
    frame = Frame(doc.leftMargin, doc.bottomMargin,
                  doc.width, doc.height, id="main")
    doc.addPageTemplates([PageTemplate(id="all", frames=[frame],
                                       onPage=header_footer)])

    story = []
    story.append(Spacer(1, 80))
    story.append(Paragraph(cover_title, STYLES["title"]))
    if cover_subtitle:
        story.append(Paragraph(cover_subtitle, STYLES["subtitle"]))
    if cover_publisher:
        story.append(Spacer(1, 100))
        story.append(Paragraph(cover_publisher, STYLES["subtitle"]))
    story.append(PageBreak())
    story.extend(md_to_flowables(md_text))

    doc.build(story)
    return pdf_path
