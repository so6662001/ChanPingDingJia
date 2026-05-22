# -*- coding: utf-8 -*-
"""把《销售作战手册 · MVP 版》Markdown 渲染为 PDF。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from md_to_pdf import build_pdf  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    md = os.path.join(ROOT, "deliverables", "销售作战手册_MVP版.md")
    pdf = os.path.join(ROOT, "deliverables", "销售作战手册_MVP版.pdf")
    build_pdf(
        md_path=md,
        pdf_path=pdf,
        header_left="销售作战手册 · MVP 版（Stage 0 限定）",
        cover_title="销售作战手册 · MVP 版",
        cover_subtitle="6 个月 5 家中型客户付费签约 = 唯一 KPI",
        cover_publisher="营销总裁室 · 销售管理部 联合发布",
    )
    print(f"[OK] {pdf}")


if __name__ == "__main__":
    main()
