# -*- coding: utf-8 -*-
"""生成《产品价格白皮书 v2》PDF 版本。"""
import os
import sys

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, PageBreak, Paragraph, Spacer
)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pricing_data as D  # noqa: E402
from pdf_utils import make_table, styles, USABLE_W, MARGIN, CHINESE_FONT  # noqa: E402


S = styles()


def P(text, style="body"):
    return Paragraph(text, S[style])


def cell(text, center=False):
    return Paragraph(str(text), S["tbl_cell_c"] if center else S["tbl_cell"])


def head_cell(text):
    return Paragraph(str(text), S["tbl_head"])


def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont(CHINESE_FONT, 8.5)
    canvas.setFillColor(colors.HexColor("#595959"))
    canvas.drawString(MARGIN, A4[1] - 12 * mm,
                      "钢铁行业数字化 · 产品价格白皮书 v2")
    canvas.drawRightString(A4[0] - MARGIN, A4[1] - 12 * mm,
                           "营销总裁室 · 产品中心 · 销售管理部")
    canvas.setStrokeColor(colors.HexColor("#1F3864"))
    canvas.line(MARGIN, A4[1] - 14 * mm, A4[0] - MARGIN, A4[1] - 14 * mm)
    canvas.setFont(CHINESE_FONT, 8.5)
    canvas.drawCentredString(A4[0] / 2, 10 * mm, f"— 第 {doc.page} 页 —")
    canvas.restoreState()


def section(story, title):
    story.append(Spacer(1, 6))
    story.append(Paragraph(title, S["h1"]))


def subsection(story, title):
    story.append(Paragraph(title, S["h2"]))


def note(story, text):
    story.append(Paragraph("说明：" + text, S["note"]))
    story.append(Spacer(1, 4))


def build_table(headers, rows, widths):
    data = [[head_cell(h) for h in headers]]
    for row in rows:
        data.append([cell(v) for v in row])
    widths_pts = [w if w >= 100 else w * mm for w in widths]
    return make_table(data, col_widths=widths_pts)


def kv_table(rows, widths):
    data = []
    for k, v in rows:
        data.append([Paragraph(k, S["tbl_cell"]), Paragraph(v, S["tbl_cell"])])
    t = make_table(data, col_widths=widths, head_rows=0)
    return t


def cover(story):
    story.append(Spacer(1, 60))
    story.append(P("钢铁行业数字化", "title"))
    story.append(P("《产品价格白皮书 v2》", "title"))
    story.append(Spacer(1, 6))
    story.append(P("营销总裁室 · 产品中心 · 销售管理部 联合发布", "subtitle"))
    story.append(Spacer(1, 30))
    story.append(P("三大目标", "h1"))
    targets = [
        ("便于市场开发", "用 3 套餐 + 6 单品压缩 SKU，让一线 30 秒讲完。"),
        ("有利可图",     "盈利品锚定『年利中位数 × 3%』；折扣红线 65 折封顶。"),
        ("自我裂变",     "老带新返现 15% / 案例升级 / 积分 / 峰会券 四件套全员开放。"),
    ]
    rows = [(k, v) for k, v in targets]
    story.append(kv_table(rows, [40 * mm, USABLE_W - 40 * mm]))
    story.append(Spacer(1, 20))
    story.append(P("文档目录", "h1"))
    toc = [
        ("01 客群分级",       "五级客群定义与对应主推档"),
        ("02 产品全景",       "引流 / 粘性 / 盈利三类产品定位"),
        ("03 单品价目",       "全部 SKU 五档 v2 价格（含此前留白部分）"),
        ("04 套餐价目",       "双宝 / 全家桶 / 工业全家桶"),
        ("05 折扣与续费",     "用户数 / 多年付 / 红线规则"),
        ("06 裂变四件套",     "返现 / 案例升级 / 积分 / 峰会券"),
        ("07 销售话术索引",   "12 个场景导航至《销售作战手册（一线版）》"),
        ("08 v1 → v2 修订对照", "本次价格白皮书所有变更点"),
        ("09 内部毛利红线",   "销售总监审批用：综合折扣 65 折底线"),
    ]
    story.append(kv_table(toc, [42 * mm, USABLE_W - 42 * mm]))
    story.append(PageBreak())


def section_tier(story):
    section(story, "01 客群分级（v2，主战场为中型）")
    headers = ["客群", "月销", "日均", "单据量", "对应主推档"]
    widths = [38 * mm, 32 * mm, 30 * mm, 36 * mm, USABLE_W - 136 * mm]
    story.append(build_table(headers, D.CUSTOMER_TIERS, widths))
    note(story, "微型客户以货袋子免费会员留存为主，禁止在合同里直接报付费版本。")


def section_product_map(story):
    section(story, "02 产品全景：引流 / 粘性 / 盈利")
    headers = ["分类", "产品", "价格形态", "市场角色"]
    widths = [22 * mm, 50 * mm, 38 * mm, USABLE_W - 110 * mm]
    story.append(build_table(headers, D.PRODUCT_MAP, widths))
    note(story, "提货宝定位为唯一引流主钩；AI 录入降级为提货宝插件，禁止单独商务报价。")


def block(story, sub_title, headers, rows, widths, n=None):
    subsection(story, sub_title)
    story.append(build_table(headers, rows, widths))
    if n:
        note(story, n)


def section_skus(story):
    section(story, "03 单品价目（v2 全量）")

    block(story, "3.1 货袋子会员（v2 新填）",
          ["版本", "客群", "年费", "日均", "月上限", "权益"],
          D.HUODAIZI,
          [22, 18, 22, 14, 22, USABLE_W / mm - 98],
          "顶级版赠送线下峰会订位 1 次；可累计获得峰会门票券。")

    block(story, "3.2 提货宝",
          ["版本", "客群", "年费", "日均", "月单上限"], D.TIHUOBAO,
          [30, 28, 30, 24, USABLE_W / mm - 112])

    block(story, "3.3 对账宝",
          ["版本", "客群", "年费", "日均", "月对账上限"], D.DUIZHANGBAO,
          [30, 28, 30, 24, USABLE_W / mm - 112])

    block(story, "3.4 双宝套餐（v2 让利 ≥30%）",
          ["客群", "单买价", "套餐价", "节省金额/比例", "货袋子权益"],
          D.DOUBLE_BUNDLE,
          [22, 22, 22, 35, USABLE_W / mm - 101],
          "双宝套餐让利与全家桶折扣阶梯（34%/42%/41%/45%）一致。")

    block(story, "3.5 库存宝（v2 新填）",
          ["版本", "客群", "年费", "日均", "月配货量上限"], D.KUCUNBAO,
          [30, 28, 30, 24, USABLE_W / mm - 112],
          "库存宝定价低于 AI 出纳 / AI 结算，作为粘性品引导客户向 WMS/AI 升级。")

    block(story, "3.6 AI 结算（v2 集团版上调至 ¥58,000）",
          ["版本", "客群", "年费", "日均", "节省(h)", "节省(¥)", "月单上限"],
          D.AI_SETTLE,
          [22, 18, 22, 16, 22, 24, USABLE_W / mm - 124],
          "价值/价格比维持在 7-10 倍区间最易成交。")

    block(story, "3.7 AI 出纳（v2 集团版上调）",
          ["版本", "客群", "年费", "日均", "节省(h)", "节省(¥)", "月单上限"],
          D.AI_CASH,
          [22, 18, 22, 16, 22, 24, USABLE_W / mm - 124])

    block(story, "3.8 AI 绩效",
          ["版本", "客群", "年费", "月均", "销售员人数上限"], D.AI_PERF,
          [30, 28, 30, 24, USABLE_W / mm - 112])

    block(story, "3.9 AI 报价（v2 新填）",
          ["版本", "客群", "年费", "日均", "月上限"], D.AI_QUOTE,
          [30, 28, 30, 24, USABLE_W / mm - 112],
          "AI 报价是销售人员日常用得最多的粘性品。")

    block(story, "3.10 AI 经营分析（v2 新增定价）",
          ["舱位", "受众", "年费", "内容"], D.AI_ANALYSIS,
          [38, 30, 32, USABLE_W / mm - 100],
          "加购模块：" + " / ".join(f"{k} {v}" for k, v in D.AI_ANALYSIS_ADDONS))

    block(story, "3.11 库准 WMS（推荐 C 业务量模式）",
          ["档位", "适配", "月吞吐", "用户", "仓库", "年订阅", "等效一次性"],
          D.WMS_C,
          [22, 30, 28, 14, 14, 26, USABLE_W / mm - 134],
          "月吞吐 = 入库 + 出库 + 移库 + 盘点单数之和。")

    block(story, "3.12 钢贸宝 ERP（订阅 = 一次性 ÷ 4）",
          ["版本", "一次性/用户", "订阅/年", "云端费/年", "次年服务费"], D.GANGMAOBAO,
          [30, 30, 26, 26, USABLE_W / mm - 112])

    block(story, "3.13 钢企通 ERP（v2 改为一次性 + 订阅二选一）",
          ["版本", "适配", "一次性/用户", "订阅/年", "云端费/年", "次年服务费"],
          D.GANGQITONG,
          [25, 26, 28, 24, 24, USABLE_W / mm - 127])

    block(story, "3.14 MES（v2 新模型：按产线档一口价）",
          ["档位", "适配", "年订阅（含云+服务+升级）", "等效一次性"], D.MES,
          [25, 50, 50, USABLE_W / mm - 125],
          "所有 MES 客户均不再使用『2,000 元/天 + 0.6 元/吨』口径。")

    block(story, "3.15 钢信用增值（v2 新填）",
          ["档位", "价格", "权益"], D.GANGXINYONG,
          [30, 30, USABLE_W / mm - 60])


def section_bundles(story):
    section(story, "04 套餐价目（v2 主推：3 个套餐覆盖 80% 成交）")
    block(story, "4.1 贸易加工类全家桶（4 档）",
          ["套餐", "包含内容", "年订阅价", "折扣 vs 单买"], D.TRADE_BUNDLE,
          [30, USABLE_W / mm - 90, 30, 30])

    subsection(story, "4.2 工业全家桶（v2 新模型）")
    story.append(P(D.INDUSTRY_BUNDLE_FORMULA, "callout"))
    story.append(Spacer(1, 4))
    block(story, "  示例对照表",
          ["示例客户", "规模", "单买推算", "首年套餐价", "次年起续费"],
          D.INDUSTRY_BUNDLE_EXAMPLE,
          [22, 28, USABLE_W / mm - 122, 28, 24],
          "口诀：『主品打底，整桶七折，次年一八』。")


def section_discount(story):
    section(story, "05 折扣与续费规则（v2 红线版）")
    block(story, "5.1 用户数阶梯折扣",
          ["用户数", "折扣"], D.USER_TIER_DISCOUNT,
          [40, USABLE_W / mm - 40])
    block(story, "5.2 多年付折扣（仅订阅版）",
          ["付费方式", "折扣", "附加权益"], D.MULTI_YEAR_DISCOUNT,
          [30, 20, USABLE_W / mm - 50])

    subsection(story, "5.3 折扣红线（v2 新增，最关键）")
    data = [[head_cell("规则"), head_cell("内容")]]
    for k, v in D.DISCOUNT_REDLINE:
        data.append([Paragraph(k, S["callout"]), Paragraph(v, S["tbl_cell"])])
    story.append(make_table(data, col_widths=[42 * mm, USABLE_W - 42 * mm]))


def section_fission(story):
    section(story, "06 客户自我裂变机制 v2（裂变四件套）")
    block(story, "全员可执行的 4 种裂变动作",
          ["机制", "规则"], D.FISSION_KIT,
          [38, USABLE_W / mm - 38])
    note(story,
         "执行口径：返现走『市场费用-推荐返佣』；案例升级由市场部审批；"
         "积分由货袋子产品自动结算；峰会门票券由销售总监统一发放。")


def section_scenes(story):
    section(story, "07 销售话术索引（详见《销售作战手册（一线版）》）")
    block(story, "12 个场景一览",
          ["编号", "场景", "适用客群", "首选报价"], D.SALES_SCENES,
          [16, USABLE_W / mm - 96, 32, 48])


def section_diff(story):
    section(story, "08 v1 → v2 修订对照（共 15 处）")
    block(story, "本次主要修订",
          ["#", "条目", "v1 原版", "v2 修订"], D.V1_V2_DIFF,
          [10, 30, USABLE_W / mm - 100, 60])
    note(story, "未改动钢贸宝 ERP / 提货宝 / 对账宝原档位价格；主要修复留白、价值/价格倒挂、并新增折扣红线与裂变机制。")


def section_redline(story):
    section(story, "09 内部毛利红线（销售总监 / 总裁审批用）")
    rows = [
        ("中型·标准包",       "¥19,800",  "¥12,870",  "¥13,800",  "¥12,870",  "≤13,800 销售签；≤12,870 总裁签"),
        ("大型·专业包",       "¥58,000",  "¥37,700",  "¥40,600",  "¥37,700",  "同上"),
        ("超大·集团包",       "¥158,000", "¥102,700", "¥110,000", "¥102,700", "另需私有化交付预算评估"),
        ("工业全家桶·中型 ⭐", "¥93,600",  "¥60,840",  "¥65,000",  "¥60,840",  "需对齐用户数 + 产线数双口径"),
        ("钢企通 M7 高端",    "¥11,800",  "¥7,670",   "¥8,260",   "¥7,670",   "按用户数线性折后再叠加红线"),
        ("MES 标准版",        "¥88,000",  "¥57,200",  "¥61,600",  "¥57,200",  "需含至少 1 次现场调试"),
    ]
    block(story, "审批红线表",
          ["套餐 / 单品", "目录价", "65 折底线", "销售签", "总裁签", "审批要求"],
          rows, [30, 22, 22, 22, 22, USABLE_W / mm - 118])
    note(story, "任何低于 65 折成交需在审批单中明确：①客户体量 ②竞争对手报价 ③战略价值。")


def main():
    out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "deliverables", "产品价格白皮书_v2.pdf")
    doc = BaseDocTemplate(out, pagesize=A4,
                          leftMargin=MARGIN, rightMargin=MARGIN,
                          topMargin=18 * mm, bottomMargin=14 * mm,
                          title="产品价格白皮书 v2")
    frame = Frame(doc.leftMargin, doc.bottomMargin,
                  doc.width, doc.height, id="main")
    doc.addPageTemplates([PageTemplate(id="all", frames=[frame], onPage=header_footer)])

    story = []
    cover(story)
    section_tier(story)
    section_product_map(story)
    story.append(PageBreak())
    section_skus(story)
    story.append(PageBreak())
    section_bundles(story)
    story.append(PageBreak())
    section_discount(story)
    section_fission(story)
    story.append(PageBreak())
    section_scenes(story)
    story.append(PageBreak())
    section_diff(story)
    story.append(PageBreak())
    section_redline(story)

    doc.build(story)
    print(f"[OK] {out}")


if __name__ == "__main__":
    main()
