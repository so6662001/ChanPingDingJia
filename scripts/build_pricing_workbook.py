# -*- coding: utf-8 -*-
"""生成《产品价格白皮书 v2》Excel 工作簿。"""
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pricing_data as D  # noqa: E402


THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

TITLE_FILL = PatternFill("solid", fgColor="1F3864")
TITLE_FONT = Font(name="DengXian", size=16, bold=True, color="FFFFFF")
H1_FILL = PatternFill("solid", fgColor="2E75B6")
H1_FONT = Font(name="DengXian", size=12, bold=True, color="FFFFFF")
H2_FILL = PatternFill("solid", fgColor="DDEBF7")
H2_FONT = Font(name="DengXian", size=11, bold=True, color="1F3864")
HEAD_FILL = PatternFill("solid", fgColor="305496")
HEAD_FONT = Font(name="DengXian", size=11, bold=True, color="FFFFFF")
ZEBRA_FILL = PatternFill("solid", fgColor="F2F2F2")
NOTE_FONT = Font(name="DengXian", size=10, italic=True, color="595959")
CELL_FONT = Font(name="DengXian", size=11)
WRAP_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def title_row(ws, row, text, ncols, fill=TITLE_FILL, font=TITLE_FONT, height=28):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.fill = fill
    c.font = font
    c.alignment = WRAP_CENTER
    ws.row_dimensions[row].height = height


def head_row(ws, row, headers):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = WRAP_CENTER
        c.border = BORDER
    ws.row_dimensions[row].height = 26


def data_rows(ws, start_row, rows, money_cols=()):
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row, 1):
            c = ws.cell(row=start_row + ri, column=ci, value=val)
            c.font = CELL_FONT
            c.alignment = WRAP_CENTER if ci != len(row) else WRAP_LEFT
            c.border = BORDER
            if ri % 2 == 1:
                c.fill = ZEBRA_FILL
            if ci in money_cols and isinstance(val, (int, float)) and val:
                c.number_format = "¥#,##0"
        ws.row_dimensions[start_row + ri].height = 22
    return start_row + len(rows)


def note_row(ws, row, text, ncols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = NOTE_FONT
    c.alignment = WRAP_LEFT
    ws.row_dimensions[row].height = 32


def section_title(ws, row, text, ncols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.fill = H2_FILL
    c.font = H2_FONT
    c.alignment = WRAP_LEFT
    ws.row_dimensions[row].height = 24


# ============================================================
# Sheet 1：封面 / 导读
# ============================================================
def build_cover(wb):
    ws = wb.active
    ws.title = "封面与导读"
    set_widths(ws, [4, 22, 60, 18])
    ws.row_dimensions[1].height = 6

    ws.merge_cells("B2:D2")
    c = ws["B2"]
    c.value = "钢铁行业数字化\n《产品价格白皮书 v2》"
    c.font = Font(name="DengXian", size=22, bold=True, color="1F3864")
    c.alignment = WRAP_CENTER
    ws.row_dimensions[2].height = 64

    ws.merge_cells("B3:D3")
    c = ws["B3"]
    c.value = "营销总裁室 · 产品中心 · 销售管理部 联合发布"
    c.font = Font(name="DengXian", size=12, italic=True, color="595959")
    c.alignment = WRAP_CENTER
    ws.row_dimensions[3].height = 22

    ws.row_dimensions[4].height = 8

    ws.merge_cells("B5:D5")
    c = ws["B5"]
    c.value = "三大目标"
    c.fill = H1_FILL
    c.font = H1_FONT
    c.alignment = WRAP_LEFT
    ws.row_dimensions[5].height = 26

    targets = [
        ("便于市场开发", "用 3 套餐 + 6 单品压缩 SKU，让一线 30 秒讲完。"),
        ("有利可图",     "盈利品锚定『年利中位数 × 3%』；折扣红线 65 折封顶。"),
        ("自我裂变",     "老带新返现 15% / 案例升级 / 积分 / 峰会券 四件套全员开放。"),
    ]
    r = 6
    for k, v in targets:
        ws.cell(row=r, column=2, value=k).font = Font(name="DengXian", size=12, bold=True, color="1F3864")
        ws.cell(row=r, column=2).alignment = WRAP_LEFT
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        c = ws.cell(row=r, column=3, value=v)
        c.font = CELL_FONT
        c.alignment = WRAP_LEFT
        ws.row_dimensions[r].height = 24
        r += 1

    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    c = ws.cell(row=r, column=2, value="工作簿目录")
    c.fill = H1_FILL
    c.font = H1_FONT
    c.alignment = WRAP_LEFT
    ws.row_dimensions[r].height = 26
    r += 1

    toc = [
        ("01 客群分级",       "五级客群定义与对应主推档"),
        ("02 产品全景",       "引流 / 粘性 / 盈利三类产品定位"),
        ("03 单品价目",       "全部 SKU 五档 v2 价格（含此前留白部分）"),
        ("04 套餐价目",       "双宝 / 全家桶 / 工业全家桶"),
        ("05 折扣与续费",     "用户数 / 多年付 / 红线规则"),
        ("06 裂变四件套",     "返现 / 案例升级 / 积分 / 峰会券"),
        ("07 销售话术索引",   "12 个场景导航至《销售作战手册（一线版）》"),
        ("08 v1→v2 修订对照", "本次价格白皮书所有变更点"),
        ("09 内部毛利红线",   "销售总监审批用：综合折扣 65 折底线"),
    ]
    for k, v in toc:
        ws.cell(row=r, column=2, value=k).font = Font(name="DengXian", size=11, bold=True)
        ws.cell(row=r, column=2).alignment = WRAP_LEFT
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        c = ws.cell(row=r, column=3, value=v)
        c.font = CELL_FONT
        c.alignment = WRAP_LEFT
        ws.row_dimensions[r].height = 22
        r += 1


# ============================================================
# Sheet 2：客群分级
# ============================================================
def build_tier(wb):
    ws = wb.create_sheet("01 客群分级")
    set_widths(ws, [22, 22, 22, 22, 26])
    title_row(ws, 1, "客群分级（v2，主战场为中型）", 5)
    head_row(ws, 2, ["客群", "月销", "日均", "单据量", "对应主推档"])
    data_rows(ws, 3, D.CUSTOMER_TIERS)
    note_row(ws, 3 + len(D.CUSTOMER_TIERS),
             "说明：『单据量』口径取销售/采购/出库/入库总单数，作为 SaaS 容量分档依据；微型客户以货袋子免费会员留存为主，禁止在合同里直接报付费版本。", 5)


# ============================================================
# Sheet 3：产品全景
# ============================================================
def build_product_map(wb):
    ws = wb.create_sheet("02 产品全景")
    set_widths(ws, [12, 32, 28, 60])
    title_row(ws, 1, "产品全景：引流 / 粘性 / 盈利", 4)
    head_row(ws, 2, ["分类", "产品", "价格形态", "市场角色"])
    end = data_rows(ws, 3, D.PRODUCT_MAP)
    note_row(ws, end, "提货宝定位为唯一引流主钩；AI 录入降级为提货宝插件，禁止单独商务报价。", 4)


# ============================================================
# Sheet 4：单品价目（合并所有产品）
# ============================================================
def build_skus(wb):
    ws = wb.create_sheet("03 单品价目")
    set_widths(ws, [16, 12, 16, 12, 18, 30, 28])
    r = 1
    title_row(ws, r, "单品价目（v2 全量）", 7); r += 1

    def block(title, headers, rows, money_cols=(), note=None, ncols=7):
        nonlocal r
        section_title(ws, r, title, ncols); r += 1
        head_row(ws, r, headers + [""] * (ncols - len(headers))); r += 1
        for row in rows:
            padded = list(row) + [""] * (ncols - len(row))
            for ci, val in enumerate(padded, 1):
                c = ws.cell(row=r, column=ci, value=val)
                c.font = CELL_FONT
                c.alignment = WRAP_CENTER
                c.border = BORDER
                if ci in money_cols and isinstance(val, (int, float)) and val:
                    c.number_format = "¥#,##0"
            ws.row_dimensions[r].height = 22
            r += 1
        if note:
            note_row(ws, r, note, ncols); r += 1
        r += 1

    block("3.1 货袋子会员（v2 新填）",
          ["版本", "客群", "年费", "日均", "月上限", "权益"], D.HUODAIZI,
          money_cols=(3,),
          note="顶级版赠送线下峰会订位 1 次；可累计获得峰会门票券（裂变四件套）。")
    block("3.2 提货宝", ["版本", "客群", "年费", "日均", "月单上限"],
          D.TIHUOBAO, money_cols=(3,),
          note="提货宝是引流主钩；销售必须在第一通电话即给客户开通对应免费版。")
    block("3.3 对账宝", ["版本", "客群", "年费", "日均", "月对账上限"],
          D.DUIZHANGBAO, money_cols=(3,))
    block("3.4 双宝套餐（v2 让利 ≥30%）",
          ["客群", "单买价", "套餐价", "节省金额/比例", "货袋子权益"],
          D.DOUBLE_BUNDLE, money_cols=(2, 3),
          note="双宝套餐让利与全家桶折扣阶梯（34%/42%/41%/45%）一致，避免内部矛盾。")
    block("3.5 库存宝（v2 新填）",
          ["版本", "客群", "年费", "日均", "月配货量上限"], D.KUCUNBAO,
          money_cols=(3,),
          note="库存宝定价低于 AI 出纳 / AI 结算，作为粘性品引导客户向 WMS/AI 升级。")
    block("3.6 AI 结算（v2 集团版上调）",
          ["版本", "客群", "年费", "日均", "节省(h)", "节省(¥)", "月单上限"],
          D.AI_SETTLE, money_cols=(3,),
          note="价值/价格比维持在 7-10 倍区间最易成交；过高反而引起『不像真的』的怀疑。")
    block("3.7 AI 出纳（v2 集团版上调）",
          ["版本", "客群", "年费", "日均", "节省(h)", "节省(¥)", "月单上限"],
          D.AI_CASH, money_cols=(3,))
    block("3.8 AI 绩效", ["版本", "客群", "年费", "月均", "销售员人数上限"],
          D.AI_PERF, money_cols=(3,))
    block("3.9 AI 报价（v2 新填）",
          ["版本", "客群", "年费", "日均", "月上限"], D.AI_QUOTE,
          money_cols=(3,),
          note="AI 报价是销售人员日常用得最多的粘性品，定价介于 AI 出纳与 AI 结算之间。")
    block("3.10 AI 经营分析（v2 新增定价）",
          ["舱位", "受众", "年费", "内容"], D.AI_ANALYSIS,
          note="加购模块：" + " / ".join(f"{k} {v}" for k, v in D.AI_ANALYSIS_ADDONS))
    block("3.11 库准 WMS（推荐 C 业务量模式）",
          ["档位", "适配", "月吞吐上限", "用户数上限", "仓库数上限", "年订阅", "等效一次性"],
          D.WMS_C,
          note="月吞吐 = 入库 + 出库 + 移库 + 盘点；一次性版次年起服务费按订阅 25% 收取。")
    block("3.12 钢贸宝 ERP（订阅 = 一次性 ÷ 4）",
          ["版本", "一次性/用户", "订阅/用户/年", "云端费/用户/年", "次年服务费"],
          D.GANGMAOBAO, money_cols=(2, 3, 4))
    block("3.13 钢企通 ERP（v2 改为一次性 + 订阅二选一）",
          ["版本", "适配", "一次性/用户", "订阅/用户/年", "云端费/用户/年", "次年服务费"],
          D.GANGQITONG, money_cols=(3, 4, 5))
    block("3.14 MES（v2 新模型：按产线档一口价）",
          ["档位", "适配", "年订阅（含云+服务+升级）", "等效一次性"], D.MES,
          note="所有 MES 客户均不再使用『2,000 元/天交付 + 0.6 元/吨摊销』口径。")
    block("3.15 钢信用增值（v2 新填）",
          ["档位", "价格", "权益"], D.GANGXINYONG)


# ============================================================
# Sheet 5：套餐价目
# ============================================================
def build_bundles(wb):
    ws = wb.create_sheet("04 套餐价目")
    set_widths(ws, [22, 60, 24, 28])
    r = 1
    title_row(ws, r, "套餐价目（v2 主推：3 个套餐覆盖 80% 成交）", 4); r += 1

    section_title(ws, r, "4.1 贸易加工类全家桶（4 档）", 4); r += 1
    head_row(ws, r, ["套餐", "包含内容", "年订阅价", "折扣 vs 单买"]); r += 1
    r = data_rows(ws, r, D.TRADE_BUNDLE)

    r += 1
    section_title(ws, r, "4.2 工业全家桶（v2 新模型）", 4); r += 1
    note_row(ws, r, D.INDUSTRY_BUNDLE_FORMULA, 4); r += 1
    head_row(ws, r, ["示例客户", "规模", "单买推算", "首年套餐价", "次年起续费"]); r += 1
    set_widths(ws, [22, 24, 60, 24, 22])
    for row in D.INDUSTRY_BUNDLE_EXAMPLE:
        for ci, val in enumerate(row, 1):
            c = ws.cell(row=r, column=ci, value=val)
            c.font = CELL_FONT
            c.alignment = WRAP_LEFT if ci == 3 else WRAP_CENTER
            c.border = BORDER
        ws.row_dimensions[r].height = 38
        r += 1

    r += 1
    note_row(ws,
        r,
        "公式记忆口诀：『主品打底，整桶七折，次年一八』——主品按用户/产线/吞吐档算单价，整体 70% 即首年套餐价，次年按首年 18% 续费，含全部升级与服务。",
        4)


# ============================================================
# Sheet 6：折扣与续费
# ============================================================
def build_discount(wb):
    ws = wb.create_sheet("05 折扣与续费")
    set_widths(ws, [22, 22, 60])
    r = 1
    title_row(ws, r, "折扣与续费规则（v2 红线版）", 3); r += 1

    section_title(ws, r, "5.1 用户数阶梯折扣（适用于一次性 + 订阅）", 3); r += 1
    head_row(ws, r, ["用户数", "折扣", ""]); r += 1
    r = data_rows(ws, r, [(a, b, "") for a, b in D.USER_TIER_DISCOUNT])

    r += 1
    section_title(ws, r, "5.2 多年付折扣（仅订阅版）", 3); r += 1
    head_row(ws, r, ["付费方式", "折扣", "附加权益"]); r += 1
    r = data_rows(ws, r, D.MULTI_YEAR_DISCOUNT)

    r += 1
    section_title(ws, r, "5.3 折扣红线（v2 新增，最关键）", 3); r += 1
    head_row(ws, r, ["规则", "内容", ""]); r += 1
    for rule, desc in D.DISCOUNT_REDLINE:
        ws.cell(row=r, column=1, value=rule).font = Font(name="DengXian", size=11, bold=True, color="C00000")
        ws.cell(row=r, column=1).alignment = WRAP_LEFT
        ws.cell(row=r, column=1).border = BORDER
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=2, value=desc)
        c.font = CELL_FONT
        c.alignment = WRAP_LEFT
        c.border = BORDER
        ws.row_dimensions[r].height = 28
        r += 1


# ============================================================
# Sheet 7：裂变四件套
# ============================================================
def build_fission(wb):
    ws = wb.create_sheet("06 裂变四件套")
    set_widths(ws, [22, 80])
    r = 1
    title_row(ws, r, "客户自我裂变机制 v2", 2); r += 1
    head_row(ws, r, ["机制", "规则"]); r += 1
    r = data_rows(ws, r, D.FISSION_KIT)
    r += 1
    note_row(ws, r,
             "执行口径：返现走『市场费用-推荐返佣』科目；案例升级由市场部审批；积分由货袋子产品自动结算；峰会门票券由销售总监统一发放。",
             2)


# ============================================================
# Sheet 8：销售话术索引
# ============================================================
def build_scenes_index(wb):
    ws = wb.create_sheet("07 销售话术索引")
    set_widths(ws, [10, 38, 22, 38])
    title_row(ws, 1, "12 场景脚本索引（详见《销售作战手册（一线版）》）", 4)
    head_row(ws, 2, ["编号", "场景", "适用客群", "首选报价"])
    data_rows(ws, 3, D.SALES_SCENES)


# ============================================================
# Sheet 9：v1 → v2 修订对照
# ============================================================
def build_diff(wb):
    ws = wb.create_sheet("08 v1 v2 修订对照")
    set_widths(ws, [6, 18, 38, 38])
    title_row(ws, 1, "v1 → v2 修订对照（共 15 处）", 4)
    head_row(ws, 2, ["#", "条目", "v1 原版", "v2 修订"])
    end = data_rows(ws, 3, D.V1_V2_DIFF)
    note_row(ws, end, "本次修订未改动钢贸宝 ERP / 提货宝 / 对账宝原档位价格，主要修复留白、上调价值/价格倒挂部分、并新增折扣红线与裂变机制。", 4)


# ============================================================
# Sheet 10：内部毛利红线
# ============================================================
def build_redline(wb):
    ws = wb.create_sheet("09 内部毛利红线")
    set_widths(ws, [24, 22, 22, 22, 22, 38])
    r = 1
    title_row(ws, r, "内部毛利红线（销售总监 / 总裁审批用）", 6); r += 1
    head_row(ws, r, ["套餐 / 单品", "目录价", "65 折底线", "销售可签上限", "总裁可签上限", "审批要求"]); r += 1

    rows = [
        ("中型·标准包",   19800,  12870, 13800, 12870, "≤13,800 销售签；≤12,870 总裁签；低于此价不签"),
        ("大型·专业包",   58000,  37700, 40600, 37700, "同上"),
        ("超大·集团包",   158000, 102700, 110000, 102700, "另需私有化交付预算评估"),
        ("工业全家桶·中型 ⭐", 93600, 60840, 65000, 60840, "需对齐用户数 + 产线数双口径"),
        ("钢企通 M7 高端", 11800, 7670,  8260,  7670, "按用户数线性折后再叠加红线"),
        ("MES 标准版",    88000, 57200, 61600, 57200, "需含至少 1 次现场调试"),
    ]
    for row in rows:
        for ci, val in enumerate(row, 1):
            c = ws.cell(row=r, column=ci, value=val)
            c.font = CELL_FONT
            c.alignment = WRAP_CENTER if ci != 6 else WRAP_LEFT
            c.border = BORDER
            if ci in (2, 3, 4, 5):
                c.number_format = "¥#,##0"
        ws.row_dimensions[r].height = 24
        r += 1
    r += 1
    note_row(ws, r,
             "本表为内部使用；任何低于 65 折成交需在审批单中明确：①客户体量 ②竞争对手报价 ③战略价值；私有化与定制开发不在此红线内，单独走项目报备。",
             6)


def main():
    wb = Workbook()
    build_cover(wb)
    build_tier(wb)
    build_product_map(wb)
    build_skus(wb)
    build_bundles(wb)
    build_discount(wb)
    build_fission(wb)
    build_scenes_index(wb)
    build_diff(wb)
    build_redline(wb)

    out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "deliverables", "产品价格白皮书_v2.xlsx")
    wb.save(out)
    print(f"[OK] {out}")


if __name__ == "__main__":
    main()
