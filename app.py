from flask import Flask, render_template, request, jsonify, send_file
from llm_parser import parse_plan
from normalizer import normalize
from rule_engine import calculate_total, PHASE_RATIOS
from config_loader import load_config
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.chart.series import DataPoint
import io
import json
import os
from datetime import datetime

app = Flask(__name__)
CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config_params.json")


# ─── 样式常量 ────────────────────────────────────────────────

C_BLUE_DARK   = "1A3C6E"   # 深蓝  - 主标题
C_BLUE        = "1a56db"   # 品牌蓝 - 表头
C_BLUE_LIGHT  = "D6E4FF"   # 浅蓝  - 交替行
C_ORANGE      = "E67E22"   # 橙色  - 重点数字
C_GREEN       = "1E8449"   # 绿色  - 小计/合计
C_GRAY        = "F5F7FA"   # 浅灰  - 背景
C_WHITE       = "FFFFFF"
C_TEXT        = "1A1A2E"
C_GOLD        = "F39C12"   # 金色  - 特别提示


def _side(color="CCCCCC", style="thin"):
    return Side(style=style, color=color)


def _border(color="CCCCCC"):
    s = _side(color)
    return Border(left=s, right=s, top=s, bottom=s)


def _header_border():
    s = _side("FFFFFF", "medium")
    return Border(left=s, right=s, top=s, bottom=s)


def _font(bold=False, size=11, color=C_TEXT, name="微软雅黑"):
    return Font(bold=bold, size=size, color=color, name=name)


def _fill(color):
    return PatternFill("solid", fgColor=color)


def _center(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)


def _left(wrap=False):
    return Alignment(horizontal="left", vertical="center", wrap_text=wrap)


def _right():
    return Alignment(horizontal="right", vertical="center")


def _num(val, decimals=2):
    return round(float(val), decimals)


def _merge_style(ws, cell_range, value="", bold=False, size=11,
                 color=C_TEXT, bg=None, align="center", wrap=False):
    ws.merge_cells(cell_range)
    top_left = ws[cell_range.split(":")[0]]
    top_left.value = value
    top_left.font = _font(bold=bold, size=size, color=color)
    if bg:
        top_left.fill = _fill(bg)
    a = _center(wrap) if align == "center" else _left(wrap) if align == "left" else _right()
    top_left.alignment = a
    return top_left


def _row_bg(ws, row, start_col, end_col, color):
    for c in range(start_col, end_col + 1):
        ws.cell(row, c).fill = _fill(color)


def _apply_header_row(ws, row_idx, headers, fills=None):
    """渲染表头行，fills 为每列背景色列表（可选）"""
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=h)
        cell.font = _font(bold=True, size=10, color=C_WHITE)
        cell.fill = _fill(fills[col_idx - 1] if fills else C_BLUE)
        cell.alignment = _center(wrap=True)
        cell.border = _header_border()


def _apply_data_row(ws, row_idx, values, alt=False, bold=False, color=None):
    bg = C_BLUE_LIGHT if alt else C_WHITE
    for col_idx, v in enumerate(values, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=v)
        cell.font = _font(bold=bold, size=10, color=color or C_TEXT)
        cell.fill = _fill(bg)
        cell.alignment = _center() if isinstance(v, (int, float)) else _left()
        cell.border = _border()


# ─── 路由 ────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/config", methods=["GET"])
def get_config():
    return jsonify(load_config())


@app.route("/config", methods=["POST"])
def save_config():
    data = request.json
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return jsonify({"ok": True})


@app.route("/parse", methods=["POST"])
def parse():
    data = request.json
    result = parse_plan(data.get("md", ""))
    return jsonify({"result": result})


@app.route("/batch_evaluate", methods=["POST"])
def batch_evaluate():
    req = request.json
    files = req.get("files", [])
    config = load_config()
    results = []

    for f in files:
        name = f.get("name", "未命名")
        content = f.get("content", "")
        try:
            parsed_str = parse_plan(content)
            cleaned = parsed_str.strip()
            if cleaned.startswith("```"):
                lines = cleaned.split("\n")
                cleaned = "\n".join(lines[1:])
            if cleaned.endswith("```"):
                cleaned = cleaned.rsplit("```", 1)[0]
            parsed = json.loads(cleaned)
            normalized = normalize(parsed)
            evaluated = calculate_total(normalized, config)
            results.append({"name": name, "status": "ok", "data": evaluated})
        except Exception as e:
            results.append({"name": name, "status": "error", "error": str(e)})

    return jsonify({"results": results})


@app.route("/export_excel", methods=["POST"])
def export_excel():
    data = request.json
    file_stream = generate_excel(data)
    return send_file(
        file_stream,
        as_attachment=True,
        download_name="信创适配费用评估报告.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ─── Excel 生成 ───────────────────────────────────────────────

def generate_excel(batch_results):
    wb = Workbook()
    ok_results = [r for r in batch_results if r.get("status") == "ok"]

    _sheet_cover(wb, ok_results)
    _sheet_summary(wb, ok_results)
    _sheet_cost_breakdown(wb, ok_results)
    _sheet_phase(wb, ok_results)
    for r in ok_results:
        _sheet_system_detail(wb, r)
    _sheet_params(wb)

    # 删除默认 Sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream


def _sheet_cover(wb, ok_results):
    """封面页"""
    ws = wb.create_sheet("封面", 0)
    ws.sheet_view.showGridLines = False

    # 行高列宽
    for r in range(1, 40):
        ws.row_dimensions[r].height = 22
    for c in range(1, 12):
        ws.column_dimensions[get_column_letter(c)].width = 14

    # 顶部色块
    for r in range(1, 8):
        _row_bg(ws, r, 1, 11, C_BLUE_DARK)

    _merge_style(ws, "B2:J3", "国产化适配改造费用评估报告",
                 bold=True, size=22, color=C_WHITE, align="center")
    _merge_style(ws, "B4:J4", "Localization Adaptation Cost Assessment Report",
                 bold=False, size=12, color="A8C4E8", align="center")

    # 分隔线
    for r in range(8, 10):
        _row_bg(ws, r, 1, 11, C_GRAY)

    # 基本信息区
    total_projects = len(ok_results)
    total_cost = sum(r["data"]["cost"]["total_incl_tax"] for r in ok_results)
    total_work = sum(r["data"]["total_workload"] for r in ok_results)
    total_systems = sum(len(r["data"]["systems"]) for r in ok_results)

    info_items = [
        ("评估项目数", str(total_projects)),
        ("涉及子系统数", str(total_systems)),
        ("合计工作量（人天）", f"{round(total_work, 1)}"),
        ("合计费用（含税）", f"¥ {total_cost:,.2f}"),
        ("报告生成日期", datetime.now().strftime("%Y年%m月%d日")),
        ("评估依据", "GBT 16680、信创适配指导规范"),
    ]

    for i, (label, val) in enumerate(info_items):
        row = 11 + i * 2
        _merge_style(ws, f"B{row}:D{row}", label, bold=True, size=11,
                     color=C_BLUE_DARK, align="left")
        _merge_style(ws, f"E{row}:J{row}", val, bold=False, size=11,
                     color=C_TEXT, align="left")

    # 技术栈对照（汇总所有项目）
    all_existing = {}
    all_target = {}
    for r in ok_results:
        d = r["data"]
        for k, v in d.get("existing_tech_stack", {}).items():
            if v and k not in all_existing:
                all_existing[k] = v
        for k, v in d.get("target_tech_stack", {}).items():
            if v and k not in all_target:
                all_target[k] = v

    if all_existing or all_target:
        tech_start = 24
        _merge_style(ws, f"B{tech_start}:J{tech_start}",
                     "技术栈改造对照",
                     bold=True, size=11, color=C_WHITE, bg=C_BLUE, align="center")
        tech_start += 1
        _merge_style(ws, f"B{tech_start}:E{tech_start}", "改造维度",
                     bold=True, size=10, color=C_WHITE, bg=C_BLUE_DARK, align="center")
        _merge_style(ws, f"F{tech_start}:G{tech_start}", "现有技术栈",
                     bold=True, size=10, color=C_WHITE, bg=C_BLUE_DARK, align="center")
        _merge_style(ws, f"H{tech_start}:J{tech_start}", "改造目标（信创）",
                     bold=True, size=10, color=C_WHITE, bg=C_BLUE_DARK, align="center")
        tech_start += 1
        label_map = {"database": "数据库", "os": "操作系统", "middleware": "中间件",
                     "hardware": "硬件平台", "frontend": "前端框架", "etl": "ETL平台"}
        all_keys = list(dict.fromkeys(list(all_existing.keys()) + list(all_target.keys())))
        for ki, key in enumerate(all_keys):
            bg = C_BLUE_LIGHT if ki % 2 == 0 else C_WHITE
            label = label_map.get(key, key)
            _merge_style(ws, f"B{tech_start}:E{tech_start}", label, bg=bg, align="left")
            _merge_style(ws, f"F{tech_start}:G{tech_start}", all_existing.get(key, "—"), bg=bg, align="left")
            _merge_style(ws, f"H{tech_start}:J{tech_start}", all_target.get(key, "—"), bg=bg, color="1E8449", align="left")
            tech_start += 1

    # 免责声明
    note_row = max(tech_start + 1, 34)
    _merge_style(ws, f"B{note_row}:J{note_row + 2}",
                 "【说明】本报告依据建设方案文档，结合信创改造工作量模型自动生成，"
                 "仅作为项目预算参考依据，最终费用以正式合同及详细方案为准。",
                 bold=False, size=9, color="888888", bg=C_GRAY, align="left", wrap=True)


def _sheet_summary(wb, ok_results):
    """汇总表"""
    ws = wb.create_sheet("费用汇总")
    ws.sheet_view.showGridLines = False

    col_widths = [6, 28, 10, 10, 14, 16, 14, 12, 14, 16, 16]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 16
    ws.row_dimensions[2].height = 36
    ws.row_dimensions[3].height = 36

    # 大标题
    _merge_style(ws, "A1:K1", "国产化适配费用评估汇总表",
                 bold=True, size=14, color=C_WHITE, bg=C_BLUE_DARK, align="center")

    headers = [
        "序号", "项目/文件名称", "子系统数", "总工作量\n(人天)",
        "人工费(元)", "测试费(元)", "小计(元)",
        "项目管理费(元)", "风险费(元)", "不含税合计(元)", "含税合计(元)"
    ]
    _apply_header_row(ws, 2, headers)

    grand = {k: 0.0 for k in ["total_workload", "labor_cost", "test_cost",
                                "subtotal", "management_cost", "risk_cost",
                                "total_excl_tax", "total_incl_tax"]}

    for idx, r in enumerate(ok_results, 1):
        d = r["data"]
        c = d["cost"]
        alt = (idx % 2 == 0)
        row = [
            idx, r["name"], len(d["systems"]),
            _num(d["total_workload"]),
            _num(c["labor_cost"]), _num(c["test_cost"]), _num(c["subtotal"]),
            _num(c["management_cost"]), _num(c["risk_cost"]),
            _num(c["total_excl_tax"]), _num(c["total_incl_tax"])
        ]
        _apply_data_row(ws, 3 + idx - 1, row, alt=alt)
        for k in grand:
            grand[k] += d.get(k, 0) if k == "total_workload" else c.get(k, 0)

    # 合计行
    total_row = 3 + len(ok_results)
    _row_bg(ws, total_row, 1, 11, C_BLUE_DARK)
    total_values = [
        "", "合 计", "", _num(grand["total_workload"]),
        _num(grand["labor_cost"]), _num(grand["test_cost"]), _num(grand["subtotal"]),
        _num(grand["management_cost"]), _num(grand["risk_cost"]),
        _num(grand["total_excl_tax"]), _num(grand["total_incl_tax"])
    ]
    for col_idx, v in enumerate(total_values, 1):
        cell = ws.cell(row=total_row, column=col_idx, value=v)
        cell.font = _font(bold=True, size=11, color=C_WHITE)
        cell.fill = _fill(C_BLUE_DARK)
        cell.alignment = _center()
        cell.border = _header_border()

    # 大写金额提示
    note_row = total_row + 2
    _merge_style(ws, f"A{note_row}:K{note_row}",
                 f"含税合计（大写）：{_amount_cn(grand['total_incl_tax'])}",
                 bold=True, size=11, color=C_ORANGE, bg="FFF8E7", align="left")


def _sheet_cost_breakdown(wb, ok_results):
    """费用构成说明表"""
    ws = wb.create_sheet("费用构成说明")
    ws.sheet_view.showGridLines = False

    col_widths = [6, 24, 20, 20, 14, 14, 14, 14, 16, 16]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 16

    _merge_style(ws, "A1:J1", "费用构成明细说明",
                 bold=True, size=14, color=C_WHITE, bg=C_BLUE_DARK, align="center")

    headers = [
        "序号", "系统名称", "改造类型", "主要技术栈替换",
        "基础工作量\n(人天)", "改造附加量\n(人天)", "小计工作量\n(人天)",
        "级别系数", "复杂度系数", "最终工作量\n(人天)"
    ]
    _apply_header_row(ws, 2, headers)

    row_idx = 3
    seq = 1
    for r in ok_results:
        for sys in r["data"]["systems"]:
            alt = (seq % 2 == 0)
            adaptation_str = "、".join(sys.get("adaptation_type", [])) or "—"
            tech_str = "、".join(sys.get("tech_stack", [])) or "—"
            base = sys.get("base_work", 0)
            adapt = sys.get("adaptation_work", 0)
            subtotal_w = round(base + adapt, 2)
            row = [
                seq, sys["name"], adaptation_str, tech_str,
                _num(base), _num(adapt), _num(subtotal_w),
                sys.get("level_factor", 1.0), sys.get("complexity", 1.0),
                _num(sys["total_work"])
            ]
            _apply_data_row(ws, row_idx, row, alt=alt)
            row_idx += 1
            seq += 1

    # 说明区
    row_idx += 1
    notes = [
        "【级别系数说明】核心系统×1.3，重要系统×1.1，一般系统×1.0",
        "【复杂度系数说明】基础1.0，模块>20增加0.15，接口>30增加0.10，改造类型≥4增加0.20，≥2增加0.10（可叠加）",
        "【改造附加量说明】数据库替换+80人天，ETL平台适配+60人天，数据迁移+40人天，硬件架构适配+25人天，操作系统适配+20人天，安全加固+15人天，中间件替换+15人天，前端框架适配+12人天，接口改造+8人天/系统；外部对接接口按接口数×8人天单独计入项目级工作量",
    ]
    for note in notes:
        _merge_style(ws, f"A{row_idx}:J{row_idx}", note,
                     bold=False, size=9, color="555555", bg=C_GRAY,
                     align="left", wrap=True)
        ws.row_dimensions[row_idx].height = 18
        row_idx += 1


def _sheet_phase(wb, ok_results):
    """工时阶段分解表"""
    ws = wb.create_sheet("工时阶段分解")
    ws.sheet_view.showGridLines = False

    phases = list(PHASE_RATIOS.keys())
    col_widths = [6, 28] + [14] * len(phases) + [16]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 16

    _merge_style(ws, f"A1:{get_column_letter(len(col_widths))}1",
                 "工时阶段分解表（各阶段工作量单位：人天）",
                 bold=True, size=14, color=C_WHITE, bg=C_BLUE_DARK, align="center")

    headers = ["序号", "系统名称"] + phases + ["合计(人天)"]
    phase_colors = [C_BLUE, "1565C0", "0D47A1", "1976D2", "1565C0"]
    fills_h = [C_BLUE_DARK, C_BLUE_DARK] + phase_colors + [C_GREEN.replace("1E8449", "1B5E20")]
    _apply_header_row(ws, 2, headers, fills=[C_BLUE_DARK] * len(headers))

    row_idx = 3
    seq = 1
    for r in ok_results:
        for sys in r["data"]["systems"]:
            alt = (seq % 2 == 0)
            phase_vals = [_num(sys.get("phases", {}).get(p, 0)) for p in phases]
            row = [seq, sys["name"]] + phase_vals + [_num(sys["total_work"])]
            _apply_data_row(ws, row_idx, row, alt=alt)
            row_idx += 1
            seq += 1

    # 合计行
    _row_bg(ws, row_idx, 1, len(headers), C_GREEN)
    ws.cell(row_idx, 1, "").fill = _fill(C_GREEN)
    ws.cell(row_idx, 2).value = "合 计"
    ws.cell(row_idx, 2).font = _font(bold=True, color=C_WHITE)
    ws.cell(row_idx, 2).fill = _fill(C_GREEN)
    ws.cell(row_idx, 2).alignment = _center()
    for pi, phase in enumerate(phases):
        total_phase = sum(
            sys.get("phases", {}).get(phase, 0)
            for r in ok_results for sys in r["data"]["systems"]
        )
        cell = ws.cell(row_idx, 3 + pi, round(total_phase, 2))
        cell.font = _font(bold=True, color=C_WHITE)
        cell.fill = _fill(C_GREEN)
        cell.alignment = _center()

    total_all = sum(r["data"]["total_workload"] for r in ok_results)
    total_cell = ws.cell(row_idx, len(headers), round(total_all, 2))
    total_cell.font = _font(bold=True, color=C_WHITE)
    total_cell.fill = _fill(C_GREEN)
    total_cell.alignment = _center()

    # 比例说明
    row_idx += 2
    _merge_style(ws, f"A{row_idx}:{get_column_letter(len(headers))}{row_idx}",
                 "阶段比例：调研分析15% / 环境搭建10% / 代码适配45% / 测试验证20% / 上线部署10%",
                 bold=False, size=9, color="555555", bg=C_GRAY, align="left")


def _sheet_system_detail(wb, r):
    """每个项目的系统详细信息 Sheet"""
    sheet_name = r["name"].replace(".md", "")[:28] + "-明细"
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    d = r["data"]
    project_name = d.get("project_name", r["name"])
    background = d.get("project_background", "")

    col_widths = [6, 22, 8, 8, 8, 10, 12, 12, 10, 10, 10, 10, 14, 14, 14, 14]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # 项目标题
    _merge_style(ws, f"A1:{get_column_letter(len(col_widths))}1",
                 f"项目名称：{project_name}",
                 bold=True, size=13, color=C_WHITE, bg=C_BLUE_DARK, align="left")
    if background:
        _merge_style(ws, f"A2:{get_column_letter(len(col_widths))}2",
                     f"项目背景：{background}",
                     bold=False, size=9, color="444444", bg=C_GRAY,
                     align="left", wrap=True)
        ws.row_dimensions[2].height = 30

    # 技术改造说明
    _merge_style(ws, f"A3:{get_column_letter(len(col_widths))}3",
                 "技术改造难点说明",
                 bold=True, size=11, color=C_WHITE, bg=C_BLUE, align="center")
    row_idx = 4
    for idx, sys in enumerate(d["systems"], 1):
        note = sys.get("complexity_note", "")
        if note:
            _merge_style(ws, f"A{row_idx}:{get_column_letter(len(col_widths))}{row_idx}",
                         f"  {idx}. {sys['name']}：{note}",
                         bold=False, size=9, color=C_TEXT, bg=C_BLUE_LIGHT,
                         align="left", wrap=True)
            ws.row_dimensions[row_idx].height = 20
            row_idx += 1

    # 系统明细表头
    headers = [
        "序号", "系统名称", "级别", "模块数", "接口数", "数据表数",
        "数据量(GB)", "改造类型",
        "基础工作量", "改造附加量", "级别系数", "复杂度系数",
        "人工费(元)", "测试费(元)", "不含税(元)", "含税(元)"
    ]
    _apply_header_row(ws, row_idx, headers)
    row_idx += 1

    config = load_config()
    price_per_day = config.get("price_per_day", 1800)
    test_rate = config.get("test_rate", 0.12)
    management_rate = config.get("management_rate", 0.12)
    risk_rate = config.get("risk_rate", 0.08)
    tax_rate = config.get("tax_rate", 0.09)

    for idx, sys in enumerate(d["systems"], 1):
        labor = sys["total_work"] * price_per_day
        test_c = labor * test_rate
        sub = labor + test_c
        mgmt = sub * management_rate
        risk = sub * risk_rate
        excl = sub + mgmt + risk
        incl = excl * (1 + tax_rate)

        adaptation_str = "、".join(sys.get("adaptation_type", [])) or "—"
        alt = (idx % 2 == 0)
        row = [
            idx, sys["name"], sys.get("level", "—"),
            sys.get("modules", 0), sys.get("interfaces", 0), sys.get("databases", 0),
            _num(sys.get("data_size_gb", 0)),
            adaptation_str,
            _num(sys.get("base_work", 0)), _num(sys.get("adaptation_work", 0)),
            sys.get("level_factor", 1.0), sys.get("complexity", 1.0),
            _num(labor), _num(test_c), _num(excl), _num(incl)
        ]
        _apply_data_row(ws, row_idx, row, alt=alt)
        row_idx += 1

    # 费用汇总块
    row_idx += 1
    cost = d["cost"]
    cost_items = [
        ("人工费（工作量 × 日费率）", cost["labor_cost"]),
        ("测试费（人工费 × 测试费率）", cost["test_cost"]),
        ("人工+测试小计", cost["subtotal"]),
        ("项目管理费（小计 × 管理费率）", cost["management_cost"]),
        ("风险费（小计 × 风险费率）", cost["risk_cost"]),
        ("不含税合计", cost["total_excl_tax"]),
        ("含税合计（税率9%）", cost["total_incl_tax"]),
    ]
    _merge_style(ws, f"A{row_idx}:{get_column_letter(len(col_widths))}{row_idx}",
                 "费用汇总",
                 bold=True, size=11, color=C_WHITE, bg=C_BLUE, align="center")
    row_idx += 1
    for label, val in cost_items:
        is_total = "合计" in label
        bg = C_BLUE_DARK if is_total else (C_GRAY if "小计" in label else C_WHITE)
        fc = C_WHITE if is_total else (C_GREEN if "小计" in label else C_TEXT)
        _merge_style(ws, f"A{row_idx}:M{row_idx}", f"  {label}",
                     bold=is_total, color=fc, bg=bg, align="left")
        cell = ws.cell(row_idx, 14)
        cell.value = f"¥ {val:,.2f}"
        cell.font = _font(bold=is_total, color=C_ORANGE if is_total else fc)
        cell.fill = _fill(bg)
        cell.alignment = _right()
        _merge_style(ws, f"O{row_idx}:{get_column_letter(len(col_widths))}{row_idx}",
                     "", bg=bg)
        row_idx += 1

    # 大写金额
    _merge_style(ws, f"A{row_idx}:{get_column_letter(len(col_widths))}{row_idx}",
                 f"含税合计（大写）：{_amount_cn(cost['total_incl_tax'])}",
                 bold=True, size=11, color=C_ORANGE, bg="FFF8E7", align="left")


def _sheet_params(wb):
    """评估参数说明 Sheet"""
    config = load_config()
    ws = wb.create_sheet("评估参数说明")
    ws.sheet_view.showGridLines = False

    for i, w in enumerate([6, 28, 16, 40], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    _merge_style(ws, "A1:D1", "评估参数说明",
                 bold=True, size=14, color=C_WHITE, bg=C_BLUE_DARK, align="center")
    _apply_header_row(ws, 2, ["序号", "参数名称", "当前值", "说明"])

    params = [
        ("模块工作量系数 module_factor",
         config.get("module_factor", 2.0),
         "每个功能模块的基础工作量（人天），含功能开发+单元测试"),
        ("接口工作量系数 interface_factor",
         config.get("interface_factor", 0.5),
         "每个接口的改造工作量（人天）"),
        ("数据表工作量系数 table_factor",
         config.get("table_factor", 0.3),
         "每张数据库表的迁移/改造工作量（人天）"),
        ("数据量系数 data_factor",
         config.get("data_factor", 0.005),
         "每GB数据的迁移工作量（人天）"),
        ("人天费率 price_per_day",
         config.get("price_per_day", 2200),
         "每人天综合费率（元），含人员工资、社保、管理分摊；信创项目人才紧缺，建议2000~2500元"),
        ("测试费率 test_rate",
         f"{config.get('test_rate', 0.12)*100:.0f}%",
         "测试费 = 人工费 × 测试费率，含功能测试、兼容性测试"),
        ("项目管理费率 management_rate",
         f"{config.get('management_rate', 0.12)*100:.0f}%",
         "项目管理费 = (人工费+测试费) × 管理费率"),
        ("风险费率 risk_rate",
         f"{config.get('risk_rate', 0.08)*100:.0f}%",
         "风险不可预见费 = (人工费+测试费) × 风险费率"),
        ("增值税率 tax_rate",
         f"{config.get('tax_rate', 0.09)*100:.0f}%",
         "含税价 = 不含税价 × (1 + 税率)，信息服务增值税率"),
    ]
    adapt_factors = config.get("adaptation_factors", {})
    for k, v in adapt_factors.items():
        params.append((f"改造附加量-{k}", v, f"包含[{k}]改造时，每个系统额外增加的工作量（人天）"))

    for idx, (name, val, desc) in enumerate(params, 1):
        alt = (idx % 2 == 0)
        _apply_data_row(ws, 2 + idx, [idx, name, val, desc], alt=alt)


def _amount_cn(amount):
    """金额转中文大写（简化版）"""
    units = ["", "万", "亿"]
    amount = round(amount, 2)
    if amount >= 1_0000_0000:
        return f"人民币{amount/1_0000_0000:.4f}亿元整"
    elif amount >= 1_0000:
        wan = amount / 10000
        return f"人民币{wan:.2f}万元整"
    else:
        return f"人民币{amount:.2f}元整"


if __name__ == "__main__":
    app.run(debug=True)
