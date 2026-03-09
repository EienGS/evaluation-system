from flask import Flask, render_template, request, jsonify, send_file
from llm_parser import parse_plan
from normalizer import normalize
from rule_engine import calculate_total
from config_loader import load_config
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import json
import os

app = Flask(__name__)

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config_params.json")


@app.route("/")
def index():
    return render_template("index.html")


# ---------- 参数配置接口 ----------

@app.route("/config", methods=["GET"])
def get_config():
    config = load_config()
    return jsonify(config)


@app.route("/config", methods=["POST"])
def save_config():
    data = request.json
    allowed_keys = {
        "module_factor", "interface_factor", "table_factor",
        "data_factor", "user_factor", "price_per_day",
        "management_rate", "risk_rate"
    }
    filtered = {k: v for k, v in data.items() if k in allowed_keys}
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(filtered, f, ensure_ascii=False, indent=2)
    return jsonify({"ok": True})


# ---------- 单文件解析 ----------

@app.route("/parse", methods=["POST"])
def parse():
    data = request.json
    md_text = data.get("md")
    result = parse_plan(md_text)
    return jsonify({"result": result})


# ---------- 批量处理：解析 + 评估一体化 ----------

@app.route("/batch_evaluate", methods=["POST"])
def batch_evaluate():
    """
    接收多个MD文件内容，依次解析+评估，出错跳过。
    请求体: { "files": [{"name": "xxx.md", "content": "..."}, ...] }
    返回:   { "results": [{"name":"xxx.md","status":"ok","data":{...}}, ...] }
    """
    req = request.json
    files = req.get("files", [])
    config = load_config()
    results = []

    for f in files:
        name = f.get("name", "未命名")
        content = f.get("content", "")
        try:
            parsed_str = parse_plan(content)
            # 清理 LLM 返回的 markdown 代码块标记
            cleaned = parsed_str.strip()
            if cleaned.startswith("```"):
                cleaned = cleaned.split("\n", 1)[1] if "\n" in cleaned else cleaned
            if cleaned.endswith("```"):
                cleaned = cleaned.rsplit("```", 1)[0]
            parsed = json.loads(cleaned)
            normalized = normalize(parsed)
            evaluated = calculate_total(normalized, config)
            results.append({
                "name": name,
                "status": "ok",
                "data": evaluated
            })
        except Exception as e:
            results.append({
                "name": name,
                "status": "error",
                "error": str(e)
            })

    return jsonify({"results": results})


# ---------- 导出 Excel ----------

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


def generate_excel(batch_results):
    wb = Workbook()

    # ---- 汇总Sheet ----
    ws_summary = wb.active
    ws_summary.title = "汇总"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1a56db")
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    summary_headers = ["文件名", "系统数", "总工作量(人天)", "开发费用(元)", "管理费用(元)", "风险费用(元)", "合计费用(元)"]
    ws_summary.append(summary_headers)
    for col_idx, _ in enumerate(summary_headers, 1):
        cell = ws_summary.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    grand_total_workload = 0
    grand_total_cost = 0

    ok_results = [r for r in batch_results if r["status"] == "ok"]

    for r in ok_results:
        d = r["data"]
        cost = d.get("cost", {})
        row = [
            r["name"],
            len(d.get("systems", [])),
            round(d.get("total_workload", 0), 2),
            round(cost.get("dev_cost", 0), 2),
            round(cost.get("management_cost", 0), 2),
            round(cost.get("risk_cost", 0), 2),
            round(cost.get("total_cost", 0), 2),
        ]
        ws_summary.append(row)
        grand_total_workload += d.get("total_workload", 0)
        grand_total_cost += cost.get("total_cost", 0)
        for col_idx in range(1, len(row) + 1):
            ws_summary.cell(row=ws_summary.max_row, column=col_idx).border = border

    # 合计行
    total_row_idx = ws_summary.max_row + 1
    ws_summary.cell(total_row_idx, 1, "合计").font = Font(bold=True)
    ws_summary.cell(total_row_idx, 3, round(grand_total_workload, 2)).font = Font(bold=True)
    ws_summary.cell(total_row_idx, 7, round(grand_total_cost, 2)).font = Font(bold=True)

    # 列宽
    col_widths = [30, 10, 18, 18, 18, 18, 18]
    for i, w in enumerate(col_widths, 1):
        ws_summary.column_dimensions[ws_summary.cell(1, i).column_letter].width = w

    # ---- 各文件明细Sheet ----
    detail_headers = [
        "系统名称", "模块数", "接口数", "数据表数",
        "模块工作量", "接口工作量", "数据库工作量", "数据工作量", "用户工作量",
        "复杂度系数", "总工作量(人天)"
    ]

    for r in ok_results:
        sheet_name = r["name"].replace(".md", "")[:31]
        ws = wb.create_sheet(title=sheet_name)
        ws.append(detail_headers)
        for col_idx, _ in enumerate(detail_headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border

        for sys in r["data"].get("systems", []):
            row = [
                sys.get("name", ""),
                sys.get("modules", 0),
                sys.get("interfaces", 0),
                sys.get("databases", 0),
                round(sys.get("module_work", 0), 2),
                round(sys.get("interface_work", 0), 2),
                round(sys.get("db_work", 0), 2),
                round(sys.get("data_work", 0), 2),
                round(sys.get("user_work", 0), 2),
                sys.get("complexity", 1),
                round(sys.get("total_work", 0), 2),
            ]
            ws.append(row)
            for col_idx in range(1, len(row) + 1):
                ws.cell(row=ws.max_row, column=col_idx).border = border

        cost = r["data"].get("cost", {})
        ws.append([])
        ws.append(["开发费用(元)", "", "", "", "", "", "", "", "", "", round(cost.get("dev_cost", 0), 2)])
        ws.append(["管理费用(元)", "", "", "", "", "", "", "", "", "", round(cost.get("management_cost", 0), 2)])
        ws.append(["风险费用(元)", "", "", "", "", "", "", "", "", "", round(cost.get("risk_cost", 0), 2)])
        ws.append(["合计费用(元)", "", "", "", "", "", "", "", "", "", round(cost.get("total_cost", 0), 2)])

        for i, w in enumerate([20, 8, 8, 8, 14, 14, 14, 14, 14, 12, 16], 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = w

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream


if __name__ == "__main__":
    app.run(debug=True)
