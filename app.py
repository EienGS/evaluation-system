from flask import Flask, render_template, request, jsonify, send_file
from llm_parser import parse_plan
from normalizer import normalize
from rule_engine import calculate_total
from config_loader import load_config
from docx_converter import docx_to_markdown
from openpyxl import Workbook
import io
import json
import os

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config_params.json")

app = Flask(__name__)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/evaluate_file", methods=["POST"])
def evaluate_file():
    """接收单个 md 或 docx 文件（multipart），串联解析+评估，返回完整结果"""
    uploaded = request.files.get("file")
    if not uploaded:
        return jsonify({"success": False, "error": "未收到文件"}), 200

    filename = uploaded.filename or "未知文件"
    ext = os.path.splitext(filename)[1].lower()

    try:
        raw_bytes = uploaded.read()
        if ext in (".docx",):
            md_text = docx_to_markdown(raw_bytes)
        else:
            md_text = raw_bytes.decode("utf-8", errors="replace")

        parsed_str = parse_plan(md_text)
        parsed = json.loads(parsed_str)
        normalized = normalize(parsed)
        config = load_config()
        result = calculate_total(normalized, config)
        result["filename"] = filename
        return jsonify({"success": True, "data": result})
    except Exception as e:
        return jsonify({"success": False, "filename": filename, "error": str(e)}), 200


@app.route("/config", methods=["GET"])
def get_config():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            config = json.load(f)
        return jsonify(config)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/config", methods=["POST"])
def save_config():
    try:
        new_config = request.json
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(new_config, f, ensure_ascii=False, indent=2)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/export_excel", methods=["POST"])
def export_excel():
    data = request.json
    file_stream = generate_excel(data)
    return send_file(
        file_stream,
        as_attachment=True,
        download_name="国产化适配费用评估.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def generate_excel(results):
    """
    results: list of evaluate_file 返回的 data 对象
    每个对象包含 filename, systems[], total_workload, total_cost 等
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "费用评估汇总"

    ws.append([
        "文件名",
        "系统名称",
        "模块工作量(人天)",
        "接口工作量(人天)",
        "数据库工作量(人天)",
        "数据对接工作量(人天)",
        "权限工作量(人天)",
        "复杂度系数",
        "系统工作量(人天)",
        "人工费(元)",
        "管理费(元)",
        "风险费(元)",
        "测试费(元)",
        "小计(元)",
        "不含税总价(元)",
        "含税总价(元)"
    ])

    for item in results:
        filename = item.get("filename", "")
        cost = item.get("cost", {})
        for sys in item.get("systems", []):
            ws.append([
                filename,
                sys.get("name", ""),
                sys.get("module_work", ""),
                sys.get("interface_work", ""),
                sys.get("db_work", ""),
                sys.get("data_work", ""),
                sys.get("user_work", ""),
                sys.get("complexity", ""),
                sys.get("total_work", ""),
                "", "", "", "", "", "", ""
            ])
        ws.append([
            filename,
            "【汇总】",
            "", "", "", "", "", "",
            item.get("total_workload", ""),
            cost.get("labor_cost", ""),
            cost.get("management_cost", ""),
            cost.get("risk_cost", ""),
            cost.get("test_cost", ""),
            cost.get("subtotal", ""),
            cost.get("total_excl_tax", ""),
            cost.get("total_incl_tax", "")
        ])
        ws.append([])

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream


if __name__ == "__main__":
    app.run(debug=True)
