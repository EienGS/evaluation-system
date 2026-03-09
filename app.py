from flask import Flask,render_template,request,jsonify
from llm_parser import parse_plan
from normalizer import normalize
from rule_engine import calculate_total
from config_loader import load_config
from openpyxl import Workbook
from flask import send_file
import io

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/parse",methods=["POST"])
def parse():

    data = request.json
    md_text = data.get("md")

    result = parse_plan(md_text)

    return jsonify({"result":result})


@app.route("/evaluate",methods=["POST"])
def evaluate():

    data = request.json

    normalized = normalize(data)

    config = load_config()

    result = calculate_total(normalized, config)

    return jsonify(result)

@app.route("/export_excel", methods=["POST"])
def export_excel():

    data = request.json

    file_stream = generate_excel(data)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name="workload_evaluation.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def generate_excel(data):

    wb = Workbook()
    ws = wb.active
    ws.title = "工作量评估"

    ws.append([
        "系统名称",
        "模块工作量",
        "接口工作量",
        "数据库工作量",
        "数据对接工作量",
        "权限工作量",
        "复杂度系数",
        "总工作量(人天)"
    ])

    for sys in data["systems"]:

        ws.append([
            sys["name"],
            sys["module_work"],
            sys["interface_work"],
            sys["db_work"],
            sys["data_work"],
            sys["user_work"],
            sys["complexity"],
            sys["total_work"]
        ])

    ws.append([])
    ws.append(["总工作量", data["total_workload"]])

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return file_stream

if __name__ == "__main__":
    app.run(debug=True)
