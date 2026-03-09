from flask import Flask,render_template,request,jsonify
from llm_parser import parse_plan
from normalizer import normalize
from rule_engine import calculate_total
from config_loader import load_config

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

if __name__ == "__main__":
    app.run(debug=True)