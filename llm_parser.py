import requests

DEEPSEEK_API = "https://api.deepseek.com/v1/chat/completions"
API_KEY = "sk-43fc1671aaf54214b04c7742c7ba4d9a"


def parse_plan(md_text):

    prompt = f"""你是信息化项目国产化适配改造评估专家，拥有丰富的信创项目经验。

请从以下建设方案文档中，提取用于评估国产化适配改造工作量的结构化参数，以JSON格式返回。

## 返回字段说明

- project_name: 项目名称（字符串）
- project_background: 项目背景简述（1-2句话）
- systems: 子系统列表，每个系统包含以下字段：
  - name: 系统名称
  - level: 系统级别，"核心"/"重要"/"一般" 之一，根据系统重要性判断
  - modules: 功能模块数量（整数，若描述不明确请合理估算）
  - interfaces: 接口数量（整数，含内外部接口）
  - tables: 数据库表数量（整数）
  - data_size_gb: 数据量（数字，单位GB，若未提及填0）
  - tech_stack: 需要替换的技术栈列表，如 ["Oracle", "x86服务器", "Windows Server"]
  - adaptation_type: 改造类型列表，可包含以下一项或多项：
    ["数据库替换", "操作系统适配", "中间件替换", "硬件架构适配", "前端框架适配", "接口改造", "安全加固"]
  - complexity_note: 一句话描述本系统改造的主要技术难点

## 注意事项
- 若方案中未明确给出数量，请根据上下文合理估算，不要填0
- 数据库替换（如Oracle→达梦/人大金仓）通常工作量最大，应充分体现
- tech_stack 和 adaptation_type 直接影响工作量评估，务必尽量完整提取

建设方案原文：

{md_text}

只返回JSON，不要有任何其他文字。
"""

    payload = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.1
    }

    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }

    r = requests.post(DEEPSEEK_API, json=payload, headers=headers, timeout=120)
    r.raise_for_status()
    result = r.json()
    return result["choices"][0]["message"]["content"]
