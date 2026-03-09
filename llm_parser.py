import requests

DEEPSEEK_API = "https://api.deepseek.com/v1/chat/completions"
API_KEY = "sk-43fc1671aaf54214b04c7742c7ba4d9a"


def parse_plan(md_text):

    prompt = f"""
你是信息化项目评估专家。

请从以下建设方案中提取系统规模参数，并输出JSON。

字段：

systems:
- name
- modules
- interfaces
- tables
- data_size_gb
- users

建设方案：

{md_text}

只返回JSON。
"""

    payload = {
        "model": "deepseek-chat",
        "messages":[
            {"role":"user","content":prompt}
        ],
        "temperature":0.2
    }

    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type":"application/json"
    }

    r = requests.post(DEEPSEEK_API,json=payload,headers=headers)

    result = r.json()

    return result["choices"][0]["message"]["content"]