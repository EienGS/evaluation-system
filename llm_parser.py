import requests
import re

DEEPSEEK_API = "https://api.deepseek.com/v1/chat/completions"
API_KEY = "sk-43fc1671aaf54214b04c7742c7ba4d9a"


def parse_plan(md_text):

    prompt = f"""你是信息化项目国产化适配改造评估专家，拥有丰富的信创项目交付经验。

请从以下建设方案文档中，提取用于评估国产化适配改造工作量的结构化参数，以JSON格式返回。

## 提取策略（按优先级）

1. **优先查找"工作量清单"/"建设内容"/"功能模块"章节**，将其中每一个明确列出的子系统/模块作为独立的 system 条目
2. 若无明确清单，则从各章节的功能描述中识别独立子系统
3. **主动扫描全文**，识别所有提及的技术替换要求（数据库品牌、操作系统、中间件、硬件平台等），即便未在同一段落出现

## 返回字段说明

- project_name: 项目名称
- project_background: 项目背景简述（1-2句话，说明业务场景）
- existing_tech_stack: 现有技术栈总结，如 {{"database": "Oracle", "os": "Windows Server", "middleware": "Tomcat"}}
- target_tech_stack: 改造目标技术栈，如 {{"database": "达梦/人大金仓", "os": "麒麟/统信", "middleware": "东方通/金蝶"}}
- external_interfaces: 需要对接的外部系统数量（整数，从方案中提取，如"与XX系统对接"）
- systems: 子系统/模块列表，每项包含：
  - name: 模块/子系统名称（直接使用方案中的命名）
  - level: "核心"/"重要"/"一般"（根据业务重要性判断，核心=直接支撑主业务/领导决策）
  - modules: 功能点/子模块数量（整数，从功能描述中数，宁多勿少）
  - interfaces: 本系统涉及的接口数量（整数，含内外部接口）
  - tables: 数据库表数量（整数，可根据功能点合理估算，每个功能点平均3~8张表）
  - data_size_gb: 数据量（数字，未提及填0）
  - tech_stack: 当前使用的、需要替换的技术栈（从全文扫描，如 ["Oracle", "x86服务器", "Windows Server", "Tomcat"]）
  - adaptation_type: 该模块涉及的改造类型，从以下选择（可多选）：
    ["数据库替换", "操作系统适配", "中间件替换", "硬件架构适配", "前端框架适配", "接口改造", "安全加固", "ETL平台适配", "数据迁移"]
  - complexity_note: 一句话描述该模块改造的主要难点

## 改造类型判断规则
- 提到达梦/人大金仓/神通/金仓/GBase/瀚高 → "数据库替换"
- 提到麒麟/统信/中标/银河 → "操作系统适配"
- 提到东方通/金蝶/中创/国产中间件 → "中间件替换"
- 提到龙芯/飞腾/鲲鹏/ARM → "硬件架构适配"
- 提到ETL/数据转换平台 且需国产化 → "ETL平台适配"
- 提到数据迁移/数据汇聚/数据入库 → "数据迁移"
- 提到等保/安全加固/密码改造 → "安全加固"
- 提到外部接口对接/数据交换/API对接 → "接口改造"
- 前端展示类系统（大屏/可视化/门户）→ "前端框架适配"

## 重要提示
- 数量宁多勿少，这是粗报价，低估风险大
- ETL平台和可视化中间件的国产化是独立大项，单独列为系统条目
- 若有数据库，请始终包含"数据库替换"和"数据迁移"

建设方案原文：

{md_text}

只返回JSON，不要有任何其他文字。
"""

    payload = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.1,
        "max_tokens": 4096
    }

    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }

    r = requests.post(DEEPSEEK_API, json=payload, headers=headers, timeout=180)
    r.raise_for_status()
    result = r.json()
    raw = result["choices"][0]["message"]["content"]

    # 去掉可能的 markdown 代码块包裹
    raw = re.sub(r"^```[a-z]*\n?", "", raw.strip(), flags=re.IGNORECASE)
    raw = re.sub(r"\n?```$", "", raw.strip())
    return raw
