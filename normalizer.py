import re


def to_float(val, default=0.0):
    """将任意值安全转换为 float，提取字符串中的数字部分。"""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        m = re.search(r"[\d.]+", val)
        if m:
            return float(m.group())
    return default


def to_int(val, default=0):
    return int(to_float(val, default))


def normalize(data):
    systems = []

    raw_systems = data.get("systems", [])
    if not isinstance(raw_systems, list):
        raw_systems = []

    for s in raw_systems:
        if not isinstance(s, dict):
            continue

        system = {}
        system["name"] = s.get("name", "未知系统")
        system["level"] = s.get("level", "一般")
        system["complexity_note"] = s.get("complexity_note", "")

        # 技术栈与改造类型
        tech_stack = s.get("tech_stack", [])
        system["tech_stack"] = tech_stack if isinstance(tech_stack, list) else []

        adaptation_type = s.get("adaptation_type", [])
        system["adaptation_type"] = adaptation_type if isinstance(adaptation_type, list) else []

        # 数量字段：兼容列表或整数
        modules = s.get("modules", [])
        system["module_count"] = len(modules) if isinstance(modules, list) else to_int(modules)

        interfaces = s.get("interfaces", [])
        system["interface_count"] = len(interfaces) if isinstance(interfaces, list) else to_int(interfaces)

        tables = s.get("tables", [])
        system["table_count"] = len(tables) if isinstance(tables, list) else to_int(tables)

        # data_size_gb 默认 0（避免凭空增加工作量）
        system["data_size_gb"] = to_float(s.get("data_size_gb"), default=0.0)

        # users
        users = s.get("users", {})
        if isinstance(users, dict):
            system["users"] = to_int(users.get("online_users", 0))
            system["concurrent_users"] = to_int(users.get("concurrent_users", 0))
        else:
            system["users"] = to_int(users)
            system["concurrent_users"] = 0

        systems.append(system)

    return {
        "project_name": data.get("project_name", ""),
        "project_background": data.get("project_background", ""),
        "existing_tech_stack": data.get("existing_tech_stack", {}),
        "target_tech_stack": data.get("target_tech_stack", {}),
        "external_interfaces": to_int(data.get("external_interfaces", 0)),
        "systems": systems
    }
