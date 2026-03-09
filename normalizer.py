import re


def to_float(val, default=0.0):
    """将任意值安全转换为 float，提取字符串中的数字部分。"""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        # 提取第一段数字（含小数点）
        m = re.search(r"[\d.]+", val)
        if m:
            return float(m.group())
    return default


def to_int(val, default=0):
    """将任意值安全转换为 int。"""
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

        # modules / interfaces / tables 可能是列表或整数
        modules = s.get("modules", [])
        system["module_count"] = len(modules) if isinstance(modules, list) else to_int(modules)

        interfaces = s.get("interfaces", [])
        system["interface_count"] = len(interfaces) if isinstance(interfaces, list) else to_int(interfaces)

        tables = s.get("tables", [])
        system["table_count"] = len(tables) if isinstance(tables, list) else to_int(tables)

        # data_size_gb 可能是数字、字符串或 None
        data_size = s.get("data_size_gb")
        system["data_size_gb"] = to_float(data_size, default=50.0)

        # users 可能是字典、整数或 None
        users = s.get("users", {})
        if isinstance(users, dict):
            system["users"] = to_int(users.get("online_users", 0))
            system["concurrent_users"] = to_int(users.get("concurrent_users", 0))
        else:
            # 直接给了一个数字
            system["users"] = to_int(users)
            system["concurrent_users"] = 0

        systems.append(system)

    return {"systems": systems}
