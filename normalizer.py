def normalize(data):

    systems = []

    for s in data["systems"]:

        system = {}

        system["name"] = s.get("name", "未知系统")

        def safe_count(val):
            if isinstance(val, list):
                return len(val)
            if isinstance(val, (int, float)):
                return int(val)
            return 0

        system["module_count"] = safe_count(s.get("modules", []))
        system["interface_count"] = safe_count(s.get("interfaces", []))
        system["table_count"] = safe_count(s.get("tables", []))

        data_size = s.get("data_size_gb")
        try:
            data_size = float(data_size) if data_size is not None else 50
        except (TypeError, ValueError):
            data_size = 50
        system["data_size_gb"] = data_size

        users = s.get("users", {})
        # LLM 可能将 users 返回为整数而非 dict，做类型保护
        if isinstance(users, dict):
            system["concurrent_users"] = int(users.get("concurrent_users", 0) or 0)
        elif isinstance(users, (int, float)):
            system["concurrent_users"] = int(users)
        else:
            system["concurrent_users"] = 0

        systems.append(system)

    return {"systems":systems}
