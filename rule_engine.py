def calculate_system_workload(system, config):

    module_factor = config["module_factor"]
    interface_factor = config["interface_factor"]
    table_factor = config["table_factor"]
    data_factor = config["data_factor"]
    user_factor = config["user_factor"]

    module_count = system.get("module_count", 0)
    interface_count = system.get("interface_count", 0)
    table_count = system.get("table_count", 0)
    data_size = system.get("data_size_gb", 0)
    concurrent_users = system.get("concurrent_users", 0)

    module_work = module_count * module_factor
    interface_work = interface_count * interface_factor
    db_work = table_count * table_factor
    data_work = data_size * data_factor
    user_work = concurrent_users * user_factor

    base_work = module_work + interface_work + db_work + data_work + user_work

    complexity = 1.0
    if module_count > 20:
        complexity += 0.2
    if interface_count > 10:
        complexity += 0.1
    if data_size > 500:
        complexity += 0.2

    total_work = base_work * complexity

    return {
        "name": system.get("name", ""),
        # 原始数量（供Excel展示）
        "modules": module_count,
        "interfaces": interface_count,
        "databases": table_count,
        # 工作量分项
        "module_work": module_work,
        "interface_work": interface_work,
        "db_work": db_work,
        "data_work": data_work,
        "user_work": user_work,
        "complexity": complexity,
        "total_work": total_work
    }


def calculate_total(data, config):

    results = []
    total_workload = 0

    for system in data["systems"]:
        r = calculate_system_workload(system, config)
        total_workload += r["total_work"]
        results.append(r)

    price_per_day = config["price_per_day"]
    management_rate = config["management_rate"]
    risk_rate = config["risk_rate"]

    dev_cost = total_workload * price_per_day
    management_cost = dev_cost * management_rate
    risk_cost = dev_cost * risk_rate
    total_cost = dev_cost + management_cost + risk_cost

    return {
        "systems": results,
        "total_workload": total_workload,
        "cost": {
            "dev_cost": dev_cost,
            "management_cost": management_cost,
            "risk_cost": risk_cost,
            "total_cost": total_cost
        }
    }
