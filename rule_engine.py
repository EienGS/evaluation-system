def calculate_system_workload(system, config):

    # 读取参数
    module_factor = config["module_factor"]
    interface_factor = config["interface_factor"]
    table_factor = config["table_factor"]
    data_factor = config["data_factor"]
    user_factor = config["user_factor"]

    # 基础工作量
    module_work = system["module_count"] * module_factor
    interface_work = system["interface_count"] * interface_factor
    db_work = system["table_count"] * table_factor
    data_work = system["data_size_gb"] * data_factor
    user_work = system["concurrent_users"] * user_factor

    base_work = module_work + interface_work + db_work + data_work + user_work

    # 系统复杂度系数
    complexity = 1

    if system["module_count"] > 20:
        complexity += 0.2

    if system["interface_count"] > 10:
        complexity += 0.1

    if system["data_size_gb"] > 500:
        complexity += 0.2

    total_work = base_work * complexity

    return {
        "name": system["name"],
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

    # 成本计算
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