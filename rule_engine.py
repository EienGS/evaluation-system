"""
国产化适配改造工作量 & 费用评估引擎

工作量计算逻辑：
  基础工作量 = 模块数×module_factor + 接口数×interface_factor + 数据表数×table_factor + 数据量GB×data_factor
  技术改造附加量 = 每种改造类型对应的人天附加值之和
  系统级别系数  = 核心1.3 / 重要1.1 / 一般1.0
  复杂度系数    = 根据模块数、接口数、改造类型数动态计算
  小计工作量    = (基础 + 技术附加) × 级别系数 × 复杂度系数

费用结构（符合信息化项目报价规范）：
  人工费     = 总工作量 × 日费率
  测试费     = 人工费 × test_rate
  小计       = 人工费 + 测试费
  项目管理费 = 小计 × management_rate
  风险费     = 小计 × risk_rate
  合计（不含税）= 小计 + 管理费 + 风险费
  含税价     = 合计 × (1 + tax_rate)

各阶段工时分布（用于Excel展示）：
  调研分析 15% / 环境搭建 10% / 代码适配 45% / 测试验证 20% / 上线部署 10%
"""

# 每种改造类型的额外工作量（人天），可在 config_params.json 中覆盖
DEFAULT_ADAPTATION_FACTORS = {
    "数据库替换": 30,
    "操作系统适配": 15,
    "中间件替换": 10,
    "硬件架构适配": 20,
    "前端框架适配": 8,
    "接口改造": 5,
    "安全加固": 10,
}

LEVEL_FACTORS = {"核心": 1.3, "重要": 1.1, "一般": 1.0}

PHASE_RATIOS = {
    "调研分析": 0.15,
    "环境搭建": 0.10,
    "代码适配": 0.45,
    "测试验证": 0.20,
    "上线部署": 0.10,
}


def _complexity_factor(system):
    """动态复杂度系数：基于模块数、接口数、改造类型数综合评定。"""
    factor = 1.0
    if system.get("module_count", 0) > 20:
        factor += 0.15
    if system.get("interface_count", 0) > 30:
        factor += 0.10
    elif system.get("interface_count", 0) > 10:
        factor += 0.05
    adaptation_count = len(system.get("adaptation_type", []))
    if adaptation_count >= 4:
        factor += 0.20
    elif adaptation_count >= 2:
        factor += 0.10
    return round(factor, 2)


def calculate_system_workload(system, config):
    module_factor = config.get("module_factor", 2)
    interface_factor = config.get("interface_factor", 0.5)
    table_factor = config.get("table_factor", 0.3)
    data_factor = config.get("data_factor", 0.005)

    module_count = system.get("module_count", 0)
    interface_count = system.get("interface_count", 0)
    table_count = system.get("table_count", 0)
    data_size = system.get("data_size_gb", 0.0)

    module_work = module_count * module_factor
    interface_work = interface_count * interface_factor
    db_work = table_count * table_factor
    data_work = data_size * data_factor
    base_work = module_work + interface_work + db_work + data_work

    # 技术改造附加工作量
    adaptation_factors = config.get("adaptation_factors", DEFAULT_ADAPTATION_FACTORS)
    adaptation_list = system.get("adaptation_type", [])
    adaptation_work = sum(adaptation_factors.get(t, 5) for t in adaptation_list)

    level = system.get("level", "一般")
    level_factor = LEVEL_FACTORS.get(level, 1.0)
    complexity = _complexity_factor(system)

    subtotal = (base_work + adaptation_work) * level_factor * complexity

    # 各阶段工时
    phases = {phase: round(subtotal * ratio, 2) for phase, ratio in PHASE_RATIOS.items()}

    return {
        "name": system.get("name", ""),
        "level": level,
        "tech_stack": system.get("tech_stack", []),
        "adaptation_type": adaptation_list,
        "complexity_note": system.get("complexity_note", ""),
        # 数量
        "modules": module_count,
        "interfaces": interface_count,
        "databases": table_count,
        "data_size_gb": data_size,
        # 工作量分项
        "module_work": round(module_work, 2),
        "interface_work": round(interface_work, 2),
        "db_work": round(db_work, 2),
        "data_work": round(data_work, 2),
        "adaptation_work": round(adaptation_work, 2),
        "base_work": round(base_work, 2),
        "level_factor": level_factor,
        "complexity": complexity,
        "total_work": round(subtotal, 2),
        "phases": phases,
    }


def calculate_total(data, config):
    results = []
    total_workload = 0.0

    for system in data.get("systems", []):
        r = calculate_system_workload(system, config)
        total_workload += r["total_work"]
        results.append(r)

    price_per_day = config.get("price_per_day", 1800)
    test_rate = config.get("test_rate", 0.12)
    management_rate = config.get("management_rate", 0.12)
    risk_rate = config.get("risk_rate", 0.08)
    tax_rate = config.get("tax_rate", 0.09)

    labor_cost = total_workload * price_per_day
    test_cost = labor_cost * test_rate
    subtotal = labor_cost + test_cost
    management_cost = subtotal * management_rate
    risk_cost = subtotal * risk_rate
    total_excl_tax = subtotal + management_cost + risk_cost
    total_incl_tax = total_excl_tax * (1 + tax_rate)

    return {
        "project_name": data.get("project_name", ""),
        "project_background": data.get("project_background", ""),
        "systems": results,
        "total_workload": round(total_workload, 2),
        "cost": {
            "labor_cost": round(labor_cost, 2),
            "test_cost": round(test_cost, 2),
            "subtotal": round(subtotal, 2),
            "management_cost": round(management_cost, 2),
            "risk_cost": round(risk_cost, 2),
            "total_excl_tax": round(total_excl_tax, 2),
            "total_incl_tax": round(total_incl_tax, 2),
        }
    }
