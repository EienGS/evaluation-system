function switchTab(name) {
    document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"))
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"))
    document.getElementById("tab-" + name).classList.add("active")
    event.target.classList.add("active")

    if (name === "config") {
        loadConfig()
    }
}

function loadConfig() {
    fetch("/config")
        .then(r => r.json())
        .then(data => {
            document.getElementById("module_factor").value = data.module_factor
            document.getElementById("interface_factor").value = data.interface_factor
            document.getElementById("table_factor").value = data.table_factor
            document.getElementById("data_factor").value = data.data_factor
            document.getElementById("user_factor").value = data.user_factor
            document.getElementById("price_per_day").value = data.price_per_day
            document.getElementById("management_rate").value = data.management_rate
            document.getElementById("risk_rate").value = data.risk_rate
        })
        .catch(err => {
            document.getElementById("configErrMsg").innerText = "加载配置失败：" + err.message
        })
}

function saveConfig() {
    document.getElementById("configMsg").innerText = ""
    document.getElementById("configErrMsg").innerText = ""

    const config = {
        module_factor: parseFloat(document.getElementById("module_factor").value),
        interface_factor: parseFloat(document.getElementById("interface_factor").value),
        table_factor: parseFloat(document.getElementById("table_factor").value),
        data_factor: parseFloat(document.getElementById("data_factor").value),
        user_factor: parseFloat(document.getElementById("user_factor").value),
        price_per_day: parseFloat(document.getElementById("price_per_day").value),
        management_rate: parseFloat(document.getElementById("management_rate").value),
        risk_rate: parseFloat(document.getElementById("risk_rate").value)
    }

    fetch("/config", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(config)
    })
        .then(r => r.json())
        .then(data => {
            if (data.success) {
                document.getElementById("configMsg").innerText = "配置已保存"
            } else {
                document.getElementById("configErrMsg").innerText = "保存失败：" + data.error
            }
        })
        .catch(err => {
            document.getElementById("configErrMsg").innerText = "保存失败：" + err.message
        })
}

function parsePlan() {
    const md = document.getElementById("mdInput").value

    fetch("/parse", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ md: md })
    })
        .then(r => r.json())
        .then(data => {
            document.getElementById("result").innerText = data.result
        })
}

function evaluate() {
    const jsonText = document.getElementById("result").innerText

    fetch("/evaluate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: jsonText
    })
        .then(r => r.json())
        .then(data => {
            document.getElementById("evaluation").innerText = JSON.stringify(data, null, 2)
        })
}

function exportExcel() {
    const evalText = document.getElementById("evaluation").innerText

    fetch("/export_excel", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: evalText
    })
        .then(res => res.blob())
        .then(blob => {
            const url = window.URL.createObjectURL(blob)
            const a = document.createElement("a")
            a.href = url
            a.download = "workload_evaluation.xlsx"
            a.click()
            window.URL.revokeObjectURL(url)
        })
}
