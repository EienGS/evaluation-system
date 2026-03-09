// ── Tab 切换 ──────────────────────────────────────────────
function switchTab(name, btn) {
    document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"))
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"))
    document.getElementById("tab-" + name).classList.add("active")
    btn.classList.add("active")
    if (name === "config") loadConfig()
}

// ── 文件管理 ──────────────────────────────────────────────
let fileQueue = []       // { file, name, status, error }
let evalResults = []     // 成功的评估结果

function addFiles(fileList) {
    Array.from(fileList).forEach(file => {
        // 去重
        if (fileQueue.find(f => f.name === file.name)) return
        fileQueue.push({ file, name: file.name, status: "pending", error: "" })
    })
    document.getElementById("fileInput").value = ""
    renderFileList()
    updateButtons()
}

function removeFile(idx) {
    fileQueue.splice(idx, 1)
    renderFileList()
    updateButtons()
}

function clearAll() {
    fileQueue = []
    evalResults = []
    renderFileList()
    updateButtons()
    document.getElementById("resultSection").style.display = "none"
    document.getElementById("resultBody").innerHTML = ""
    document.getElementById("progressMsg").innerText = ""
}

function renderFileList() {
    const list = document.getElementById("fileList")
    if (fileQueue.length === 0) { list.innerHTML = ""; return }

    list.innerHTML = fileQueue.map((f, i) => `
        <div class="file-item">
            <span class="file-name" title="${f.name}">${f.name}</span>
            ${f.error ? `<span class="error-detail" title="${f.error}">${f.error}</span>` : ""}
            <span class="file-status status-${f.status}">${statusLabel(f.status)}</span>
            <button class="remove-btn" onclick="removeFile(${i})" title="移除">&#x2715;</button>
        </div>
    `).join("")
}

function statusLabel(s) {
    return { pending: "待处理", running: "处理中", success: "成功", error: "已跳过" }[s] || s
}

function updateButtons() {
    const hasPending = fileQueue.some(f => f.status === "pending")
    const hasResult  = evalResults.length > 0
    document.getElementById("runBtn").disabled    = fileQueue.length === 0
    document.getElementById("exportBtn").disabled = !hasResult
}

// ── 拖拽上传 ─────────────────────────────────────────────
const dropZone = document.getElementById("dropZone")
dropZone.addEventListener("dragover", e => { e.preventDefault(); dropZone.classList.add("drag-over") })
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("drag-over"))
dropZone.addEventListener("drop", e => {
    e.preventDefault()
    dropZone.classList.remove("drag-over")
    const files = Array.from(e.dataTransfer.files).filter(f => f.name.match(/\.(md|markdown)$/i))
    addFiles(files)
})

// ── 批量评估 ──────────────────────────────────────────────
async function runAll() {
    evalResults = []
    document.getElementById("resultBody").innerHTML = ""
    document.getElementById("resultSection").style.display = "none"
    document.getElementById("runBtn").disabled = true
    document.getElementById("exportBtn").disabled = true

    // 重置所有状态为 pending
    fileQueue.forEach(f => { f.status = "pending"; f.error = "" })
    renderFileList()

    let successCount = 0
    let errorCount = 0

    for (let i = 0; i < fileQueue.length; i++) {
        const item = fileQueue[i]
        item.status = "running"
        renderFileList()
        setProgress(`正在处理 ${i + 1} / ${fileQueue.length}：${item.name}`)

        try {
            const md = await readFileText(item.file)
            const resp = await fetch("/evaluate_file", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ md, filename: item.name })
            })
            const json = await resp.json()

            if (json.success) {
                item.status = "success"
                console.log("[v0] evaluate result:", JSON.stringify(json.data))
                evalResults.push(json.data)
                appendResultRows(json.data)
                successCount++
            } else {
                item.status = "error"
                item.error = json.error || "未知错误"
                errorCount++
            }
        } catch (err) {
            item.status = "error"
            item.error = err.message
            errorCount++
        }

        renderFileList()
    }

    setProgress(`完成：${successCount} 个成功，${errorCount} 个跳过`)

    if (evalResults.length > 0) {
        document.getElementById("resultSection").style.display = "block"
        document.getElementById("exportBtn").disabled = false
    }
    document.getElementById("runBtn").disabled = false
    updateButtons()
}

function readFileText(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader()
        reader.onload = e => resolve(e.target.result)
        reader.onerror = () => reject(new Error("文件读取失败"))
        reader.readAsText(file, "utf-8")
    })
}

function setProgress(msg) {
    document.getElementById("progressMsg").innerText = msg
}

function appendResultRows(data) {
    const tbody = document.getElementById("resultBody")
    const filename = data.filename || ""
    const cost = data.cost || {}

    // 文件分组标题行
    const header = document.createElement("tr")
    header.className = "row-file-header"
    header.innerHTML = `<td colspan="13">${filename}</td>`
    tbody.appendChild(header)

    // 各系统行
    ;(data.systems || []).forEach(sys => {
        const tr = document.createElement("tr")
        tr.innerHTML = [
            "", sys.name,
            fmt(sys.module_work), fmt(sys.interface_work),
            fmt(sys.db_work), fmt(sys.data_work), fmt(sys.user_work),
            sys.complexity,
            fmt(sys.total_work),
            "", "", "", ""
        ].map(v => `<td>${v}</td>`).join("")
        tbody.appendChild(tr)
    })

    // 汇总行
    const summary = document.createElement("tr")
    summary.className = "row-summary"
    summary.innerHTML = [
        "", "【汇总】",
        "", "", "", "", "", "",
        fmt(data.total_workload),
        fmtCost(cost.dev_cost),
        fmtCost(cost.management_cost),
        fmtCost(cost.risk_cost),
        fmtCost(cost.total_cost)
    ].map(v => `<td>${v}</td>`).join("")
    tbody.appendChild(summary)
}

function fmt(v) { return v !== undefined ? (+v).toFixed(2) : "" }
function fmtCost(v) { return v !== undefined ? (+v).toFixed(0) : "" }

// ── 导出 Excel ───────────────────────────────────────────
function exportExcel() {
    fetch("/export_excel", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(evalResults)
    })
        .then(res => res.blob())
        .then(blob => {
            const url = window.URL.createObjectURL(blob)
            const a = document.createElement("a")
            a.href = url
            a.download = "国产化适配费用评估.xlsx"
            a.click()
            window.URL.revokeObjectURL(url)
        })
}

// ── 参数配置 ──────────────────────────────────────────────
function loadConfig() {
    fetch("/config")
        .then(r => r.json())
        .then(data => {
            document.getElementById("module_factor").value    = data.module_factor
            document.getElementById("interface_factor").value = data.interface_factor
            document.getElementById("table_factor").value     = data.table_factor
            document.getElementById("data_factor").value      = data.data_factor
            document.getElementById("user_factor").value      = data.user_factor
            document.getElementById("price_per_day").value    = data.price_per_day
            document.getElementById("management_rate").value  = data.management_rate
            document.getElementById("risk_rate").value        = data.risk_rate
        })
        .catch(err => {
            document.getElementById("configErrMsg").innerText = "加载配置失败：" + err.message
        })
}

function saveConfig() {
    document.getElementById("configMsg").innerText    = ""
    document.getElementById("configErrMsg").innerText = ""

    const config = {
        module_factor:    parseFloat(document.getElementById("module_factor").value),
        interface_factor: parseFloat(document.getElementById("interface_factor").value),
        table_factor:     parseFloat(document.getElementById("table_factor").value),
        data_factor:      parseFloat(document.getElementById("data_factor").value),
        user_factor:      parseFloat(document.getElementById("user_factor").value),
        price_per_day:    parseFloat(document.getElementById("price_per_day").value),
        management_rate:  parseFloat(document.getElementById("management_rate").value),
        risk_rate:        parseFloat(document.getElementById("risk_rate").value)
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
                setTimeout(() => { document.getElementById("configMsg").innerText = "" }, 3000)
            } else {
                document.getElementById("configErrMsg").innerText = "保存失败：" + data.error
            }
        })
        .catch(err => {
            document.getElementById("configErrMsg").innerText = "保存失败：" + err.message
        })
}
