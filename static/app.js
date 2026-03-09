/* ========== 状态 ========== */
let selectedFiles = [];   // { name, size, content }
let batchResults = [];    // 批量评估结果

/* ========== Tab 切换 ========== */
function switchTab(name) {
  document.querySelectorAll(".tab-pane").forEach(p => p.classList.remove("active"));
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
  document.getElementById("pane-" + name).classList.add("active");
  document.getElementById("tab-" + name).classList.add("active");
  if (name === "config") loadConfigToForm();
}

/* ========== Toast ========== */
function showToast(msg, type) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.className = "show toast-" + (type || "ok");
  clearTimeout(t._timer);
  t._timer = setTimeout(() => { t.className = ""; }, 3000);
}

/* ========== 文件上传 ========== */
(function () {
  const zone = document.getElementById("uploadZone");
  zone.addEventListener("dragover", e => { e.preventDefault(); zone.classList.add("dragover"); });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
  zone.addEventListener("drop", e => {
    e.preventDefault();
    zone.classList.remove("dragover");
    onFilesSelected(e.dataTransfer.files);
  });
})();

function onFilesSelected(fileList) {
  const existing = new Set(selectedFiles.map(f => f.name));
  Array.from(fileList).forEach(file => {
    if (existing.has(file.name)) return;
    existing.add(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      selectedFiles.push({ name: file.name, size: file.size, content: e.target.result });
      renderFileList();
    };
    reader.readAsText(file, "UTF-8");
  });
}

function renderFileList() {
  const container = document.getElementById("fileList");
  container.innerHTML = "";
  selectedFiles.forEach((f, idx) => {
    const status = f.status || "pending";
    const statusLabels = { pending: "待处理", processing: "处理中", ok: "成功", error: "失败" };
    const div = document.createElement("div");
    div.className = "file-item";
    div.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#6b7280" stroke-width="2">
        <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/>
      </svg>
      <span class="file-name">${escHtml(f.name)}</span>
      <span class="file-size">${formatSize(f.size)}</span>
      ${f.error ? `<span class="file-error" title="${escHtml(f.error)}">${escHtml(f.error.substring(0, 60))}…</span>` : ""}
      <span class="file-status status-${status}">${statusLabels[status]}</span>
      ${status === "pending" ? `<button class="file-remove" onclick="removeFile(${idx})" title="移除">×</button>` : ""}
    `;
    container.appendChild(div);
  });
  document.getElementById("btnRun").disabled = selectedFiles.length === 0;
}

function removeFile(idx) {
  selectedFiles.splice(idx, 1);
  renderFileList();
}

function clearFiles() {
  selectedFiles = [];
  batchResults = [];
  renderFileList();
  document.getElementById("resultSection").style.display = "none";
  document.getElementById("btnExport").disabled = true;
  document.getElementById("progressWrap").style.display = "none";
}

function formatSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  return (bytes / 1024).toFixed(1) + " KB";
}

function escHtml(str) {
  return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/* ========== 批量评估 ========== */
async function runBatch() {
  if (selectedFiles.length === 0) return;

  batchResults = [];
  selectedFiles.forEach(f => { f.status = "pending"; delete f.error; });

  const btnRun = document.getElementById("btnRun");
  btnRun.disabled = true;
  btnRun.innerHTML = `<span class="spinner"></span> 评估中...`;

  const progressWrap = document.getElementById("progressWrap");
  const progressBar = document.getElementById("progressBar");
  const progressLabel = document.getElementById("progressLabel");
  progressWrap.style.display = "block";
  progressBar.style.width = "0%";

  for (let i = 0; i < selectedFiles.length; i++) {
    const f = selectedFiles[i];
    f.status = "processing";
    renderFileList();
    progressLabel.textContent = `正在处理 ${i + 1}/${selectedFiles.length}：${f.name}`;
    progressBar.style.width = `${Math.round((i / selectedFiles.length) * 100)}%`;

    try {
      const res = await fetch("/batch_evaluate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ files: [{ name: f.name, content: f.content }] })
      });
      const json = await res.json();
      const result = json.results[0];
      if (result.status === "ok") {
        f.status = "ok";
      } else {
        f.status = "error";
        f.error = result.error;
      }
      batchResults.push(result);
    } catch (e) {
      f.status = "error";
      f.error = e.message;
      batchResults.push({ name: f.name, status: "error", error: e.message });
    }
    renderFileList();
  }

  const okCount = selectedFiles.filter(f => f.status === "ok").length;
  const errCount = selectedFiles.filter(f => f.status === "error").length;
  progressBar.style.width = "100%";
  progressLabel.textContent = `处理完成：${okCount} 个成功，${errCount} 个失败`;

  btnRun.disabled = false;
  btnRun.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg> 重新评估`;

  renderResults();
  showToast(`评估完成，${okCount} 个成功`, okCount > 0 ? "ok" : "err");
}

function renderResults() {
  const okResults = batchResults.filter(r => r.status === "ok");
  const section = document.getElementById("resultSection");

  if (batchResults.length === 0) { section.style.display = "none"; return; }
  section.style.display = "block";

  let totalWork = 0, totalCost = 0;
  okResults.forEach(r => {
    totalWork += r.data.total_workload || 0;
    totalCost += (r.data.cost && r.data.cost.total_incl_tax) || 0;
  });
  document.getElementById("statOk").textContent = okResults.length;
  document.getElementById("statWork").textContent = totalWork.toFixed(1);
  document.getElementById("statCost").textContent = (totalCost / 10000).toFixed(2);

  const tbody = document.getElementById("resultBody");
  tbody.innerHTML = "";

  batchResults.forEach(r => {
    const tr = document.createElement("tr");
    if (r.status === "ok") {
      const d = r.data;
      const cost = d.cost || {};
      tr.innerHTML = `
        <td>${escHtml(r.name)}</td>
        <td>${(d.systems || []).length}</td>
        <td class="td-total">${(d.total_workload || 0).toFixed(1)}</td>
        <td>${fmt(cost.labor_cost)}</td>
        <td>${fmt(cost.test_cost)}</td>
        <td>${fmt(cost.management_cost)}</td>
        <td>${fmt(cost.risk_cost)}</td>
        <td>${fmt(cost.total_excl_tax)}</td>
        <td class="td-cost">${fmt(cost.total_incl_tax)}</td>
        <td><span class="file-status status-ok">成功</span></td>
      `;
    } else {
      tr.innerHTML = `
        <td>${escHtml(r.name)}</td>
        <td colspan="8" class="td-err">处理失败：${escHtml(r.error || "未知错误")}</td>
        <td><span class="file-status status-error">失败</span></td>
      `;
    }
    tbody.appendChild(tr);
  });

  document.getElementById("btnExport").disabled = okResults.length === 0;
}

function fmt(num) {
  if (num == null) return "-";
  return Math.round(num).toLocaleString("zh-CN");
}

/* ========== 导出 Excel ========== */
function exportExcel() {
  const okResults = batchResults.filter(r => r.status === "ok");
  if (okResults.length === 0) { showToast("没有可导出的成功结果", "err"); return; }

  fetch("/export_excel", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(batchResults)
  })
  .then(res => {
    if (!res.ok) throw new Error("导出失败");
    return res.blob();
  })
  .then(blob => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "信创适配费用评估报告.xlsx";
    a.click();
    URL.revokeObjectURL(url);
    showToast("导出成功", "ok");
  })
  .catch(e => showToast("导出失败：" + e.message, "err"));
}

/* ========== 参数配置 ========== */
function loadConfigToForm() {
  fetch("/config")
    .then(r => r.json())
    .then(cfg => {
      // 基础参数
      const basicKeys = ["module_factor","interface_factor","table_factor","data_factor",
                         "price_per_day","test_rate","management_rate","risk_rate","tax_rate"];
      basicKeys.forEach(k => {
        const el = document.getElementById("p_" + k);
        if (el) el.value = cfg[k] !== undefined ? cfg[k] : "";
      });
      // 改造附加量参数
      const af = cfg.adaptation_factors || {};
      Object.entries(af).forEach(([k, v]) => {
        const el = document.getElementById("af_" + k);
        if (el) el.value = v;
      });
    })
    .catch(e => showToast("加载配置失败：" + e.message, "err"));
}

function saveConfig() {
  // 先加载当前配置（保留 adaptation_factors 结构）
  fetch("/config")
    .then(r => r.json())
    .then(cfg => {
      // 覆盖基础参数
      const basicKeys = ["module_factor","interface_factor","table_factor","data_factor",
                         "price_per_day","test_rate","management_rate","risk_rate","tax_rate"];
      basicKeys.forEach(k => {
        const el = document.getElementById("p_" + k);
        if (el) cfg[k] = parseFloat(el.value);
      });
      // 覆盖改造附加量
      if (!cfg.adaptation_factors) cfg.adaptation_factors = {};
      const afKeys = Object.keys(cfg.adaptation_factors);
      afKeys.forEach(k => {
        const el = document.getElementById("af_" + k);
        if (el) cfg.adaptation_factors[k] = parseFloat(el.value);
      });

      return fetch("/config", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(cfg)
      });
    })
    .then(r => r.json())
    .then(() => showToast("参数保存成功", "ok"))
    .catch(e => showToast("保存失败：" + e.message, "err"));
}
