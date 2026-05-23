/* global XLSX, chrome */
const TTL_MS = 24 * 60 * 60 * 1000; // 24h

const els = {
  file: document.getElementById("file"),
  clear: document.getElementById("clear"),
  meta: document.getElementById("fileMeta"),
  sheet: document.getElementById("sheet"),
  search: document.getElementById("search"),
  tableWrap: document.getElementById("tableWrap"),
  fill: document.getElementById("fill"),
  selMeta: document.getElementById("selMeta"),
  status: document.getElementById("status"),
};

let state = {
  fileName: null,
  loadedAt: 0,
  sheets: {}, // { sheetName: { headers: [], rows: [[...]] } }
  activeSheet: null,
  selectedRowIdx: null,
};

function setStatus(msg, kind) {
  els.status.className = "status " + (kind || "");
  els.status.textContent = msg || "";
  if (!msg) els.status.className = "";
}

async function save() {
  await chrome.storage.local.set({ sheetState: state });
}

async function load() {
  const { sheetState } = await chrome.storage.local.get("sheetState");
  if (!sheetState) return;
  if (Date.now() - sheetState.loadedAt > TTL_MS) {
    await chrome.storage.local.remove("sheetState");
    return;
  }
  state = sheetState;
  render();
}

function parseWorkbook(arrayBuffer, name) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheets = {};
  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false });
    if (!aoa.length) continue;
    // Find first row with the most non-empty cells as header (in case of merged title rows)
    let headerIdx = 0;
    let best = 0;
    for (let i = 0; i < Math.min(aoa.length, 5); i++) {
      const c = aoa[i].filter((v) => String(v).trim() !== "").length;
      if (c > best) { best = c; headerIdx = i; }
    }
    const headers = aoa[headerIdx].map((h) => String(h).trim());
    const rows = aoa.slice(headerIdx + 1).filter((r) => r.some((v) => String(v).trim() !== ""));
    sheets[sheetName] = { headers, rows };
  }
  state = {
    fileName: name,
    loadedAt: Date.now(),
    sheets,
    activeSheet: Object.keys(sheets)[0] || null,
    selectedRowIdx: null,
  };
}

function render() {
  if (!state.fileName) {
    els.meta.textContent = "No file loaded.";
    els.sheet.innerHTML = "";
    els.tableWrap.innerHTML = "";
    els.fill.disabled = true;
    els.selMeta.textContent = "No row selected.";
    return;
  }
  const expires = new Date(state.loadedAt + TTL_MS).toLocaleString();
  els.meta.textContent = `${state.fileName} — remembered until ${expires}`;

  els.sheet.innerHTML = "";
  for (const name of Object.keys(state.sheets)) {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    if (name === state.activeSheet) opt.selected = true;
    els.sheet.appendChild(opt);
  }
  renderTable();
}

function renderTable() {
  const s = state.sheets[state.activeSheet];
  if (!s) { els.tableWrap.innerHTML = ""; return; }
  const q = els.search.value.trim().toLowerCase();
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  const numTh = document.createElement("th"); numTh.textContent = "#"; trh.appendChild(numTh);
  s.headers.forEach((h) => { const th = document.createElement("th"); th.textContent = h; trh.appendChild(th); });
  thead.appendChild(trh);
  table.appendChild(thead);
  const tbody = document.createElement("tbody");
  s.rows.forEach((row, idx) => {
    const text = row.join(" ").toLowerCase();
    if (q && !text.includes(q)) return;
    const tr = document.createElement("tr");
    if (idx === state.selectedRowIdx) tr.classList.add("selected");
    const numTd = document.createElement("td"); numTd.textContent = idx + 1; tr.appendChild(numTd);
    s.headers.forEach((_, i) => { const td = document.createElement("td"); td.textContent = row[i] ?? ""; tr.appendChild(td); });
    tr.addEventListener("click", () => {
      state.selectedRowIdx = idx;
      els.fill.disabled = false;
      els.selMeta.textContent = `Row ${idx + 1} selected.`;
      save();
      renderTable();
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  els.tableWrap.innerHTML = "";
  els.tableWrap.appendChild(table);
}

els.file.addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  setStatus("Reading file…");
  try {
    const buf = await f.arrayBuffer();
    parseWorkbook(buf, f.name);
    await save();
    render();
    setStatus("Loaded. Pick a row, then click Fill.", "ok");
  } catch (err) {
    setStatus("Could not parse file: " + err.message, "err");
  }
});

els.clear.addEventListener("click", async () => {
  await chrome.storage.local.remove("sheetState");
  state = { fileName: null, loadedAt: 0, sheets: {}, activeSheet: null, selectedRowIdx: null };
  render();
  setStatus("Cleared.", "ok");
});

els.sheet.addEventListener("change", () => {
  state.activeSheet = els.sheet.value;
  state.selectedRowIdx = null;
  els.fill.disabled = true;
  els.selMeta.textContent = "No row selected.";
  save();
  renderTable();
});

els.search.addEventListener("input", renderTable);

els.fill.addEventListener("click", async () => {
  const s = state.sheets[state.activeSheet];
  if (!s || state.selectedRowIdx == null) return;
  const row = s.rows[state.selectedRowIdx];
  const data = {};
  // For duplicate headers (e.g. "Crop" appears twice), the LATER value wins —
  // the right-hand "Crop"/"Variety" columns hold the human-readable names that
  // match the form labels.
  s.headers.forEach((h, i) => {
    const key = String(h).trim();
    if (!key) return;
    const val = row[i];
    if (val === "" || val == null) return;
    data[key] = val;
  });

  setStatus("Filling…");
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const [{ result }] = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: fillFormInPage,
      args: [data],
    });
    setStatus(
      result.filled.length
        ? `Filled ${result.filled.length} field(s): ${result.filled.join(", ")}${result.missed.length ? ` · Missed: ${result.missed.join(", ")}` : ""}`
        : "No matching fields found on this page.",
      result.filled.length ? "ok" : "err",
    );
  } catch (err) {
    setStatus("Fill failed: " + err.message, "err");
  }
});

// This function runs in the page context.
function fillFormInPage(data) {
  const norm = (s) => String(s || "").toLowerCase().replace(/\([^)]*\)/g, "").replace(/[^a-z0-9]+/g, " ").trim();

  // Build candidate field list: every input/select/textarea + its best label text.
  const fields = [];
  document.querySelectorAll("input, select, textarea").forEach((el) => {
    if (el.type === "hidden" || el.disabled) return;
    const labels = [];
    if (el.id) {
      const lab = document.querySelector(`label[for="${CSS.escape(el.id)}"]`);
      if (lab) labels.push(lab.innerText);
    }
    const wrapLabel = el.closest("label");
    if (wrapLabel) labels.push(wrapLabel.innerText);
    // Walk up a few ancestors and look at preceding text (h6, label, div, span)
    let node = el.parentElement;
    for (let i = 0; i < 5 && node; i++, node = node.parentElement) {
      const heads = node.querySelectorAll("h1,h2,h3,h4,h5,h6,label,legend,.label,[class*='label']");
      heads.forEach((h) => labels.push(h.innerText));
      // previous sibling text
      let prev = node.previousElementSibling;
      if (prev) labels.push(prev.innerText);
    }
    labels.push(el.placeholder || "", el.name || "", el.getAttribute("aria-label") || "");
    fields.push({ el, labels: labels.map(norm).filter(Boolean) });
  });

  const filled = [];
  const missed = [];

  const setValue = (el, value) => {
    const v = String(value);
    const tag = el.tagName.toLowerCase();
    if (tag === "select") {
      const opts = Array.from(el.options);
      const match = opts.find((o) => norm(o.text) === norm(v) || norm(o.value) === norm(v))
        || opts.find((o) => norm(o.text).includes(norm(v)));
      if (match) el.value = match.value; else el.value = v;
    } else if (el.type === "checkbox" || el.type === "radio") {
      el.checked = ["true", "yes", "1", "on", "y", "√", "✓"].includes(v.toLowerCase().trim());
    } else {
      // React/Angular controlled inputs need a native setter
      const proto = tag === "textarea" ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
      const setter = Object.getOwnPropertyDescriptor(proto, "value")?.set;
      if (setter) setter.call(el, v); else el.value = v;
    }
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    el.dispatchEvent(new Event("blur", { bubbles: true }));
  };

  for (const [key, value] of Object.entries(data)) {
    const k = norm(key);
    if (!k) continue;
    // Score each field
    let best = null;
    let bestScore = 0;
    for (const f of fields) {
      let score = 0;
      for (const l of f.labels) {
        if (!l) continue;
        if (l === k) score = Math.max(score, 100);
        else if (l.split(" ").join("") === k.split(" ").join("")) score = Math.max(score, 95);
        else if (l.includes(k) || k.includes(l)) score = Math.max(score, 60 + Math.min(30, Math.min(l.length, k.length)));
      }
      if (score > bestScore) { bestScore = score; best = f; }
    }
    if (best && bestScore >= 60) {
      try { setValue(best.el, value); filled.push(key); } catch { missed.push(key); }
    } else {
      missed.push(key);
    }
  }
  return { filled, missed };
}

load();
