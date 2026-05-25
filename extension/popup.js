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
  complete: document.getElementById("complete"),
  download: document.getElementById("download"),
  selMeta: document.getElementById("selMeta"),
  status: document.getElementById("status"),
};

let state = {
  fileName: null,
  loadedAt: 0,
  sheets: {}, // { sheetName: { headers: [], rows: [[...]], completed: [idx,...] } }
  activeSheet: null,
  selectedRowIdx: null,
};

function isCompleted(sheet, idx) {
  return Array.isArray(sheet?.completed) && sheet.completed.includes(idx);
}

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
    sheets[sheetName] = { headers, rows, completed: [] };
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
  const norm = (s) =>
    String(s || "")
      .toLowerCase()
      .replace(/\([^)]*\)/g, " ")
      .replace(/[^a-z0-9]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  const toks = (s) => norm(s).split(" ").filter(Boolean);
  // Stop-words that should NOT drive a match on their own.
  const STOP = new Set(["no", "id", "of", "the", "a", "an", "to", "from", "by", "for", "in", "on", "type", "name", "code", "date", "mrp"]);

  // Build candidates with a TIGHT label set — closest preceding label-like element only.
  const fields = [];
  document.querySelectorAll("input, select, textarea").forEach((el) => {
    if (el.type === "hidden" || el.disabled || el.readOnly) return;
    const rect = el.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) return;
    const labels = [];
    if (el.id) {
      const lab = document.querySelector(`label[for="${CSS.escape(el.id)}"]`);
      if (lab) labels.push(lab.innerText);
    }
    const wrapLabel = el.closest("label");
    if (wrapLabel) labels.push(wrapLabel.innerText);
    // Walk up at most 4 ancestors; at each level look at the IMMEDIATELY PRECEDING
    // sibling that has short text (a label, not a whole section).
    let node = el;
    for (let i = 0; i < 4 && node; i++) {
      let prev = node.previousElementSibling;
      while (prev) {
        const txt = (prev.innerText || "").trim();
        if (txt && txt.length > 0 && txt.length < 80 && !txt.includes("\n")) {
          labels.push(txt);
          break;
        }
        prev = prev.previousElementSibling;
      }
      node = node.parentElement;
    }
    labels.push(el.placeholder || "", el.name || "", el.id || "", el.getAttribute("aria-label") || "");
    const labelTokSets = labels.map(toks).filter((t) => t.length);
    if (!labelTokSets.length) return;
    fields.push({ el, labelTokSets, used: false });
  });

  const setValue = (el, value) => {
    const v = String(value);
    const tag = el.tagName.toLowerCase();
    try { el.focus(); } catch {}
    if (tag === "select") {
      const opts = Array.from(el.options);
      const match = opts.find((o) => norm(o.text) === norm(v) || norm(o.value) === norm(v))
        || opts.find((o) => norm(o.text).includes(norm(v)));
      if (match) el.value = match.value; else el.value = v;
    } else if (el.type === "checkbox" || el.type === "radio") {
      el.checked = ["true", "yes", "1", "on", "y", "√", "✓"].includes(v.toLowerCase().trim());
    } else {
      const proto = tag === "textarea" ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
      const setter = Object.getOwnPropertyDescriptor(proto, "value")?.set;
      if (setter) setter.call(el, v); else el.value = v;
    }
    // Fire a wide net of events for React / Angular / autocomplete widgets
    for (const type of ["keydown", "keypress", "input", "keyup", "change", "blur"]) {
      const ev = type.startsWith("key")
        ? new KeyboardEvent(type, { bubbles: true, key: "a" })
        : new Event(type, { bubbles: true });
      try { el.dispatchEvent(ev); } catch {}
    }
  };

  // Token-based score between a data-key (tokens) and one label (tokens).
  const score = (keyT, labT) => {
    if (!keyT.length || !labT.length) return 0;
    const keyStr = keyT.join(" ");
    const labStr = labT.join(" ");
    if (keyStr === labStr) return 100;
    const keyCore = keyT.filter((t) => !STOP.has(t));
    const labCore = labT.filter((t) => !STOP.has(t));
    if (!keyCore.length || !labCore.length) return 0;
    const labSet = new Set(labT);
    const matched = keyCore.filter((t) => labSet.has(t)).length;
    if (matched === 0) return 0;
    // Require ALL core key tokens to be present in the label.
    if (matched < keyCore.length) return 0;
    const extraCore = labCore.filter((t) => !keyT.includes(t)).length;
    let s = 75 + (matched / keyCore.length) * 20 - extraCore * 10;
    return Math.max(0, s);
  };

  const filled = [];
  const missed = [];
  // Match more-specific (more-token) keys first so they claim their field.
  const entries = Object.entries(data).sort((a, b) => toks(b[0]).length - toks(a[0]).length);

  for (const [key, value] of entries) {
    const keyT = toks(key);
    if (!keyT.length) continue;
    let best = null;
    let bestScore = 0;
    for (const f of fields) {
      if (f.used) continue;
      let s = 0;
      for (const lt of f.labelTokSets) s = Math.max(s, score(keyT, lt));
      if (s > bestScore) { bestScore = s; best = f; }
    }
    if (best && bestScore >= 70) {
      try { setValue(best.el, value); best.used = true; filled.push(key); }
      catch { missed.push(key); }
    } else {
      missed.push(key);
    }
  }
  return { filled, missed };
}

load();
