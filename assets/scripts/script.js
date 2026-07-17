// ── DEFAULT DATA ──────────────────────────────────────────────────────────────
const DEFAULT_INCOMES = [
  { id: "islamhadaya", name: "Islamhadaya", amount: 15000000 },
  { id: "tashe", name: "Tashe", amount: 1500000 },
  { id: "yaraneh", name: "Yaraneh", amount: 600000 },
];

const DEFAULT_CATEGORIES = [
  { id: "food", label: "🛒 Food & Drinks", target: 6000000 },
  { id: "family", label: "👨‍👩‍👧 Family", target: 5000000 },
  { id: "housing", label: "🏗️ Housing & Construction", target: 5000000 },
  { id: "work", label: "💼 Work & Growth", target: 2000000 },
  { id: "project", label: "💻 Project Expenses", target: 5000000 },
  { id: "saving", label: "💰 Saving & Investment", target: 2227196 },
  { id: "transport", label: "🚗 Transportation", target: 500000 },
  { id: "loans", label: "🏦 Loans & Debts", target: 5000000 },
  { id: "bills", label: "📡 Bills, Internet, VPN", target: 1000000 },
  { id: "clothing", label: "👕 Clothing", target: 1000000 },
  { id: "charity", label: "🕌 Charity & Donation", target: 500000 },
  { id: "health", label: "💊 Health & Beauty", target: 500000 },
];

// ── STATE ─────────────────────────────────────────────────────────────────────
let state = {
  incomes: [],
  expenses: [], // { id, categoryId, desc, amount, date }
  categories: [], // { id, label, target }
};

// ── PERSISTENCE ───────────────────────────────────────────────────────────────
function load() {
  const saved = localStorage.getItem("daramd_v1");
  if (saved) {
    state = JSON.parse(saved);
    if (!state.categories || state.categories.length === 0) {
      state.categories = DEFAULT_CATEGORIES.map((c) => ({ ...c }));
      save();
    }
  } else {
    state.incomes = DEFAULT_INCOMES.map((i) => ({ ...i }));
    state.expenses = [];
    state.categories = DEFAULT_CATEGORIES.map((c) => ({ ...c }));
    save();
  }
}
function save() {
  localStorage.setItem("daramd_v1", JSON.stringify(state));
}
function clearAllData() {
  if (!confirm("Reset all data to defaults?")) return;
  localStorage.removeItem("daramd_v1");
  load();
  render();
}

// ── FORMAT ────────────────────────────────────────────────────────────────────
function fmt(n) {
  return Math.abs(n).toLocaleString("en-US");
}
function esc(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// ── TOAST ─────────────────────────────────────────────────────────────────────
function showToast(msg, type = "success") {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.className = "show " + type;
  clearTimeout(t._timer);
  t._timer = setTimeout(() => {
    t.className = "";
  }, 4000);
}

// ── EXCEL IMPORT (Khordad Budget structure) ──────────────────────────────────
// Expected layout per the "Khordad Budget" sheet (0-indexed columns):
//   col1(B)=category name   col2(C)=budget total   col3(D)=spent
//   col6(G)=income name     col7(H)=income total
//   col9(J)=tx id  col10(K)=tx name  col11(L)=tx amount  col12(M)=type(want/need/saving)  col13(N)=tx category
let pendingWorkbook = null;
let pendingFileName = "";

function openImportModal() {
  openModal(
    `
    <div class="section-label" style="margin-bottom:18px;">Import from Excel</div>
    <p style="font-size:13px;color:var(--muted);line-height:1.5;margin-bottom:18px;">
      Upload your finance workbook. Sheets matching the <strong style="color:var(--text);">Budget</strong> structure
      (category totals, income sources, and a transaction list) will be parsed automatically.
    </p>
    <div class="upload-zone" id="uploadZone">
      <input type="file" id="excelFileInput" accept=".xlsx,.xls" onchange="handleFileSelect(event)">
      <div style="font-size:14px;font-weight:600;color:var(--jade);">📂 Choose .xlsx file</div>
      <div style="font-size:11px;color:var(--muted);margin-top:5px;">or drag &amp; drop here</div>
    </div>
    <div style="display:flex;gap:10px;margin-top:22px;">
      <button class="btn-ghost" style="flex:1;padding:13px;" onclick="closeModal()">Cancel</button>
    </div>
  `,
    true,
  );

  setTimeout(() => {
    const zone = document.getElementById("uploadZone");
    if (!zone) return;
    zone.addEventListener("dragover", (e) => {
      e.preventDefault();
      zone.classList.add("dragover");
    });
    zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
    zone.addEventListener("drop", (e) => {
      e.preventDefault();
      zone.classList.remove("dragover");
      const file = e.dataTransfer.files[0];
      if (file && /\.(xlsx|xls)$/i.test(file.name)) {
        readWorkbookFile(file);
      } else {
        showToast("Please drop an .xlsx or .xls file", "error");
      }
    });
  }, 0);
}

function handleFileSelect(e) {
  const file = e.target.files[0];
  if (!file) return;
  readWorkbookFile(file);
}

function readWorkbookFile(file) {
  pendingFileName = file.name;
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      pendingWorkbook = XLSX.read(ev.target.result, { type: "array" });
      const sheetNames = pendingWorkbook.SheetNames;
      if (sheetNames.length === 1) {
        importFromSheet(sheetNames[0]);
      } else {
        showSheetPicker(sheetNames);
      }
    } catch (err) {
      showToast("Could not read file: " + err.message, "error");
    }
  };
  reader.onerror = () => showToast("Failed to read file", "error");
  reader.readAsArrayBuffer(file);
}

function showSheetPicker(sheetNames) {
  const options = sheetNames
    .map((name) => {
      const lower = name.toLowerCase();
      let tag =
        '<span class="tag" style="background:var(--ink-line);color:var(--muted);">Other</span>';
      if (lower.includes("budget"))
        tag = '<span class="tag tag-jade">Budget</span>';
      else if (lower.includes("accounting"))
        tag =
          '<span class="tag" style="background:var(--brass-soft);color:var(--brass);">Ledger</span>';
      return `<div class="sheet-option" onclick="importFromSheet('${esc(name)}')">
      <span style="font-size:13px;font-weight:600;">${esc(name)}</span>
      ${tag}
    </div>`;
    })
    .join("");

  openModal(
    `
    <div class="section-label" style="margin-bottom:8px;">Multiple sheets found</div>
    <p style="font-size:13px;color:var(--muted);margin-bottom:18px;">Choose which sheet to import from <strong style="color:var(--text);">${esc(pendingFileName)}</strong>.</p>
    <div style="max-height:340px;overflow-y:auto;" class="scrollable">${options}</div>
    <div style="display:flex;gap:10px;margin-top:18px;">
      <button class="btn-ghost" style="flex:1;padding:13px;" onclick="closeModal()">Cancel</button>
    </div>
  `,
    true,
  );
}

const IMPORT_CAT_MAP = {
  "food & drinks": "food",
  "food and drinks": "food",
  food: "food",
  family: "family",
  "housing & construction": "housing",
  "housing and construction": "housing",
  home: "housing",
  "work & growth": "work",
  "work and growth": "work",
  "work expenses": "work",
  "project expenses": "project",
  project: "project",
  "saving & investment": "saving",
  "saving and investment": "saving",
  saving: "saving",
  transportation: "transport",
  transport: "transport",
  "loans & debts": "loans",
  "loans and debts": "loans",
  "debts and loans": "loans",
  "bills, internet, vpn": "bills",
  bills: "bills",
  internet: "bills",
  "charity & donation": "charity",
  "charity and donation": "charity",
  donation: "charity",
  "gift and donation": "charity",
  "health & beauty": "health",
  "health and beauty": "health",
  health: "health",
  clothing: "clothing",
  cloth: "clothing",
  banking: "work",
  "fun & entertainment": "work",
  charge: "bills",
  wedding: "family",
};

function matchImportCategory(raw) {
  if (!raw) return null;
  const key = String(raw).toLowerCase().trim();
  return IMPORT_CAT_MAP[key] || null;
}

function ensureCategory(catId, rawLabel) {
  let cat = state.categories.find((c) => c.id === catId);
  if (!cat) {
    cat = { id: catId, label: rawLabel || catId, target: 0 };
    state.categories.push(cat);
  }
  return cat;
}

function importFromSheet(sheetName) {
  closeModal();
  if (!pendingWorkbook) {
    showToast("No file loaded", "error");
    return;
  }
  const ws = pendingWorkbook.Sheets[sheetName];
  if (!ws) {
    showToast("Sheet not found: " + sheetName, "error");
    return;
  }

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  // Locate the header row: column 1 = "name", column 2 = "total"
  let headerIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 12); i++) {
    const r = rows[i];
    if (
      r &&
      String(r[1] || "")
        .toLowerCase()
        .trim() === "name" &&
      String(r[2] || "")
        .toLowerCase()
        .trim() === "total"
    ) {
      headerIdx = i;
      break;
    }
  }

  if (headerIdx < 0) {
    showToast(
      "This sheet doesn't match the expected Budget structure",
      "error",
    );
    return;
  }

  let newIncomeCount = 0,
    newExpenseCount = 0,
    updatedBudgets = 0;
  const newIncomeMap = {};

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;

    // Budget categories: col B (idx 1) name, col C (idx 2) total budget
    const catName = r[1];
    const catTotal = parseFloat(r[2]);
    if (catName && !isNaN(catTotal) && catTotal > 0) {
      const lowerCat = String(catName).toLowerCase().trim();
      const isSummaryRow =
        lowerCat === "want:" ||
        lowerCat === "need:" ||
        lowerCat === "saving:" ||
        lowerCat === "balance:";
      if (!isSummaryRow) {
        const catId = matchImportCategory(catName);
        if (catId) {
          const cat = ensureCategory(catId, catName);
          cat.target = catTotal;
          updatedBudgets++;
        }
      }
    }

    // Income sources: col G (idx 6) name, col H (idx 7) total
    const incName = r[6];
    const incAmount = parseFloat(r[7]);
    if (incName && !isNaN(incAmount) && incAmount > 0) {
      const lowerName = String(incName).toLowerCase().trim();
      if (
        lowerName !== "total icome:" &&
        lowerName !== "total income:" &&
        lowerName !== "goal:"
      ) {
        const id = lowerName.replace(/\s+/g, "_");
        newIncomeMap[id] = {
          id,
          name: String(incName),
          amount: incAmount,
        };
        newIncomeCount++;
      }
    }

    // Transaction list: col J (idx 9) id, col K (idx 10) name, col L (idx 11) amount, col M (idx 12) type, col N (idx 13) category
    const txName = r[10];
    const txAmount = parseFloat(r[11]);
    const txCategory = r[13];
    if (
      txName &&
      !isNaN(txAmount) &&
      txAmount > 0 &&
      String(txName).toLowerCase().trim() !== "name"
    ) {
      let catId = matchImportCategory(txCategory);
      if (!catId)
        catId = ensureCategory(
          "cat_" +
            String(txCategory || "misc")
              .toLowerCase()
              .replace(/[^a-z0-9]+/g, "_"),
          txCategory,
        ).id;
      state.expenses.push({
        id:
          Date.now().toString() + "_" + Math.random().toString(36).slice(2, 8),
        categoryId: catId,
        desc: String(txName),
        amount: txAmount,
        date: sheetName,
      });
      newExpenseCount++;
    }
  }

  // Merge income streams: update existing, add new
  const existingIds = new Set(state.incomes.map((i) => i.id));
  Object.values(newIncomeMap).forEach((inc) => {
    if (existingIds.has(inc.id)) {
      const ex = state.incomes.find((i) => i.id === inc.id);
      ex.amount = inc.amount;
      ex.name = inc.name;
    } else {
      state.incomes.push(inc);
    }
  });

  save();
  render();

  if (newExpenseCount === 0 && newIncomeCount === 0 && updatedBudgets === 0) {
    showToast('No matching data found in "' + sheetName + '"', "error");
  } else {
    showToast(
      `✓ Imported "${sheetName}": ${newExpenseCount} transactions, ${newIncomeCount} income streams, ${updatedBudgets} budgets`,
      "success",
    );
  }
}

// ── ACTIONS ───────────────────────────────────────────────────────────────────
function saveIncome() {
  const name = document.getElementById("incomeName").value.trim();
  const amount = parseFloat(document.getElementById("incomeAmount").value);
  if (!name || isNaN(amount) || amount <= 0)
    return alert("Enter a name and valid amount.");
  const id = name.toLowerCase().replace(/\s+/g, "_");
  const existing = state.incomes.find((i) => i.id === id);
  if (existing) {
    existing.amount = amount;
  } else {
    state.incomes.push({ id, name, amount });
  }
  document.getElementById("incomeName").value = "";
  document.getElementById("incomeAmount").value = "";
  save();
  render();
}

function deleteIncome(id) {
  state.incomes = state.incomes.filter((i) => i.id !== id);
  save();
  render();
}

function saveExpense() {
  const catId = document.getElementById("expenseCategory").value;
  const desc = document.getElementById("expenseDesc").value.trim();
  const amount = parseFloat(document.getElementById("expenseAmount").value);
  if (!catId) return alert("Select a category.");
  if (isNaN(amount) || amount <= 0) return alert("Enter a valid amount.");
  state.expenses.push({
    id: Date.now().toString(),
    categoryId: catId,
    desc: desc || "",
    amount,
    date: new Date().toLocaleDateString("en-US", {
      month: "short",
      day: "numeric",
    }),
  });
  document.getElementById("expenseCategory").value = "";
  document.getElementById("expenseDesc").value = "";
  document.getElementById("expenseAmount").value = "";
  save();
  render();
}

function deleteExpense(id) {
  state.expenses = state.expenses.filter((e) => e.id !== id);
  save();
  render();
}

function saveNewBudget() {
  const emoji = document.getElementById("newCatEmoji").value.trim();
  const name = document.getElementById("newCatName").value.trim();
  const target = parseFloat(document.getElementById("newCatTarget").value);
  if (!name) return alert("Enter a category name.");
  if (isNaN(target) || target <= 0)
    return alert("Enter a valid budget amount.");
  const id = "cat_" + Date.now().toString(36);
  const label = emoji ? `${emoji} ${name}` : name;
  state.categories.push({ id, label, target });
  closeModal();
  save();
  render();
}

function saveEditedIncome(id) {
  const name = document.getElementById("editIncomeName").value.trim();
  const amount = parseFloat(document.getElementById("editIncomeAmount").value);
  if (!name || isNaN(amount) || amount <= 0)
    return alert("Enter a name and valid amount.");
  const inc = state.incomes.find((i) => i.id === id);
  if (inc) {
    inc.name = name;
    inc.amount = amount;
  }
  closeModal();
  save();
  render();
}

function saveEditedExpense(id) {
  const catId = document.getElementById("editExpenseCategory").value;
  const desc = document.getElementById("editExpenseDesc").value.trim();
  const amount = parseFloat(document.getElementById("editExpenseAmount").value);
  if (!catId) return alert("Select a category.");
  if (isNaN(amount) || amount <= 0) return alert("Enter a valid amount.");
  const exp = state.expenses.find((e) => e.id === id);
  if (exp) {
    exp.categoryId = catId;
    exp.desc = desc;
    exp.amount = amount;
  }
  closeModal();
  save();
  render();
}

// ── CONTEXT MENU ──────────────────────────────────────────────────────────────
let activeContextTarget = null;

function showContextMenu(e, type, id) {
  e.stopPropagation();
  closeContextMenu();
  activeContextTarget = { type, id };
  const menu = document.createElement("div");
  menu.className = "context-menu";
  menu.id = "activeContextMenu";
  menu.innerHTML = `
    <div class="context-menu-item" onclick="handleContextEdit()">✏️ Edit</div>
    <div class="context-menu-divider"></div>
    <div class="context-menu-item danger" onclick="handleContextDelete()">🗑️ Delete</div>
  `;
  document.body.appendChild(menu);
  const pad = 8;
  const rect = menu.getBoundingClientRect();
  let x = e.clientX,
    y = e.clientY;
  if (x + rect.width + pad > window.innerWidth)
    x = window.innerWidth - rect.width - pad;
  if (y + rect.height + pad > window.innerHeight)
    y = window.innerHeight - rect.height - pad;
  menu.style.left = x + "px";
  menu.style.top = y + "px";
}

function closeContextMenu() {
  const m = document.getElementById("activeContextMenu");
  if (m) m.remove();
  activeContextTarget = null;
}

function handleContextEdit() {
  if (!activeContextTarget) return;
  const { type, id } = activeContextTarget;
  closeContextMenu();
  if (type === "income") openEditIncomeModal(id);
  else if (type === "expense") openEditExpenseModal(id);
}

function handleContextDelete() {
  if (!activeContextTarget) return;
  const { type, id } = activeContextTarget;
  closeContextMenu();
  if (type === "income") deleteIncome(id);
  else if (type === "expense") deleteExpense(id);
}

document.addEventListener("click", closeContextMenu);
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    closeContextMenu();
    closeModal();
  }
});

// ── MODALS ────────────────────────────────────────────────────────────────────
function openModal(html, wide) {
  closeContextMenu();
  closeModal();
  const overlay = document.createElement("div");
  overlay.id = "modalOverlay";
  overlay.className = "modal-overlay";
  overlay.innerHTML = `<div class="modal-box${wide ? " wide" : ""}" onclick="event.stopPropagation()">${html}</div>`;
  overlay.onclick = closeModal;
  document.body.appendChild(overlay);
}

function closeModal() {
  const overlay = document.getElementById("modalOverlay");
  if (overlay) overlay.remove();
}

function openAddBudgetModal() {
  openModal(`
    <div class="section-label" style="margin-bottom:18px;">Add Budget Category</div>
    <input type="text" id="newCatEmoji" placeholder="Emoji (optional, e.g. 🎮)" class="mb-3" maxlength="4">
    <input type="text" id="newCatName" placeholder="Category name (e.g. Entertainment)" class="mb-3">
    <input type="number" id="newCatTarget" placeholder="Monthly budget in Toman" class="mb-5">
    <div style="display:flex;gap:10px;">
      <button class="btn-ghost" style="flex:1;padding:13px;" onclick="closeModal()">Cancel</button>
      <button class="btn-primary" style="flex:1;" onclick="saveNewBudget()">Add Category</button>
    </div>
  `);
}

function openEditIncomeModal(id) {
  const inc = state.incomes.find((i) => i.id === id);
  if (!inc) return;
  openModal(`
    <div class="section-label" style="margin-bottom:18px;">Edit Income</div>
    <input type="text" id="editIncomeName" value="${esc(inc.name)}" class="mb-3" placeholder="Stream name">
    <input type="number" id="editIncomeAmount" value="${inc.amount}" class="mb-5" placeholder="Amount in Toman">
    <div style="display:flex;gap:10px;">
      <button class="btn-ghost" style="flex:1;padding:13px;" onclick="closeModal()">Cancel</button>
      <button class="btn-primary" style="flex:1;" onclick="saveEditedIncome('${id}')">Save Changes</button>
    </div>
  `);
}

function openEditExpenseModal(id) {
  const exp = state.expenses.find((e) => e.id === id);
  if (!exp) return;
  const catOptions = state.categories
    .map(
      (c) =>
        `<option value="${c.id}" ${c.id === exp.categoryId ? "selected" : ""}>${c.label}</option>`,
    )
    .join("");
  openModal(`
    <div class="section-label" style="margin-bottom:18px;">Edit Expense</div>
    <select id="editExpenseCategory" class="mb-3">${catOptions}</select>
    <input type="text" id="editExpenseDesc" value="${esc(exp.desc || "")}" class="mb-3" placeholder="Description (optional)">
    <input type="number" id="editExpenseAmount" value="${exp.amount}" class="mb-5" placeholder="Amount in Toman">
    <div style="display:flex;gap:10px;">
      <button class="btn-ghost" style="flex:1;padding:13px;" onclick="closeModal()">Cancel</button>
      <button class="btn-primary" style="flex:1;" onclick="saveEditedExpense('${id}')">Save Changes</button>
    </div>
  `);
}

// ── PERFORATION STRIP ─────────────────────────────────────────────────────────
function renderPerfStrip() {
  const row = document.getElementById("perfRow");
  if (!row) return;
  const width = row.parentElement.clientWidth || 1200;
  const count = Math.max(8, Math.floor(width / 16));
  row.innerHTML = Array(count).fill('<div class="perf-dot"></div>').join("");
}

// ── SCROLL REVEAL ─────────────────────────────────────────────────────────────
function initReveal() {
  const els = document.querySelectorAll(".reveal");
  const io = new IntersectionObserver(
    (entries) => {
      entries.forEach((e) => {
        if (e.isIntersecting) {
          e.target.classList.add("in");
          io.unobserve(e.target);
        }
      });
    },
    { threshold: 0.1 },
  );
  els.forEach((el) => io.observe(el));
}

// ── RENDER ────────────────────────────────────────────────────────────────────
let incomeChart = null,
  spendChart = null;

function render() {
  const totalIncome = state.incomes.reduce((s, i) => s + i.amount, 0);
  const totalBudget = state.categories.reduce((s, c) => s + c.target, 0);
  const totalSpent = state.expenses.reduce((s, e) => s + e.amount, 0);
  const balance = totalIncome - totalSpent;

  // Metrics
  document.getElementById("metricIncome").textContent = fmt(totalIncome);
  document.getElementById("metricBudget").textContent = fmt(totalBudget);
  document.getElementById("metricSpent").textContent = fmt(totalSpent);
  const balEl = document.getElementById("metricBalance");
  balEl.textContent = (balance < 0 ? "−" : "") + fmt(balance);
  balEl.style.color = balance < 0 ? "#FF9577" : "#4FE3AC";

  // Income list
  const il = document.getElementById("incomeList");
  il.innerHTML =
    state.incomes.length === 0
      ? `<div style="color:var(--muted-2);font-size:12px;text-align:center;padding:20px;">No income streams yet</div>`
      : state.incomes
          .map(
            (inc) => `
      <div class="fade-in row-item" style="padding:13px 14px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--ink-line);" onclick="showContextMenu(event,'income','${inc.id}')">
        <div>
          <div style="font-size:13px;font-weight:600;">${esc(inc.name)}</div>
        </div>
        <div style="display:flex;align-items:center;gap:10px;">
          <span class="num" style="font-size:13px;font-weight:700;color:#4FE3AC;">${fmt(inc.amount)}</span>
        </div>
      </div>`,
          )
          .join("");

  // Category select
  const sel = document.getElementById("expenseCategory");
  const cur = sel.value;
  sel.innerHTML =
    '<option value="">— Select Category —</option>' +
    state.categories
      .map(
        (c) =>
          `<option value="${c.id}" ${c.id === cur ? "selected" : ""}>${c.label}</option>`,
      )
      .join("");

  // Transaction table
  const tbody = document.getElementById("txBody");
  const sorted = [...state.expenses].reverse();
  if (sorted.length === 0) {
    tbody.innerHTML = `<tr><td colspan="4" style="color:var(--muted-2);font-size:12px;padding:16px 0;text-align:center;">No expenses logged yet</td></tr>`;
  } else {
    tbody.innerHTML = sorted
      .map((e) => {
        const cat = state.categories.find((c) => c.id === e.categoryId);
        return `<tr class="fade-in row-item" style="border-top:1px solid var(--ink-line);" onclick="showContextMenu(event,'expense','${e.id}')">
        <td style="padding:11px 0;font-size:11px;color:var(--muted);">${cat ? esc(cat.label.substring(0, 14)) : esc(e.categoryId)}</td>
        <td style="padding:11px 4px;font-size:11px;max-width:80px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${esc(e.desc) || "—"}</td>
        <td class="num" style="padding:11px 0;text-align:right;font-size:12px;">${fmt(e.amount)}</td>
        <td style="padding:11px 0;text-align:right;color:var(--muted-2);font-size:14px;">⋯</td>
      </tr>`;
      })
      .join("");
  }

  // Budget tracker
  const bt = document.getElementById("budgetTracker");
  bt.innerHTML = state.categories
    .map((cat) => {
      const spent = state.expenses
        .filter((e) => e.categoryId === cat.id)
        .reduce((s, e) => s + e.amount, 0);
      const pct = cat.target > 0 ? (spent / cat.target) * 100 : 0;
      const pctDisplay = Math.round(pct);
      const fillPct = Math.min(pct, 100);
      const over = spent > cat.target;
      const ovBy = spent - cat.target;
      return `
    <div style="margin-bottom:22px;">
      <div style="display:flex;align-items:baseline;justify-content:space-between;margin-bottom:9px;gap:10px;">
        <span style="font-size:12.5px;font-weight:600;">${esc(cat.label)}</span>
        <span class="num" style="font-size:11px;color:${over ? "var(--rust)" : "var(--muted)"};white-space:nowrap;font-weight:600;">${pctDisplay}% / ${fmt(cat.target)}</span>
      </div>
      <div class="progress-track">
        <div class="progress-fill" style="width:${fillPct}%;background:${over ? "var(--rust)" : "var(--jade)"};"></div>
      </div>
      ${over ? `<div style="display:flex;justify-content:flex-end;margin-top:9px;"><span class="tag tag-rust">↑ ${fmt(ovBy)} over</span></div>` : ""}
    </div>`;
    })
    .join("");

  // Charts
  try {
    renderCharts(totalIncome);
  } catch (err) {
    console.error("Chart render failed:", err);
  }
}

function renderCharts(totalIncome) {
  // Income donut
  const incCtx = document.getElementById("incomeChart").getContext("2d");
  const incomeColors = [
    "#1FAE7E",
    "#C9A35F",
    "#E2633E",
    "#5C8AE6",
    "#9B6DFF",
    "#06B6D4",
    "#F0B429",
    "#4ADE80",
  ];
  if (incomeChart) incomeChart.destroy();

  if (state.incomes.length === 0) {
    incomeChart = null;
  } else {
    incomeChart = new Chart(incCtx, {
      type: "doughnut",
      data: {
        labels: state.incomes.map((i) => i.name),
        datasets: [
          {
            data: state.incomes.map((i) => i.amount),
            backgroundColor: incomeColors.slice(0, state.incomes.length),
            borderColor: "#151B26",
            borderWidth: 3,
            hoverOffset: 8,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        cutout: "68%",
        plugins: {
          legend: {
            position: "bottom",
            labels: {
              color: "#7C8493",
              font: { family: "Inter", size: 11 },
              padding: 14,
              usePointStyle: true,
              pointStyle: "circle",
            },
          },
          tooltip: {
            backgroundColor: "#0A0E14",
            borderColor: "#232B3A",
            borderWidth: 1,
            padding: 12,
            titleFont: { family: "Inter", weight: "600" },
            bodyFont: { family: "JetBrains Mono" },
            callbacks: {
              label: (ctx) =>
                ` ${ctx.label}: ${ctx.raw.toLocaleString()} Toman`,
            },
          },
        },
      },
    });
  }

  // Spending bar chart
  const spCtx = document.getElementById("spendChart").getContext("2d");
  if (spendChart) spendChart.destroy();

  const catTotals = state.categories
    .map((cat) => ({
      label: cat.label,
      spent: state.expenses
        .filter((e) => e.categoryId === cat.id)
        .reduce((s, e) => s + e.amount, 0),
      target: cat.target,
    }))
    .filter((c) => c.spent > 0)
    .sort((a, b) => b.spent - a.spent);

  if (catTotals.length === 0) {
    spendChart = null;
  } else {
    spendChart = new Chart(spCtx, {
      type: "bar",
      data: {
        labels: catTotals.map((c) => c.label),
        datasets: [
          {
            label: "Spent",
            data: catTotals.map((c) => c.spent),
            backgroundColor: catTotals.map((c) =>
              c.spent > c.target
                ? "rgba(226,99,62,0.75)"
                : "rgba(31,174,126,0.75)",
            ),
            borderRadius: 3,
            borderSkipped: false,
            barThickness: 16,
          },
          {
            label: "Budget",
            data: catTotals.map((c) => c.target),
            backgroundColor: "rgba(124,132,147,0.18)",
            borderRadius: 3,
            borderSkipped: false,
            barThickness: 16,
          },
        ],
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            labels: {
              color: "#7C8493",
              font: { family: "Inter", size: 11 },
              usePointStyle: true,
              pointStyle: "rect",
            },
          },
          tooltip: {
            backgroundColor: "#0A0E14",
            borderColor: "#232B3A",
            borderWidth: 1,
            padding: 12,
            titleFont: { family: "Inter", weight: "600" },
            bodyFont: { family: "JetBrains Mono" },
            callbacks: {
              label: (ctx) =>
                ` ${ctx.dataset.label}: ${ctx.raw.toLocaleString()} Toman`,
            },
          },
        },
        scales: {
          x: {
            ticks: {
              color: "#4D5563",
              font: { size: 10, family: "JetBrains Mono" },
            },
            grid: { color: "#1A212E" },
          },
          y: {
            ticks: {
              color: "#7C8493",
              font: { size: 11, family: "Inter" },
            },
            grid: { color: "transparent" },
          },
        },
      },
    });
  }
}

// ── INIT ──────────────────────────────────────────────────────────────────────
load();
render();
renderPerfStrip();
window.addEventListener("resize", renderPerfStrip);
initReveal();
