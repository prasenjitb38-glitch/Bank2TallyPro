let rows = [];
let statementSummaries = [];
let statementSequence = 0;
let filter = "All";
const forceFreshStatement = new URLSearchParams(window.location.search).get("fresh") === "1";
const $ = selector => document.querySelector(selector);
document.title = "Bank2Tally Suite v1.7.0";
const brandNode = document.querySelector(".brand");
if (brandNode) brandNode.innerHTML = 'Bank<span>2</span>Tally <small>Suite</small>';
const brandTagline = document.querySelector(".brand-area p");
if (brandTagline) brandTagline.textContent = "Offline statement review, GST reconciliation and TallyPrime preparation";
const escapeHtml = value => String(value ?? "").replace(/[&<>"']/g, character => ({
  "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
}[character]));
const input = $("#fileInput");
const card = $("#uploadCard");
const message = $("#message");
const bankLedgerInput = $("#bankLedger");
const processingStatus = $("#processingStatus");
const processingText = $("#processingText");
const processingPercent = $("#processingPercent");
const progressBar = $("#progressBar");
let processingTimer = null;
let licenseStatus = null;
let selectedPlan = "200 pages - Rs 100";
let selectedCoupon = "";
let currentAccount = {};
let accountAlreadyRegistered = false;
let gstRows = [];
let gstDatasets = {};
const itcTallyExcluded = new Set();
const itcDifferenceInitialized = new Set();
const itcDifferenceOpenMonths = new Set();
let gstSalesOriginalRows = [];
let tallyMasters = { connected: false, company: "", ledgers: [], items: [] };
let hsnMasterRows = [];
let companyStatementRows = [];
let companyReconcileRows = [];
let activeGstModule = "reconciliation";
let gstReconRows = [];
let gstReconTallyRows = [];
let gstReconResults = [];
let gstReconGstr1Rows = [];
let gstReconTallySalesRows = [];
let gstReconSalesResults = [];
let salesReconDashboard = null;
let gstr3bDashboard = null;
let activeReconTab = "overview";
let salesReconPage = 1;
let gstReconSharedPeriod = "ALL";
let gstReconPeriodUserChosen = false;
const SALES_RECON_PAGE_SIZE = 100;
const PURCHASE_RECON_PAGE_SIZE = 100;
const PURCHASE_2A_PAGE_SIZE = 100;
let purchaseReconPages = { matched: 1, only2a: 1, only2b: 1, notes: 1, mismatch: 1 };
let purchase2aPages = { b2b: 1, credit: 1, debit: 1, amendment: 1 };
let purchase2aFilters = {
  b2b: { q: "", rate: "" },
  credit: { q: "", rate: "" },
  debit: { q: "", rate: "" },
  amendment: { q: "", rate: "" },
};
let gstReconDatasetCounts = {};
let gstReconDatasetsLoaded = new Set();
let gstReconSessionLoadSeq = 0;

if (forceFreshStatement) {
  window.addEventListener("pageshow", () => {
    rows = [];
    statementSummaries = [];
    statementSequence = 0;
    history.replaceState(null, "", window.location.pathname);
    render();
    showMessage("New statement session ready.");
  }, { once: true });
}

function ensureGstReconPanelVisible() {
  if (activeGstModule !== "threeway") return;
  const panel = $("#gstThreeWayPanel");
  if (panel) panel.classList.remove("hidden");
  if ($("#reconOverviewPane") && activeReconTab === "overview") {
    $("#reconOverviewPane").classList.remove("hidden");
  }
  // Keep recon tabs in view after long XLSX imports (file picker / layout can leave scroll at top).
  const dialog = $("#gstDialog");
  if (panel && dialog && dialog.open) {
    try {
      const panelTop = panel.getBoundingClientRect().top;
      const dialogRect = dialog.getBoundingClientRect();
      if (panelTop > dialogRect.bottom - 80 || panelTop < dialogRect.top + 40) {
        panel.scrollIntoView({ block: "nearest", inline: "nearest" });
      }
    } catch (_) {}
  }
}

function debounce(fn, waitMs = 250) {
  let timer = 0;
  return (...args) => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => fn(...args), waitMs);
  };
}

function purchasePagerHtml(sectionKey, page, pages, total) {
  return `<div class="sales-recon-pager purchase-recon-pager" data-purchase-page-section="${escapeHtml(sectionKey)}">
    <button type="button" class="secondary purchase-page-prev" ${page <= 1 ? "disabled" : ""}>Previous</button>
    <span>Page ${page} / ${pages} (${Number(total || 0).toLocaleString("en-IN")} rows)</span>
    <button type="button" class="secondary purchase-page-next" ${page >= pages ? "disabled" : ""}>Next</button>
  </div>`;
}

function paginateEntries(entries, page, pageSize) {
  const total = entries.length;
  const pages = Math.max(1, Math.ceil(total / pageSize));
  const safePage = Math.min(Math.max(1, page || 1), pages);
  const start = (safePage - 1) * pageSize;
  return {
    page: safePage,
    pages,
    total,
    start,
    pageEntries: entries.slice(start, start + pageSize),
  };
}

function normalizeGstUiPeriod(value) {
  const raw = String(value || "").trim();
  if (!raw) return "ALL";
  const upper = raw.toUpperCase();
  if (upper === "ALL" || upper.startsWith("ALL") || upper.includes("FY 2025")) return "ALL";
  const digits = raw.replace(/\D/g, "").slice(0, 6);
  return digits.length === 6 ? digits : "ALL";
}

function getGstReconPeriod() {
  return normalizeGstUiPeriod(
    gstReconSharedPeriod
    || $("#reconPeriodFilter")?.value
    || $("#salesReconPeriod")?.value
    || $("#gstr3bReconPeriod")?.value
    || "ALL"
  );
}

function setGstReconPeriod(value, options = {}) {
  const period = normalizeGstUiPeriod(value);
  if (options.userChosen) gstReconPeriodUserChosen = true;
  gstReconSharedPeriod = period;
  ["reconPeriodFilter", "salesReconPeriod", "gstr3bReconPeriod"].forEach((id) => {
    const node = document.getElementById(id);
    if (!node) return;
    if (node.tagName === "SELECT") {
      const has = Array.from(node.options).some((option) => option.value === period);
      node.value = has ? period : "ALL";
      if (!has) gstReconSharedPeriod = "ALL";
    } else {
      node.value = period === "ALL" ? "ALL" : period;
    }
  });
  if (options.silent || options.refresh === false) return gstReconSharedPeriod;
  refreshActiveReconPeriodViews(false);
  return gstReconSharedPeriod;
}

function refreshActiveReconPeriodViews(showError = false) {
  // Period change: refresh only the visible recon tab (avoid triple heavy API calls).
  if (activeGstModule !== "threeway") return;
  if (activeReconTab === "sales") {
    salesReconPage = 1;
    gstReconSalesResults = [];
    refreshSalesReconDashboard(showError);
    return;
  }
  if (activeReconTab === "gstr3b") {
    refreshGstr3bDashboard(showError);
    return;
  }
  if (activeReconTab === "purchase") {
    if (typeof refreshItcDashboard === "function") refreshItcDashboard(false);
    return;
  }
  refreshGstReconOverview(showError);
}

$("#licenseKey").maxLength = 24;
$("#licenseKey").placeholder = "XXXX-XXXX-XXXX-XXXX-XXXX";
$("#licenseKey").parentElement.childNodes[0].textContent = "Activation Code";
$(".license-separator span").textContent = "Already have an Activation Code?";

async function refreshLicense() {
  const response = await fetch("/api/license/status", { method: "POST", headers: { "Content-Type": "application/json" }, body: "{}" });
  licenseStatus = await response.json();
  $("#creditRemaining").textContent = Number(licenseStatus.remaining || 0).toLocaleString("en-IN");
  $("#licensePlan").textContent = licenseStatus.plan || "Trial";
  $("#deviceId").textContent = licenseStatus.device_id || "";
}

function showLoggedInCustomer(account) {
  currentAccount = account || {};
  $("#customerName").textContent = account.customer_name || "";
  $("#customerBusiness").textContent = account.business_name || "";
  $("#customerBadge").classList.remove("hidden");
  $("#logoutBtn").classList.remove("hidden");
}

async function initializeAccount() {
  const response = await fetch("/api/account/status", { method: "POST", headers: { "Content-Type": "application/json" }, body: "{}" });
  const state = await response.json();
  if (state.logged_in) {
    showLoggedInCustomer(state.account || {});
    await refreshLicense();
    return;
  }
  const registered = state.registered === true;
  accountAlreadyRegistered = registered;
  $("#accountDialog").classList.toggle("registered-account", registered);
  $("#accountTitle").textContent = registered ? "Welcome Back" : "Create Customer Account";
  $("#accountSubtitle").textContent = registered ? "Enter your PIN to open Bank2Tally Suite." : "Register once to start your 50-credit trial.";
  $("#registrationFields").classList.toggle("hidden", registered);
  $("#accountSubmit").textContent = registered ? "Login to Bank2Tally Suite" : "Create Account & Start Trial";
  $("#accountPin").value = "";
  $("#accountError").classList.add("hidden");
  setAccountMode(registered ? "login" : "signup");
  $("#accountDialog").showModal();
  $("#accountPin").focus();
}

function setAccountMode(mode) {
  const signup = mode === "signup";
  if (signup && accountAlreadyRegistered) {
    $("#accountError").textContent = "An account is already registered on this computer. Please Login.";
    $("#accountError").classList.remove("hidden");
    return false;
  }
  $("#registrationFields").classList.toggle("hidden", !signup);
  $("#accountTitle").textContent = signup ? "Create Customer Account" : "Welcome Back";
  $("#accountSubtitle").textContent = signup ? "Register once to start your trial." : "Enter your PIN to open Bank2Tally Suite.";
  $("#accountSubmit").textContent = signup ? "Create Account & Start Trial" : "Login to Bank2Tally Suite";
  $("#accountLoginBtn").classList.toggle("active", !signup);
  $("#accountSignupBtn").classList.toggle("active", signup);
  $("#accountError").classList.add("hidden");
  return true;
}

$("#accountLoginBtn").onclick = () => {
  if (!accountAlreadyRegistered) {
    $("#accountError").textContent = "No customer account found. Please Sign Up first.";
    $("#accountError").classList.remove("hidden");
    return;
  }
  setAccountMode("login");
  $("#accountSubmit").click();
};

$("#accountSignupBtn").onclick = () => {
  if (!setAccountMode("signup")) return;
  const complete = $("#accountCustomerName").value.trim() && $("#accountMobile").value.trim()
    && $("#accountBusinessName").value.trim() && $("#accountPin").value.trim();
  if (complete) $("#accountSubmit").click();
  else $("#accountCustomerName").focus();
};

document.querySelectorAll("[data-dashboard-open]").forEach(button => {
  button.addEventListener("click", () => {
    const target = $(button.dataset.dashboardOpen);
    if (target) target.click();
  });
});

$("#accountSubmit").onclick = async () => {
  const registered = $("#registrationFields").classList.contains("hidden");
  const body = registered
    ? { pin: $("#accountPin").value }
    : { customerName: $("#accountCustomerName").value, mobile: $("#accountMobile").value, businessName: $("#accountBusinessName").value, pin: $("#accountPin").value };
  const response = await fetch(registered ? "/api/account/login" : "/api/account/register", {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body)
  });
  const result = await response.json();
  if (!response.ok) {
    $("#accountError").textContent = result.error || "Login failed.";
    $("#accountError").classList.remove("hidden");
    return;
  }
  $("#accountDialog").close();
  showLoggedInCustomer(result.account || {});
  // Fresh login: drop any in-browser portal import cache from a previous session.
  gstDatasets["GSTR-2B"] = undefined;
  gstDatasets["GSTR-1"] = undefined;
  gstDatasets["GSTR-3B"] = undefined;
  gstReconPortalPageReady = false;
  try { resetGstReconWorkspace(); } catch (_) {}
  await refreshLicense();
};
$("#accountPin").addEventListener("keydown", event => {
  if (event.key === "Enter") { event.preventDefault(); $("#accountSubmit").click(); }
});
$("#accountDialog").addEventListener("cancel", event => event.preventDefault());
$("#logoutBtn").onclick = async () => {
  await fetch("/api/account/logout", { method: "POST", headers: { "Content-Type": "application/json" }, body: "{}" });
  $("#customerBadge").classList.add("hidden");
  $("#logoutBtn").classList.add("hidden");
  await initializeAccount();
};
initializeAccount().catch(error => showMessage(error.message, true));
$("#licenseBtn").onclick = async () => {
  await refreshLicense();
  $("#licenseKey").value = "";
  $("#licenseError").classList.add("hidden");
  selectedPlan = "200 pages - Rs 100";
  selectedCoupon = "";
  $("#couponCode").value = "";
  document.querySelectorAll(".plan-option").forEach(button => button.classList.toggle("selected", button.dataset.plan === selectedPlan));
  $("#paymentArea").classList.add("hidden");
  $("#paymentNote").classList.add("hidden");
  $("#whatsappBtn").classList.add("hidden");
  $("#paymentDoneBtn").classList.remove("hidden");
  $("#proceedPaymentBtn").classList.remove("hidden");
  $("#licenseDialog").showModal();
};
document.querySelectorAll(".plan-option").forEach(button => button.onclick = () => {
  selectedPlan = button.dataset.plan;
  document.querySelectorAll(".plan-option").forEach(item => item.classList.toggle("selected", item === button));
  $("#paymentArea").classList.add("hidden");
  $("#paymentNote").classList.add("hidden");
  $("#whatsappBtn").classList.add("hidden");
  $("#paymentDoneBtn").classList.remove("hidden");
  $("#proceedPaymentBtn").classList.remove("hidden");
});
$("#proceedPaymentBtn").onclick = () => {
  selectedCoupon = $("#couponCode").value.trim().toUpperCase();
  $("#paymentPlan").textContent = selectedPlan;
  $("#paymentCoupon").textContent = selectedCoupon ? `Coupon: ${selectedCoupon} — seller will confirm the discount` : "";
  $("#paymentCoupon").classList.toggle("hidden", !selectedCoupon);
  $("#paymentArea").classList.remove("hidden");
  $("#proceedPaymentBtn").classList.add("hidden");
};
$("#paymentDoneBtn").onclick = () => {
  $("#whatsappBtn").classList.remove("hidden");
  $("#paymentNote").classList.remove("hidden");
  $("#paymentDoneBtn").classList.add("hidden");
};
$("#copyDeviceBtn").onclick = async () => {
  await navigator.clipboard.writeText($("#deviceId").textContent);
  $("#copyDeviceBtn").textContent = "Copied";
  setTimeout(() => $("#copyDeviceBtn").textContent = "Copy Device ID", 1200);
};
$("#copyUpiBtn").onclick = async () => {
  await navigator.clipboard.writeText($("#upiId").textContent);
  $("#copyUpiBtn").textContent = "UPI ID Copied";
  setTimeout(() => $("#copyUpiBtn").textContent = "Copy UPI ID", 1200);
};
$("#whatsappBtn").onclick = () => {
  const couponLine = selectedCoupon ? `%0ACoupon Code: ${encodeURIComponent(selectedCoupon)}` : "";
  const customerLine = `%0ACustomer Name: ${encodeURIComponent(currentAccount.customer_name || "")}%0AMobile Number: ${encodeURIComponent(currentAccount.mobile || "")}%0ABusiness Name: ${encodeURIComponent(currentAccount.business_name || "")}`;
  const text = `Bank2Tally credit purchase${customerLine}%0ASelected plan: ${encodeURIComponent(selectedPlan)}${couponLine}%0ADevice ID: ${encodeURIComponent($("#deviceId").textContent)}%0AUPI ID: pinkub0000-1@okaxis%0AI will attach my payment screenshot.`;
  window.open(`https://wa.me/919508773595?text=${text}`, "_blank", "noopener");
};
$("#activateLicenseBtn").onclick = async event => {
  event.preventDefault();
  const response = await fetch("/api/license/activate", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ licenseKey: $("#licenseKey").value.trim() })
  });
  const result = await response.json();
  if (!response.ok) {
    $("#licenseError").textContent = result.error || "License activation failed.";
    $("#licenseError").classList.remove("hidden");
    return;
  }
  $("#licenseDialog").close();
  await refreshLicense();
  showMessage(`License activated. ${result.remaining} credits available.`);
};

$("#refreshBtn").onclick = () => {
  rows = [];
  statementSummaries = [];
  statementSequence = 0;
  $("#fileInput").value = "";
  render();
  showMessage("New statement session ready.");
  window.scrollTo({ top: 0, behavior: "smooth" });
};
$("#gstRefreshBtn").onclick = async event => {
  event.preventDefault();
  event.stopPropagation();
  const module = activeGstModule;
  resetGstModuleWorkspace();
  updateGstWorkflowGuide();
  [0,5,12,18,28].forEach(rate => {
    $(`#salesLess${rate}`).value = "0";
    $(`#salesAdd${rate}`).value = "0";
  });
  if (module === "threeway") {
    // Force a clean portal session — do not reload old SQLite GSTR files after Refresh.
    try { await resetGstReconPortalSession(true); } catch (_) {}
    try { await loadGstReconSession({ restorePortal: false }); } catch (_) {}
  }
  try { await loadHsnMaster(); } catch (_) {}
};
async function closeBank2Tally(button, askConfirmation = true) {
  if (askConfirmation && !confirm("Close Bank2Tally? Unsaved review changes will be lost.")) return;
  button.disabled = true;
  button.textContent = "Closing...";
  try {
    await fetch("/api/shutdown", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: "{}"
    });
  } catch (_) {
    // The local server may close before the browser receives the response.
  }
  document.body.innerHTML = `<main class="closed-screen"><div><h1>Bank2Tally closed</h1><p>You can close this browser tab now.</p></div></main>`;
  setTimeout(() => window.close(), 250);
}
$("#closeBtn").onclick = () => closeBank2Tally($("#closeBtn"));
$("#accountCloseBtn").onclick = () => closeBank2Tally($("#accountCloseBtn"), false);

$("#openBankModule").onclick = () => {
  $("#moduleDashboard").classList.add("hidden");
  $("#bankModule").classList.remove("hidden");
};
async function loadHsnMaster(query = "") {
  const response = await fetch("/api/hsn/list", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ query, limit: 5000 })
  });
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || "HSN Master could not be loaded.");
  hsnMasterRows = result.rows || [];
  $("#hsnRows").innerHTML = hsnMasterRows.map((row, index) => `<tr class="hsn-record" data-index="${index}">
    <td>${escapeHtml(row.hsn_code)}</td><td>${escapeHtml(row.item_name)}</td><td>${escapeHtml(row.description)}</td>
    <td>${Number(row.gst_rate || 0)}%</td><td>${escapeHtml(row.uqc)}</td><td>${escapeHtml(row.category)}</td></tr>`).join("");
  document.querySelectorAll(".hsn-record").forEach(record => record.onclick = () => {
    const row = hsnMasterRows[Number(record.dataset.index)];
    $("#hsnCode").value = row.hsn_code || ""; $("#hsnItemName").value = row.item_name || "";
    $("#hsnDescription").value = row.description || ""; $("#hsnGstRate").value = row.gst_rate || 0;
    $("#hsnUqc").value = row.uqc || ""; $("#hsnCategory").value = row.category || "";
  });
}
$("#openHsnMaster").onclick = async () => {
  $("#hsnMasterDialog").showModal();
  try { await loadHsnMaster(); } catch (failure) { $("#hsnMessage").textContent = failure.message; }
};
$("#hsnSearchBtn").onclick = async () => {
  try { await loadHsnMaster($("#hsnSearch").value.trim()); } catch (failure) { $("#hsnMessage").textContent = failure.message; }
};
$("#hsnSearch").onkeydown = event => { if (event.key === "Enter") { event.preventDefault(); $("#hsnSearchBtn").click(); } };
$("#hsnSaveBtn").onclick = async () => {
  try {
    const response = await fetch("/api/hsn/save", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({
      hsn_code:$("#hsnCode").value, item_name:$("#hsnItemName").value, description:$("#hsnDescription").value,
      gst_rate:$("#hsnGstRate").value, uqc:$("#hsnUqc").value, category:$("#hsnCategory").value
    })});
    const result = await response.json(); if (!response.ok) throw new Error(result.error || "Save failed.");
    $("#hsnMessage").textContent = "HSN Item saved successfully."; await loadHsnMaster($("#hsnSearch").value.trim());
  } catch (failure) { $("#hsnMessage").textContent = failure.message; }
};
$("#hsnImportBtn").onclick = async () => {
  const file = $("#hsnFileInput").files[0]; if (!file) return alert("Select an HSN Master Excel or CSV file.");
  const button = $("#hsnImportBtn"); button.disabled = true; button.textContent = "Importing...";
  try {
    const response = await fetch("/api/hsn/import", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({
      file:{name:file.name,data:await fileToBase64(file)}
    })});
    const result = await response.json(); if (!response.ok) throw new Error(result.error || "Import failed.");
    $("#hsnMessage").textContent = `${result.saved} HSN Item imported; ${result.skipped} row skipped.`;
    await loadHsnMaster();
  } catch (failure) { $("#hsnMessage").textContent = failure.message; }
  finally { button.disabled = false; button.textContent = "Import Excel / CSV"; }
};
async function syncTallyMasters() {
  const button = $("#tallySyncBtn");
  const error = $("#tallyConnectorError");
  const headerBtn = $("#openTallyConnector");
  const headerStatus = $("#dashboardTallyStatus");
  if (button) {
    button.disabled = true;
    button.textContent = "Syncing...";
  }
  if (headerStatus) headerStatus.textContent = "Connecting...";
  if (error) error.classList.add("hidden");
  try {
    const response = await fetch("/api/tally/sync", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: "{}"
    });
    let result = {};
    try {
      result = await response.json();
    } catch (_) {
      throw new Error(`Tally sync returned invalid JSON (HTTP ${response.status}).`);
    }
    if (response.status === 401) {
      const message = result.error || "Login required. Enter your PIN, then Connect Tally again.";
      if (error) {
        error.textContent = message;
        error.classList.remove("hidden");
      }
      if (headerStatus) headerStatus.textContent = "Connect Tally";
      try { $("#tallyConnectorDialog")?.close(); } catch (_) {}
      await initializeAccount();
      return;
    }
    if (!response.ok) throw new Error(result.error || `Tally sync failed (HTTP ${response.status}).`);
    tallyMasters = result;
    if ($("#tallyCompanyName")) $("#tallyCompanyName").textContent = result.company || "Open Tally company";
    if ($("#tallySyncTime")) $("#tallySyncTime").textContent = `Synced: ${new Date(result.synced_at).toLocaleString()}`;
    if (headerStatus) headerStatus.textContent = "Tally Connected";
    if (headerBtn) headerBtn.classList.add("connected");
    if ($("#tallyStatusDot")) $("#tallyStatusDot").classList.add("connected");
    if ($("#tallyLedgerCount")) $("#tallyLedgerCount").textContent = Number(result.counts?.ledgers || 0).toLocaleString("en-IN");
    if ($("#tallyItemCount")) $("#tallyItemCount").textContent = Number(result.counts?.items || 0).toLocaleString("en-IN");
    if ($("#tallyVoucherCount")) $("#tallyVoucherCount").textContent = Number(result.counts?.voucher_types || 0).toLocaleString("en-IN");
    if ($("#tallyLedgerPreview")) {
      $("#tallyLedgerPreview").innerHTML = (result.ledgers || []).slice(0, 30).map(item => `<span>${escapeHtml(item.name)}<small>${escapeHtml(item.parent || "")}</small></span>`).join("");
    }
    if ($("#tallyLedgerList")) {
      $("#tallyLedgerList").innerHTML = (result.ledgers || []).map(item => `<option value="${escapeHtml(item.name)}"></option>`).join("");
    }
    if ($("#tallyItemList")) {
      $("#tallyItemList").innerHTML = (result.items || []).map(item => `<option value="${escapeHtml(item.name)}">${escapeHtml(item.hsn || "")}</option>`).join("");
    }
    if ($("#tallyVoucherTypeList")) {
      $("#tallyVoucherTypeList").innerHTML = (result.voucher_types || []).map(item => `<option value="${escapeHtml(item.name)}">${escapeHtml(item.parent || "")}</option>`).join("");
    }
    if ($("#tallyItemPreview")) {
      $("#tallyItemPreview").innerHTML = (result.items || []).slice(0, 30).map(item => `<span>${escapeHtml(item.name)}<small>${escapeHtml(item.hsn || item.parent || "")}</small></span>`).join("");
    }
    if ($("#tallySyncCounts")) $("#tallySyncCounts").classList.remove("hidden");
    if ($("#tallyMasterPreview")) $("#tallyMasterPreview").classList.remove("hidden");
    if ($("#tallyDisconnectBtn")) $("#tallyDisconnectBtn").classList.remove("hidden");
    if (rows.length) await autoMatchBankLedgers();
    if (gstRows.length) {
      renderSalesRows();
      renderSalesNoteRows();
    }
  } catch (failure) {
    const message = failure?.message || String(failure) || "Tally sync failed.";
    if (error) {
      error.textContent = message;
      error.classList.remove("hidden");
    } else {
      alert(message);
    }
    if (headerStatus) headerStatus.textContent = "Connect Tally";
    if (headerBtn) headerBtn.classList.remove("connected");
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = tallyMasters.company ? "Reconnect" : "Connect";
    }
  }
}
if ($("#tallySyncBtn")) $("#tallySyncBtn").onclick = syncTallyMasters;
if ($("#openTallyConnector")) {
  $("#openTallyConnector").onclick = async () => {
    const error = $("#tallyConnectorError");
    if (error) error.classList.add("hidden");
    const dialog = $("#tallyConnectorDialog");
    if (dialog && typeof dialog.showModal === "function" && !dialog.open) {
      dialog.showModal();
    }
    await syncTallyMasters();
  };
}
$("#tallyDisconnectBtn").onclick = () => {
  tallyMasters = {company:"", ledgers:[], items:[], voucher_types:[], counts:{}};
  $("#tallyCompanyName").textContent = "Not connected";
  $("#tallySyncTime").textContent = "Click Connect to read Tally masters.";
  $("#dashboardTallyStatus").textContent = "Connect Tally";
  $("#openTallyConnector").classList.remove("connected");
  $("#tallyStatusDot").classList.remove("connected");
  $("#tallySyncCounts").classList.add("hidden");
  $("#tallyMasterPreview").classList.add("hidden");
  $("#tallyLedgerList").innerHTML = "";
  $("#tallyItemList").innerHTML = "";
  $("#tallyVoucherTypeList").innerHTML = "";
  $("#tallySyncBtn").textContent = "Connect";
  $("#tallyDisconnectBtn").classList.add("hidden");
};
function updateMargImportOptions() {
  const returnType = $("#gstReturnType").value;
  const isMarg = returnType.toUpperCase().startsWith("MARG");
  const isGstr3b = returnType === "GSTR-3B";
  if (activeGstModule === "reconciliation") $("#gstFileLabel").textContent = isGstr3b ? "GSTR-3B PDF / File" : "GST Reconciliation File";
  $("#gstFileInput").accept = isGstr3b ? ".pdf,.json,.zip,.xlsx,.xlsm" : ".json,.xlsx,.xlsm,.xls,.csv,.zip,.mbk,.pdf";
  $("#margImportOptions").classList.toggle("hidden", !isMarg);
  $("#margMonthField").classList.toggle("hidden", $("#margPeriodType").value !== "monthly");
  if (!isMarg) {
    $("#margPasswordField").classList.add("hidden");
    $("#margBackupPassword").value = "";
  }
}

function resetGstModuleWorkspace() {
  gstRows = [];
  gstDatasets = {};
  gstSalesOriginalRows = [];
  $("#gstFileInput").value = "";
  $("#gstResults").classList.add("hidden");
  $("#gstError").classList.add("hidden");
  $("#gstSalesPanel").classList.add("hidden");
  $("#gstTallyPanel").classList.add("hidden");
  $("#gstTableFilters").classList.add("hidden");
  $("#gstRateSummaries").classList.add("hidden");
  $("#gstr1SummaryPanel").classList.add("hidden");
  $("#gstNotesSection").classList.add("hidden");
  $("#gstMatchCounts").classList.add("hidden");
  $("#gstPurchaseReconcilePanel").classList.toggle("hidden", activeGstModule !== "reconciliation");
  $("#gstThreeWayPanel").classList.toggle("hidden", activeGstModule !== "threeway");
  $("#gstPaymentReviewPanel").classList.toggle("hidden", activeGstModule !== "payment");
  $("#gstPurchaseResultBoxes").classList.add("hidden");
  $("#gstr2SummaryPanel").classList.add("hidden");
  $("#purchaseGstr3bComparePanel").classList.add("hidden");
  $("#purchaseGstr3bCompareRows").innerHTML = "";
  $("#purchaseGstr2aSummaryPanel").classList.add("hidden");
  $("#purchaseGstr2aSummaryRows").innerHTML = "";
  $("#purchaseGstr3bStatus").textContent = "GSTR-3B: Not loaded";
  $("#gstr2aLoadStatus").textContent = "GSTR-2A: Not loaded";
  $("#gstr2bLoadStatus").textContent = "GSTR-2B: Not loaded";
  $("#gstReconcileBtn").disabled = true;
  if (activeGstModule === "threeway") resetGstReconWorkspace();
  $("#gstRows").innerHTML = "";
  $("#gstNoteRows").innerHTML = "";
  if ($("#purchase2aEmpty")) {
    $("#purchase2aEmpty").classList.remove("hidden");
    $("#purchase2aContent").classList.add("hidden");
    $("#purchase2aSummaryRows").innerHTML = "";
    $("#purchase2aBoxes").innerHTML = "";
  }
  ["gstInvoices","gstInvoiceValue","gstTaxable","gstIgst","gstCgst","gstSgst"].forEach(id => {
    $(`#${id}`).textContent = id === "gstInvoices" ? "0" : "0.00";
  });
  updateMargImportOptions();
}

function updateGstWorkflowGuide() {
  const guide = $("#gstWorkflowGuide");
  if (activeGstModule === "reconciliation") {
    guide.innerHTML = "";
    guide.classList.add("hidden");
  } else if (activeGstModule === "threeway") {
    guide.innerHTML = `<strong>GST Reconciliation:</strong><span>1. Overview: Purchase + Sales + GSTR-3B / Net Payable.</span><span>2. Purchase: GSTR-2B vs Tally + ITC.</span><span>3. Sales: GSTR-1 vs Tally Output GST.</span><span>4. GSTR-3B tab: Liability, ITC claim, utilisation, net payable.</span><span>5. GST Payment &amp; Ledger: payments, cash ledger, ITC ledger vs Tally.</span><span>6. One Click GST Sync refreshes portal return sides.</span>`;
    guide.classList.remove("hidden");
  } else guide.classList.add("hidden");
}

function openGstWorkspace(mode) {
  activeGstModule = mode;
  $(".gst-import-row").classList.remove("hidden");
  if (mode === "sales") {
    $("#gstModuleTitle").textContent = "GST Sales / GSTR-1";
    $("#gstModuleSubtitle").textContent = "Import GSTR-1 JSON, Excel or MARG sales data and prepare reviewed Tally sales entries.";
    $("#gstFileLabel").textContent = "GST Sales File";
    $("#gstReturnType").innerHTML = `<option>GSTR-1 / Sales Register</option><option>GSTR-1 JSON (GST Portal)</option><option>MARG Backup / Sales Register</option>`;
  } else if (mode === "payment") {
    $("#gstModuleTitle").textContent = "GST Payment & ITC";
    $("#gstModuleSubtitle").textContent = "Import GST payment and ledger reports for review. Sales/GSTR-1 data is kept separate.";
    $("#gstFileLabel").textContent = "GST Payment File";
    $("#gstReturnType").innerHTML = `<option>GST Payment List</option><option>GST Cash Ledger</option><option>GST Credit Ledger</option><option>GSTR-3B</option><option>Liability & ITC Comparison</option>`;
  } else if (mode === "threeway") {
    $("#gstModuleTitle").textContent = "GST Reconciliation";
    $("#gstModuleSubtitle").textContent = "Purchase (GSTR-2B) and Sales (GSTR-1) reconciliation with Tally — Phase 1 + Phase 2.";
    $("#gstFileLabel").textContent = "GST Reconciliation File";
    $("#gstReturnType").innerHTML = `<option>GSTR-2B</option><option>GSTR-3B</option><option>GSTR-1 JSON (GST Portal)</option><option>GSTR-1 / Sales Register</option>`;
    setReconTab("overview", { refresh: false });
  } else {
    $("#gstModuleTitle").textContent = "Purchase / GSTR-2";
    $("#gstModuleSubtitle").textContent = "Import GSTR-2A + GSTR-2B, match invoices, then post Purchase vouchers with correct voucher vs supplier invoice dates.";
    $("#gstFileLabel").textContent = "GSTR-2A / GSTR-2B File";
    $("#gstReturnType").innerHTML = `<option>GSTR-2A</option><option>GSTR-2B</option>`;
    $("#gstFileInput").multiple = true;
  }
  $("#gstReturnType").selectedIndex = 0;
  resetGstModuleWorkspace();
  updateGstWorkflowGuide();
  $("#gstDialog").showModal();
}
$("#gstReturnType").onchange = updateMargImportOptions;
$("#margPeriodType").onchange = updateMargImportOptions;
$("#gstFileInput").addEventListener("change", () => {
  $("#margPasswordField").classList.add("hidden");
  $("#margBackupPassword").value = "";
});
$("#openSalesModule").onclick = async () => {
  openGstWorkspace("sales");
  try { await loadHsnMaster(); } catch (_) {}
};
// Mobile invoice photo entry: OCR creates a draft, while the user remains in control
// of the reviewed values before a single Sales Invoice is posted to Tally.
const photoInvoiceIds = ["photoInvoiceNo","photoInvoiceDate","photoInvoiceParty","photoInvoiceGstin","photoInvoiceItem","photoInvoiceHsn","photoInvoiceQty","photoInvoiceTaxable","photoInvoiceRate","photoInvoiceTaxType","photoInvoiceTotal"];
function photoInvoiceSetMessage(text, error=false) { const el=$("#photoInvoiceMessage"); if(!el)return; el.textContent=text; el.style.color=error?"#b42318":""; }
function photoInvoiceReset() { photoInvoiceIds.forEach(id=>{const el=$("#"+id); if(el) el.value="";}); $("#photoInvoiceQty").value="1"; $("#photoInvoiceItem").value="Items"; $("#photoInvoiceRate").value="12"; $("#photoInvoicePreview").classList.add("hidden"); $("#photoInvoiceImage").removeAttribute("src"); $("#photoInvoiceOcrText").textContent="Choose a photo to read invoice details."; photoInvoiceSetMessage("Review the draft, then send it to the connected Tally company."); }
function photoInvoiceCalculateTotal() { const taxable=Number($("#photoInvoiceTaxable")?.value||0), rate=Number($("#photoInvoiceRate")?.value||0); if(taxable) $("#photoInvoiceTotal").value=(taxable*(1+rate/100)).toFixed(2); }
function photoInvoiceDraftFromText(text) {
  const clean=String(text||"").replace(/\r/g,"");
  const inv=(clean.match(/(?:invoice|inv|bill)\s*(?:no|number|#)\s*[:\-]?\s*([A-Z0-9][A-Z0-9\/-]{2,})/i)||[])[1] || (clean.match(/\b(?:INV|BILL)[\s-]*([A-Z0-9][A-Z0-9\/-]{2,})\b/i)||[])[1];
  const invoiceCandidate=inv && !/^(dated|date|number|no|e)$/i.test(inv) ? inv : (clean.match(/\b[A-Z]{2,}\/\d{2,}\/\d{2}-\d{2}\b/i)||[])[0];
  const gstin=(clean.match(/\b\d{2}[A-Z]{5}\d{4}[A-Z]\d[Z]\w\b/i)||[])[0];
  const hsn=(clean.match(/(?:hsn|sac)[^0-9]{0,15}(\d{4,8})/i)||[])[1];
  const date=(clean.match(/\b(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})\b/)||[])[1] || (clean.match(/\b(\d{1,2})\s*[- ]\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s*[- ]\s*(\d{2,4})\b/i)||[]).slice(1).join("-");
  const partyMatch=clean.match(/(?:buyer\s*\(\s*bill\s*to\s*\)|bill\s*to|billed\s*to|customer|party)\s*[:\-]?\s*(?:\n\s*)?([^\n]{3,60})/i);
  const party=partyMatch?.[1]?.replace(/\s{2,}/g," ").trim();
  const taxableMatch=clean.match(/taxable\s*(?:value)?\D{0,25}(\d[\d,]*(?:\.\d{1,2})?)/i);
  const totalMatches=[...clean.matchAll(/(?:grand\s*total|invoice\s*total|total\s*amount|amount\s*payable|net\s*amount|\btotal\b)\D{0,30}(?:₹|rs\.?|inr)?\s*(\d[\d,]*(?:\.\d{1,2})?)/gi)];
  const totalMatch=totalMatches[totalMatches.length-1];
  const amounts=[taxableMatch?.[1],totalMatch?.[1]].map(v=>Number(String(v||"").replace(/,/g,""))).filter(v=>v>0);
  const detectedRate=(clean.match(/(?:cgst|sgst|igst)[^%]{0,50}(\d{1,2})\s*%/i)||[])[1];
  const anyRate=(clean.match(/\b(5|12|18|28)\s*%/i)||[])[1];
  if(invoiceCandidate) $("#photoInvoiceNo").value=invoiceCandidate;
  if(gstin) $("#photoInvoiceGstin").value=gstin.toUpperCase();
  if(hsn) $("#photoInvoiceHsn").value=hsn;
  const partyFallback=(clean.match(/\b([A-Z][A-Z ]{3,}\s+(?:GALLERY|TRADERS|ENTERPRISES|ENTERPRISE|STORES|SERVICES))\b/i)||[])[1];
  if((party && !/^(no|number|date|gstin|invoice|dated)$/i.test(party)) || partyFallback) $("#photoInvoiceParty").value=(partyFallback||party).trim();
  const itemLine=(clean.split("\n").map(line=>line.replace(/\s+/g," ").trim()).find(line=>/(rental|service|goods|product|particulars)/i.test(line)&&!/(invoice|customer|buyer|taxable|total|gstin|dated)/i.test(line)&&line.length>=6)||"");
  if(itemLine) $("#photoInvoiceItem").value=itemLine.replace(/^(particulars|item)\s*[:\-]?\s*/i,"").trim();
  if(date){const parts=date.split(/[\/-]/); if(parts.length===3 && /^\d+$/.test(parts[1])){const p=parts.map(Number); const d=p[2]<100?2000+p[2]:p[2]; $("#photoInvoiceDate").value=`${d}-${String(p[1]).padStart(2,"0")}-${String(p[0]).padStart(2,"0")}`;} else {const m={jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12}[String(parts[1]).slice(0,3).toLowerCase()]; const d=Number(parts[2])<100?2000+Number(parts[2]):Number(parts[2]); if(m&&d) $("#photoInvoiceDate").value=`${d}-${String(m).padStart(2,"0")}-${String(Number(parts[0])).padStart(2,"0")}`;}}
  if(detectedRate || anyRate){ const rate=detectedRate ? taxTypeRate(Number(detectedRate)) : Number(anyRate); $("#photoInvoiceRate").value=String(rate); }
  if(taxableMatch) $("#photoInvoiceTaxable").value=Number(taxableMatch[1].replace(/,/g,"")).toFixed(2);
  if(totalMatch) $("#photoInvoiceTotal").value=Number(totalMatch[1].replace(/,/g,"")).toFixed(2);
  else if(amounts.length) $("#photoInvoiceTotal").value=Math.max(...amounts).toFixed(2);
  if(!taxableMatch){const numeric=[...clean.matchAll(/\b\d[\d,]*(?:\.\d{1,2})?\b/g)].map(m=>Number(m[0].replace(/,/g,""))).filter(v=>v>100&&v<500000); if(numeric.length>1){const sorted=[...new Set(numeric)].sort((a,b)=>a-b); $("#photoInvoiceTaxable").value=sorted[sorted.length-2].toFixed(2); if(!$("#photoInvoiceTotal").value) $("#photoInvoiceTotal").value=sorted[sorted.length-1].toFixed(2);}}
  const taxableValue=Number($("#photoInvoiceTaxable").value||0), totalValue=Number($("#photoInvoiceTotal").value||0), selectedRate=Number($("#photoInvoiceRate").value||0);
  if(taxableValue>0 && (totalValue<taxableValue || totalValue>taxableValue*1.5)) $("#photoInvoiceTotal").value=(taxableValue*(1+selectedRate/100)).toFixed(2);
}
function taxTypeRate(componentRate){ const r=Number(componentRate||0); return r>0&&r<=14 ? r*2 : r; }
$("#openPhotoInvoiceModule").onclick = () => { $("#photoInvoiceDialog").showModal(); };
$("#photoInvoiceClearBtn").onclick = photoInvoiceReset;
["photoInvoiceTaxable","photoInvoiceRate"].forEach(id=>$("#"+id)?.addEventListener("input",photoInvoiceCalculateTotal));
async function handlePhotoInvoiceFile(event) {
  const file=event.target.files?.[0]; if(!file)return;
  const preview=$("#photoInvoicePreview"), image=$("#photoInvoiceImage"), status=$("#photoInvoiceOcrStatus"), out=$("#photoInvoiceOcrText");
  preview.classList.remove("hidden"); image.src=URL.createObjectURL(file); status.textContent="Reading photo…"; out.textContent="OCR is preparing the invoice draft.";
  try {
    if(!window.Tesseract) throw new Error("OCR library is still loading. Please try again in a moment.");
    const result=await window.Tesseract.recognize(file,"eng",{logger:msg=>{if(msg.status)status.textContent=`${msg.status} ${Math.round((msg.progress||0)*100)}%`;}});
    const text=String(result?.data?.text||"").trim(); photoInvoiceDraftFromText(text); status.textContent="Draft ready — review it"; out.textContent=text||"No text found. Enter the invoice values manually."; photoInvoiceSetMessage("OCR draft created. Check the invoice number, date, party, item and totals before sending.");
  } catch(err) { status.textContent="Manual entry"; out.textContent=err.message||"OCR unavailable"; photoInvoiceSetMessage("Photo loaded. Enter the invoice values manually, then send."); }
}
$("#photoInvoiceFile")?.addEventListener("change", handlePhotoInvoiceFile);
$("#photoInvoiceGallery")?.addEventListener("change", handlePhotoInvoiceFile);
$("#photoInvoiceSendBtn")?.addEventListener("click", async event => {
  event.preventDefault();
  const invoiceNo=$("#photoInvoiceNo").value.trim(), party=$("#photoInvoiceParty").value.trim(), invoiceDate=$("#photoInvoiceDate").value, taxable=Number($("#photoInvoiceTaxable").value||0), total=Number($("#photoInvoiceTotal").value||0), rate=Number($("#photoInvoiceRate").value||0), qty=Number($("#photoInvoiceQty").value||1), taxType=$("#photoInvoiceTaxType").value;
  if(!invoiceNo||!invoiceDate||!party||taxable<=0||total<=0) return photoInvoiceSetMessage("Invoice no., date, party, taxable value and total are required.",true);
  const d=invoiceDate.split("-"); const displayDate=`${d[2]}-${d[1]}-${d[0]}`; const igst=taxType==="igst"?taxable*rate/100:0, cgst=taxType==="local"?taxable*rate/200:0, sgst=cgst;
  const row={selected:true,ready_for_sales_tally:true,document_type:"Sales Invoice",invoice_no:invoiceNo,invoice_date:displayDate,party_name:party,party_ledger:party,gstin:$("#photoInvoiceGstin").value.trim(),invoice_value:total,taxable_value:taxable, sales_allocations:[{item_name:$("#photoInvoiceItem").value.trim()||"Items",hsn:$("#photoInvoiceHsn").value.trim(),quantity:qty,unit:"Pcs",rate,taxable_value:taxable,igst,cgst,sgst,cess:0}]};
  const button=event.currentTarget; button.disabled=true; button.textContent="Sending…"; photoInvoiceSetMessage("Sending the reviewed invoice to Tally…");
  try { const response=await fetch("/api/gst/sales/tally/send-one",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({row,ledgers:{}})}); const data=await response.json(); if(!response.ok) throw new Error(data.error||data.message||"Tally rejected the invoice."); photoInvoiceSetMessage(data.status==="CREATED"?"Invoice created in Tally successfully.":(data.message||"Tally response received.")); }
  catch(err){ photoInvoiceSetMessage(err.message||"Could not send to Tally.",true); }
  finally { button.disabled=false; button.textContent="Send Invoice to Tally"; }
});
$("#openGstPaymentModule").onclick = () => {
  openGstWorkspace("payment");
};
$("#openExcelVoucherModule").onclick = () => {
  $("#moduleDashboard").classList.add("hidden");
  $("#bankModule").classList.remove("hidden");
  showMessage("Choose an Excel file. Map Date, Particulars, Debit/Credit and review the suggested voucher and ledger before sending to Tally.");
};
$("#openCompanyLedgerModule").onclick = () => {
  $("#companyError").classList.add("hidden");
  $("#companyLedgerDialog").showModal();
};
function renderCompanyRows(sourceRows) {
  $("#companyRows").innerHTML = sourceRows.map((row, index) => {
    const selectable = row.match_status === "Only in Statement" && row.status === "Ready";
    return `<tr><td>${selectable ? `<input class="company-select" data-index="${index}" type="checkbox" ${row.selected ? "checked" : ""}>` : "—"}</td>
      <td>${escapeHtml(row.date || "")}</td><td>${escapeHtml(row.particulars || "")}</td><td>${escapeHtml(row.reference || "")}</td>
      <td class="money">${Number(row.debit || 0).toLocaleString("en-IN",{minimumFractionDigits:2})}</td>
      <td class="money">${Number(row.credit || 0).toLocaleString("en-IN",{minimumFractionDigits:2})}</td>
      <td>${escapeHtml(row.voucher_type || "")}</td><td><span class="gst-status ${row.match_status === "Matched" ? "matched" : "review"}">${escapeHtml(row.match_status || row.status || "Loaded")}</span></td></tr>`;
  }).join("");
  document.querySelectorAll(".company-select").forEach(input => input.onchange = () => {
    companyReconcileRows[Number(input.dataset.index)].selected = input.checked;
  });
}
$("#companyParseBtn").onclick = async () => {
  const files = [...$("#companyFileInput").files];
  const error = $("#companyError");
  error.classList.add("hidden");
  if (!files.length) return alert("Select a company statement file.");
  const button = $("#companyParseBtn");
  button.disabled = true; button.textContent = "Reading...";
  try {
    const packed = [];
    for (const file of files) packed.push({name:file.name,data:await fileToBase64(file)});
    const response = await fetch("/api/company/parse", {
      method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({partyLedger:$("#companyPartyLedger").value,files:packed})
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "Statement reading failed.");
    companyStatementRows = result.rows || [];
    companyReconcileRows = companyStatementRows.map(row => ({...row, match_status:"Loaded", selected:false}));
    renderCompanyRows(companyReconcileRows);
    $("#companyReconcileBtn").disabled = false;
    $("#companyCounts").innerHTML = `<span>Statement Rows<strong>${companyStatementRows.length.toLocaleString("en-IN")}</strong></span>`;
    $("#companyCounts").classList.remove("hidden");
  } catch (failure) {
    error.textContent = failure.message; error.classList.remove("hidden");
  } finally { button.disabled=false;button.textContent="Read Statement"; }
};
$("#companyReconcileBtn").onclick = async () => {
  const error = $("#companyError"); error.classList.add("hidden");
  const partyLedger = $("#companyPartyLedger").value.trim();
  if (!partyLedger) return alert("Select the matching Tally party ledger.");
  const button=$("#companyReconcileBtn");button.disabled=true;button.textContent="Comparing...";
  try {
    const response=await fetch("/api/company/reconcile",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({statementRows:companyStatementRows,partyLedger,tolerance:$("#companyTolerance").value})});
    const result=await response.json();if(!response.ok)throw new Error(result.error||"Tally comparison failed.");
    companyReconcileRows=result.rows||[];renderCompanyRows(companyReconcileRows);
    $("#companyCounts").innerHTML=Object.entries(result.counts||{}).map(([name,count])=>`<span>${escapeHtml(name)}<strong>${Number(count).toLocaleString("en-IN")}</strong></span>`).join("");
    $("#companySendBtn").classList.remove("hidden");
  } catch(failure){error.textContent=failure.message;error.classList.remove("hidden");}
  finally{button.disabled=false;button.textContent="Compare with Tally";}
};
$("#companySendBtn").onclick = async () => {
  const selected=companyReconcileRows.filter(row=>row.selected&&row.match_status==="Only in Statement");
  if(!selected.length)return alert("Select at least one reviewed missing entry.");
  if(!confirm(`Take a Tally backup first. Create ${selected.length} selected voucher(s) in the open Tally company?`))return;
  const button=$("#companySendBtn");button.disabled=true;button.textContent="Sending...";
  try{
    const response=await fetch("/api/company/tally/send",{method:"POST",headers:{"Content-Type":"application/json"},
      body:JSON.stringify({rows:companyReconcileRows,partyLedger:$("#companyPartyLedger").value.trim(),counterLedger:$("#companyCounterLedger").value.trim()})});
    const result=await response.json();if(!response.ok)throw new Error(result.error||"Tally import failed.");
    alert(`${result.created||0} company-ledger voucher(s) created in Tally.`);
  }catch(failure){$("#companyError").textContent=failure.message;$("#companyError").classList.remove("hidden");}
  finally{button.disabled=false;button.textContent="Send Selected to Tally";}
};
$("#openGstModule").onclick = async () => {
  openGstWorkspace("reconciliation");
  try { await loadHsnMaster(); } catch (_) {}
  try {
    await loadGstReconSession({ restorePortal: true });
    const has2a = Boolean((gstDatasets["GSTR-2A"] || []).length);
    const has2b = Boolean((gstDatasets["GSTR-2B"] || []).length);
    if (has2a && has2b) {
      const reconciled = await reconcileGst();
      if (reconciled && gstRows.length && $("#purchase2aWorkspace")) {
        $("#purchase2aWorkspace").classList.add("hidden");
        $("#purchase2aWorkspace").parentElement?.classList.add("hidden");
      }
    } else if (has2a) {
      renderPurchase2aWorkspace();
    }
    updatePurchaseImportStatus();
  } catch (_) {
    updatePurchaseImportStatus();
  }
};
$("#openGstReconModule").onclick = async () => {
  openGstWorkspace("threeway");
  ensureGstReconPanelVisible();
  if (!gstReconPeriodUserChosen) setGstReconPeriod("ALL", { refresh: false, silent: true });
  try {
    await loadGstReconSession();
  } catch (_) {
    // Still show the recon shell even if a background refresh fails.
  } finally {
    ensureGstReconPanelVisible();
    if (activeReconTab === "overview") {
      try { await refreshGstReconOverview(false); } catch (_) {}
      ensureGstReconPanelVisible();
    }
  }
};
$("#bankBackBtn").onclick = () => {
  $("#bankModule").classList.add("hidden");
  $("#moduleDashboard").classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
};
function salesRateLabel(row) {
  if (row.sales_allocations?.length) return [...new Set(row.sales_allocations.map(item => `${item.rate}%`))].join(", ");
  if (row.items?.length) return [...new Set(row.items.map(item => `${itemGstRate(item)}%`))].join(", ");
  const taxable = Number(row.taxable_value || 0);
  if (!taxable) return "0%";
  const calculated = 100 * (Number(row.igst || 0) + Number(row.cgst || 0) + Number(row.sgst || 0) + Number(row.cess || 0)) / taxable;
  return `${[0,5,12,18,28].reduce((best, rate) => Math.abs(rate-calculated) < Math.abs(best-calculated) ? rate : best, 0)}%`;
}
function salesHsnLabel(row) {
  const parts = row.sales_allocations?.length ? row.sales_allocations : (row.items || []);
  return [...new Set(parts.map(item => String(item.hsn || "").trim()).filter(Boolean))].join(", ");
}
function salesItemLabel(row) {
  if (row.expense_ledger) return row.expense_ledger;
  const parts = row.sales_allocations?.length ? row.sales_allocations : (row.items || []);
  const names = parts.map(item => {
    const hsn = String(item.hsn || "").trim();
    const rate = item.rate !== undefined && row.sales_allocations?.length ? Number(item.rate || 0) : itemGstRate(item);
    const tallyHsnItems = hsn ? (tallyMasters.items || []).filter(master => normalizedHsn(master.hsn) === normalizedHsn(hsn)) : [];
    const offlineHsnItems = hsn ? hsnMasterRows.filter(master => normalizedHsn(master.hsn_code) === normalizedHsn(hsn) && master.active !== 0) : [];
    const automaticTallyItem = tallyHsnItems.length === 1 ? tallyHsnItems[0].name : "";
    const automaticOfflineItem = !tallyHsnItems.length && offlineHsnItems.length === 1 ? offlineHsnItems[0].item_name : "";
    const multipleMatches = tallyHsnItems.length > 1 || (!tallyHsnItems.length && offlineHsnItems.length > 1);
    return String(item.item_name || item.name || automaticTallyItem || automaticOfflineItem || (hsn && multipleMatches ? `Select HSN ${hsn} Item` : (hsn ? `HSN ${hsn} Items` : `${rate}% Items`))).trim();
  }).filter(Boolean);
  return [...new Set(names)].join(", ");
}
function setRowItemName(index, itemName, explicitRow, options = {}) {
  const targets = explicitRow
    ? [explicitRow]
    : [gstRows[index], gstSalesOriginalRows[index]].filter(Boolean);
  targets.forEach(row => {
    if(/expenses?/i.test(itemName)) {
      const taxable=Number(row.taxable_value||0), calculated=taxable?100*(Number(row.igst||0)+Number(row.cgst||0)+Number(row.sgst||0)+Number(row.cess||0))/taxable:0;
      row.expense_ledger=itemName;
      row.gst_rate=[0,5,12,18,28].reduce((best,value)=>Math.abs(value-calculated)<Math.abs(best-calculated)?value:best,0);
      row.sales_allocations=[];
      return;
    }
    row.expense_ledger="";
    if(!row.sales_allocations?.length) {
      const sourceItems = Array.isArray(row.items) && row.items.length ? row.items : [row];
      row.sales_allocations = sourceItems.map(part => {
        const taxable=Number(part.taxable_value??row.taxable_value??0), calculated=taxable?100*(Number(part.igst||0)+Number(part.cgst||0)+Number(part.sgst||0)+Number(part.cess||0))/taxable:0;
        const rawRate=Number(part.rate??part.gst_rate??calculated), rate=[0,5,12,18,28].reduce((best,value)=>Math.abs(value-rawRate)<Math.abs(best-rawRate)?value:best,0);
        return {rate,hsn:String(part.hsn_code||part.hsn||row.hsn_code||""),item_name:itemName,quantity:Number(part.quantity??row.quantity??1)||1,unit:String(part.unit||part.uqc||row.unit||row.uqc||"Pcs"),taxable_value:taxable,igst:Number(part.igst||0),cgst:Number(part.cgst||0),sgst:Number(part.sgst||0),cess:Number(part.cess||0)};
      }).filter(part=>part.taxable_value>0);
    }
    const parts = row.sales_allocations?.length ? row.sales_allocations : (row.items || []);
    parts.forEach(item => { item.item_name = itemName; });
  });
  if (activeGstModule === "reconciliation" && targets[0] && options.syncSources !== false) {
    syncPurchaseItemMapToSources(targets[0]);
  }
}

/** Same Apply Item formula as GSTR-1: selected gstRows → setRowItemName → refresh. */
async function applyItemToSelectedRows({
  rows = null,
  itemInputId = "",
  itemInputEl = null,
  predicate = () => true,
  refresh = null,
  emptySelectMessage = "Select at least one visible row.",
  successNoun = "selected row(s)",
} = {}) {
  const input = itemInputEl || (itemInputId ? $(`#${itemInputId}`) : null);
  const itemName = String(input?.value || "").trim();
  const sourceRows = rows || gstRows;
  const selected = sourceRows.map((row, index) => ({ row, index })).filter(({ row }) => row.selected && predicate(row));
  if (!selected.length) return alert(emptySelectMessage);
  if (!itemName) return alert("Select or enter the Item Name.");
  const bulkPurchase = activeGstModule === "reconciliation" && selected.length > 100;
  for (let position = 0; position < selected.length; position += 1) {
    const { index, row } = selected[position];
    setRowItemName(index, itemName, rows ? row : undefined, { syncSources: !bulkPurchase });
    // Let the browser paint/respond while thousands of invoices are updated.
    if (bulkPurchase && position > 0 && position % 250 === 0) {
      await new Promise(resolve => requestAnimationFrame(resolve));
    }
  }
  if (bulkPurchase) syncPurchaseItemMapsToSources(selected.map(({ row }) => row));
  if (typeof refresh === "function") refresh();
  if (activeGstModule === "reconciliation") {
    // Start persistence after the UI has painted; JSON serialization of a
    // full 2A/2B year must not freeze the Apply Item click.
    setTimeout(() => persistPurchaseItemMappings(), 0);
  }
  alert(`${selected.length} ${successNoun} changed to ${itemName}.`);
  return selected.length;
}

function purchaseItemMapKey(row) {
  if (!row) return "";
  return [
    String(row.gstin || "").replace(/\s+/g, "").toUpperCase(),
    String(row.invoice_no || "").toUpperCase().replace(/[^A-Z0-9]/g, ""),
    String(row.invoice_date || row.original_invoice_date || "").replace(/\//g, "-"),
  ].join("|");
}

function syncPurchaseItemMapToSources(row) {
  if (!row) return;
  const mapping = {
    sales_allocations: row.sales_allocations ? JSON.parse(JSON.stringify(row.sales_allocations)) : [],
    expense_ledger: row.expense_ledger || "",
    item_name: salesItemLabel(row) || "",
  };
  if (row.gstr2a) Object.assign(row.gstr2a, mapping);
  if (row.gstr2b) Object.assign(row.gstr2b, mapping);
  const key = purchaseItemMapKey(row);
  if (!key || key.startsWith("||")) return;
  ["GSTR-2A", "GSTR-2B"].forEach((datasetKey) => {
    (gstDatasets[datasetKey] || []).forEach((source) => {
      if (purchaseItemMapKey(source) === key) Object.assign(source, mapping);
    });
  });
}

function restorePurchaseItemMappings(rows) {
  const maps = new Map();
  ["GSTR-2A", "GSTR-2B"].forEach((datasetKey) => {
    (gstDatasets[datasetKey] || []).forEach((source) => {
      if (!(source.sales_allocations?.length || source.expense_ledger)) return;
      maps.set(purchaseItemMapKey(source), {
        sales_allocations: source.sales_allocations,
        expense_ledger: source.expense_ledger || "",
      });
    });
  });
  (rows || []).forEach((row) => {
    const hit = maps.get(purchaseItemMapKey(row))
      || maps.get(purchaseItemMapKey(row.gstr2a))
      || maps.get(purchaseItemMapKey(row.gstr2b));
    if (!hit) return;
    row.sales_allocations = hit.sales_allocations ? JSON.parse(JSON.stringify(hit.sales_allocations)) : [];
    row.expense_ledger = hit.expense_ledger || "";
    if (row.gstr2a) {
      row.gstr2a.sales_allocations = row.sales_allocations;
      row.gstr2a.expense_ledger = row.expense_ledger;
    }
    if (row.gstr2b) {
      row.gstr2b.sales_allocations = row.sales_allocations;
      row.gstr2b.expense_ledger = row.expense_ledger;
    }
  });
}

async function persistPurchaseItemMappings() {
  try {
    if ((gstDatasets["GSTR-2A"] || []).length) await saveGstReconSession({ gstr2a: gstDatasets["GSTR-2A"] });
    if ((gstDatasets["GSTR-2B"] || []).length) await saveGstReconSession({ gstr2b: gstDatasets["GSTR-2B"] });
  } catch (_) {}
}

function salesQuantity(row) {
  const parts = row.sales_allocations?.length ? row.sales_allocations : (row.items || []);
  return parts.reduce((total, item) => total + Number(item.quantity || 0), 0);
}
function normalizedHsn(value) {
  return String(value || "").replace(/^HSN\s*/i, "").replace(/\s+/g, "").trim().toLowerCase();
}
function itemSuggestionsForRow(row) {
  const hsns = new Set(salesHsnLabel(row).split(",").map(normalizedHsn).filter(Boolean));
  const tallyMatched = (tallyMasters.items || []).filter(item => hsns.has(normalizedHsn(item.hsn)));
  const offlineMatched = hsnMasterRows.filter(item => item.active !== 0 && hsns.has(normalizedHsn(item.hsn_code)))
    .map(item => ({name:item.item_name, hsn:item.hsn_code, parent:item.description || item.category || "Offline HSN Master"}));
  const allItems = [
    ...tallyMatched,
    ...offlineMatched,
    ...(tallyMasters.items || [])
  ];
  // Matching HSN items stay at the top, but every synced Tally item remains
  // available so the user can always choose a different stock item manually.
  return allItems.filter((item,index,all) =>
    item.name && all.findIndex(other => String(other.name).toLowerCase() === String(item.name).toLowerCase()) === index
  );
}
function showItemSuggestions(row) {
  $("#tallyItemList").innerHTML = itemSuggestionsForRow(row).map(item =>
    `<option value="${escapeHtml(item.name)}">${escapeHtml(item.hsn ? `HSN ${item.hsn}` : (item.parent || ""))}</option>`
  ).join("");
}
function itemGstRate(item) {
  const taxable = Number(item.taxable_value || 0);
  if (!taxable) return 0;
  const calculated = 100 * (Number(item.igst || 0) + Number(item.cgst || 0) + Number(item.sgst || 0) + Number(item.cess || 0)) / taxable;
  return [0,5,12,18,28].reduce((best, rate) => Math.abs(rate-calculated) < Math.abs(best-calculated) ? rate : best, 0);
}
function prepareSalesRowsForTally(sourceRows) {
  sourceRows.forEach(row => {
    const note = isSalesNote(row);
    // Capture Excel document totals before allocation rebuild. Credit Note
    // item/rate lines can disagree with Note Value / Taxable columns; the
    // document row must win for GSTR-1 summary and Output GST.
    const docTotals = {
      taxable: Math.abs(Number(row.taxable_value || 0)),
      igst: Math.abs(Number(row.igst || 0)),
      cgst: Math.abs(Number(row.cgst || 0)),
      sgst: Math.abs(Number(row.sgst || 0)),
      cess: Math.abs(Number(row.cess || 0)),
      invoice: Math.abs(Number(row.invoice_value || 0)),
    };
    const parts = row.items?.length ? row.items : [row];
    row.sales_allocations = parts.filter(part => {
      const tv = Number(part.taxable_value || 0);
      return note ? Math.abs(tv) > 0.005 : tv > 0;
    }).map(part => {
      const tv = Number(part.taxable_value || 0);
      const absTv = note ? Math.abs(tv) : tv;
      const rate = itemGstRate(part);
      const hsn = String(part.hsn || "").trim();
      return {
        rate, hsn, quantity: Math.abs(Number(part.quantity || 0)) || 0, item_rate: Math.abs(Number(part.rate || 0)), unit: String(part.unit || part.uqc || "Pcs"),
        item_name: String(part.item_name || part.name || ""),
        taxable_value: absTv, igst: Math.abs(Number(part.igst || 0)),
        cgst: Math.abs(Number(part.cgst || 0)), sgst: Math.abs(Number(part.sgst || 0)), cess: Math.abs(Number(part.cess || 0))
      };
    });
    if (note) {
      const allocTaxable = row.sales_allocations.reduce((s, a) => s + Number(a.taxable_value || 0), 0);
      const allocCgst = row.sales_allocations.reduce((s, a) => s + Number(a.cgst || 0), 0);
      const allocSgst = row.sales_allocations.reduce((s, a) => s + Number(a.sgst || 0), 0);
      const mismatch = docTotals.taxable > 0.005 && (
        Math.abs(allocTaxable - docTotals.taxable) > 0.05
        || Math.abs(allocCgst - docTotals.cgst) > 0.05
        || Math.abs(allocSgst - docTotals.sgst) > 0.05
      );
      if (mismatch || !row.sales_allocations.length) {
        const rate = itemGstRate({
          taxable_value: docTotals.taxable, igst: docTotals.igst,
          cgst: docTotals.cgst, sgst: docTotals.sgst, cess: docTotals.cess,
        });
        const seed = row.sales_allocations[0] || {};
        row.sales_allocations = docTotals.taxable > 0.005 ? [{
          rate,
          hsn: String(seed.hsn || row.hsn_code || "").trim(),
          quantity: Number(seed.quantity || 1) || 1,
          item_rate: 0,
          unit: String(seed.unit || "Pcs"),
          item_name: String(seed.item_name || ""),
          taxable_value: docTotals.taxable,
          igst: docTotals.igst,
          cgst: docTotals.cgst,
          sgst: docTotals.sgst,
          cess: docTotals.cess,
        }] : [];
      }
      if (docTotals.taxable > 0.005) {
        row.taxable_value = docTotals.taxable;
        row.igst = docTotals.igst;
        row.cgst = docTotals.cgst;
        row.sgst = docTotals.sgst;
        row.cess = docTotals.cess;
        row.invoice_value = docTotals.invoice > 0.005
          ? docTotals.invoice
          : (docTotals.taxable + docTotals.igst + docTotals.cgst + docTotals.sgst + docTotals.cess);
      } else if (row.sales_allocations.length) {
        row.invoice_value = row.sales_allocations.reduce((s, a) =>
          s + a.taxable_value + a.igst + a.cgst + a.sgst + a.cess, 0);
        row.taxable_value = row.sales_allocations.reduce((s, a) => s + a.taxable_value, 0);
        row.igst = row.sales_allocations.reduce((s, a) => s + a.igst, 0);
        row.cgst = row.sales_allocations.reduce((s, a) => s + a.cgst, 0);
        row.sgst = row.sales_allocations.reduce((s, a) => s + a.sgst, 0);
        row.cess = row.sales_allocations.reduce((s, a) => s + a.cess, 0);
      }
    }
    row.is_sales_note = note;
    row.ready_for_sales_tally = Boolean(!note && row.sales_allocations.length && row.invoice_no);
    row.ready_for_note_tally = Boolean(note && row.sales_allocations.length && row.invoice_no);
  });
}
function gstRateBreakdown(sourceRows) {
  const totals = Object.fromEntries([0,5,12,18,28].map(rate => [rate,{taxable:0,igst:0,cgst:0,sgst:0}]));
  sourceRows.forEach(row => {
    const hasAdjustedAllocations = Boolean(row.sales_allocations?.length);
    const parts = hasAdjustedAllocations ? row.sales_allocations : (row.items?.length ? row.items : [row]);
    parts.forEach(part => {
      const rate = hasAdjustedAllocations ? Number(part.rate || 0) : itemGstRate(part);
      const target = totals[rate] || (totals[rate] = {taxable:0,igst:0,cgst:0,sgst:0});
      target.taxable += Number(part.taxable_value || 0);
      target.igst += Number(part.igst || 0);
      target.cgst += Number(part.cgst || 0);
      target.sgst += Number(part.sgst || 0);
    });
  });
  return totals;
}
function rateSummaryHtml(totals) {
  const money = value => Number(value || 0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2});
  const rates = [0,5,12,18,28];
  const overall = rates.reduce((sum,rate) => {
    Object.keys(sum).forEach(key => { sum[key] += Number(totals[rate]?.[key] || 0); });
    return sum;
  },{taxable:0,igst:0,cgst:0,sgst:0});
  return `<div class="gst-rate-summary-grid"><span class="rate-head">GST Rate</span><span class="rate-head">Taxable</span><span class="rate-head">IGST</span><span class="rate-head">CGST</span><span class="rate-head">SGST</span>
    ${rates.map(rate => `<span>${rate}%</span><span>${money(totals[rate]?.taxable)}</span><span>${money(totals[rate]?.igst)}</span><span>${money(totals[rate]?.cgst)}</span><span>${money(totals[rate]?.sgst)}</span>`).join("")}
    <span>Total</span><span>${money(overall.taxable)}</span><span>${money(overall.igst)}</span><span>${money(overall.cgst)}</span><span>${money(overall.sgst)}</span></div>`;
}
function renderGstRateSummaries() {
  const notes = gstRows.filter(row => isSalesNote(row));
  // Rate-wise Gross and Net must follow the adjusted invoice rows displayed
  // in the table after Apply Invoice Amendment.
  const sales = gstRows.filter(row => !isSalesNote(row));
  $("#gstSalesRateSummary").innerHTML = rateSummaryHtml(gstRateBreakdown(sales));
  $("#gstNotesRateSummary").innerHTML = rateSummaryHtml(gstRateBreakdown(notes));
  let netHost = $("#gstNetRateSummary");
  if (!netHost) {
    const wrap = $("#gstRateSummaries");
    if (wrap) {
      const details = document.createElement("details");
      details.open = true;
      details.innerHTML = "<summary>Rate-wise Net Total (Output GST)</summary><div id=\"gstNetRateSummary\"></div>";
      wrap.appendChild(details);
      netHost = $("#gstNetRateSummary");
    }
  }
  if (netHost) netHost.innerHTML = rateSummaryHtml(gstr1NetRateBreakdown(sales, notes));
  $("#gstRateSummaries").classList.remove("hidden");
  renderGstr1Summary();
}
function componentTotals(sourceRows) {
  const breakdown = gstRateBreakdown(sourceRows);
  return Object.values(breakdown).reduce((total, values) => {
    ["taxable","igst","cgst","sgst"].forEach(key => { total[key] += Number(values[key] || 0); });
    return total;
  }, {taxable:0,igst:0,cgst:0,sgst:0});
}
const GSTR1_TOTAL_KEYS = ["taxable", "igst", "cgst", "sgst"];
function gstr1DocumentKind(row) {
  if (!isSalesNote(row)) return "invoice";
  const doc = String(row.document_type || "").toLowerCase();
  const invNo = String(row.invoice_no || "").trim().toUpperCase();
  if (/debit/.test(doc) || /^(DN|DR)/.test(invNo)) return "debit";
  if (/return|refund|sales return/.test(doc)) return "return";
  if (/credit/.test(doc) || /^(CN|CR)/.test(invNo)) return "credit";
  return "credit";
}
function gstr1MagnitudeTotals(rows) {
  const breakdown = gstRateBreakdown(rows || []);
  return GSTR1_TOTAL_KEYS.reduce((total, key) => {
    total[key] = Object.values(breakdown).reduce(
      (sum, values) => sum + Math.abs(Number(values[key] || 0)),
      0,
    );
    return total;
  }, { taxable: 0, igst: 0, cgst: 0, sgst: 0 });
}
function gstr1NetTotals(grossTotals, noteRows) {
  const notes = noteRows || [];
  const creditReturn = gstr1MagnitudeTotals(notes.filter(row =>
    ["credit", "return"].includes(gstr1DocumentKind(row))));
  const debit = gstr1MagnitudeTotals(notes.filter(row => gstr1DocumentKind(row) === "debit"));
  const net = addTotals(subtractTotals(grossTotals, creditReturn), debit);
  return { creditReturn, debit, net };
}
function gstr1NetRateBreakdown(salesRows, noteRows) {
  const rates = [0, 5, 12, 18, 28];
  const sales = gstRateBreakdown(salesRows || []);
  const creditReturn = gstRateBreakdown((noteRows || []).filter(row =>
    ["credit", "return"].includes(gstr1DocumentKind(row))));
  const debit = gstRateBreakdown((noteRows || []).filter(row => gstr1DocumentKind(row) === "debit"));
  const net = Object.fromEntries(rates.map(rate => [rate, { taxable: 0, igst: 0, cgst: 0, sgst: 0 }]));
  rates.forEach(rate => {
    GSTR1_TOTAL_KEYS.forEach(key => {
      net[rate][key] = round2(
        Number(sales[rate]?.[key] || 0)
        - Math.abs(Number(creditReturn[rate]?.[key] || 0))
        + Math.abs(Number(debit[rate]?.[key] || 0)),
      );
    });
  });
  return net;
}
function round2(value) {
  return Math.round(Number(value || 0) * 100) / 100;
}
function subtractTotals(left, right) {
  return Object.fromEntries(["taxable","igst","cgst","sgst"].map(key => [key, Number(left[key] || 0) - Number(right[key] || 0)]));
}
function addTotals(left, right) {
  return Object.fromEntries(["taxable","igst","cgst","sgst"].map(key => [key, Number(left[key] || 0) + Number(right[key] || 0)]));
}
function amendmentPreviewTotals() {
  const lessTotal = {taxable:0,igst:0,cgst:0,sgst:0};
  const addTotal = {taxable:0,igst:0,cgst:0,sgst:0};
  const sourceRows = (gstSalesOriginalRows.length ? gstSalesOriginalRows : gstRows).filter(row => !isSalesNote(row));
  const breakdown = gstRateBreakdown(sourceRows);
  [0,5,12,18,28].forEach(rate => {
    const less = Math.max(0, Number($(`#salesLess${rate}`)?.value || 0));
    const add = Math.max(0, Number($(`#salesAdd${rate}`)?.value || 0));
    const available = Number(breakdown[rate]?.taxable || 0);
    const sourceIgst = Number(breakdown[rate]?.igst || 0);
    const sourceCgst = Number(breakdown[rate]?.cgst || 0);
    const sourceSgst = Number(breakdown[rate]?.sgst || 0);
    const sourceTax = sourceIgst + sourceCgst + sourceSgst;
    const localSourceTax = sourceCgst + sourceSgst;
    const calculate = taxable => {
      const fullTax = taxable * rate / 100;
      const igst = sourceTax > 0 ? fullTax * sourceIgst / sourceTax : 0;
      const localTax = fullTax - igst;
      const cgst = localSourceTax > 0 ? localTax * sourceCgst / localSourceTax : localTax / 2;
      return {taxable, igst, cgst, sgst:localTax - cgst};
    };
    const lessPreview = calculate(Math.min(less, available));
    const addPreview = calculate(add);
    Object.keys(lessTotal).forEach(key => {
      lessTotal[key] += lessPreview[key];
      addTotal[key] += addPreview[key];
    });
    const netPreview = subtractTotals(lessPreview, addPreview);
    const money = value => Number(value || 0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2});
    const target = $(`#salesTaxPreview${rate}`);
    if (target) target.innerHTML = rate === 0 ? "" :
      `<span>IGST<strong>${money(netPreview.igst)}</strong></span><span>CGST<strong>${money(netPreview.cgst)}</strong></span><span>SGST<strong>${money(netPreview.sgst)}</strong></span>`;
  });
  return {less:lessTotal, add:addTotal};
}
function renderGstr1Summary() {
  if (!gstRows.length) return;
  const currentSales = gstRows.filter(row => !isSalesNote(row));
  const originalRows = gstSalesOriginalRows.length ? gstSalesOriginalRows : gstRows;
  const originalSales = originalRows.filter(row => !isSalesNote(row));
  const notes = gstRows.filter(row => isSalesNote(row));
  const b2b = componentTotals(originalSales.filter(row => String(row.gstin || "").trim()));
  const b2c = componentTotals(originalSales.filter(row => !String(row.gstin || "").trim()));
  const total = componentTotals(originalSales);
  const { creditReturn, debit, net: afterLess } = gstr1NetTotals(total, notes);
  const liveAmendment = amendmentPreviewTotals();
  const lessAmendment = liveAmendment.less;
  const addAmendment = liveAmendment.add;
  const grossTotal = addTotals(subtractTotals(afterLess, lessAmendment), addAmendment);
  const summaryRows = [
    ["B2B", b2b], ["B2C", b2c], ["Total", total], ["Less Credit Note", creditReturn],
    ["After Less Total", afterLess], ["Less Invoice Amendment", lessAmendment],
    ["Add Invoice Amendment", addAmendment], ["Gross Total", grossTotal],
  ];
  if (GSTR1_TOTAL_KEYS.some(key => Math.abs(Number(debit[key] || 0)) > 0.005)) {
    summaryRows.splice(4, 0, ["Add Debit Note", debit]);
  }
  const money = value => Number(value || 0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2});
  $("#gstr1SummaryRows").innerHTML = summaryRows.map(([label, values]) =>
    `<tr class="summary-${label.toLowerCase().replaceAll(" ","-")}"><th>${label}</th><td>${money(values.taxable)}</td><td>${money(values.igst)}</td><td>${money(values.cgst)}</td><td>${money(values.sgst)}</td></tr>`
  ).join("");
  $("#gstr1SummaryPanel").classList.remove("hidden");
}
["salesLess0","salesLess5","salesLess12","salesLess18","salesLess28","salesAdd0","salesAdd5","salesAdd12","salesAdd18","salesAdd28"].forEach(id => {
  $(`#${id}`).addEventListener("input", renderGstr1Summary);
});
function isSalesNote(row) {
  return row.is_sales_note || /note|refund|return/i.test(String(row.document_type || ""));
}
function filteredSalesNotes() {
  const values = {
    party: $("#gstNoteFilterParty").value.trim().toLowerCase(),
    number: $("#gstNoteFilterNumber").value.trim().toLowerCase(),
    date: $("#gstNoteFilterDate").value.trim().toLowerCase(),
    rate: $("#gstNoteFilterRate").value,
    hsn: $("#gstNoteFilterHsn").value.trim().toLowerCase()
  };
  return gstRows.map((row,index) => ({row,index})).filter(({row}) =>
    isSalesNote(row) &&
    (!values.party || String(row.party_ledger || row.party_name || "").toLowerCase().includes(values.party)) &&
    (!values.number || String(row.invoice_no || "").toLowerCase().includes(values.number)) &&
    (!values.date || String(row.invoice_date || "").toLowerCase().includes(values.date)) &&
    (!values.rate || salesRateLabel(row).split(", ").includes(values.rate)) &&
    (!values.hsn || salesHsnLabel(row).toLowerCase().includes(values.hsn))
  );
}
function filteredSalesRows() {
  const values = {
    gstin: $("#gstFilterGstin").value.trim().toLowerCase(), party: $("#gstFilterParty").value.trim().toLowerCase(),
    invoice: $("#gstFilterInvoice").value.trim().toLowerCase(), date: $("#gstFilterDate").value.trim().toLowerCase(),
    rate: $("#gstFilterRate").value, hsn: $("#gstFilterHsn").value.trim().toLowerCase(),
    status: $("#gstFilterStatus").value.trim().toLowerCase()
  };
  return gstRows.map((row, index) => ({row, index})).filter(({row}) =>
    !isSalesNote(row) && (!values.gstin || String(row.gstin || "").toLowerCase().includes(values.gstin)) &&
    (!values.party || String(row.party_ledger || row.party_name || "").toLowerCase().includes(values.party)) &&
    (!values.invoice || String(row.invoice_no || "").toLowerCase().includes(values.invoice)) &&
    (!values.date || String(row.invoice_date || "").toLowerCase().includes(values.date)) &&
    (!values.rate || salesRateLabel(row).split(", ").includes(values.rate)) &&
    (!values.hsn || salesHsnLabel(row).toLowerCase().includes(values.hsn)) &&
    (!values.status || `${salesRateLabel(row)} ready`.toLowerCase().includes(values.status))
  );
}
function rateFilteredAmounts(row, rateValue = "", hsnValue = "") {
  if (!rateValue && !hsnValue) return {
    quantity: salesQuantity(row), taxable: Number(row.taxable_value || 0),
    igst: Number(row.igst || 0), cgst: Number(row.cgst || 0), sgst: Number(row.sgst || 0)
  };
  const adjusted = Boolean(row.sales_allocations?.length);
  const parts = adjusted ? row.sales_allocations : (row.items?.length ? row.items : [row]);
  const wantedRate = rateValue ? Number(String(rateValue).replace("%", "")) : null;
  const wantedHsn = normalizedHsn(hsnValue);
  return parts.filter(part => {
    const rate = adjusted ? Number(part.rate || 0) : itemGstRate(part);
    return (wantedRate === null || rate === wantedRate) &&
      (!wantedHsn || normalizedHsn(part.hsn).includes(wantedHsn));
  }).reduce((totals, part) => {
    totals.quantity += Number(part.quantity || 0);
    totals.taxable += Number(part.taxable_value || 0);
    totals.igst += Number(part.igst || 0);
    totals.cgst += Number(part.cgst || 0);
    totals.sgst += Number(part.sgst || 0);
    return totals;
  }, {quantity:0,taxable:0,igst:0,cgst:0,sgst:0});
}
function renderSalesRows() {
  const money = value => Number(value || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const filtered = filteredSalesRows();
  const selectedRate = $("#gstFilterRate").value;
  const selectedHsn = $("#gstFilterHsn").value.trim();
  $("#gstRows").innerHTML = filtered.slice(0, 1000).map(({row, index}) => {
    const amounts = rateFilteredAmounts(row, selectedRate, selectedHsn);
    return `<tr>
    <td><input class="gst-sales-select" data-index="${index}" type="checkbox" ${row.selected ? "checked" : ""}></td>
    <td>${escapeHtml(row.gstin || "")}</td><td><input class="gst-sales-party" data-index="${index}" value="${escapeHtml(row.party_ledger || row.party_name || "Cash")}"></td>
    <td>${escapeHtml(row.invoice_no || "")}</td><td>${escapeHtml(row.invoice_date || "")}</td>
    <td class="money">${money(row.invoice_value)}</td><td><input class="gst-sales-item" data-index="${index}" list="tallyItemList" value="${escapeHtml(salesItemLabel(row))}"></td>
    <td>${escapeHtml(salesHsnLabel(row) || "—")}</td><td class="money">${money(amounts.quantity)}</td>
    <td class="money">${money(amounts.taxable)}</td><td>${escapeHtml(salesRateLabel(row))}</td>
    <td class="money">${money(amounts.igst)}</td><td class="money">${money(amounts.cgst)}</td><td class="money">${money(amounts.sgst)}</td>
    <td><span class="gst-status matched">${salesRateLabel(row)} Ready</span></td></tr>`;
  }).join("");
  const filteredAmounts = filtered.reduce((total, item) => {
    const amounts = rateFilteredAmounts(item.row, selectedRate, selectedHsn);
    ["taxable","igst","cgst","sgst"].forEach(key => { total[key] += amounts[key]; });
    return total;
  },{taxable:0,igst:0,cgst:0,sgst:0});
  $("#gstFilteredTaxable").textContent = money(filteredAmounts.taxable);
  $("#gstFilteredIgst").textContent = money(filteredAmounts.igst);
  $("#gstFilteredCgst").textContent = money(filteredAmounts.cgst);
  $("#gstFilteredSgst").textContent = money(filteredAmounts.sgst);
  document.querySelectorAll(".gst-sales-select").forEach(input => input.onchange = () => {
    gstRows[Number(input.dataset.index)].selected = input.checked;
  });
  document.querySelectorAll(".gst-sales-party").forEach(input => input.oninput = () => {
    gstRows[Number(input.dataset.index)].party_ledger = input.value.trim();
  });
  document.querySelectorAll(".gst-sales-item").forEach(input => input.onchange = () => {
    setRowItemName(Number(input.dataset.index), input.value.trim());
  });
  document.querySelectorAll(".gst-sales-item").forEach(input => input.onfocus = () => {
    showItemSuggestions(gstRows[Number(input.dataset.index)] || {});
  });
}
function noteValidationErrors(row) {
  const errors = [];
  if (!String(row.party_ledger || row.party_name || "").trim()) errors.push("Party ledger missing");
  if (!String(row.invoice_no || "").trim()) errors.push("Note number missing");
  if (!String(row.invoice_date || "").trim()) errors.push("Date missing");
  if (!row.sales_allocations?.length) errors.push("No taxable line items");
  if (row.sales_allocations?.length) {
    const hasItem = row.sales_allocations.some(a => String(a.item_name || "").trim());
    if (!hasItem) errors.push("Stock item missing");
  }
  const tv = Math.abs(Number(row.taxable_value || 0));
  if (tv < 0.01) errors.push("Taxable value is zero");
  return errors;
}
function renderSalesNoteRows() {
  const money = value => Number(value || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const voucherType = $("#gstNoteVoucherType").value;
  const allNotes = gstRows.map((row,index) => ({row,index})).filter(({row}) => isSalesNote(row));
  const notes = filteredSalesNotes();
  const selectedRate = $("#gstNoteFilterRate").value;
  const selectedHsn = $("#gstNoteFilterHsn").value.trim();
  $("#gstNotesSection").classList.toggle("hidden", !allNotes.length);
  $("#gstNoteRows").innerHTML = notes.map(({row,index}) => {
    const amounts = rateFilteredAmounts(row, selectedRate, selectedHsn);
    const missing = noteValidationErrors(row);
    const statusLabel = missing.length ? `<span class="gst-status review" title="${escapeHtml(missing.join(', '))}">${escapeHtml(missing[0])}</span>` : '<span class="gst-status matched">Ready</span>';
    return `<tr>
    <td><input class="gst-note-select" data-index="${index}" type="checkbox" ${row.selected ? "checked" : ""}></td>
    <td><input class="gst-note-party" data-index="${index}" value="${escapeHtml(row.party_ledger || row.party_name || "Cash")}"></td>
    <td>${escapeHtml(row.invoice_no || "")}</td><td>${escapeHtml(row.invoice_date || "")}</td><td class="money">${money(Math.abs(Number(row.invoice_value || 0)))}</td>
    <td><input class="gst-note-item" data-index="${index}" list="tallyItemList" value="${escapeHtml(salesItemLabel(row))}"></td>
    <td>${escapeHtml(salesHsnLabel(row) || "—")}</td><td class="money">${money(Math.abs(amounts.quantity))}</td><td class="money">${money(Math.abs(amounts.taxable))}</td>
    <td>${escapeHtml(salesRateLabel(row))}</td><td class="money">${money(Math.abs(amounts.igst))}</td><td class="money">${money(Math.abs(amounts.cgst))}</td><td class="money">${money(Math.abs(amounts.sgst))}</td>
    <td>${escapeHtml(voucherType)}</td><td>${statusLabel}</td></tr>`;
  }).join("");
  const noteAmounts = notes.reduce((total,item) => {
    const amounts = rateFilteredAmounts(item.row, selectedRate, selectedHsn);
    ["taxable","igst","cgst","sgst"].forEach(key => { total[key] += amounts[key]; });
    return total;
  },{taxable:0,igst:0,cgst:0,sgst:0});
  $("#gstNoteFilteredTaxable").textContent = money(noteAmounts.taxable);
  $("#gstNoteFilteredIgst").textContent = money(noteAmounts.igst);
  $("#gstNoteFilteredCgst").textContent = money(noteAmounts.cgst);
  $("#gstNoteFilteredSgst").textContent = money(noteAmounts.sgst);
  document.querySelectorAll(".gst-note-select").forEach(input => input.onchange = () => {
    gstRows[Number(input.dataset.index)].selected = input.checked;
    $("#gstSendNotesTallyBtn").disabled = false;
  });
  document.querySelectorAll(".gst-note-party").forEach(input => input.oninput = () => {
    gstRows[Number(input.dataset.index)].party_ledger = input.value.trim();
  });
  document.querySelectorAll(".gst-note-item").forEach(input => input.onchange = () => {
    setRowItemName(Number(input.dataset.index), input.value.trim());
  });
  document.querySelectorAll(".gst-note-item").forEach(input => input.onfocus = () => {
    showItemSuggestions(gstRows[Number(input.dataset.index)] || {});
  });
  $("#gstSendNotesTallyBtn").disabled = false;
}
["gstNoteFilterParty","gstNoteFilterNumber","gstNoteFilterDate","gstNoteFilterHsn"].forEach(id => {
  $(`#${id}`).oninput = renderSalesNoteRows;
});
$("#gstNoteFilterRate").onchange = renderSalesNoteRows;
$("#gstClearNoteFilters").onclick = () => {
  ["gstNoteFilterParty","gstNoteFilterNumber","gstNoteFilterDate","gstNoteFilterHsn"].forEach(id => { $(`#${id}`).value = ""; });
  $("#gstNoteFilterRate").value = "";
  renderSalesNoteRows();
};
$("#toggleInvoiceAmendment").onclick = () => {
  const body = $("#invoiceAmendmentBody");
  const open = body.classList.contains("hidden");
  body.classList.toggle("hidden", !open);
  $("#toggleInvoiceAmendment").textContent = open ? "⌃" : "⌄";
  $("#toggleInvoiceAmendment").setAttribute("aria-expanded", String(open));
  $("#toggleInvoiceAmendment").title = open ? "Close Invoice Amendment" : "Open Invoice Amendment";
};
["gstFilterGstin","gstFilterParty","gstFilterInvoice","gstFilterDate","gstFilterHsn","gstFilterStatus"].forEach(id => {
  $(`#${id}`).oninput = renderSalesRows;
});
$("#gstFilterRate").onchange = renderSalesRows;
$("#gstClearFilters").onclick = () => {
  ["gstFilterGstin","gstFilterParty","gstFilterInvoice","gstFilterDate","gstFilterHsn","gstFilterStatus"].forEach(id => { $(`#${id}`).value = ""; });
  $("#gstFilterRate").value = "";
  renderSalesRows();
};
$("#gstImportBtn").onclick = async () => {
  const files = [...$("#gstFileInput").files];
  const button = $("#gstImportBtn");
  const error = $("#gstError");
  error.classList.add("hidden");
  if (!files.length) {
    error.textContent = "Select a GST Portal JSON, Excel or ZIP file.";
    error.classList.remove("hidden");
    return;
  }
  button.disabled = true;
  button.textContent = "Importing...";
  try {
    const packed = [];
    for (const file of files) packed.push({ name: file.name, data: await fileToBase64(file) });
    const returnType = $("#gstReturnType").value;
    if (["GSTR-2A", "GSTR-2B"].includes(returnType)) {
      const combinedNames = files.map(file => file.name.toUpperCase()).join(" ");
      const looks2a = /GSTR?2A|R2A/.test(combinedNames), looks2b = /GSTR?2B|R2B/.test(combinedNames);
      if ((returnType === "GSTR-2A" && looks2b && !looks2a) || (returnType === "GSTR-2B" && looks2a && !looks2b)) {
        throw new Error(`Selected Return Type is ${returnType}, but the file name appears to be ${returnType === "GSTR-2A" ? "GSTR-2B" : "GSTR-2A"}. Select the correct Return Type and import again.`);
      }
    }
    const financial = ["GST Cash Ledger", "GST Credit Ledger", "GST Payment List", "Liability & ITC Comparison"].includes(returnType);
    const isGstr3b = ["threeway", "reconciliation", "payment"].includes(activeGstModule) && returnType === "GSTR-3B";
    const isGstr1Recon = activeGstModule === "threeway" && String(returnType).startsWith("GSTR-1");
    const endpoint = isGstr1Recon ? "/api/gst/recon/gstr1-import"
      : (isGstr3b ? "/api/gst/gstr3b-import" : (financial ? "/api/gst/financial-import" : "/api/gst/import"));
    const response = await fetch(endpoint, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify(financial ? { reportType: returnType, files: packed } : {
        returnType, files: packed,
        periodType: $("#margPeriodType").value,
        periodMonth: $("#margMonth").value,
        financialYear: $("#margFinancialYear").value.trim(),
        backupPassword: $("#margBackupPassword").value,
        returnPeriod: getGstReconPeriod(),
      })
    });
    const result = await response.json();
    if (!response.ok) {
      if (result.password_required) {
        $("#margPasswordField").classList.remove("hidden");
        $("#margBackupPassword").focus();
      }
      throw new Error(result.error || "GST file import failed.");
    }
    if (isGstr1Recon) {
      gstReconGstr1Rows = result.rows || [];
      gstDatasets["GSTR-1"] = gstReconGstr1Rows;
      // Do not change shared Period / Month on import — keep ALL / FY or user month.
      renderSalesReconStatus();
      updateSalesReconReady();
      setReconTab("sales");
      if (result.duplicate) {
        error.textContent = (result.imports || []).map(item => item.message).filter(Boolean).join(" ") || "Duplicate GSTR-1 import skipped.";
        error.classList.remove("hidden");
      }
      await refreshSalesReconDashboard(false);
      await refreshGstReconOverview(false);
      return;
    }
    if (isGstr3b) {
      const importedPeriods = result.imported_periods || Object.keys(result.periods || {});
      gstDatasets["GSTR-3B"] = {
        ...(result.summary || {}),
        ...(result.totals || {}),
        net_itc: result.net_itc || result.totals || {},
        periods: result.periods || {},
        net_periods: result.net_periods || {},
        outward_periods: result.outward_periods || {},
        nil_periods: result.nil_periods || {},
        reverse_charge: result.reverse_charge || {},
        imported_periods: importedPeriods,
        return_period: result.return_period || (importedPeriods.length > 1 ? "ALL" : (importedPeriods[0] || "")),
      };
      if (activeGstModule === "threeway") {
        await saveGstReconSession({ gstr3b: gstDatasets["GSTR-3B"] });
        renderGstReconStatus();
        populateGstr3bPeriodSelector(importedPeriods, getGstReconPeriod());
        if ($("#reconGstr3bPeriodStatus")) {
          const count = importedPeriods.length;
          $("#reconGstr3bPeriodStatus").textContent = result.duplicate
            ? `GSTR-3B: ${count} period(s) on file (duplicate skipped)`
            : `GSTR-3B: ${count} period(s) loaded`;
        }
        const fileErrors = (result.errors || []).map(item => `${item.file || "file"}: ${item.error || item.message || "failed"}`).filter(Boolean);
        if (fileErrors.length || result.duplicate) {
          const notes = [];
          if (fileErrors.length) notes.push(fileErrors.join(" "));
          if (result.duplicate) {
            notes.push((result.imports || []).map(item => item.message).filter(Boolean).join(" ") || "Duplicate GSTR-3B import skipped.");
          }
          error.textContent = notes.filter(Boolean).join(" ");
          error.classList.remove("hidden");
        }
        // Restore Import button before heavy Books/Tally dashboard refreshes so UI never stays on "Importing...".
        button.disabled = false;
        button.textContent = "Import GST File";
        setReconTab("gstr3b", { refresh: false });
        try { await refreshGstr3bDashboard(false); } catch (_) {}
        try { await refreshGstReconOverview(false); } catch (_) {}
        ensureGstReconPanelVisible();
        return;
      } else if(activeGstModule === "reconciliation") {
        $("#purchaseGstr3bStatus").textContent = `GSTR-3B: ${files.length} file(s) loaded`;
        populateGstr3bPeriodSelector(importedPeriods, getGstReconPeriod());
        renderPurchaseGstr3bComparison();
      } else { $("#gstResults").classList.remove("hidden");$("#gstMatchCounts").classList.remove("hidden");$("#gstMatchCounts").innerHTML=`<span>GSTR-3B Loaded<strong>${files.length}</strong></span><span>ITC Available<strong>${Number((result.totals||{}).igst||0)+Number((result.totals||{}).cgst||0)+Number((result.totals||{}).sgst||0)}</strong></span>`;$(".gst-next-note").textContent="GSTR-3B loaded for GST payment, input adjustment and available ITC review.";renderGstPaymentReview(); }
      return;
    }
    if (financial) {
      const report = (result.reports || [])[0] || {};
      gstDatasets[returnType]=report;
      const records = report.records || [];
      const periods = report.periods || {};
      $("#gstResults").classList.remove("hidden");
      $("#gstTallyPanel").classList.add("hidden");
      $("#gstMatchCounts").classList.remove("hidden");
      $("#gstMatchCounts").innerHTML = `<span>${escapeHtml(report.kind || returnType)}<strong>${Number(records.length || Object.keys(periods).length).toLocaleString("en-IN")}</strong></span>${report.total ? `<span>Total Payment<strong>${Number(report.total).toLocaleString("en-IN",{minimumFractionDigits:2})}</strong></span>` : ""}${report.rcm_records ? `<span>Reverse Charge<strong>${report.rcm_records.length}</strong></span>` : ""}`;
      $("#gstRows").innerHTML = records.slice(0, 500).map(row => `<tr><td>—</td><td>${escapeHtml(row.reference || row.CPIN || "")}</td><td>${escapeHtml(row.description || row["Mode of Payment"] || "")}</td><td>${escapeHtml(row.tax_period || "")}</td><td>${escapeHtml(row.date || row["Deposit Date"] || "")}</td><td class="money">${Number(row.Amount || 0).toLocaleString("en-IN")}</td><td colspan="5">${escapeHtml(row.transaction_type || row["Deposit Status"] || "")}</td></tr>`).join("");
      $(".gst-next-note").textContent = report.kind === "Tax Liability and ITC Comparison"
        ? `${Object.keys(periods).length} tax periods loaded for GSTR-1, GSTR-3B, GSTR-2B ITC and RCM comparison.`
        : "GST ledger/payment data loaded. Review it before preparing Tally entries.";
      if(activeGstModule==="payment")renderGstPaymentReview();
      return;
    }
    gstRows = result.rows || [];
    // Prefer API return_type; fall back to the selected Return Type (XLSX / ZIP imports).
    const resolvedReturnType = String(result.return_type || returnType || "").trim();
    const isSales = resolvedReturnType.startsWith("GSTR-1") || resolvedReturnType.toLowerCase().includes("sales");
    const datasetKey = resolvedReturnType.startsWith("GSTR-2A") ? "GSTR-2A"
      : (resolvedReturnType.startsWith("GSTR-2B") ? "GSTR-2B" : resolvedReturnType);
    gstDatasets[datasetKey] = gstRows;
    if (datasetKey === "GSTR-2A") gstReconDatasetsLoaded.add("gstr2a");
    if (datasetKey === "GSTR-2B") gstReconDatasetsLoaded.add("gstr2b");
    if (activeGstModule === "threeway") {
      // Invalidate any in-flight summary session load so it cannot wipe this import.
      gstReconSessionLoadSeq += 1;
      if (datasetKey === "GSTR-2B") {
        gstReconRows = gstRows;
        gstDatasets["GSTR-2B"] = gstRows;
        gstReconDatasetsLoaded.add("gstr2b");
        gstReconDatasetCounts.gstr2b = gstRows.length;
        await saveGstReconSession({ gstr2b: gstReconRows });
        try { renderGstRecon2bSummary(gstReconRows); } catch (_) {}
        // Do not change shared Period / Month on GSTR-2B import.
      }
      if (datasetKey === "GSTR-1" || String(datasetKey).startsWith("GSTR-1")) {
        gstReconGstr1Rows = gstRows;
        gstDatasets["GSTR-1"] = gstRows;
        gstReconDatasetsLoaded.add("gstr1");
        gstReconDatasetCounts.gstr1 = gstRows.length;
      }
      ensureGstReconPanelVisible();
      setReconTab("overview", { refresh: false });
      renderGstReconStatus();
      updateGstReconReady();
      try { await refreshGstReconOverview(false); } catch (_) {}
      ensureGstReconPanelVisible();
      return;
    }
    if (activeGstModule === "reconciliation") {
      if (datasetKey === "GSTR-2A") {
        await saveGstReconSession({ gstr2a: gstRows });
      }
      if (datasetKey === "GSTR-2B") {
        gstReconRows = gstRows;
        await saveGstReconSession({ gstr2b: gstReconRows });
      }
      updatePurchaseImportStatus();
      const bothLoaded = Boolean(gstDatasets["GSTR-2A"]?.length && gstDatasets["GSTR-2B"]?.length);
      if ($("#gstReconcileBtn")) $("#gstReconcileBtn").disabled = !bothLoaded;
      $("#gstResults").classList.remove("hidden");
      $("#gstSalesPanel").classList.add("hidden");
      $("#gstTableFilters").classList.add("hidden");
      $("#gstRateSummaries").classList.add("hidden");
      // If both returns are loaded, skip the very large standalone 2A table.
      // Reconcile directly and render the smaller reviewed result groups.
      if (datasetKey === "GSTR-2A" && !bothLoaded) renderPurchase2aWorkspace();
      if (datasetKey === "GSTR-2B") renderGstr2Summary([], [], [], []);
      if (bothLoaded) {
        // Once both returns are available, show reconciliation immediately
        // instead of leaving it below the long standalone GSTR-2A register.
        const reconciled = await reconcileGst();
        if (reconciled && gstRows.length && $("#purchase2aWorkspace")) {
          $("#purchase2aWorkspace").classList.add("hidden");
          $("#purchase2aWorkspace").parentElement?.classList.add("hidden");
        }
        const target = reconciled && gstRows.length ? ($("#purchase2a2bDashboard") || $("#gstr2SummaryPanel")) : null;
        if (target) {
          target.scrollIntoView({ behavior: "auto", block: "start" });
          setTimeout(() => target.scrollIntoView({ behavior: "auto", block: "start" }), 0);
        }
      } else {
        if ($("#purchase2aWorkspace")) {
          $("#purchase2aWorkspace").classList.remove("hidden");
          $("#purchase2aWorkspace").parentElement?.classList.remove("hidden");
        }
        error.textContent = `${datasetKey} imported (${gstRows.length.toLocaleString("en-IN")} document(s)). Import the other return to match.`;
        error.classList.remove("hidden");
        setPurchaseSheetView(datasetKey === "GSTR-2B" ? "2b" : "2a");
      }
      return;
    }
    const summary = result.summary || {};
    const money = value => Number(value || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    $("#gstInvoices").textContent = Number(summary.invoices || 0).toLocaleString("en-IN");
    $("#gstInvoiceValue").textContent = money(summary.invoice_value);
    $("#gstTaxable").textContent = money(summary.taxable_value);
    $("#gstIgst").textContent = money(summary.igst);
    $("#gstCgst").textContent = money(summary.cgst);
    $("#gstSgst").textContent = money(summary.sgst);
    $("#gstRows").innerHTML = gstRows.slice(0, 1000).map(row => `<tr>
      <td>—</td><td>${escapeHtml(row.gstin)}</td><td>${escapeHtml(row.party_name || "")}</td><td>${escapeHtml(row.invoice_no)}</td><td>${escapeHtml(row.invoice_date)}</td>
      <td class="money">${money(row.invoice_value)}</td><td class="money">${money(row.taxable_value)}</td>
      <td class="money">${money(row.igst)}</td><td class="money">${money(row.cgst)}</td><td class="money">${money(row.sgst)}</td>
      <td><span class="gst-imported">Imported</span></td></tr>`).join("");
    $("#gstTallyPanel").classList.add("hidden");
    $("#gstSalesPanel").classList.toggle("hidden", !isSales);
    $("#gstTableFilters").classList.toggle("hidden", !isSales);
    if (isSales) {
      gstSalesOriginalRows = JSON.parse(JSON.stringify(gstRows));
      prepareSalesRowsForTally(gstRows);
      gstRows.forEach(row => {
        row.selected = false;
        row.party_ledger = (!row.party_name || ["cash sales and purchase", "cash sales & purchase", "cash sale", "cash sales"].includes(String(row.party_name).trim().toLowerCase()))
          ? "Cash" : row.party_name;
      });
      // A new file must not retain the selected-count label from the
      // previously imported month/file.
      $("#gstSelectVisibleItems").textContent = "Select All Visible Rows";
      $("#gstSelectVisibleNotes").textContent = "Select All Visible Notes";
      renderSalesRows();
      renderSalesNoteRows();
      renderGstRateSummaries();
      $("#gstSalesAdjustmentNote").textContent = "Enter amendment amounts if required, then click Apply Invoice Amendment. Use zero if no reduction is needed.";
      $("#gstSendSalesTallyBtn").disabled = false;
    } else $("#gstRateSummaries").classList.add("hidden");
    $("#gstResults").classList.remove("hidden");
    $("#gstMatchCounts").classList.add("hidden");
  } catch (failure) {
    error.textContent = failure.message;
    error.classList.remove("hidden");
  } finally {
    button.disabled = false;
    button.textContent = "Import GST File";
  }
};
$("#gstApplySalesAdjustment").onclick = async () => {
  const button = $("#gstApplySalesAdjustment");
  const error = $("#gstError");
  error.classList.add("hidden");
  button.disabled = true;
  button.textContent = "Applying...";
  try {
    const selectedInvoiceKeys = new Set(gstRows.filter(row => row.selected && !isSalesNote(row)).map(row =>
      `${String(row.invoice_no || "").trim()}|${String(row.invoice_date || "").trim()}|${String(row.gstin || row.party_name || "").trim()}`
    ));
    if (!selectedInvoiceKeys.size) throw new Error("Select the invoices that should share the amendment.");
    const adjustmentRows = gstSalesOriginalRows.map(row => {
      const copy = {...row};
      const key = `${String(copy.invoice_no || "").trim()}|${String(copy.invoice_date || "").trim()}|${String(copy.gstin || copy.party_name || "").trim()}`;
      copy.selected = !isSalesNote(copy) && selectedInvoiceKeys.has(key);
      return copy;
    });
    const response = await fetch("/api/gst/sales/adjust", {
      method: "POST", headers: {"Content-Type": "application/json"},
      body: JSON.stringify({rows: adjustmentRows, reductions: {
        0: $("#salesLess0").value, 5: $("#salesLess5").value, 12: $("#salesLess12").value,
        18: $("#salesLess18").value, 28: $("#salesLess28").value
      }, additions: {
        0: $("#salesAdd0").value, 5: $("#salesAdd5").value, 12: $("#salesAdd12").value,
        18: $("#salesAdd18").value, 28: $("#salesAdd28").value
      }})
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "Rate-wise adjustment failed.");
    const requestedRates = [0,5,12,18,28].filter(rate =>
      Number($(`#salesLess${rate}`).value || 0) > 0 || Number($(`#salesAdd${rate}`).value || 0) > 0
    );
    const unsupportedRate = requestedRates.find(rate => !Object.prototype.hasOwnProperty.call(result.applied || {}, String(rate)));
    if (unsupportedRate !== undefined) {
      throw new Error(`Old Bank2Tally server is still running and does not support ${unsupportedRate}% amendment. Double-click Start_Bank2Tally.bat to restart the updated app.`);
    }
    gstRows = result.rows || [];
    gstRows.forEach(row => {
      const key = `${String(row.invoice_no || "").trim()}|${String(row.invoice_date || "").trim()}|${String(row.gstin || row.party_name || "").trim()}`;
      row.selected = !isSalesNote(row) && selectedInvoiceKeys.has(key);
    });
    renderSalesRows();
    renderSalesNoteRows();
    renderGstRateSummaries();
    $("#gstSendSalesTallyBtn").disabled = false;
    const applied = result.applied || {};
    $("#gstSalesAdjustmentNote").textContent = [0,5,12,18,28].map(rate =>
      `${rate}% Less ₹${Number(applied[String(rate)]?.requested || 0).toLocaleString("en-IN")} / Add ₹${Number(applied[String(rate)]?.added || 0).toLocaleString("en-IN")}`
    ).join(" • ") + ". Select reviewed bills, then send to Tally.";
    const summary = result.summary || {};
    $("#gstInvoiceValue").textContent = Number(summary.invoice_value || 0).toLocaleString("en-IN",{minimumFractionDigits:2});
    $("#gstTaxable").textContent = Number(summary.taxable_value || 0).toLocaleString("en-IN",{minimumFractionDigits:2});
    $("#gstIgst").textContent = Number(summary.igst || 0).toLocaleString("en-IN",{minimumFractionDigits:2});
    $("#gstCgst").textContent = Number(summary.cgst || 0).toLocaleString("en-IN",{minimumFractionDigits:2});
    $("#gstSgst").textContent = Number(summary.sgst || 0).toLocaleString("en-IN",{minimumFractionDigits:2});
  } catch (failure) {
    error.textContent = failure.message;
    error.classList.remove("hidden");
  } finally {
    button.disabled = false;
    button.textContent = "Apply Invoice Amendment";
  }
};
$("#gstSelectVisibleItems").onclick = () => {
  const visible = filteredSalesRows();
  const select = !visible.every(({row}) => row.selected);
  visible.forEach(({row}) => { row.selected = select; });
  renderSalesRows();
  $("#gstSelectVisibleItems").textContent = select ? `Selected ${visible.length} Visible Rows` : "Select All Visible Rows";
};
$("#gstSelectedItemName").onfocus = () => {
  const selected = gstRows.filter(row => row.selected && !isSalesNote(row));
  if (selected.length) showItemSuggestions(selected[0]);
};
$("#gstApplySelectedItem").onclick = () => {
  applyItemToSelectedRows({
    itemInputId: "gstSelectedItemName",
    predicate: (row) => !isSalesNote(row),
    refresh: () => { renderSalesRows(); renderSalesNoteRows(); },
  });
};
$("#gstNoteVoucherType").onchange = renderSalesNoteRows;
$("#gstSelectAllNotes").onclick = () => {
  const notes = gstRows.filter(row => isSalesNote(row));
  const select = !notes.every(row => row.selected);
  notes.forEach(row => { row.selected = select; });
  renderSalesNoteRows();
};
$("#gstSelectVisibleNotes").onclick = () => {
  const visible = filteredSalesNotes();
  const select = !visible.every(({row}) => row.selected);
  visible.forEach(({row}) => { row.selected = select; });
  renderSalesNoteRows();
  $("#gstSelectVisibleNotes").textContent = select ? `Selected ${visible.length} Visible Notes` : "Select All Visible Notes";
};
$("#gstSelectedNoteItemName").onfocus = () => {
  const selected = gstRows.filter(row => row.selected && isSalesNote(row));
  if (selected.length) showItemSuggestions(selected[0]);
};
$("#gstApplySelectedNoteItem").onclick = () => {
  applyItemToSelectedRows({
    itemInputId: "gstSelectedNoteItemName",
    predicate: (row) => isSalesNote(row),
    refresh: () => renderSalesNoteRows(),
    emptySelectMessage: "Select at least one visible Note/Return.",
    successNoun: "selected Note/Return row(s)",
  });
};
let tallySendTimer = null;
let tallySendValue = 0;
let lastTallySendAction = null;
function startTallySendProgress(title, retryAction) {
  clearInterval(tallySendTimer);
  tallySendValue = 2;
  lastTallySendAction = retryAction;
  const overlay = $("#tallySendOverlay");
  const card = overlay.querySelector(".tally-send-card");
  card.classList.remove("success", "failure");
  $("#tallySendTitle").textContent = title;
  $("#tallySendMessage").textContent = "Connecting and sending vouchers. Keep TallyPrime and the correct company open.";
  $("#tallySendClose").classList.add("hidden");
  $("#tallySendAgain").classList.add("hidden");
  if (!overlay.open) overlay.showModal();
  updateTallySendProgress(tallySendValue);
  tallySendTimer = setInterval(() => {
    const step = tallySendValue < 45 ? 5 : tallySendValue < 75 ? 2 : 1;
    tallySendValue = Math.min(92, tallySendValue + step);
    updateTallySendProgress(tallySendValue);
  }, 480);
}
function updateTallySendProgress(value) {
  $("#tallySendPercent").textContent = `${value}%`;
  $("#tallySendRing").style.setProperty("--progress", `${value * 3.6}deg`);
}
function finishTallySendProgress(success, message, options = {}) {
  clearInterval(tallySendTimer);
  tallySendTimer = null;
  updateTallySendProgress(100);
  const card = $("#tallySendOverlay").querySelector(".tally-send-card");
  const retryCount = Number(options.retryCount || 0);
  const counts = options.counts || null;
  card.classList.remove("success", "failure");
  if (counts) {
    const created = counts.created_confirmed != null ? counts.created_confirmed : (counts.created || 0);
    const alreadyExists = counts.already_exists_count || 0;
    const attempted = counts.missing_sent_count || 0;
    const stillMissing = counts.still_missing_count || 0;
    const failed = stillMissing || (counts.errors || 0);
    const exceptions = counts.exceptions || 0;
    const missingOnly = Boolean(counts.missing_only_mode);
    const isSuccess = failed === 0 && exceptions === 0;
    card.classList.add(isSuccess ? "success" : "failure");
    $("#tallySendTitle").textContent = failed > 0 || exceptions > 0
      ? (missingOnly ? "Retry Missing Only — Review Required" : "Tally Import — Review Required")
      : created > 0
        ? (missingOnly ? "Missing Sales Sent to Tally" : "Successfully Sent to Tally")
        : alreadyExists > 0 && attempted === 0
          ? "All Selected Sales Already Exist in Tally"
          : "Successfully Sent to Tally";
    const lines = [];
    // Never report "Successfully Created: 161" from already-in-Tally skips.
    if (counts.tally_sales_count_before != null) {
      lines.push(`Tally Sales before retry: ${Number(counts.tally_sales_count_before).toLocaleString("en-IN")}`);
    }
    if (missingOnly || attempted > 0) {
      lines.push(`Missing invoices attempted: ${Number(attempted).toLocaleString("en-IN")}`);
    }
    lines.push(`Successfully Created / Verified (this run): ${Number(created).toLocaleString("en-IN")}`);
    if (alreadyExists > 0 && !missingOnly) {
      lines.push(`Already Exists (verified in Tally, not resent): ${alreadyExists.toLocaleString("en-IN")}`);
    }
    if (stillMissing > 0) {
      lines.push(`Still MISSING IN TALLY: ${Number(stillMissing).toLocaleString("en-IN")}`);
    }
    if ((counts.master_missing_count || 0) > 0) {
      lines.push(`MASTER MISSING: ${Number(counts.master_missing_count).toLocaleString("en-IN")}`);
    }
    lines.push(`Failed: ${Number(failed).toLocaleString("en-IN")}`);
    if (exceptions > 0) {
      lines.push(`Tally exceptions (not treated as Already Exists): ${exceptions.toLocaleString("en-IN")}`);
    }
    if (counts.tally_sales_count_after != null) {
      lines.push(`Tally Sales count after send: ${Number(counts.tally_sales_count_after).toLocaleString("en-IN")}`);
    }
    if (!missingOnly) {
      lines.push(`Total selected: ${Number(counts.selected_total || 0).toLocaleString("en-IN")}`);
    }
    if ((counts.invoice_results || []).length) {
      lines.push("");
      lines.push("Per-invoice Tally results:");
      counts.invoice_results.forEach((item, index) => {
        const master = item.missing_master
          ? ` | MASTER MISSING: ${item.missing_master.name || item.missing_master.message || ""}`
          : "";
        lines.push(
          `${index + 1}. ${item.invoice_no || "—"} | ${item.party || "—"} | ${item.send_result || item.status || "—"} | ` +
          `CREATED=${item.created || 0} ALTERED=${item.altered || 0} ERRORS=${item.errors || 0} EXCEPTIONS=${item.exceptions || 0}` +
          `${item.lineerror ? ` | LINEERROR: ${item.lineerror}` : ""}${master}`
        );
      });
    } else {
      const report = counts.skipped_report || [];
      const missingRows = report.filter(item => !item.found_in_tally);
      if (missingRows.length) {
        lines.push("");
        lines.push("MISSING IN TALLY:");
        missingRows.slice(0, 30).forEach((item, i) => {
          lines.push(
            `  ${i + 1}. ${item.invoice_no || "—"} | ${item.invoice_date || "—"} | ${item.party || "—"} | ₹${Number(item.amount || 0).toLocaleString("en-IN", { minimumFractionDigits: 2 })} | Found=No`
          );
        });
      }
    }
    if ((counts.details || []).length && !(counts.invoice_results || []).length) {
      lines.push("");
      lines.push("Tally Error Details:");
      counts.details.forEach((detail, i) => lines.push(`  ${i + 1}. ${detail}`));
    }
    $("#tallySendMessage").textContent = lines.join("\n");
    $("#tallySendMessage").style.whiteSpace = "pre-line";
  } else {
    const acceptedMatch = String(message || "").match(/accepted\s+([\d,]+)\s+voucher/i);
    const acceptedCount = acceptedMatch ? Number(acceptedMatch[1].replace(/,/g, "")) : 0;
    const partial = !success && acceptedCount > 0;
    card.classList.add(success ? "success" : "failure");
    $("#tallySendTitle").textContent = success
      ? "Successfully Sent to Tally"
      : partial ? "Partially Sent to Tally" : "Not Sent to Tally";
    $("#tallySendMessage").textContent = message;
    $("#tallySendMessage").style.whiteSpace = "";
  }
  $("#tallySendClose").classList.remove("hidden");
  const again = $("#tallySendAgain");
  if (retryCount > 0 && options.retryAction) {
    lastTallySendAction = options.retryAction;
    again.textContent = `Retry Missing Only (${retryCount.toLocaleString()})`;
    again.classList.remove("hidden");
  } else if (success || counts) {
    again.classList.add("hidden");
  } else {
    again.textContent = "Try Again";
    again.classList.remove("hidden");
  }
}
$("#tallySendClose").onclick = () => $("#tallySendOverlay").close();
$("#tallySendAgain").onclick = () => {
  $("#tallySendOverlay").close();
  if (lastTallySendAction) setTimeout(lastTallySendAction, 50);
};

async function sendTallyInBatches(endpoint, rows, extraPayload, kindLabel) {
  const batchSize = 100;
  const batches = [];
  for (let index = 0; index < rows.length; index += batchSize) batches.push(rows.slice(index, index + batchSize));
  let created = 0;
  let altered = 0;
  let exceptions = 0;
  let errors = 0;
  let alreadyExists = 0;
  let missingSent = 0;
  let createdConfirmed = 0;
  let stillMissing = 0;
  let tallySalesAfter = null;
  const warnings = [];
  const details = [];
  const skippedReport = [];
  const failedReport = [];
  let ignored = 0;
  $("#tallySendMessage").textContent = "Checking Party Ledgers and GSTIN in TallyPrime...";
  const ledgerResponse = await fetch("/api/gst/party-ledgers/ensure", {
    method:"POST", headers:{"Content-Type":"application/json"},
    body:JSON.stringify({rows, ledgers:extraPayload.ledgers || {}})
  });
  const ledgerResult = await ledgerResponse.json();
  if (!ledgerResponse.ok) throw new Error(ledgerResult.error || "Party Ledger creation failed.");
  const ledgerMappings = ledgerResult.mappings || {};
  rows.forEach(row => {
    const sourceName = String(row.party_ledger || row.party_name || "").trim().toLowerCase();
    if (ledgerMappings[sourceName]) row.party_ledger = ledgerMappings[sourceName];
  });
  const salesLedgers = ledgerResult.salesLedgers || {};
  Object.entries(salesLedgers).forEach(([field, ledgerName]) => {
    if (extraPayload.ledgers && ledgerName) extraPayload.ledgers[field] = ledgerName;
    const input = $(`#${field}`);
    if (input && ledgerName) input.value = ledgerName;
  });
  const taxLedgers = ledgerResult.taxLedgers || {};
  Object.entries(taxLedgers).forEach(([field, ledgerName]) => {
    if (extraPayload.ledgers && ledgerName) extraPayload.ledgers[field] = ledgerName;
    const fieldToInput = {
      igstLedger:"salesIgstLedger", cgstLedger:"salesCgstLedger",
      sgstLedger:"salesSgstLedger", roundLedger:"salesRoundLedger"
    };
    const input = fieldToInput[field] ? $(`#${fieldToInput[field]}`) : null;
    if (input && ledgerName) input.value = ledgerName;
  });
  if (ledgerResult.created) {
    $("#tallySendMessage").textContent = `${ledgerResult.created} missing Party Ledger(s) created with GSTIN.`;
  }
  for (let index = 0; index < batches.length; index += 1) {
    $("#tallySendMessage").textContent = kindLabel === "Sales"
      ? `Verifying Tally Sales, then sending missing batch ${index + 1} of ${batches.length}...`
      : `Sending ${kindLabel} batch ${index + 1} of ${batches.length} to TallyPrime...`;
    const response = await fetch(endpoint, {
      method:"POST", headers:{"Content-Type":"application/json"},
      body:JSON.stringify({...extraPayload, rows:batches[index]})
    });
    const result = await response.json();
    if (!response.ok) throw new Error(`Batch ${index + 1}: ${result.error || "Tally import failed."}`);
    created += result.created || 0;
    altered += result.altered || 0;
    ignored += result.ignored || 0;
    exceptions += result.exceptions || 0;
    errors += result.errors || 0;
    alreadyExists += result.already_exists_count || 0;
    missingSent += result.missing_sent_count || 0;
    createdConfirmed += result.created_confirmed || 0;
    stillMissing += result.still_missing_count || 0;
    if (result.tally_sales_count_after != null) tallySalesAfter = result.tally_sales_count_after;
    if (result.warning) warnings.push(result.warning);
    if (result.details) result.details.forEach(d => { if (!details.includes(d)) details.push(d); });
    if (result.skipped_report) skippedReport.push(...result.skipped_report);
    if (result.failed_report) failedReport.push(...result.failed_report);
    // Mark local sales rows from verification report.
    if (kindLabel === "Sales" && Array.isArray(result.skipped_report)) {
      result.skipped_report.forEach(item => {
        const hit = rows.find(row => String(row.invoice_no || "") === String(item.invoice_no || ""));
        if (!hit) return;
        if (item.found_in_tally) {
          hit.tally_status = "Already in Tally";
          hit.tally_voucher_no = item.matching_tally_voucher_no || "";
        } else {
          hit.tally_status = "MISSING IN TALLY";
          hit.ready_for_sales_tally = true;
        }
      });
    }
    if (kindLabel === "Sales" && Array.isArray(result.failed_report)) {
      result.failed_report.forEach(item => {
        const hit = rows.find(row => String(row.invoice_no || "") === String(item.invoice_no || ""));
        if (!hit) return;
        hit.tally_status = "MISSING IN TALLY";
        hit.ready_for_sales_tally = true;
      });
    }
    clearInterval(tallySendTimer);
    tallySendTimer = null;
    updateTallySendProgress(Math.round(((index + 1) / batches.length) * 100));
  }
  return {
    created, altered, ignored, exceptions, errors, warnings, details,
    already_exists_count: alreadyExists,
    missing_sent_count: missingSent,
    created_confirmed: createdConfirmed,
    still_missing_count: stillMissing,
    tally_sales_count_after: tallySalesAfter,
    skipped_report: skippedReport,
    failed_report: failedReport,
    selected_total: rows.length,
  };
}

async function sendNotesToTally() {
  const voucherType = $("#gstNoteVoucherType").value;
  if (voucherType === "Do not import") return alert("Choose Credit Note, Debit Note or Journal.");
  const selectedRate = $("#gstNoteFilterRate").value;
  const selectedHsn = $("#gstNoteFilterHsn").value.trim();
  const checkedNotes = filteredSalesNotes().map(({row}) => row).filter(row => row.selected);
  if (!checkedNotes.length) return alert("Select at least one Credit/Debit Note or Sales Return by checking the checkbox.");
  const invalid = checkedNotes.filter(row => noteValidationErrors(row).length);
  if (invalid.length === checkedNotes.length) {
    const firstErrors = noteValidationErrors(invalid[0]);
    return alert(`Cannot send: ${firstErrors.join(", ")} for note ${invalid[0].invoice_no || "(no number)"}.`);
  }
  const selected = checkedNotes.filter(row => !noteValidationErrors(row).length).map(row => {
      if (!selectedRate && !selectedHsn) return row;
      const copy = JSON.parse(JSON.stringify(row));
      const wantedRate = selectedRate ? Number(selectedRate.replace("%", "")) : null;
      const wantedHsn = normalizedHsn(selectedHsn);
      copy.sales_allocations = (copy.sales_allocations || []).filter(item =>
        (wantedRate === null || Number(item.rate || 0) === wantedRate) &&
        (!wantedHsn || normalizedHsn(item.hsn).includes(wantedHsn))
      );
      copy.taxable_value = copy.sales_allocations.reduce((sum,item) => sum + Number(item.taxable_value || 0), 0);
      copy.igst = copy.sales_allocations.reduce((sum,item) => sum + Number(item.igst || 0), 0);
      copy.cgst = copy.sales_allocations.reduce((sum,item) => sum + Number(item.cgst || 0), 0);
      copy.sgst = copy.sales_allocations.reduce((sum,item) => sum + Number(item.sgst || 0), 0);
      copy.cess = copy.sales_allocations.reduce((sum,item) => sum + Number(item.cess || 0), 0);
      copy.invoice_value = copy.taxable_value + copy.igst + copy.cgst + copy.sgst + copy.cess;
      return copy;
    })
    .filter(row => !selectedRate && !selectedHsn || row.sales_allocations?.length);
  if (!selected.length) return alert("No valid notes to send after validation. Check the Review Status column for details.");
  if (!confirm(`Take a Tally backup first. Create ${selected.length} ${voucherType} voucher(s) in Tally?`)) return;
  const button = $("#gstSendNotesTallyBtn");
  button.disabled = true;
  button.textContent = "Sending...";
  startTallySendProgress(`Sending ${selected.length} ${voucherType} voucher(s)`, sendNotesToTally);
  try {
    const result = await sendTallyInBatches("/api/gst/notes/tally/send", selected, {voucherType,ledgers:{
        salesLedger0:$("#salesLedger0").value, salesLedger5:$("#salesLedger5").value, salesLedger12:$("#salesLedger12").value,
        salesLedger18:$("#salesLedger18").value, salesLedger28:$("#salesLedger28").value,
        igstLedger:$("#salesIgstLedger").value, cgstLedger:$("#salesCgstLedger").value,
        sgstLedger:$("#salesSgstLedger").value, roundLedger:$("#salesRoundLedger").value
      }}, "Note/Return");
    const tallyFailed = (result.errors || 0) > 0;
    finishTallySendProgress(!tallyFailed, "", {
      counts: {
        created: result.created || 0,
        exceptions: result.exceptions || 0,
        errors: result.errors || 0,
        ignored: result.ignored || 0,
        details: result.details || [],
      },
    });
  } catch (failure) {
    $("#gstError").textContent = failure.message;
    $("#gstError").classList.remove("hidden");
    finishTallySendProgress(false, failure.message);
  } finally {
    button.disabled = false;
    button.textContent = "Send Notes to Tally";
  }
}
$("#gstSendNotesTallyBtn").onclick = sendNotesToTally;

let lastMissingSalesVouchers = [];
let lastSalesVerifyPeriod = "";

function salesTallyLedgersPayload() {
  return {
    salesLedger0: $("#salesLedger0").value,
    salesLedger5: $("#salesLedger5").value,
    salesLedger12: $("#salesLedger12").value,
    salesLedger18: $("#salesLedger18").value,
    salesLedger28: $("#salesLedger28").value,
    igstLedger: $("#salesIgstLedger").value,
    cgstLedger: $("#salesCgstLedger").value,
    sgstLedger: $("#salesSgstLedger").value,
    roundLedger: $("#salesRoundLedger").value,
  };
}

function prepareSelectedSalesRowsForSend() {
  const selectedRate = $("#gstFilterRate").value;
  const selectedHsn = $("#gstFilterHsn").value.trim();
  const checkedRows = filteredSalesRows().map(({ row }) => row).filter(row => row.selected && !isSalesNote(row));
  return checkedRows.filter(row => row.ready_for_sales_tally !== false).map(row => {
    if (!selectedRate && !selectedHsn) return row;
    const copy = JSON.parse(JSON.stringify(row));
    const wantedRate = selectedRate ? Number(selectedRate.replace("%", "")) : null;
    const wantedHsn = normalizedHsn(selectedHsn);
    copy.sales_allocations = (copy.sales_allocations || []).filter(item =>
      (wantedRate === null || Number(item.rate || 0) === wantedRate) &&
      (!wantedHsn || normalizedHsn(item.hsn).includes(wantedHsn))
    );
    copy.taxable_value = copy.sales_allocations.reduce((sum, item) => sum + Number(item.taxable_value || 0), 0);
    copy.igst = copy.sales_allocations.reduce((sum, item) => sum + Number(item.igst || 0), 0);
    copy.cgst = copy.sales_allocations.reduce((sum, item) => sum + Number(item.cgst || 0), 0);
    copy.sgst = copy.sales_allocations.reduce((sum, item) => sum + Number(item.sgst || 0), 0);
    copy.cess = copy.sales_allocations.reduce((sum, item) => sum + Number(item.cess || 0), 0);
    copy.invoice_value = copy.taxable_value + copy.igst + copy.cgst + copy.sgst + copy.cess;
    return copy;
  }).filter(row => row.sales_allocations?.length && Number(row.invoice_value || 0) > 0.005);
}

function inferSalesReturnPeriod(rows) {
  const periods = new Set(
    (rows || [])
      .map(row => String(row.source_period || "").replace(/\D/g, "").slice(0, 6))
      .filter(value => value.length === 6)
  );
  if (periods.size === 1) return [...periods][0];
  // Fallback from invoice dates (DD-MM-YYYY / YYYY-MM-DD)
  const fromDates = new Set();
  (rows || []).forEach(row => {
    const raw = String(row.invoice_date || "").replace(/\//g, "-");
    let match = raw.match(/^(\d{2})-(\d{2})-(\d{4})$/);
    if (match) {
      fromDates.add(`${match[2]}${match[3]}`);
      return;
    }
    match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (match) fromDates.add(`${match[2]}${match[1]}`);
  });
  if (fromDates.size === 1) return [...fromDates][0];
  return "";
}

async function ensureSalesPartyLedgers(rows) {
  $("#tallySendMessage").textContent = "Checking Party Ledgers and GSTIN in TallyPrime...";
  const ledgers = salesTallyLedgersPayload();
  const ledgerResponse = await fetch("/api/gst/party-ledgers/ensure", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ rows, ledgers }),
  });
  const ledgerResult = await ledgerResponse.json();
  if (!ledgerResponse.ok) throw new Error(ledgerResult.error || "Party Ledger creation failed.");
  const mappings = ledgerResult.mappings || {};
  rows.forEach(row => {
    const key = String(row.party_ledger || row.party_name || "").trim().toLowerCase();
    if (mappings[key]) row.party_ledger = mappings[key];
  });
  Object.entries(ledgerResult.salesLedgers || {}).forEach(([field, ledgerName]) => {
    if (ledgerName && $(`#${field}`)) $(`#${field}`).value = ledgerName;
  });
  return salesTallyLedgersPayload();
}

async function fetchMissingSalesVouchers(candidateRows) {
  const returnPeriod = inferSalesReturnPeriod(candidateRows);
  lastSalesVerifyPeriod = returnPeriod;
  $("#tallySendMessage").textContent = "Fetching July/period Sales vouchers from Tally and comparing with GSTR-1...";
  const response = await fetch("/api/gst/sales/tally/verify", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      rows: candidateRows.map(row => ({ ...row, selected: true })),
      returnPeriod,
      tolerance: 1,
    }),
  });
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || "Could not verify Sales against Tally.");
  const missing = result.missing_sales_vouchers || result.missing_rows || [];
  lastMissingSalesVouchers = missing.map(row => ({ ...row, selected: true, ready_for_sales_tally: true }));
  return {
    missing: lastMissingSalesVouchers,
    already_exists_count: result.already_exists_count || 0,
    tally_sales_count: result.tally_sales_count || 0,
    period: result.period || returnPeriod || "ALL",
    report: result.skipped_report || [],
  };
}

async function sendMissingSalesIndividually(missingRows, ledgers) {
  if (!missingRows.length) {
    return {
      created_confirmed: 0,
      still_missing_count: 0,
      master_missing_count: 0,
      missing_sent_count: 0,
      already_exists_count: 0,
      tally_sales_count_after: null,
      invoice_results: [],
      failed_report: [],
    };
  }
  const results = [];
  let createdConfirmed = 0;
  let masterMissing = 0;
  let stillMissing = 0;
  let tallySalesAfter = null;
  for (let index = 0; index < missingRows.length; index += 1) {
    const row = missingRows[index];
    $("#tallySendMessage").textContent =
      `Sending missing voucher ${index + 1} of ${missingRows.length}: ${row.invoice_no || ""}`;
    updateTallySendProgress(Math.round(((index) / Math.max(missingRows.length, 1)) * 100));
    const response = await fetch("/api/gst/sales/tally/send-one", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        row: { ...row, selected: true, ready_for_sales_tally: true },
        ledgers,
        returnPeriod: lastSalesVerifyPeriod,
        tolerance: 1,
      }),
    });
    const result = await response.json();
    if (!response.ok) {
      results.push({
        invoice_no: row.invoice_no,
        party: row.party_ledger || row.party_name,
        send_result: "FAILED",
        status: "FAILED",
        created: 0,
        altered: 0,
        errors: 1,
        exceptions: 0,
        lineerror: result.error || "Send failed",
        found_in_tally_after: false,
      });
      stillMissing += 1;
      continue;
    }
    results.push(result);
    if (result.tally_sales_count != null) tallySalesAfter = result.tally_sales_count;
    if (result.found_in_tally_after) {
      createdConfirmed += 1;
      const hit = gstRows.find(item => String(item.invoice_no || "") === String(result.invoice_no || ""));
      if (hit) {
        hit.tally_status = "Sent to Tally";
        hit.tally_voucher_no = result.matching_tally_voucher_no || "";
      }
    } else if (result.status === "MASTER MISSING") {
      masterMissing += 1;
      stillMissing += 1;
      const hit = gstRows.find(item => String(item.invoice_no || "") === String(result.invoice_no || ""));
      if (hit) {
        hit.tally_status = "MASTER MISSING";
        hit.ready_for_sales_tally = true;
      }
    } else {
      stillMissing += 1;
      const hit = gstRows.find(item => String(item.invoice_no || "") === String(result.invoice_no || ""));
      if (hit) {
        hit.tally_status = "MISSING IN TALLY";
        hit.ready_for_sales_tally = true;
      }
    }
  }
  // STEP 10: re-fetch Tally Sales and recompute remaining missing among this set.
  $("#tallySendMessage").textContent = "Re-fetching Tally Sales count after missing-only send...";
  let finalMissingCount = stillMissing;
  try {
    const finalVerify = await fetchMissingSalesVouchers(
      missingRows.map(row => ({ ...row, selected: true, ready_for_sales_tally: true }))
    );
    finalMissingCount = finalVerify.missing.length;
    stillMissing = finalMissingCount;
    createdConfirmed = Math.max(0, missingRows.length - finalMissingCount);
    tallySalesAfter = finalVerify.tally_sales_count;
    lastMissingSalesVouchers = finalVerify.missing;
  } catch (_) {
    const remainingMissing = results.filter(item => !item.found_in_tally_after);
    lastMissingSalesVouchers = remainingMissing.map(item => {
      const source = missingRows.find(row => String(row.invoice_no || "") === String(item.invoice_no || ""));
      return source ? { ...source, selected: true, ready_for_sales_tally: true, tally_status: item.status } : null;
    }).filter(Boolean);
  }
  updateTallySendProgress(100);
  return {
    created_confirmed: createdConfirmed,
    still_missing_count: stillMissing,
    master_missing_count: masterMissing,
    missing_sent_count: missingRows.length,
    already_exists_count: 0,
    tally_sales_count_after: tallySalesAfter,
    invoice_results: results,
    failed_report: lastMissingSalesVouchers,
  };
}

function syncPurchaseItemMapsToSources(rows) {
  const maps = new Map();
  (rows || []).forEach((row) => {
    if (!row) return;
    const mapping = {
      sales_allocations: row.sales_allocations ? JSON.parse(JSON.stringify(row.sales_allocations)) : [],
      expense_ledger: row.expense_ledger || "",
      item_name: salesItemLabel(row) || "",
    };
    if (row.gstr2a) Object.assign(row.gstr2a, mapping);
    if (row.gstr2b) Object.assign(row.gstr2b, mapping);
    const key = purchaseItemMapKey(row);
    if (key && !key.startsWith("||")) maps.set(key, mapping);
  });
  if (!maps.size) return;
  ["GSTR-2A", "GSTR-2B"].forEach((datasetKey) => {
    (gstDatasets[datasetKey] || []).forEach((source) => {
      const mapping = maps.get(purchaseItemMapKey(source));
      if (mapping) Object.assign(source, mapping);
    });
  });
}

async function sendMissingSalesBulkFast(missingRows, ledgers) {
  const batchCount = Math.ceil(missingRows.length / 250);
  $("#tallySendMessage").textContent = `Sending ${missingRows.length} missing vouchers safely in ${batchCount} batch${batchCount === 1 ? "" : "es"}...`;
  updateTallySendProgress(75);
  const response = await fetch("/api/gst/sales/tally/send-bulk-fast", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      rows: missingRows.map(row => ({ ...row, selected: true, ready_for_sales_tally: true })),
      ledgers,
      returnPeriod: lastSalesVerifyPeriod,
      tolerance: 1,
    }),
  });
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || "Bulk Sales send failed.");
  lastMissingSalesVouchers = (result.missing_rows || []).map(row => ({
    ...row, selected: true, ready_for_sales_tally: true,
  }));
  return result;
}

function finishMissingSalesProgress(verifyInfo, sendInfo) {
  const attempted = sendInfo.missing_sent_count || 0;
  const created = sendInfo.created_confirmed || 0;
  const stillMissing = sendInfo.still_missing_count || 0;
  const masterMissing = sendInfo.master_missing_count || 0;
  const success = stillMissing === 0;
  finishTallySendProgress(success, "", {
    retryCount: stillMissing,
    retryAction: stillMissing ? retryMissingSalesOnly : null,
    counts: {
      created,
      created_confirmed: created,
      already_exists_count: 0, // do not surface 161 as created/selected on retry
      missing_sent_count: attempted,
      still_missing_count: stillMissing,
      master_missing_count: masterMissing,
      errors: stillMissing,
      exceptions: 0,
      ignored: 0,
      details: (sendInfo.invoice_results || []).map(item => item.lineerror).filter(Boolean),
      // Show only vouchers skipped because they already existed before send.
      // Pre-send MISSING rows must not reappear on a successful result screen.
      skipped_report: (verifyInfo.report || []).filter(item => item.found_in_tally),
      failed_report: sendInfo.failed_report || [],
      invoice_results: sendInfo.invoice_results || [],
      tally_sales_count_before: verifyInfo.tally_sales_count,
      tally_sales_count_after: sendInfo.tally_sales_count_after,
      selected_total: attempted,
      missing_only_mode: true,
    },
  });
}

async function retryMissingSalesOnly() {
  // CRITICAL: never call the full Send Sales flow / never resend all 181.
  const button = $("#gstSendSalesTallyBtn");
  button.disabled = true;
  button.textContent = "Retrying missing...";
  startTallySendProgress("Retry Missing Only — verifying against Tally Sales", retryMissingSalesOnly);
  try {
    // Compare full selected GSTR-1 Sales vs live Tally to rebuild missing_sales_vouchers.
    // Then SEND only that missing array — never the full selected set.
    const compareRows = prepareSelectedSalesRowsForSend();
    const seedMissing = lastMissingSalesVouchers.length
      ? lastMissingSalesVouchers.map(row => ({ ...row, selected: true, ready_for_sales_tally: true }))
      : [];
    const verifyPool = compareRows.length ? compareRows : seedMissing;
    if (!verifyPool.length) throw new Error("No Sales invoices available to compare against Tally.");
    const ledgers = await ensureSalesPartyLedgers(
      seedMissing.length ? seedMissing : verifyPool.slice(0, Math.min(verifyPool.length, 40))
    );
    const verifyInfo = await fetchMissingSalesVouchers(verifyPool);
    const missing = verifyInfo.missing;
    if (!missing.length) {
      finishTallySendProgress(true, `No missing Sales vouchers left. Tally Sales count: ${verifyInfo.tally_sales_count}.`);
      return;
    }
    if (missing.length >= 100 || (missing.length === verifyPool.length && verifyPool.length >= 100)) {
      throw new Error(
        `Retry aborted: missing count is ${missing.length} (full selection ${verifyPool.length}). ` +
        `Refusing to send all invoices again. Expected ~20 missing only.`
      );
    }
    const confirmText =
      `About to send ${missing.length} missing vouchers\n\n` +
      `Tally Sales before retry: ${verifyInfo.tally_sales_count}\n` +
      `Already in Tally (will NOT be resent): ${verifyInfo.already_exists_count}\n` +
      `Missing to send: ${missing.length}\n\n` +
      `Each voucher will be sent individually.`;
    if (confirmText.includes("About to send 181") || missing.length >= 100) {
      throw new Error("Retry aborted: confirm text would send 181 — wrong dataset.");
    }
    if (!confirm(confirmText)) {
      finishTallySendProgress(false, "Retry cancelled.");
      return;
    }
    $("#tallySendTitle").textContent = `Sending ${missing.length} missing Sales vouchers (one by one)`;
    const sendInfo = await sendMissingSalesIndividually(missing, ledgers);
    try { renderSalesRows(); } catch (_) {}
    finishMissingSalesProgress(verifyInfo, sendInfo);
  } catch (failure) {
    $("#gstError").textContent = failure.message;
    $("#gstError").classList.remove("hidden");
    finishTallySendProgress(false, failure.message);
  } finally {
    button.disabled = false;
    button.textContent = "Send Sales to Tally";
  }
}

async function sendSalesToTally() {
  const selected = prepareSelectedSalesRowsForSend();
  if (!selected.length) return alert("Select at least one reviewed Sales Invoice.");
  const button = $("#gstSendSalesTallyBtn");
  button.disabled = true;
  button.textContent = "Verifying...";
  // Retry must never be wired to this full-selection function.
  startTallySendProgress("Verifying Sales against Tally before send", retryMissingSalesOnly);
  try {
    const ledgers = await ensureSalesPartyLedgers(selected);
    const verifyInfo = await fetchMissingSalesVouchers(selected);
    const missing = verifyInfo.missing;
    if (!missing.length) {
      finishTallySendProgress(true, "", {
        counts: {
          created: 0,
          created_confirmed: 0,
          already_exists_count: verifyInfo.already_exists_count || selected.length,
          missing_sent_count: 0,
          still_missing_count: 0,
          errors: 0,
          exceptions: 0,
          skipped_report: verifyInfo.report || [],
          tally_sales_count_before: verifyInfo.tally_sales_count,
          tally_sales_count_after: verifyInfo.tally_sales_count,
          selected_total: 0,
          missing_only_mode: true,
        },
      });
      return;
    }
    if (!confirm(
      `About to send ${missing.length} missing vouchers\n\n` +
      `Selected Sales invoices: ${selected.length}\n` +
      `Already in Tally (will NOT be resent): ${verifyInfo.already_exists_count}\n` +
      `Tally Sales count now: ${verifyInfo.tally_sales_count}\n\n` +
      `Only the ${missing.length} missing vouchers will be sent safely in batches of 250.`
    )) {
      finishTallySendProgress(false, "Send cancelled.");
      return;
    }
    $("#tallySendTitle").textContent = `Sending ${missing.length} missing Sales vouchers (fast bulk)`;
    const sendInfo = await sendMissingSalesBulkFast(missing, ledgers);
    try { renderSalesRows(); } catch (_) {}
    finishMissingSalesProgress(
      { ...verifyInfo, already_exists_count: verifyInfo.already_exists_count },
      sendInfo,
    );
  } catch (failure) {
    $("#gstError").textContent = failure.message;
    $("#gstError").classList.remove("hidden");
    finishTallySendProgress(false, failure.message);
  } finally {
    button.disabled = false;
    button.textContent = "Send Sales to Tally";
  }
}
$("#gstSendSalesTallyBtn").onclick = sendSalesToTally;
function isPurchaseNote(row) {
  const section = String(row.section || "").trim().toUpperCase();
  return /credit\s*note|debit\s*note|amendment|return/i.test(String(row.document_type || ""))
    || ["B2BA", "CDNRA", "ECOA", "ISDA", "TDSA", "TCSA"].includes(section);
}
function purchaseLedgerMatch(row) {
  const gstin=String(row.gstin||"").replace(/\s+/g,"").toUpperCase(), party=String(row.party_name||"").trim().toLowerCase();
  const match=(tallyMasters.ledgers||[]).find(item=>gstin&&String(item.gstin||"").replace(/\s+/g,"").toUpperCase()===gstin)||(tallyMasters.ledgers||[]).find(item=>party&&String(item.name||"").trim().toLowerCase()===party);
  return (match||{}).name||row.party_name||"";
}
function resolvePurchaseLedgerForRow(row, rate) {
  if (row.expense_ledger) return { ledger: row.expense_ledger, required: false };
  const rateKey = `gstPurchaseLedger${Number(rate || 0)}`;
  const configured = String($(`#${rateKey}`)?.value || "").trim();
  const masters = (tallyMasters.ledgers || []).map((item) => String(item.name || "").trim().toLowerCase());
  if (configured && masters.length && masters.includes(configured.toLowerCase())) {
    return { ledger: configured, required: false };
  }
  if (configured && !masters.length) {
    // Tally masters not loaded yet — allow configured rate ledger but flag for review.
    return { ledger: configured, required: false };
  }
  if (configured) return { ledger: configured, required: false };
  return { ledger: "", required: true };
}
function purchaseGstRate(row) {
  const taxable = Number(row.taxable_value || 0);
  if (!taxable) return 0;
  const calculated = 100 * (Number(row.igst || 0) + Number(row.cgst || 0) + Number(row.sgst || 0) + Number(row.cess || 0)) / taxable;
  return [0, 5, 12, 18, 28].reduce((best, rate) => Math.abs(rate - calculated) < Math.abs(best - calculated) ? rate : best, 0);
}
function purchaseTotals(sourceRows) { return sourceRows.reduce((t,row)=>{t.invoices++;["taxable_value","igst","cgst","sgst"].forEach(k=>t[k]+=Number(row[k]||0));return t},{invoices:0,taxable_value:0,igst:0,cgst:0,sgst:0}); }
function purchaseSummaryAdd(a,b,sign=1) { const out={...a};["invoices","taxable_value","igst","cgst","sgst"].forEach(k=>out[k]=Number(out[k]||0)+(sign*Number(b[k]||0)));return out; }
function purchaseNoteKind(row) {
  const type=String(row.document_type||"").toLowerCase(), number=String(row.invoice_no||"").trim().toUpperCase();
  const section=String(row.section||"").trim().toUpperCase();
  if (/debit/.test(type)||/^(DN|DR)/.test(number)) return "debit";
  if (/credit|sales return|purchase return|refund/.test(type)||/^(CN|CR)/.test(number)) return "credit";
  const values=[row.taxable_value,row.igst,row.cgst,row.sgst,row.cess].map(Number);
  const isAmendment=["B2BA", "CDNRA", "ECOA", "ISDA", "TDSA", "TCSA"].includes(section)||/amendment/.test(type);
  if (isAmendment) return /decrease|reduc|downward/.test(type)||values.some(value=>value<0) ? "amendment_decrease" : "amendment_increase";
  return "b2b";
}
function purchaseAbsTotals(sourceRows) {
  return (sourceRows || []).reduce((total, row) => {
    total.invoices += 1;
    ["taxable_value", "igst", "cgst", "sgst", "cess"].forEach(key => {
      total[key] += Math.abs(Number(row[key] || 0));
    });
    return total;
  }, { invoices: 0, taxable_value: 0, igst: 0, cgst: 0, sgst: 0, cess: 0 });
}
function gstr2bDocBucket(row) {
  const kind = isPurchaseNote(row) ? purchaseNoteKind(row) : "b2b";
  if (kind === "credit") return "Credit Note";
  if (kind === "debit") return "Debit Note";
  if (kind === "amendment_increase") return "Amendment Increase";
  if (kind === "amendment_decrease") return "Amendment Decrease";
  if (kind === "amendment") return "Amendment";
  return "Invoice";
}
function roundMoney(value) {
  return Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100;
}
function signedPurchaseTax(row) {
  const fields = {
    taxable_value: Number(row.taxable_value || 0),
    igst: Number(row.igst || 0),
    cgst: Number(row.cgst || 0),
    sgst: Number(row.sgst || 0),
    cess: Number(row.cess || 0),
  };
  const bucket = gstr2bDocBucket(row);
  const alreadySigned = Boolean(row._portal_signed)
    || Number(row.document_sign) === -1
    || fields.taxable_value < 0 || fields.igst < 0 || fields.cgst < 0 || fields.sgst < 0 || fields.cess < 0;
  let sign = 1;
  if (!alreadySigned && (bucket === "Credit Note" || bucket === "Amendment Decrease")) sign = -1;
  if (sign < 0) {
    Object.keys(fields).forEach(key => { fields[key] = roundMoney(sign * fields[key]); });
  }
  fields.itc = roundMoney(fields.igst + fields.cgst + fields.sgst + fields.cess);
  fields.bucket = bucket;
  return fields;
}
function buildGstr2bGrossNetSummary(rows) {
  const empty = () => ({ count: 0, taxable_value: 0, igst: 0, cgst: 0, sgst: 0, cess: 0, itc: 0 });
  const buckets = {
    Invoice: empty(),
    "Debit Note": empty(),
    "Credit Note": empty(),
    "Amendment Increase": empty(),
    "Amendment Decrease": empty(),
    Amendment: empty(),
  };
  const net = empty();
  const invoiceByOriginalKey = new Map();
  const originalKey = row => [
    String(row.gstin || "").trim().toUpperCase(),
    String(row.invoice_no || "").replace(/[^a-z0-9]/gi, "").toUpperCase(),
    String(row.invoice_date || row.original_invoice_date || "").trim(),
  ].join("|");
  (rows || []).forEach(row => {
    if (gstr2bDocBucket(row) === "Invoice") invoiceByOriginalKey.set(originalKey(row), row);
  });
  (rows || []).forEach(row => {
    let signed = signedPurchaseTax(row);
    const bucketName = signed.bucket in buckets ? signed.bucket : "Invoice";
    if (bucketName.startsWith("Amendment")) {
      const prior = invoiceByOriginalKey.get([
        String(row.gstin || "").trim().toUpperCase(),
        String(row.original_invoice_no || row.invoice_no || "").replace(/[^a-z0-9]/gi, "").toUpperCase(),
        String(row.original_invoice_date || row.invoice_date || "").trim(),
      ].join("|"));
      if (prior) {
        const old = signedPurchaseTax(prior);
        signed = { ...signed };
        ["taxable_value", "igst", "cgst", "sgst", "cess"].forEach(key => {
          signed[key] = roundMoney(Number(signed[key] || 0) - Number(old[key] || 0));
        });
        signed.itc = roundMoney(signed.igst + signed.cgst + signed.sgst + signed.cess);
      }
    }
    const bucket = buckets[bucketName];
    bucket.count += 1;
    net.count += 1;
    ["taxable_value", "igst", "cgst", "sgst", "cess", "itc"].forEach(key => {
      bucket[key] = roundMoney(bucket[key] + Number(signed[key] || 0));
      net[key] = roundMoney(net[key] + Number(signed[key] || 0));
    });
  });
  return {
    invoice_count: buckets.Invoice.count,
    credit_note_count: buckets["Credit Note"].count,
    debit_note_count: buckets["Debit Note"].count,
    amendment_count: buckets["Amendment Increase"].count + buckets["Amendment Decrease"].count + buckets.Amendment.count,
    document_count: net.count,
    gross_invoice_itc: buckets.Invoice.itc,
    credit_note_itc: Math.abs(buckets["Credit Note"].itc),
    debit_note_itc: buckets["Debit Note"].itc,
    amendment_itc: roundMoney(
      buckets["Amendment Increase"].itc + buckets["Amendment Decrease"].itc + buckets.Amendment.itc
    ),
    buckets,
    net_taxable: net.taxable_value,
    net_igst: net.igst,
    net_cgst: net.cgst,
    net_sgst: net.sgst,
    net_cess: net.cess,
    net_itc: net.itc,
    formula: "Gross Invoice ITC − Credit Note ITC + Debit Note ITC ± Amendments",
  };
}
function purchaseGstr2Breakdown(source) {
  const summary = buildGstr2bGrossNetSummary(source || []);
  const asTotals = (bucket, countOverride) => ({
    invoices: countOverride != null ? countOverride : (bucket.count || 0),
    taxable_value: bucket.taxable_value || 0,
    igst: bucket.igst || 0,
    cgst: bucket.cgst || 0,
    sgst: bucket.sgst || 0,
    cess: bucket.cess || 0,
    itc: bucket.itc || 0,
  });
  const total = asTotals(summary.buckets.Invoice, summary.invoice_count);
  const debit = asTotals(summary.buckets["Debit Note"], summary.debit_note_count);
  const creditBucket = summary.buckets["Credit Note"];
  const credit = asTotals({
    ...creditBucket,
    taxable_value: Math.abs(creditBucket.taxable_value || 0),
    igst: Math.abs(creditBucket.igst || 0),
    cgst: Math.abs(creditBucket.cgst || 0),
    sgst: Math.abs(creditBucket.sgst || 0),
    cess: Math.abs(creditBucket.cess || 0),
    itc: Math.abs(creditBucket.itc || 0),
  }, summary.credit_note_count);
  const increase = asTotals(summary.buckets["Amendment Increase"]);
  const decrease = asTotals(summary.buckets["Amendment Decrease"]);
  const net = {
    invoices: summary.invoice_count,
    taxable_value: summary.net_taxable,
    igst: summary.net_igst,
    cgst: summary.net_cgst,
    sgst: summary.net_sgst,
    cess: summary.net_cess,
    itc: summary.net_itc,
  };
  return { total, debit, credit, increase, decrease, gross: net, net, summary };
}
function purchaseMiniTable(entries,selectable=false,editable=false) {
  const money=v=>Number(v||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2});
  if(!entries.length)return `<p class="gst-next-note">No rows in this section.</p>`;
  const edit=(index,field,value,type="text")=>`<input class="purchase-row-edit" data-index="${index}" data-field="${field}" type="${type}" value="${escapeHtml(value??"")}">`;
  return `<table><thead><tr>${selectable?"<th>Select</th>":""}<th>GSTIN</th><th>Party / Tally Ledger</th><th>Item / Tally Stock Item</th><th>Invoice No.</th><th>Original Invoice Date</th><th>GSTR-2B Period</th><th>Tally Entry Date</th><th>Invoice Value</th><th>Taxable</th><th>GST Rate</th><th>IGST</th><th>CGST</th><th>SGST</th><th>Status</th></tr></thead><tbody>${entries.map(({row,index})=>`<tr>${selectable?`<td><input class="purchase-match-select" data-index="${index}" type="checkbox" ${row.selected?"checked":""}></td>`:""}<td>${escapeHtml(row.gstin||"")}</td><td>${selectable?`<input class="purchase-party-ledger" data-index="${index}" list="tallyLedgerList" value="${escapeHtml(row.party_ledger||row.party_name||"")}">`:escapeHtml(row.party_name||"")}</td><td>${selectable?`<input class="purchase-item-name" data-index="${index}" list="tallyItemList" value="${escapeHtml(salesItemLabel(row)||"")}" placeholder="Select Tally Item">`:escapeHtml(salesItemLabel(row)||"")}</td><td>${editable?edit(index,"invoice_no",row.invoice_no):escapeHtml(row.invoice_no||"")}</td><td>${editable?edit(index,"original_invoice_date",row.original_invoice_date||row.invoice_date):escapeHtml(row.original_invoice_date||row.invoice_date||"")}</td><td>${escapeHtml(row.gstr2b_period||"—")}</td><td>${editable?edit(index,"tally_entry_date",row.tally_entry_date):escapeHtml(row.tally_entry_date||"—")}</td><td class="money">${editable?edit(index,"invoice_value",row.invoice_value,"number"):money(row.invoice_value)}</td><td class="money">${editable?edit(index,"taxable_value",row.taxable_value,"number"):money(row.taxable_value)}</td><td>${purchaseGstRate(row)}%</td><td class="money">${editable?edit(index,"igst",row.igst,"number"):money(row.igst)}</td><td class="money">${editable?edit(index,"cgst",row.cgst,"number"):money(row.cgst)}</td><td class="money">${editable?edit(index,"sgst",row.sgst,"number"):money(row.sgst)}</td><td>${escapeHtml(row.status||row.document_type||"")}</td></tr>`).join("")}</tbody></table>`;
}
function purchaseTableWithTotals(entries,selectable=false,editable=false,sectionKey=""){
  if(!entries.length)return purchaseMiniTable(entries,selectable,editable);
  const money=v=>Number(v||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2}),t=purchaseTotals(entries.map(entry=>entry.row));
  if (!sectionKey) {
    return `<div class="purchase-box-totals"><span>Taxable <strong>${money(t.taxable_value)}</strong></span><span>IGST <strong>${money(t.igst)}</strong></span><span>CGST <strong>${money(t.cgst)}</strong></span><span>SGST <strong>${money(t.sgst)}</strong></span></div>${purchaseMiniTable(entries,selectable,editable)}`;
  }
  const pageSize = PURCHASE_RECON_PAGE_SIZE;
  const currentPage = purchaseReconPages[sectionKey] || 1;
  const paged = paginateEntries(entries, currentPage, pageSize);
  purchaseReconPages[sectionKey] = paged.page;
  const pager = paged.total > pageSize
    ? purchasePagerHtml(sectionKey, paged.page, paged.pages, paged.total)
    : "";
  return `<div class="purchase-box-totals"><span>Taxable <strong>${money(t.taxable_value)}</strong></span><span>IGST <strong>${money(t.igst)}</strong></span><span>CGST <strong>${money(t.cgst)}</strong></span><span>SGST <strong>${money(t.sgst)}</strong></span></div>${purchaseMiniTable(paged.pageEntries,selectable,editable)}${pager}`;
}
function bindPurchaseReconPagers(root=document) {
  root.querySelectorAll(".purchase-recon-pager[data-purchase-page-section]").forEach((pager) => {
    const section = pager.getAttribute("data-purchase-page-section");
    const prev = pager.querySelector(".purchase-page-prev");
    const next = pager.querySelector(".purchase-page-next");
    if (prev) prev.onclick = () => {
      purchaseReconPages[section] = Math.max(1, (purchaseReconPages[section] || 1) - 1);
      renderPurchaseReconciliation();
    };
    if (next) next.onclick = () => {
      purchaseReconPages[section] = (purchaseReconPages[section] || 1) + 1;
      renderPurchaseReconciliation();
    };
  });
}
function renderPurchase2aWorkspace() {
  const source = gstDatasets["GSTR-2A"] || [];
  if (!source.length) {
    if ($("#purchase2aEmpty")) {
      $("#purchase2aEmpty").classList.remove("hidden");
      $("#purchase2aContent").classList.add("hidden");
      $("#purchase2aSummaryRows").innerHTML = "";
      $("#purchase2aBoxes").innerHTML = "";
    }
    return;
  }
  const indexed = source.map((row, index) => ({ row, index }));
  const kind = row => isPurchaseNote(row) ? purchaseNoteKind(row) : "b2b";
  const groups = [
    ["B2B Purchase", "b2b", "purchase2a-b2b"],
    ["Credit Note", "credit", "purchase2a-credit"],
    ["Debit Note", "debit", "purchase2a-debit"],
    ["Amendment Bills", "amendment", "purchase2a-amendment"]
  ];
  const groupRows = key => indexed.filter(({row}) => key === "amendment"
    ? ["amendment_increase", "amendment_decrease"].includes(kind(row))
    : kind(row) === key);
  const breakdown = purchaseReturnBreakdown(source, "GSTR-2A");
  const labels = ["B2B Purchase  +", "Debit Note  +", "Credit Note  −", "Amendment Increase  +", "Amendment Decrease  −", "Gross Total"];
  const money2 = value => Number(value || 0).toLocaleString("en-IN", {minimumFractionDigits:2, maximumFractionDigits:2});
  const rateRows = entries => entries.flatMap(({row,index}) => {
    const items = Array.isArray(row.items) && row.items.length ? row.items : [{gst_rate:purchaseGstRate(row),taxable_value:row.taxable_value,igst:row.igst,cgst:row.cgst,sgst:row.sgst}];
    return items.map((item,itemIndex)=>({row,index,item,itemIndex}));
  });
  $("#purchase2aSummaryRows").innerHTML = breakdown.map(([, total], index) => `<tr class="${index === breakdown.length - 1 ? "purchase2a-gross" : ""}"><th>${labels[index]}</th><td>${Math.max(0, Number(total.invoices || 0)).toLocaleString("en-IN")}</td><td>${money2(total.taxable_value)}</td><td>${money2(total.igst)}</td><td>${money2(total.cgst)}</td><td>${money2(total.sgst)}</td></tr>`).join("");
  $("#purchase2aBoxes").innerHTML = groups.map(([title, key, cls]) => {
    const entries = groupRows(key), totals = purchaseTotals(entries.map(entry => entry.row));
    const allRateRows = rateRows(entries).filter(({row, item}) => {
      const filter = purchase2aFilters[key] || { q: "", rate: "" };
      const rateVal = String(Number(item.gst_rate ?? purchaseGstRate(item) ?? 0));
      if (filter.rate && rateVal !== filter.rate) return false;
      if (filter.q) {
        const blob = [row.gstin, row.party_name, row.invoice_no, row.invoice_date, rateVal, row.document_type, item.hsn_code, row.hsn_code]
          .join(" ")
          .toLowerCase();
        if (!blob.includes(filter.q)) return false;
      }
      return true;
    });
    const pageSize = PURCHASE_2A_PAGE_SIZE;
    const paged = paginateEntries(allRateRows, purchase2aPages[key] || 1, pageSize);
    purchase2aPages[key] = paged.page;
    const filterState = purchase2aFilters[key] || { q: "", rate: "" };
    const rows = paged.pageEntries.map(({row,index,item}) => { const rate=Number(item.gst_rate??purchaseGstRate(item)??0),taxable=Number(item.taxable_value||0),igst=Number(item.igst||0),cgst=Number(item.cgst||0),sgst=Number(item.sgst||0),search=[row.gstin,row.party_name,row.invoice_no,row.invoice_date,rate,row.document_type,item.hsn_code,row.hsn_code].join(" ").toLowerCase(),itemName=String(item.item_name||row.item_name||row.stock_item||salesItemLabel(row)||""); return `<tr data-index="${index}" data-purchase-search="${escapeHtml(search)}" data-purchase-rate="${rate}"><td><input class="purchase2a-select" data-index="${index}" type="checkbox" ${row.selected?"checked":""}></td><td>${escapeHtml(row.gstin||"")}</td><td>${escapeHtml(row.party_name||"")}</td><td>${escapeHtml(row.invoice_no||"")}</td><td>${escapeHtml(row.invoice_date||row.original_invoice_date||"")}</td><td class="money">${money2(row.invoice_value)}</td><td><input class="purchase2a-item" data-index="${index}" list="tallyItemList" value="${escapeHtml(itemName)}" placeholder="Select Tally Item"></td><td>${escapeHtml(item.hsn_code||item.hsn||row.hsn_code||"—")}</td><td class="money">${money2(item.quantity??row.quantity??0)}</td><td class="money">${money2(taxable)}</td><td class="gst-rate-cell">${money2(rate)}%</td><td class="money">${money2(igst)}</td><td class="money">${money2(cgst)}</td><td class="money">${money2(sgst)}</td>${key==="amendment"?`<td>${purchaseNoteKind(row)==="amendment_decrease"?"Decrease":"Increase"}</td>`:""}</tr>`; }).join("") || `<tr><td colspan="15" class="purchase2a-no-data">No ${title} found.</td></tr>`;
    const pager = paged.total > pageSize
      ? `<div class="sales-recon-pager purchase-2a-pager" data-purchase2a-section="${key}"><button type="button" class="secondary purchase2a-page-prev" ${paged.page <= 1 ? "disabled" : ""}>Previous</button><span>Page ${paged.page} / ${paged.pages} (${paged.total.toLocaleString("en-IN")} rows)</span><button type="button" class="secondary purchase2a-page-next" ${paged.page >= paged.pages ? "disabled" : ""}>Next</button></div>`
      : "";
    return `<details class="purchase2a-box ${cls}" data-purchase-section="${key}" open><summary><span>${title}</span><span class="purchase2a-count">${entries.length.toLocaleString("en-IN")}</span></summary><div class="purchase2a-tools"><button type="button" class="purchase2a-select-all">Select All</button><input class="purchase2a-search" type="search" placeholder="Search GSTIN, party, invoice no., date or HSN" value="${escapeHtml(filterState.q || "")}"><input class="purchase2a-bulk-item" list="tallyItemList" placeholder="Item for selected invoices"><button type="button" class="purchase2a-apply-item">Apply Item</button><select class="purchase2a-rate-filter"><option value="">All GST Rates</option><option value="0" ${filterState.rate==="0"?"selected":""}>0%</option><option value="5" ${filterState.rate==="5"?"selected":""}>5%</option><option value="12" ${filterState.rate==="12"?"selected":""}>12%</option><option value="18" ${filterState.rate==="18"?"selected":""}>18%</option><option value="28" ${filterState.rate==="28"?"selected":""}>28%</option></select><button type="button" class="purchase2a-clear-filter">Clear</button><button type="button" class="purchase2a-send primary">Send to Tally</button></div><div class="purchase2a-table-wrap"><table><thead><tr class="purchase2a-column-totals"><th colspan="9"></th><th><small>Taxable</small><strong>₹${money2(totals.taxable_value)}</strong></th><th></th><th><small>IGST</small><strong>₹${money2(totals.igst)}</strong></th><th><small>CGST</small><strong>₹${money2(totals.cgst)}</strong></th><th><small>SGST</small><strong>₹${money2(totals.sgst)}</strong></th>${key==="amendment"?"<th></th>":""}</tr><tr><th>Select</th><th>GSTIN</th><th>Party</th><th>Invoice No.</th><th>Date</th><th>Invoice Value</th><th>Item / Tally Stock Item</th><th>HSN</th><th>Quantity</th><th>Taxable</th><th>GST Rate</th><th>IGST</th><th>CGST</th><th>SGST</th>${key==="amendment"?"<th>Type</th>":""}</tr></thead><tbody>${rows}</tbody></table></div>${pager}</details>`;
  }).join("");
  document.querySelectorAll(".purchase-2a-pager").forEach((pager) => {
    const key = pager.getAttribute("data-purchase2a-section");
    const prev = pager.querySelector(".purchase2a-page-prev");
    const next = pager.querySelector(".purchase2a-page-next");
    if (prev) prev.onclick = () => {
      purchase2aPages[key] = Math.max(1, (purchase2aPages[key] || 1) - 1);
      renderPurchase2aWorkspace();
    };
    if (next) next.onclick = () => {
      purchase2aPages[key] = (purchase2aPages[key] || 1) + 1;
      renderPurchase2aWorkspace();
    };
  });
  const debouncedPurchase2aRender = debounce(() => renderPurchase2aWorkspace(), 250);
  document.querySelectorAll(".purchase2a-box").forEach((box) => {
    const search = box.querySelector(".purchase2a-search");
    const rate = box.querySelector(".purchase2a-rate-filter");
    const key = box.dataset.purchaseSection;
    const sectionPredicate = (row) => key === "amendment"
      ? ["amendment_increase", "amendment_decrease"].includes(kind(row))
      : kind(row) === key;
    const applyFilter = () => {
      purchase2aFilters[key] = {
        q: (search.value || "").trim().toLowerCase(),
        rate: rate.value || "",
      };
      purchase2aPages[key] = 1;
      debouncedPurchase2aRender();
    };
    search.addEventListener("input", applyFilter);
    rate.addEventListener("change", () => {
      purchase2aFilters[key] = {
        q: (search.value || "").trim().toLowerCase(),
        rate: rate.value || "",
      };
      purchase2aPages[key] = 1;
      renderPurchase2aWorkspace();
    });
    box.querySelector(".purchase2a-clear-filter").addEventListener("click", () => {
      search.value = "";
      rate.value = "";
      purchase2aFilters[key] = { q: "", rate: "" };
      purchase2aPages[key] = 1;
      renderPurchase2aWorkspace();
    });
    // Same as GSTR-1: select/apply against the displayed dataset rows (GSTR-2A).
    box.querySelectorAll(".purchase2a-select").forEach((input) => input.addEventListener("change", () => {
      const row = source[Number(input.dataset.index)];
      if (!row) return;
      row.selected = input.checked;
      row.ready_for_tally = true;
      row.ready_for_purchase_note = true;
      box.querySelectorAll(`.purchase2a-select[data-index="${input.dataset.index}"]`).forEach((peer) => { peer.checked = input.checked; });
    }));
    box.querySelectorAll(".purchase2a-item").forEach((input) => {
      input.addEventListener("focus", () => {
        const row = source[Number(input.dataset.index)];
        if (row) showItemSuggestions(row);
      });
      input.addEventListener("change", () => {
        const row = source[Number(input.dataset.index)];
        if (!row) return;
        setRowItemName(Number(input.dataset.index), input.value.trim(), row);
        persistPurchaseItemMappings();
        renderPurchase2aWorkspace();
      });
    });
    box.querySelector(".purchase2a-select-all").addEventListener("click", (event) => {
      const filter = purchase2aFilters[key] || { q: "", rate: "" };
      const targets = source.filter((row, index) => {
        if (!sectionPredicate(row)) return false;
        const items = Array.isArray(row.items) && row.items.length ? row.items : [{ gst_rate: purchaseGstRate(row), hsn_code: row.hsn_code }];
        return items.some((item) => {
          const rateVal = String(Number(item.gst_rate ?? purchaseGstRate(item) ?? 0));
          if (filter.rate && rateVal !== filter.rate) return false;
          if (filter.q) {
            const blob = [row.gstin, row.party_name, row.invoice_no, row.invoice_date, rateVal, row.document_type, item.hsn_code, row.hsn_code]
              .join(" ")
              .toLowerCase();
            if (!blob.includes(filter.q)) return false;
          }
          return true;
        });
      });
      const select = targets.some((row) => !row.selected);
      targets.forEach((row) => {
        row.selected = select;
        row.ready_for_tally = true;
        row.ready_for_purchase_note = true;
      });
      event.currentTarget.textContent = select ? "Unselect All" : "Select All";
      renderPurchase2aWorkspace();
    });
    const bulkItem = box.querySelector(".purchase2a-bulk-item");
    if (bulkItem) {
      bulkItem.addEventListener("focus", () => {
        const selected = source.filter((row) => row.selected && sectionPredicate(row));
        if (selected.length) showItemSuggestions(selected[0]);
      });
    }
    box.querySelector(".purchase2a-apply-item").addEventListener("click", () => {
      applyItemToSelectedRows({
        rows: source,
        itemInputEl: bulkItem,
        predicate: sectionPredicate,
        refresh: renderPurchase2aWorkspace,
        emptySelectMessage: "Select at least one invoice in this box.",
      });
    });
    box.querySelector(".purchase2a-send").addEventListener("click", (event) => {
      // Send uses gstRows; keep them aligned with the active GSTR-2A dataset.
      gstRows = source;
      if (key === "b2b") return sendSelectedPurchaseRows(sectionPredicate, event.currentTarget, "Send to Tally");
      return sendSelectedPurchaseNotes(event.currentTarget, "Send to Tally", sectionPredicate);
    });
  });
  $("#purchase2aEmpty").classList.add("hidden");
  $("#purchase2aContent").classList.remove("hidden");
}

function renderPurchase2bWorkspace() {
  const source = gstDatasets["GSTR-2B"] || [];
  const workspace = $("#purchase2bWorkspace");
  const summaryPanel = $("#gstr2SummaryPanel");
  const sheetRows = $("#purchase2bSheetRows");
  if (workspace && summaryPanel && sheetRows && summaryPanel.parentElement !== workspace) {
    workspace.insertBefore(summaryPanel, sheetRows);
  }
  if ($("#purchase2bSheetCount")) $("#purchase2bSheetCount").textContent = source.length.toLocaleString("en-IN");
  if (sheetRows) {
    sheetRows.innerHTML = purchaseTableWithTotals(
      source.map((row, index) => ({ row, index })), false, false
    );
  }
  if (source.length) {
    renderGstr2Summary([], [], [], []);
  } else if (summaryPanel) {
    summaryPanel.classList.add("hidden");
    if ($("#gstr2SummaryRows")) $("#gstr2SummaryRows").innerHTML = "";
  }
}

function placeGstr2bSummaryForMatch() {
  const summary2a = $("#purchaseGstr2aSummaryPanel");
  const summary2b = $("#gstr2SummaryPanel");
  if (!summary2a || !summary2b || !summary2a.parentElement) return;
  summary2a.parentElement.insertBefore(summary2b, summary2a.nextSibling);
}

function setPurchaseSheetView(view = "match") {
  const show2a = view === "2a", show2b = view === "2b", showMatch = view === "match";
  const placeholder = $("#purchase2aWorkspace")?.parentElement;
  if (placeholder) placeholder.classList.toggle("hidden", !show2a);
  if ($("#purchase2aWorkspace")) $("#purchase2aWorkspace").classList.toggle("hidden", !show2a);
  if ($("#purchase2bWorkspace")) $("#purchase2bWorkspace").classList.toggle("hidden", !show2b);
  if (showMatch) placeGstr2bSummaryForMatch();
  const has2b = Boolean((gstDatasets["GSTR-2B"] || []).length);
  if ($("#gstr2SummaryPanel")) $("#gstr2SummaryPanel").classList.toggle("hidden", !(has2b && (show2b || showMatch)));
  ["purchase2a2bDashboard", "purchasePeriodFilters", "purchaseGstr2aSummaryPanel", "gstPurchaseResultBoxes"].forEach(id => {
    const element = $(`#${id}`);
    if (element) element.classList.toggle("hidden", !showMatch || !gstRows.length);
  });
  [["purchaseView2aBtn",show2a],["purchaseView2bBtn",show2b],["purchaseViewMatchBtn",showMatch]].forEach(([id,on]) => {
    $(`#${id}`)?.classList.toggle("active", on);
  });
  if ($("#gstr2aLoadStatus")) $("#gstr2aLoadStatus").classList.toggle("hidden", show2b);
  if ($("#gstr2bLoadStatus")) $("#gstr2bLoadStatus").classList.toggle("hidden", show2a);
  if ($("#gstClear2aBtn")) $("#gstClear2aBtn").classList.toggle("hidden", show2b);
  if ($("#gstClear2bBtn")) $("#gstClear2bBtn").classList.toggle("hidden", show2a);
  if ($("#gstClear2aBtn")) $("#gstClear2aBtn").disabled = false;
  if ($("#gstClear2bBtn")) $("#gstClear2bBtn").disabled = false;
  if ($("#gstReconcileBtn")) $("#gstReconcileBtn").classList.toggle("hidden", !showMatch);
  const tolerance = $("#gstTolerance")?.closest("label");
  if (tolerance) tolerance.classList.toggle("hidden", !showMatch);
  if (show2a) renderPurchase2aWorkspace();
  if (show2b) renderPurchase2bWorkspace();
  if (showMatch && has2b) renderGstr2Summary([], [], [], []);
}

if ($("#purchaseView2aBtn")) $("#purchaseView2aBtn").onclick = () => setPurchaseSheetView("2a");
if ($("#purchaseView2bBtn")) $("#purchaseView2bBtn").onclick = () => setPurchaseSheetView("2b");
if ($("#purchaseViewMatchBtn")) $("#purchaseViewMatchBtn").onclick = () => setPurchaseSheetView("match");
function renderGstr2Summary(matched,only2a,only2b,notes) {
  const money=v=>Number(v||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2});
  const sourceAll=gstDatasets["GSTR-2B"]||[];
  const portal=purchaseGstr2Breakdown(sourceAll);
  const eligible=purchaseGstr2Breakdown(sourceAll.filter(purchaseAllowedForTally));
  const rowOf=(label,t,cls="")=>{
    const totalItc=t.itc!=null?Number(t.itc):(Number(t.igst||0)+Number(t.cgst||0)+Number(t.sgst||0)+Number(t.cess||0));
    return `<tr class="${cls}"><th>${label}</th><td>${Math.max(0,Number(t.invoices||0)).toLocaleString("en-IN")}</td><td>${money(t.taxable_value)}</td><td>${money(t.igst)}</td><td>${money(t.cgst)}</td><td>${money(t.sgst)}</td><td>${money(t.cess)}</td><td class="gstr2-total-itc">${money(totalItc)}</td></tr>`;
  };
  const amendment = {
    invoices: Number(portal.increase.invoices||0)+Number(portal.decrease.invoices||0),
    taxable_value: Number(portal.increase.taxable_value||0)-Number(portal.decrease.taxable_value||0),
    igst: Number(portal.increase.igst||0)-Number(portal.decrease.igst||0),
    cgst: Number(portal.increase.cgst||0)-Number(portal.decrease.cgst||0),
    sgst: Number(portal.increase.sgst||0)-Number(portal.decrease.sgst||0),
    cess: Number(portal.increase.cess||0)-Number(portal.decrease.cess||0),
    itc: Number(portal.summary?.amendment_itc||0),
  };
  $("#gstr2SummaryRows").innerHTML = [
    rowOf("Invoice Count / Gross Invoice ITC", portal.total),
    rowOf("(+) Debit Note ITC", portal.debit),
    rowOf("(−) Credit Note ITC", portal.credit),
    rowOf("(±) Amendment Adjustment", amendment),
    rowOf("Net GSTR-2B ITC", portal.net || portal.gross, "gstr2-gross-row"),
    rowOf("Tally Eligible Net ITC", eligible.net || eligible.gross, "gstr2-gross-row"),
  ].join("");
  $("#gstr2SummaryPanel").classList.remove("hidden");
  renderPurchaseGstr2aSummary();
  if (gstDatasets["GSTR-3B"]) renderPurchaseGstr3bComparison();
}
function purchaseNetGstr2bTotals() {
  const summary = buildGstr2bGrossNetSummary(gstDatasets["GSTR-2B"] || []);
  return {
    invoices: summary.invoice_count,
    taxable_value: summary.net_taxable,
    igst: summary.net_igst,
    cgst: summary.net_cgst,
    sgst: summary.net_sgst,
    cess: summary.net_cess,
    itc: summary.net_itc,
  };
}
function purchaseReturnBreakdown(sourceRows,label) {
  const portal = purchaseGstr2Breakdown(sourceRows || []);
  const net = portal.net || portal.gross || { invoices: 0, taxable_value: 0, igst: 0, cgst: 0, sgst: 0, cess: 0 };
  const amendment = {
    invoices: Number(portal.increase.invoices||0)+Number(portal.decrease.invoices||0),
    taxable_value: Number(portal.increase.taxable_value||0)-Number(portal.decrease.taxable_value||0),
    igst: Number(portal.increase.igst||0)-Number(portal.decrease.igst||0),
    cgst: Number(portal.increase.cgst||0)-Number(portal.decrease.cgst||0),
    sgst: Number(portal.increase.sgst||0)-Number(portal.decrease.sgst||0),
  };
  return [
    [`Total ${label}  +`, portal.total],
    ["Debit Note  +", portal.debit],
    ["Credit Note  −", portal.credit],
    ["Amendment Adjustment  ±", amendment],
    ["Net Total", net],
  ];
}
function renderPurchaseGstr2aSummary() {
  const source=gstDatasets["GSTR-2A"]||[];if(!source.length)return;
  const money=value=>Number(value||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2});
  $("#purchaseGstr2aSummaryRows").innerHTML=purchaseReturnBreakdown(source,"GSTR-2A").map(([label,t])=>`<tr class="${/Net Total/.test(label)?"gstr2-gross-row":""}"><th>${label}</th><td>${Math.max(0,t.invoices).toLocaleString("en-IN")}</td><td>${money(t.taxable_value)}</td><td>${money(t.igst)}</td><td>${money(t.cgst)}</td><td>${money(t.sgst)}</td></tr>`).join("");
  $("#purchaseGstr2aSummaryPanel").classList.remove("hidden");
}
function renderPurchaseGstr3bComparison() {
  const report=gstDatasets["GSTR-3B"]||{}, net=report.net_itc||report, rcm=report.reverse_charge||{}, gstr2=purchaseNetGstr2bTotals();
  const diff={taxable_value:0,igst:Number(report.igst||0)-Number(gstr2.igst||0),cgst:Number(report.cgst||0)-Number(gstr2.cgst||0),sgst:Number(report.sgst||0)-Number(gstr2.sgst||0),cess:Number(report.cess||0)};
  const money=value=>Number(value||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2});
  const row=(label,t,cls="")=>`<tr class="${cls}"><th>${label}</th><td>${money(t.taxable_value)}</td><td>${money(t.igst)}</td><td>${money(t.cgst)}</td><td>${money(t.sgst)}</td><td>${money(t.cess)}</td></tr>`;
  $("#purchaseGstr3bCompareRows").innerHTML=row("Net GSTR-2B ITC",gstr2)+row("GSTR-3B ITC (excluding reversal)",report)+row("GSTR-3B ITC (considering reversal)",net)+row("Difference: GSTR-3B − GSTR-2B",diff,"gstr3b-difference-row")+row("GSTR-3B Reverse Charge",rcm);
  const totalDifference=Number(diff.igst||0)+Number(diff.cgst||0)+Number(diff.sgst||0)+Number(diff.cess||0);
  $("#purchaseGstr3bAdvice").textContent=Math.abs(totalDifference)<=1?"GSTR-2B ITC and GSTR-3B claimed ITC match within ₹1.":`ITC difference ₹${money(totalDifference)}. Verify eligible ITC, Credit/Debit Notes, amendments and Reverse Charge before Tally import.`;
  $("#purchaseGstr3bComparePanel").classList.remove("hidden");
  renderItcDifferenceInvoices();
  renderPurchaseGstr2aSummary();
}
function normalizedPurchasePeriod(value) {
  const digits=String(value||"").replace(/\D/g,"");
  if(digits.length!==6)return digits;
  if(Number(digits.slice(0,2))<=12)return digits.slice(2)+digits.slice(0,2);
  return digits;
}
function purchaseDocumentKey(row) {
  const source=row?.gstr2b||row||{};
  return [normalizedPurchasePeriod(source.source_period||source.gstr2b_period||row?.gstr2b_period),String(source.gstin||"").replace(/\s+/g,"").toUpperCase(),String(source.invoice_no||"").replace(/[^A-Z0-9]/gi,"").toUpperCase(),String(source.document_type||"Invoice").toLowerCase()].join("|");
}
function purchaseAllowedForTally(row) { return !itcTallyExcluded.has(purchaseDocumentKey(row)); }
function purchaseSearchMatch(row,query) {
  const source=row?.gstr2b||row?.gstr2a||row||{},needle=String(query||"").trim().toLowerCase();if(!needle)return true;
  return [source.gstin,source.party_name,row?.party_ledger,source.invoice_no,source.invoice_date,source.original_invoice_date,source.tally_entry_date,source.document_type,source.hsn_code,source.hsn,source.item_name,row?.item_name,row?.expense_ledger,source.source_period,source.gstr2b_period,source.invoice_value,source.taxable_value,source.igst,source.cgst,source.sgst,source.status].some(value=>String(value??"").toLowerCase().includes(needle));
}
function itcDifferenceInvoices() {
  const source=(gstDatasets["GSTR-2B"]||[]),periods=normalizeGstr3bPeriods((gstDatasets["GSTR-3B"]||{}).periods||{}),tolerance=Number($("#gstTolerance")?.value||1);
  const result=[];
  Object.keys(periods).forEach(period=>{
    const documents=source.filter(row=>normalizedPurchasePeriod(row.source_period||row.gstr2b_period)===period);
    if(!documents.length)return;
    const regular=documents.filter(row=>!isPurchaseNote(row)),notes=documents.filter(isPurchaseNote);
    let net=purchaseTotals(regular);
    notes.forEach(note=>{const sign=["credit","amendment_decrease"].includes(purchaseNoteKind(note))?-1:1;net=purchaseSummaryAdd(net,purchaseTotals([note]),sign);});
    const claimed=periods[period]||{},heads=["igst","cgst","sgst"],short=heads.filter(head=>Number(net[head]||0)-Number(claimed[head]||0)>tolerance);
    if(!short.length)return;
    const reason=short.map(head=>`Monthly ${head.toUpperCase()} shortfall ₹${(Number(net[head]||0)-Number(claimed[head]||0)).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2})}`).join(" · ");
    documents.filter(row=>short.some(head=>Number(row[head]||0)>0)).forEach(row=>result.push({row,period,reason}));
  });
  return result;
}
function renderItcDifferenceInvoices() {
  const panel=$("#itcDifferenceInvoicePanel"),body=$("#itcDifferenceInvoiceRows");if(!panel||!body)return;
  const allEntries=itcDifferenceInvoices(),query=$("#itcDifferenceSearch")?.value||"",money=value=>Number(value||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2});
  let addedDefaultExclusion=false;
  allEntries.forEach(({row})=>{const key=purchaseDocumentKey(row);if(!itcDifferenceInitialized.has(key)){itcDifferenceInitialized.add(key);itcTallyExcluded.add(key);addedDefaultExclusion=true;}});
  const filterValue=(id)=>String($(id)?.value||"").trim().toLowerCase(),matches=(value,id)=>{const filter=filterValue(id);return !filter||String(value??"").toLowerCase().includes(filter);};
  const entries=allEntries.filter(({row,period,reason})=>{const key=purchaseDocumentKey(row),action=filterValue("#itcFilterAction"),excluded=itcTallyExcluded.has(key),periodText=period.length===6?period.slice(4)+"/"+period.slice(0,4):period;return purchaseSearchMatch(row,query)&&matches(periodText,"#itcFilterPeriod")&&matches(row.gstin,"#itcFilterGstin")&&matches(row.party_name,"#itcFilterParty")&&matches(row.invoice_no,"#itcFilterInvoice")&&matches(row.invoice_date||row.original_invoice_date,"#itcFilterDate")&&matches(row.document_type||"Invoice","#itcFilterType")&&matches(`${row.taxable_value||0} ${money(row.taxable_value)}`,"#itcFilterTaxable")&&matches(`${row.igst||0} ${money(row.igst)}`,"#itcFilterIgst")&&matches(`${row.cgst||0} ${money(row.cgst)}`,"#itcFilterCgst")&&matches(`${row.sgst||0} ${money(row.sgst)}`,"#itcFilterSgst")&&matches(reason,"#itcFilterReason")&&(!action||(action==="exclude"?excluded:!excluded));});
  const totals=entries.reduce((sum,{row})=>{sum.taxable+=Number(row.taxable_value||0);sum.igst+=Number(row.igst||0);sum.cgst+=Number(row.cgst||0);sum.sgst+=Number(row.sgst||0);return sum;},{taxable:0,igst:0,cgst:0,sgst:0});
  $("#itcDifferenceTaxable").textContent=money(totals.taxable);$("#itcDifferenceIgst").textContent=money(totals.igst);$("#itcDifferenceCgst").textContent=money(totals.cgst);$("#itcDifferenceSgst").textContent=money(totals.sgst);
  const filterIds=["#itcDifferenceSearch","#itcFilterPeriod","#itcFilterGstin","#itcFilterParty","#itcFilterInvoice","#itcFilterDate","#itcFilterType","#itcFilterTaxable","#itcFilterIgst","#itcFilterCgst","#itcFilterSgst","#itcFilterReason","#itcFilterAction"],filterActive=filterIds.some(id=>filterValue(id));
  $("#itcDifferenceInvoiceCount").textContent=filterActive?`${entries.length.toLocaleString("en-IN")} / ${allEntries.length.toLocaleString("en-IN")}`:entries.length.toLocaleString("en-IN");
  const monthGroups=new Map();entries.forEach(entry=>{if(!monthGroups.has(entry.period))monthGroups.set(entry.period,[]);monthGroups.get(entry.period).push(entry);});
  body.innerHTML=[...monthGroups.entries()].sort(([a],[b])=>a.localeCompare(b)).map(([period,monthEntries])=>{const open=itcDifferenceOpenMonths.has(period),monthTotal=monthEntries.reduce((sum,{row})=>{sum.taxable+=Number(row.taxable_value||0);sum.igst+=Number(row.igst||0);sum.cgst+=Number(row.cgst||0);sum.sgst+=Number(row.sgst||0);return sum;},{taxable:0,igst:0,cgst:0,sgst:0}),periodText=period.length===6?period.slice(4)+"/"+period.slice(0,4):period,header=`<tr class="itc-month-row"><td colspan="12"><button type="button" class="itc-month-toggle" data-period="${escapeHtml(period)}"><span class="itc-month-arrow">${open?"▼":"▶"}</span><strong>${escapeHtml(periodText)}</strong><span>${monthEntries.length.toLocaleString("en-IN")} invoice(s)</span><span>Taxable ₹${money(monthTotal.taxable)}</span><span>IGST ₹${money(monthTotal.igst)}</span><span>CGST ₹${money(monthTotal.cgst)}</span><span>SGST ₹${money(monthTotal.sgst)}</span></button></td></tr>`;if(!open)return header;return header+monthEntries.map(({row,reason})=>{const key=purchaseDocumentKey(row),excluded=itcTallyExcluded.has(key);return `<tr class="${excluded?"itc-row-excluded":"itc-row-included"}"><td>${escapeHtml(periodText)}</td><td>${escapeHtml(row.gstin||"")}</td><td>${escapeHtml(row.party_name||"")}</td><td>${escapeHtml(row.invoice_no||"")}</td><td>${escapeHtml(row.invoice_date||row.original_invoice_date||"")}</td><td>${escapeHtml(row.document_type||"Invoice")}</td><td class="money">${money(row.taxable_value)}</td><td class="money">${money(row.igst)}</td><td class="money">${money(row.cgst)}</td><td class="money">${money(row.sgst)}</td><td class="itc-reason">${escapeHtml(reason)}</td><td><select class="itc-tally-action" data-key="${escapeHtml(key)}"><option value="exclude" ${excluded?"selected":""}>Do Not Send to Tally</option><option value="include" ${excluded?"":"selected"}>Include in Tally</option></select></td></tr>`;}).join("");}).join("")||`<tr><td colspan="12" class="empty-state">No invoice-level ITC difference detected for the loaded periods.</td></tr>`;
  // Keep this technical detail box off the Purchase screen. Its exclusions
  // still feed the Problematic / Skipped list in the Tally review popup.
  panel.classList.add("hidden");
  document.querySelectorAll(".itc-month-toggle").forEach(button=>button.onclick=()=>{const period=button.dataset.period;if(itcDifferenceOpenMonths.has(period))itcDifferenceOpenMonths.delete(period);else itcDifferenceOpenMonths.add(period);renderItcDifferenceInvoices();});
  document.querySelectorAll(".itc-tally-action").forEach(select=>select.onchange=()=>{if(select.value==="exclude")itcTallyExcluded.add(select.dataset.key);else itcTallyExcluded.delete(select.dataset.key);renderGstr2Summary([],[],[],[]);});
  if(addedDefaultExclusion)setTimeout(()=>renderGstr2Summary([],[],[],[]),0);
}
function renderPurchaseReconciliation() {
  const indexed=gstRows.map((row,index)=>({row,index})).filter(({row}) => purchasePeriodFilterMatch(row) && purchaseDashFilterMatch(row));
  const notes=indexed.filter(({row})=>isPurchaseNote(row));
  const matched=indexed.filter(({row})=>row.gstr2a&&row.gstr2b&&!isPurchaseNote(row)&&row.category!=="mismatch"&&!String(row.status||"").includes("Mismatch"));
  const mismatch=indexed.filter(({row})=>!isPurchaseNote(row)&&(row.category==="mismatch"||String(row.status||"").includes("Mismatch")));
  const only2a=indexed.filter(({row})=>row.gstr2a&&!row.gstr2b&&!isPurchaseNote(row)).map(({row,index})=>({row:{...(row.gstr2a||row),status:row.status,category:row.category,selected:row.selected,ready_for_tally:row.ready_for_tally,tally_entry_date:row.tally_entry_date,original_invoice_date:row.original_invoice_date,itc_status:row.itc_status,tally_status:row.tally_status,party_ledger:row.party_ledger,expense_ledger:row.expense_ledger,sales_allocations:row.sales_allocations,available_in_gstr2b:false,available_in_gstr2a:true,gstr2b_period:row.gstr2b_period||""},index}));
  const only2b=indexed.filter(({row})=>row.gstr2b&&!row.gstr2a&&!isPurchaseNote(row)).map(({row,index})=>({row:{...(row.gstr2b||row),status:row.status,category:row.category,selected:row.selected,ready_for_tally:row.ready_for_tally,tally_entry_date:row.tally_entry_date,original_invoice_date:row.original_invoice_date,itc_status:row.itc_status,review_required:true,party_ledger:row.party_ledger,expense_ledger:row.expense_ledger,sales_allocations:row.sales_allocations,gstr2b_period:row.gstr2b_period||row.source_period||""},index}));
  const filteredMatched=matched.filter(({row})=>purchaseSearchMatch(row,$("#gstMatchedSearch")?.value)),filteredOnly2a=only2a.filter(({row})=>purchaseSearchMatch(row,$("#gstOnly2aSearch")?.value)),filteredOnly2b=only2b.filter(({row})=>purchaseSearchMatch(row,$("#gstOnly2bSearch")?.value)),filteredNotes=notes.filter(({row})=>purchaseSearchMatch(row,$("#gstNotesSearch")?.value)),filteredMismatch=mismatch.filter(({row})=>purchaseSearchMatch(row,$("#gstMismatchSearch")?.value));
  $("#gstMatchedTitle").textContent="2A + 2B Matched";
  $("#gstMatchedRows").innerHTML=purchaseTableWithTotals(filteredMatched,true,false,"matched");
  $("#gstOnly2aRows").innerHTML=purchaseTableWithTotals(filteredOnly2a,true,false,"only2a");
  $("#gstOnly2bRows").innerHTML=purchaseTableWithTotals(filteredOnly2b,true,true,"only2b");
  $("#gstPurchaseNoteRows").innerHTML=purchaseTableWithTotals(filteredNotes,true,false,"notes");
  if ($("#gstMismatchRows")) {
    const mismatchPageSize = PURCHASE_RECON_PAGE_SIZE;
    const mismatchPaged = paginateEntries(filteredMismatch, purchaseReconPages.mismatch || 1, mismatchPageSize);
    purchaseReconPages.mismatch = mismatchPaged.page;
    $("#gstMismatchRows").innerHTML = mismatchPaged.total
      ? mismatchPaged.pageEntries.map(({row,index}) => {
          const diff = row.differences || {};
          const money = (v) => Number(v || 0).toLocaleString("en-IN", {minimumFractionDigits: 2, maximumFractionDigits: 2});
          const a = row.gstr2a || {};
          const b = row.gstr2b || {};
          return `<div class="purchase-mismatch-block"><label><input class="purchase-match-select" data-index="${index}" type="checkbox" ${row.selected?"checked":""}> Select</label> <strong>${escapeHtml(row.invoice_no || "")}</strong> · ${escapeHtml(row.gstin || "")} · ${escapeHtml(row.status || "")}
            <div class="purchase-box-actions"><input class="purchase-party-ledger" data-index="${index}" list="tallyLedgerList" value="${escapeHtml(row.party_ledger||row.party_name||"")}" placeholder="Party / Tally Ledger"><input class="purchase-item-name" data-index="${index}" list="tallyItemList" value="${escapeHtml(salesItemLabel(row)||"")}" placeholder="Item / Tally Stock Item"></div>
            <table><thead><tr><th>Particular</th><th>GSTR-2B</th><th>GSTR-2A</th><th>Tally</th><th>Difference</th></tr></thead><tbody>
            ${["taxable_value","igst","cgst","sgst","cess","invoice_value"].map((field) => {
              const tallyVal = (row.tally || {})[field];
              return `<tr><td>${field}</td><td class="money">${money(b[field])}</td><td class="money">${money(a[field])}</td><td class="money">${tallyVal == null || tallyVal === "" ? "—" : money(tallyVal)}</td><td class="money">${money(diff[field])}</td></tr>`;
            }).join("")}
            </tbody></table></div>`;
        }).join("") + (mismatchPaged.total > mismatchPageSize
          ? purchasePagerHtml("mismatch", mismatchPaged.page, mismatchPaged.pages, mismatchPaged.total)
          : "")
      : `<div class="purchase-empty">No mismatches.</div>`;
  }
  [["gstMatchedCount",filteredMatched,matched],["gstOnly2aCount",filteredOnly2a,only2a],["gstOnly2bCount",filteredOnly2b,only2b],["gstPurchaseNoteCount",filteredNotes,notes]].forEach(([id,shown,all])=>{if($(`#${id}`))$(`#${id}`).textContent=shown.length===all.length?all.length.toLocaleString("en-IN"):`${shown.length.toLocaleString("en-IN")} / ${all.length.toLocaleString("en-IN")}`;});
  if ($("#gstMismatchCount")) $("#gstMismatchCount").textContent = mismatch.length.toLocaleString("en-IN");
  [["gstSelectAllMatched",matched],["gstSelectAllOnly2a",only2a],["gstSelectAllOnly2b",only2b],["gstSelectAllMismatch",mismatch],["gstSelectAllNotes",notes]].forEach(([id,rs])=>{const button=$(`#${id}`);if(button)button.textContent=rs.length&&rs.every(({row})=>row.selected)?"Unselect All":"Select All";});
  renderGstr2Summary(matched.map(x=>x.row),only2a.map(x=>x.row),only2b.map(x=>x.row),notes.map(x=>x.row));
  renderPurchase2a2bDashboard();
  $("#gstPurchaseResultBoxes").classList.remove("hidden");
  $("#gstResults").classList.add("hidden");
  document.querySelectorAll(".purchase-match-select").forEach(input=>input.onchange=()=>gstRows[Number(input.dataset.index)].selected=input.checked);
  document.querySelectorAll(".purchase-party-ledger").forEach(input=>input.oninput=()=>gstRows[Number(input.dataset.index)].party_ledger=input.value.trim());
  document.querySelectorAll(".purchase-item-name").forEach(input=>{
    input.onfocus = () => {
      const row = gstRows[Number(input.dataset.index)];
      if (row) showItemSuggestions(row);
    };
    input.onchange = () => {
      setRowItemName(Number(input.dataset.index), input.value.trim());
      persistPurchaseItemMappings();
      renderPurchaseReconciliation();
    };
  });
  document.querySelectorAll(".purchase-row-edit").forEach(input=>input.oninput=()=>{
    const row=gstRows[Number(input.dataset.index)],field=input.dataset.field;
    row[field]=input.type==="number"?Number(input.value||0):input.value.trim();
    if(field==="original_invoice_date")row.invoice_date=row[field];
  });
  bindPurchaseReconPagers();
}

let purchase2a2bDashboard = null;
let purchaseDashFilter = "";

function purchaseMonthFromDate(value) {
  const text = String(value || "").trim();
  let match = text.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
  if (match) return match[2].padStart(2, "0");
  match = text.match(/^(\d{4})[-\/](\d{2})/);
  if (match) return match[2];
  match = text.match(/^(\d{2})(\d{4})$/);
  if (match) return match[1];
  return "";
}

function purchaseMonthFromPeriod(value) {
  const digits = String(value || "").replace(/\D/g, "");
  if (digits.length === 6) return digits.slice(0, 2);
  return purchaseMonthFromDate(value);
}

function purchasePeriodFilterMatch(row) {
  const origNeed = $("#purchaseOrigMonthFilter")?.value || "";
  const periodNeed = $("#purchase2bMonthFilter")?.value || "";
  if (origNeed) {
    const orig = purchaseMonthFromDate(row.original_invoice_date || row.invoice_date);
    if (orig !== origNeed) return false;
  }
  if (periodNeed) {
    const period = purchaseMonthFromPeriod(row.gstr2b_period || row.source_period);
    // 2A-only rows have no 2B period — exclude when filtering by 2B month.
    if (!period || period !== periodNeed) return false;
  }
  return true;
}

function purchaseDashFilterMatch(row) {
  const key = purchaseDashFilter;
  if (!key) return true;
  if (key === "matched_2a_2b") return ["matched_2a_2b", "matched_2a_2b_tally"].includes(row.category);
  if (key === "matched_missing_tally") return row.category === "matched_2a_2b" && !row.purchase_booked;
  if (key === "only_2b_review") return row.category === "only_2b_review";
  if (key === "only_2a") return row.category === "only_2a" || (row.available_in_gstr2a && !row.available_in_gstr2b);
  if (key === "mismatch") return row.category === "mismatch";
  if (key === "ready_to_send") return Boolean(row.ready_for_tally);
  if (key === "already_in_tally") return Boolean(row.purchase_booked);
  if (key === "sent_to_tally") {
    const status = String(row.tally_status || "").toLowerCase();
    return status === "sent to tally" || status === "already in tally";
  }
  return true;
}

function renderPurchase2a2bDashboard(dashboard) {
  if (dashboard) purchase2a2bDashboard = dashboard;
  const dash = purchase2a2bDashboard || {};
  const set = (id, key) => {
    const pack = dash[key] || {};
    const node = $(`#${id}`);
    if (!node) return;
    const count = Number(pack.count || 0);
    const amount = Number(pack.taxable_value || 0);
    node.innerHTML = amount
      ? `${count.toLocaleString("en-IN")}<small>₹${amount.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</small>`
      : count.toLocaleString("en-IN");
  };
  set("purchaseCardMatched", "matched_2a_2b");
  set("purchaseCardMissingTally", "matched_missing_tally");
  set("purchaseCardOnly2b", "only_2b_review");
  set("purchaseCardOnly2a", "only_2a");
  set("purchaseCardMismatch", "mismatch");
  set("purchaseCardReady", "ready_to_send");
  set("purchaseCardInTally", "already_in_tally");
  set("purchaseCardSent", "sent_to_tally");
  if ($("#purchase2a2bDashboard")) $("#purchase2a2bDashboard").classList.toggle("hidden", !gstRows.length);
  if ($("#purchasePeriodFilters")) $("#purchasePeriodFilters").classList.toggle("hidden", !gstRows.length);
  document.querySelectorAll(".purchase-dash-card").forEach((button) => {
    button.classList.toggle("active", button.dataset.purchaseFilter === purchaseDashFilter);
  });
  const note = $("#purchaseDashFilterNote");
  if (note) {
    if (purchaseDashFilter) {
      note.textContent = `Showing filter: ${purchaseDashFilter.replace(/_/g, " ")}. Click the same card again to clear.`;
      note.classList.remove("hidden");
    } else {
      note.classList.add("hidden");
    }
  }
}

function updatePurchaseImportStatus() {
  const a = gstDatasets["GSTR-2A"] || [];
  const b = gstDatasets["GSTR-2B"] || [];
  const chip = (el, ok, label) => {
    if (!el) return;
    el.textContent = label;
    el.classList.toggle("ok", !!ok);
    el.classList.toggle("warn", !ok);
  };
  chip($("#gstr2aLoadStatus"), a.length, a.length ? `GSTR-2A: Imported (${a.length.toLocaleString("en-IN")})` : "GSTR-2A: Not Imported");
  chip($("#gstr2bLoadStatus"), b.length, b.length ? `GSTR-2B: Imported (${b.length.toLocaleString("en-IN")})` : "GSTR-2B: Not Imported");
  const tallyOk = Boolean(tallyMasters.connected || (gstReconTallyRows || []).length);
  chip($("#purchaseTallyStatus"), tallyOk, tallyOk ? "Tally: Connected" : "Tally: Not Connected");
  if ($("#gstReconcileBtn")) $("#gstReconcileBtn").disabled = !(a.length && b.length);
}

async function clearPurchasePortalReturn(returnType) {
  if (!confirm(`Clear imported ${returnType} for the current company/FY? Tally data and the other return stay unchanged.`)) return;
  const error = $("#gstError");
  const button = document.querySelector(`[data-purchase-clear="${returnType}"]`);
  const sourceRows = gstDatasets[returnType] || [];
  const scopeGstin = getGstPortalGstin({
    gstr2a: returnType === "GSTR-2A" ? sourceRows : (gstDatasets["GSTR-2A"] || []),
    gstr2b: returnType === "GSTR-2B" ? sourceRows : (gstDatasets["GSTR-2B"] || []),
  });
  const scopeFinancialYear = String(
    sourceRows.find(row => row.financial_year)?.financial_year
    || gstPortalContext.financial_year
    || "2025-26"
  ).trim();
  try {
    if (button) {
      button.disabled = true;
      button.dataset.previousText = button.textContent;
      button.textContent = `Clearing ${returnType}...`;
    }
    const response = await fetch("/api/gst/recon/clear", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        returnType,
        gstin: scopeGstin,
        financialYear: scopeFinancialYear,
      }),
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(result.error || `Could not clear ${returnType}.`);
    if (returnType === "GSTR-2A") {
      gstDatasets["GSTR-2A"] = [];
      gstReconDatasetsLoaded.delete("gstr2a");
      renderPurchase2aWorkspace();
    }
    if (returnType === "GSTR-2B") {
      gstDatasets["GSTR-2B"] = [];
      gstReconRows = [];
      gstReconDatasetsLoaded.delete("gstr2b");
    }
    if (result.gstin) gstPortalContext.gstin = String(result.gstin).toUpperCase();
    if (result.financial_year) gstPortalContext.financial_year = result.financial_year;
    gstRows = [];
    purchase2a2bDashboard = null;
    purchaseDashFilter = "";
    if ($("#gstPurchaseResultBoxes")) $("#gstPurchaseResultBoxes").classList.add("hidden");
    if ($("#purchase2a2bDashboard")) $("#purchase2a2bDashboard").classList.add("hidden");
    if ($("#purchasePeriodFilters")) $("#purchasePeriodFilters").classList.add("hidden");
    if ($("#gstr2SummaryPanel")) $("#gstr2SummaryPanel").classList.add("hidden");
    if ($("#purchase2aWorkspace")) {
      $("#purchase2aWorkspace").classList.remove("hidden");
      $("#purchase2aWorkspace").parentElement?.classList.remove("hidden");
    }
    updatePurchaseImportStatus();
    setPurchaseSheetView(returnType === "GSTR-2A" ? "2a" : "2b");
  } catch (failure) {
    if (error) {
      error.textContent = failure.message || String(failure);
      error.classList.remove("hidden");
    }
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = button.dataset.previousText || `Clear ${returnType}`;
    }
  }
}
function renderStandalonePurchaseRegister(label) {
  gstRows.forEach(row=>{row.status=`${label} Ready`;row.ready_for_tally=!isPurchaseNote(row);row.ready_for_purchase_note=isPurchaseNote(row);row.selected=false;row.party_ledger=purchaseLedgerMatch(row);});
  const invoices=gstRows.map((row,index)=>({row,index})).filter(({row})=>!isPurchaseNote(row)),notes=gstRows.map((row,index)=>({row,index})).filter(({row})=>isPurchaseNote(row));
  $("#gstMatchedTitle").textContent=`${label} – Ready Purchase Invoices`;$("#gstMatchedRows").innerHTML=purchaseTableWithTotals(invoices,true);$("#gstOnly2aRows").innerHTML=purchaseTableWithTotals([]);$("#gstOnly2bRows").innerHTML=purchaseTableWithTotals([]);$("#gstPurchaseNoteRows").innerHTML=purchaseTableWithTotals(notes,true);
  $("#gstMatchedCount").textContent=invoices.length.toLocaleString("en-IN");$("#gstOnly2aCount").textContent="0";$("#gstOnly2bCount").textContent="0";$("#gstPurchaseNoteCount").textContent=notes.length.toLocaleString("en-IN");
  renderGstr2Summary(invoices.map(x=>x.row),[],[],notes.map(x=>x.row));$("#gstPurchaseResultBoxes").classList.remove("hidden");$("#gstTallyPanel").classList.remove("hidden");
  $("#gstResults").classList.add("hidden");
  document.querySelectorAll(".purchase-match-select").forEach(input=>input.onchange=()=>gstRows[Number(input.dataset.index)].selected=input.checked);
  document.querySelectorAll(".purchase-party-ledger").forEach(input=>input.oninput=()=>gstRows[Number(input.dataset.index)].party_ledger=input.value.trim());
  document.querySelectorAll(".purchase-item-name").forEach(input=>{
    input.onfocus = () => { const row = gstRows[Number(input.dataset.index)]; if (row) showItemSuggestions(row); };
    input.onchange = () => { setRowItemName(Number(input.dataset.index), input.value.trim()); persistPurchaseItemMappings(); };
  });
}
function resetGstReconWorkspace() {
  gstReconRows = [];
  gstReconTallyRows = [];
  gstReconResults = [];
  gstReconGstr1Rows = [];
  gstReconTallySalesRows = [];
  gstReconSalesResults = [];
  salesReconDashboard = null;
  gstr3bDashboard = null;
  salesReconPage = 1;
  activeReconTab = "overview";
  if ($("#recon2bStatus")) $("#recon2bStatus").textContent = "GSTR-2B: Not loaded";
  if ($("#recon3bStatus")) $("#recon3bStatus").textContent = "GSTR-3B: Not loaded";
  if ($("#reconTallyStatus")) $("#reconTallyStatus").textContent = "Tally Purchase: Not synced";
  if ($("#reconGstr3bPeriodStatus")) $("#reconGstr3bPeriodStatus").textContent = "GSTR-3B: Not loaded";
  if ($("#recon2bTallyBtn")) $("#recon2bTallyBtn").disabled = true;
  if ($("#reconItcBtn")) $("#reconItcBtn").disabled = true;
  ["recon2bSummaryPanel", "reconItcDashboardPanel", "reconVoucherReportPlaceholder", "gstr3bComparePanel", "gstr3bLiabilityPanel", "gstr3bItcPanel", "gstr3bUtilPanel"].forEach(id => {
    const node = document.getElementById(id);
    if (node) node.classList.add("hidden");
  });
  if ($("#reconVoucherMatchedCount")) $("#reconVoucherMatchedCount").textContent = "0";
  if ($("#reconVoucherMismatchCount")) $("#reconVoucherMismatchCount").textContent = "0";
  if ($("#reconItcRows")) $("#reconItcRows").innerHTML = "";
  if ($("#reconItcAdvice")) $("#reconItcAdvice").textContent = "";
  if ($("#gstr3bCompareRows")) $("#gstr3bCompareRows").innerHTML = "";
  if ($("#gstr3bLiabilityRows")) $("#gstr3bLiabilityRows").innerHTML = "";
  if ($("#gstr3bItcRows")) $("#gstr3bItcRows").innerHTML = "";
  if ($("#gstr3bUtilRows")) $("#gstr3bUtilRows").innerHTML = "";
  if ($("#gstr3bReconCards")) $("#gstr3bReconCards").classList.add("hidden");
  if ($("#gstr3bReconAdvice")) $("#gstr3bReconAdvice").textContent = "";
  renderSalesReconStatus();
  updateSalesReconReady();
  if ($("#salesReconRows")) $("#salesReconRows").innerHTML = "";
  if ($("#salesReconCards")) $("#salesReconCards").classList.add("hidden");
  if ($("#salesFyPeriodPanel")) $("#salesFyPeriodPanel").classList.add("hidden");
  if ($("#salesFyPeriodRows")) $("#salesFyPeriodRows").innerHTML = "";
  if ($("#salesFyPeriodTotals")) $("#salesFyPeriodTotals").innerHTML = "";
  if ($("#salesOutputSummaryPanel")) $("#salesOutputSummaryPanel").classList.add("hidden");
  if ($("#salesDocTypePanel")) $("#salesDocTypePanel").classList.add("hidden");
  if ($("#salesDocTypeRows")) $("#salesDocTypeRows").innerHTML = "";
  // UI-only tab switch during reset — do not kick off async dataset/overview fetches.
  setReconTab("overview", { refresh: false });
  ensureGstReconPanelVisible();
}

function renderGstReconStatus() {
  const g2b = gstReconRows.length || gstDatasets["GSTR-2B"]?.length || Number(gstReconDatasetCounts.gstr2b || 0);
  const g3b = gstDatasets["GSTR-3B"] && Object.keys(gstDatasets["GSTR-3B"]).length ? "Loaded" : "Not loaded";
  const tallyCount = gstReconTallyRows.length || Number(gstReconDatasetCounts.tally_purchase || 0);
  if ($("#recon2bStatus")) {
    $("#recon2bStatus").textContent = g2b ? `GSTR-2B: ${Number(g2b).toLocaleString("en-IN")} rows loaded` : "GSTR-2B: Not loaded";
  }
  if ($("#recon3bStatus")) {
    $("#recon3bStatus").textContent = g3b === "Loaded" ? "GSTR-3B: Loaded" : "GSTR-3B: Not loaded";
  }
  if ($("#reconTallyStatus")) {
    $("#reconTallyStatus").textContent = tallyCount
      ? `Tally Purchase: ${Number(tallyCount).toLocaleString("en-IN")} vouchers synced`
      : "Tally Purchase: Not synced";
  }
  updateGstReconReady();
}

function updateGstReconReady() {
  const ready = Boolean(gstReconRows.length && gstReconTallyRows.length);
  if ($("#recon2bTallyBtn")) $("#recon2bTallyBtn").disabled = !ready;
  if ($("#reconItcBtn")) $("#reconItcBtn").disabled = !gstReconRows.length;
}

function reconMoney(value) {
  return Number(value || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function renderGstRecon2bSummary(rows) {
  const summary = buildGstr2bGrossNetSummary(rows || []);
  const money = reconMoney;
  if ($("#recon2bCount")) $("#recon2bCount").textContent = money(summary.net_itc);
  if ($("#recon2bInvoiceCount")) $("#recon2bInvoiceCount").textContent = Number(summary.invoice_count || 0).toLocaleString("en-IN");
  if ($("#recon2bCreditCount")) $("#recon2bCreditCount").textContent = Number(summary.credit_note_count || 0).toLocaleString("en-IN");
  if ($("#recon2bDebitCount")) $("#recon2bDebitCount").textContent = Number(summary.debit_note_count || 0).toLocaleString("en-IN");
  if ($("#recon2bAmendCount")) $("#recon2bAmendCount").textContent = Number(summary.amendment_count || 0).toLocaleString("en-IN");
  if ($("#recon2bNetItc")) $("#recon2bNetItc").textContent = money(summary.net_itc);
  if ($("#recon2bTaxable")) $("#recon2bTaxable").textContent = money(summary.net_taxable);
  if ($("#recon2bIgst")) $("#recon2bIgst").textContent = money(summary.net_igst);
  if ($("#recon2bCgst")) $("#recon2bCgst").textContent = money(summary.net_cgst);
  if ($("#recon2bSgst")) $("#recon2bSgst").textContent = money(summary.net_sgst);
  if ($("#recon2bGrossNetRows")) {
    const line = (label, count, itc, extra = {}, cls = "") => {
      const bucket = extra.bucket || {};
      return `<tr class="${cls}"><td>${label}</td><td>${Number(count || 0).toLocaleString("en-IN")}</td><td class="money">${money(bucket.taxable_value)}</td><td class="money">${money(bucket.igst)}</td><td class="money">${money(bucket.cgst)}</td><td class="money">${money(bucket.sgst)}</td><td class="money">${money(bucket.cess)}</td><td class="money">${money(itc)}</td></tr>`;
    };
    const inv = summary.buckets.Invoice;
    const credit = summary.buckets["Credit Note"];
    const debit = summary.buckets["Debit Note"];
    const amd = {
      taxable_value: summary.net_taxable - inv.taxable_value + credit.taxable_value - debit.taxable_value,
      igst: 0, cgst: 0, sgst: 0, cess: 0,
    };
    $("#recon2bGrossNetRows").innerHTML = [
      line("Gross Invoice ITC", summary.invoice_count, summary.gross_invoice_itc, { bucket: inv }),
      line("(+) Debit Note ITC", summary.debit_note_count, summary.debit_note_itc, { bucket: debit }),
      line("(−) Credit Note ITC", summary.credit_note_count, summary.credit_note_itc, { bucket: credit }),
      line("(±) Amendment Adjustment", summary.amendment_count, summary.amendment_itc, { bucket: amd }),
      `<tr class="gstr2-gross-row"><td><strong>Net GSTR-2B ITC</strong></td><td>${Number(summary.invoice_count || 0).toLocaleString("en-IN")}</td><td class="money">${money(summary.net_taxable)}</td><td class="money">${money(summary.net_igst)}</td><td class="money">${money(summary.net_cgst)}</td><td class="money">${money(summary.net_sgst)}</td><td class="money">${money(summary.net_cess)}</td><td class="money"><strong>${money(summary.net_itc)}</strong></td></tr>`,
    ].join("");
  }
  $("#recon2bSummaryPanel").classList.remove("hidden");
  refreshPurchaseDocumentTypeSummary(false);
}

function renderSignedDocumentTypeSummary(pack, opts = {}) {
  const money = reconMoney;
  const panel = $(opts.panelId || "");
  const rowsEl = $(opts.rowsId || "");
  if (!panel || !rowsEl || !pack) return;
  const portal = pack.portal || {};
  const byType = pack.by_type || [];
  const portalRows = portal.rows || [];
  const net = pack.net || {};
  const portalNetSource = net.portal_net_gst != null
    ? net.portal_net_gst
    : ((portal.net || {}).net_gst || 0);
  const portalNet = Number(portalNetSource || 0);
  const tallyNet = Number(net.tally_net_gst || 0);
  const diff = Number(net.difference != null ? net.difference : (tallyNet - portalNet));
  const matched = Boolean(net.matched);
  const hasTallyCompare = byType.length > 0;

  const typeRows = hasTallyCompare
    ? byType.map(item => {
        const left = item.portal || {};
        return {
          document_type: item.document_type,
          count: left.count || item.portal_count || 0,
          taxable_value: left.taxable_value,
          igst: left.igst,
          cgst: left.cgst,
          sgst: left.sgst,
          cess: left.cess,
          signed_total: left.signed_total,
          tally_count: item.tally_count || (item.tally || {}).count || 0,
          tally_signed: (item.tally || {}).signed_total,
          diff: item.signed_total_difference,
          type_matched: item.matched,
        };
      })
    : portalRows.map(row => ({
        document_type: row.document_type,
        count: row.count,
        taxable_value: row.taxable_value,
        igst: row.igst,
        cgst: row.cgst,
        sgst: row.sgst,
        cess: row.cess,
        signed_total: row.signed_total,
        tally_count: "—",
        tally_signed: null,
        diff: null,
        type_matched: null,
      }));

  rowsEl.innerHTML = typeRows.map(row => {
    const matchCell = row.type_matched == null
      ? "—"
      : (row.type_matched ? "Matched" : "Diff");
    const matchCls = row.type_matched == null ? "" : (row.type_matched ? "matched" : "review");
    return `<tr class="${matchCls}">
      <td>${escapeHtml(row.document_type || "")}</td>
      <td>${Number(row.count || 0).toLocaleString("en-IN")}</td>
      <td class="money">${money(row.taxable_value)}</td>
      <td class="money">${money(row.igst)}</td>
      <td class="money">${money(row.cgst)}</td>
      <td class="money">${money(row.sgst)}</td>
      <td class="money">${money(row.cess)}</td>
      <td class="money">${money(row.signed_total)}</td>
      <td>${row.tally_count === "—" ? "—" : Number(row.tally_count || 0).toLocaleString("en-IN")}</td>
      <td class="money">${row.tally_signed == null ? "—" : money(row.tally_signed)}</td>
      <td class="money">${row.diff == null ? "—" : money(row.diff)}</td>
      <td>${matchCell}</td>
    </tr>`;
  }).join("") || `<tr><td colspan="12">No documents.</td></tr>`;

  if ($(opts.netLabelId)) $(opts.netLabelId).textContent = money(portalNet);
  if ($(opts.portalNetId)) $(opts.portalNetId).textContent = money(portalNet);
  if ($(opts.tallyNetId)) $(opts.tallyNetId).textContent = hasTallyCompare ? money(tallyNet) : "—";
  if ($(opts.diffId)) $(opts.diffId).textContent = hasTallyCompare ? money(diff) : "—";
  if ($(opts.statusId)) {
    $(opts.statusId).textContent = hasTallyCompare
      ? (matched ? "Matched" : "Difference")
      : "Portal only — reconcile to compare with Tally";
  }
  panel.classList.remove("hidden");
}

async function refreshPurchaseDocumentTypeSummary(showError = false) {
  // Document Type Summary UI removed (placeholder card). Keep API call so backend pack stays fresh.
  if (!gstReconRows.length) return;
  try {
    const response = await fetch("/api/gst/recon/document-summary", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        portal: gstReconRows,
        tally: gstReconTallyRows || [],
        portal_label: "GSTR-2B",
        tally_label: "Tally Purchase",
        tolerance: Number($("#reconTolerance")?.value || 1),
      }),
    });
    const pack = await response.json();
    if (!response.ok) throw new Error(pack.error || "Document type summary failed.");
    return pack;
  } catch (failure) {
    if (showError) {
      const error = $("#gstError");
      error.textContent = failure.message;
      error.classList.remove("hidden");
    }
  }
}

function renderGstReconMatchResults(rows, counts, documentSummary = null) {
  gstReconResults = rows || [];
  const matched = Number(counts?.Matched || 0);
  const only2b = Number(counts?.["Only in GSTR-2B"] || 0);
  const onlyTally = Number(counts?.["Only in Tally"] || 0);
  const mismatch = Number(counts?.["Amount/Tax Mismatch"] || 0)
    + Number(counts?.["Date Mismatch"] || 0)
    + only2b
    + onlyTally;
  if ($("#reconVoucherMatchedCount")) {
    $("#reconVoucherMatchedCount").textContent = matched.toLocaleString("en-IN");
  }
  if ($("#reconVoucherMismatchCount")) {
    $("#reconVoucherMismatchCount").textContent = mismatch.toLocaleString("en-IN");
  }
  if ($("#reconVoucherReportPlaceholder")) {
    $("#reconVoucherReportPlaceholder").classList.remove("hidden");
  }
  if ($("#reconVoucherViewReportBtn")) {
    $("#reconVoucherViewReportBtn").disabled = true;
  }
  // documentSummary retained for callers / future dedicated report page; demo tables removed.
  void documentSummary;
}

function renderItcDashboardView(dashboard) {
  const money = reconMoney;
  const summary = dashboard?.summary || {};
  const grossNet = dashboard?.gross_net || {};
  const tallyBooked = dashboard?.tally_booked || {};
  $("#reconItcGstr2b").textContent = money(summary.gstr2b_net_itc ?? summary.gstr2b_itc);
  $("#reconItcGstr3b").textContent = money(summary.gstr3b_itc);
  $("#reconItcTally").textContent = money(summary.tally_itc);
  $("#reconItcAvailable").textContent = money(summary.available_itc);
  if ($("#reconItcGrossNetBar")) {
    $("#reconItcGrossInvoice").textContent = money(summary.gstr2b_gross_invoice_itc ?? grossNet.gross_invoice_itc);
    $("#reconItcCreditNote").textContent = money(summary.gstr2b_credit_note_itc ?? grossNet.credit_note_itc);
    $("#reconItcDebitNote").textContent = money(summary.gstr2b_debit_note_itc ?? grossNet.debit_note_itc);
    $("#reconItcAmendment").textContent = money(summary.gstr2b_amendment_itc ?? grossNet.amendment_itc);
    $("#reconItcGrossNetBar").classList.remove("hidden");
  }
  if ($("#reconTallyBookedPanel")) {
    const lines = tallyBooked.lines || [
      { particulars: "Purchase ITC", sign: "+", amount: summary.tally_booked_purchase_itc, count: 0 },
      { particulars: "Less Credit Note adjustment", sign: "−", amount: summary.tally_booked_credit_note_itc, count: 0 },
      { particulars: "Less ITC Reversal", sign: "−", amount: summary.tally_booked_reversal_itc, count: 0 },
      { particulars: "Plus Debit Note", sign: "+", amount: summary.tally_booked_debit_note_itc, count: 0 },
      { particulars: "Final Tally Booked (Section 4C)", sign: "=", amount: summary.tally_itc, count: 0 },
    ];
    $("#reconTallyBookedRows").innerHTML = lines.map(line => {
      const cls = line.sign === "=" ? "gstr2-gross-row" : "";
      return `<tr class="${cls}"><td>${escapeHtml(line.particulars || "")}</td><td>${Number(line.count || 0).toLocaleString("en-IN")}</td><td class="money">${money(line.amount)}</td></tr>`;
    }).join("");
    if ($("#reconTallyBookedFormula")) {
      const igst = money(summary.tally_booked_igst ?? tallyBooked.igst);
      const cgst = money(summary.tally_booked_cgst ?? tallyBooked.cgst);
      const sgst = money(summary.tally_booked_sgst ?? tallyBooked.sgst);
      $("#reconTallyBookedFormula").textContent =
        `${tallyBooked.formula || summary.tally_booked_formula || "Purchase ITC − Credit Note − ITC Reversal + Debit Note"}. `
        + `Section 4C components: IGST ₹${igst} · CGST ₹${cgst} · SGST ₹${sgst} = ₹${money(summary.tally_itc)}. `
        + `Not a gross purchase-voucher sum.`;
    }
    $("#reconTallyBookedPanel").classList.remove("hidden");
  }
  $("#reconItcRows").innerHTML = (dashboard?.rows || []).map(row => {
    const warn = row.available_itc > 1 ? "difference-warning" : "difference-ok";
    return `<tr class="${warn}"><td>${escapeHtml(row.period_label || row.period || "")}</td><td>${Number(row.gstr2b_invoices || 0).toLocaleString("en-IN")}</td><td class="money">${money(row.gstr2b_itc)}</td><td class="money">${money(row.gstr3b_itc)}</td><td class="money">${money(row.tally_itc)}</td><td class="money">${money(row.available_itc)}</td><td>${escapeHtml(row.action || "")}</td></tr>`;
  }).join("") || `<tr><td colspan="7">Import GSTR-2B to begin the ITC dashboard.</td></tr>`;
  $("#reconItcAdvice").textContent = summary.available_itc > 1
    ? `Net GSTR-2B ITC ₹${money(summary.gstr2b_net_itc ?? summary.gstr2b_itc)}. Tally Booked (4C) ₹${money(summary.tally_itc)}. Available ITC ₹${money(summary.available_itc)}.`
    : `Net GSTR-2B ITC ₹${money(summary.gstr2b_net_itc ?? summary.gstr2b_itc)} and Tally Booked (4C) ₹${money(summary.tally_itc)} are aligned within the allowed difference.`;
  $("#reconItcDashboardPanel").classList.remove("hidden");
  if ($("#reconItcViewDiffBtn")) $("#reconItcViewDiffBtn").disabled = !(Number(summary.gstr2b_itc) || Number(summary.tally_itc));
  if ($("#reconItcDiffExportBtn")) $("#reconItcDiffExportBtn").disabled = !window.__itcDiffRecon;
}

let itcDiffReconCache = null;

function renderItcDifferenceView(recon) {
  const money = reconMoney;
  const summary = recon?.summary || {};
  const counts = recon?.counts || {};
  itcDiffReconCache = recon;
  window.__itcDiffRecon = recon;
  if ($("#itcDiffPortalNet")) $("#itcDiffPortalNet").textContent = money(summary.portal_net_itc);
  if ($("#itcDiffTallyNet")) $("#itcDiffTallyNet").textContent = money(summary.tally_booked_itc);
  if ($("#itcDiffAvailable")) $("#itcDiffAvailable").textContent = money(summary.available_itc_difference);
  if ($("#itcDiffVoucherSum")) $("#itcDiffVoucherSum").textContent = money(summary.voucher_difference_total);
  if ($("#itcDiffBalanced")) {
    $("#itcDiffBalanced").textContent = summary.balanced ? "Balanced" : "Check rows";
    $("#itcDiffBalanced").style.color = summary.balanced ? "#0b8d55" : "#b45309";
  }
  if ($("#itcDiffSummaryText")) {
    $("#itcDiffSummaryText").textContent =
      `${summary.formula || "Net GSTR-2B ITC − Tally Booked (Section 4C)"} = ₹${money(summary.available_itc_difference)}. `
      + `Sum of voucher differences = ₹${money(summary.voucher_difference_total)}.`;
  }
  if ($("#itcDiffCounts")) {
    $("#itcDiffCounts").innerHTML = Object.entries(counts)
      .sort((a, b) => b[1] - a[1])
      .map(([status, count]) => `<span>${escapeHtml(status)} <strong>${Number(count).toLocaleString("en-IN")}</strong></span>`)
      .join("") || "";
  }
  if ($("#itcDiffRows")) {
    $("#itcDiffRows").innerHTML = (recon?.rows || []).map(row => {
      const nonzero = Math.abs(Number(row.difference || 0)) > 0.005;
      const cls = row.status === "Matched" && !nonzero ? "diff-matched" : (nonzero ? "diff-nonzero" : "review");
      const blank = (value) => (value === null || value === undefined || value === "" ? "—" : money(value));
      return `<tr class="${cls}">
        <td><span class="gst-status ${row.status === "Matched" ? "matched" : "review"}">${escapeHtml(row.status || "")}</span></td>
        <td class="reason-cell">${escapeHtml(row.reason || "")}</td>
        <td>${escapeHtml(row.gstin || "")}</td>
        <td>${escapeHtml(row.party_name || "")}</td>
        <td>${escapeHtml(row.invoice_no || "")}</td>
        <td>${escapeHtml(row.invoice_date || "")}</td>
        <td class="money">${blank(row.portal_taxable)}</td>
        <td class="money">${blank(row.tally_taxable)}</td>
        <td class="money">${blank(row.portal_itc)}</td>
        <td class="money">${blank(row.tally_itc)}</td>
        <td class="money">${money(row.difference)}</td>
      </tr>`;
    }).join("") || `<tr><td colspan="11">No vouchers to compare. Import GSTR-2B and sync Tally Purchase.</td></tr>`;
  }
  if ($("#itcDiffAdvice")) {
    $("#itcDiffAdvice").textContent = summary.balanced
      ? `Voucher differences explain Available ITC ₹${money(summary.available_itc_difference)} exactly. ${summary.row_formula || ""}`
      : `Voucher sum ₹${money(summary.voucher_difference_total)} vs Available ITC ₹${money(summary.available_itc_difference)} — review unmatched rows.`;
  }
  if ($("#reconItcDiffExportBtn")) $("#reconItcDiffExportBtn").disabled = !(recon?.rows || []).length;
}

async function loadItcDifferenceRecon(showError = true) {
  const tolerance = Number($("#reconTolerance")?.value || 1);
  try {
    const response = await fetch("/api/gst/recon/itc-difference", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        gstr2b: window.gstReconRows || undefined,
        tolerance,
      }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "Could not build difference reconciliation.");
    renderItcDifferenceView(result);
    return result;
  } catch (failure) {
    if (showError) alert(failure.message || String(failure));
    throw failure;
  }
}

async function openItcDifferenceDialog() {
  const dialog = $("#itcDiffDialog");
  if (!dialog) return;
  const btn = $("#reconItcViewDiffBtn");
  if (btn) {
    btn.disabled = true;
    btn.textContent = "Loading…";
  }
  try {
    await loadItcDifferenceRecon(true);
    if (typeof dialog.showModal === "function") dialog.showModal();
    else dialog.classList.remove("hidden");
  } catch (_failure) {
    /* alert already shown */
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = "View Difference";
    }
  }
}

async function exportItcDifferenceExcel() {
  try {
    if (!itcDiffReconCache) await loadItcDifferenceRecon(true);
    const response = await fetch("/api/gst/recon/itc-difference-export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        recon: itcDiffReconCache,
        title: "Available ITC Difference",
        tolerance: Number($("#reconTolerance")?.value || 1),
      }),
    });
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err.error || "Excel export failed.");
    }
    const blob = await response.blob();
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Available_ITC_Difference.xlsx";
    link.click();
    URL.revokeObjectURL(link.href);
  } catch (failure) {
    alert(failure.message || String(failure));
  }
}

let salesDiffReconCache = null;
let salesDiffFilteredRows = [];

function salesDiffDialogRoot() {
  return document.getElementById("salesDiffDialog");
}

function salesDiffQuery(id) {
  const root = salesDiffDialogRoot();
  if (root) {
    const node = root.querySelector(`#${id}`);
    if (node) return node;
  }
  return document.getElementById(id);
}

function salesDiffSourceRows() {
  const rows = salesDiffReconCache && Array.isArray(salesDiffReconCache.rows)
    ? salesDiffReconCache.rows
    : [];
  return rows;
}

/** Exception report: non-zero Output GST difference only; never Matched / zero rows. */
function salesDiffExceptionRows(rows) {
  return (rows || []).filter(row => {
    if (String(row.status || "") === "Matched") return false;
    return Math.abs(Number(row.difference || 0)) > 0.005;
  });
}

function renderSalesDifferenceBadges(recon) {
  const counts = recon?.counts || {};
  const summary = recon?.summary || {};
  const money = reconMoney;
  if ($("#salesDiffBadgeMatched")) $("#salesDiffBadgeMatched").textContent = Number(counts.Matched || 0).toLocaleString("en-IN");
  if ($("#salesDiffBadgeOnlyG1")) $("#salesDiffBadgeOnlyG1").textContent = Number(counts["Only in GSTR-1"] || 0).toLocaleString("en-IN");
  if ($("#salesDiffBadgeOnlyTally")) $("#salesDiffBadgeOnlyTally").textContent = Number(counts["Only in Tally"] || 0).toLocaleString("en-IN");
  if ($("#salesDiffBadgeDate")) $("#salesDiffBadgeDate").textContent = Number(counts["Date Mismatch"] || 0).toLocaleString("en-IN");
  if ($("#salesDiffBadgeAmount")) $("#salesDiffBadgeAmount").textContent = Number(counts["Amount/Tax Mismatch"] || 0).toLocaleString("en-IN");
  if ($("#salesDiffBadgeOutput")) $("#salesDiffBadgeOutput").textContent = money(summary.output_gst_difference);
  if ($("#salesDiffBadgeVoucherSum")) $("#salesDiffBadgeVoucherSum").textContent = money(summary.voucher_difference_total);
  if ($("#salesDiffBadges")) $("#salesDiffBadges").classList.toggle("hidden", !(recon?.rows || []).length);
}

function salesDiffMoneyMismatch(left, right, tolerance = 1) {
  const a = left === null || left === undefined || left === "" ? null : Number(left);
  const b = right === null || right === undefined || right === "" ? null : Number(right);
  if (a === null && b === null) return false;
  if (!Number.isFinite(a) && !Number.isFinite(b)) return false;
  return Math.abs((Number.isFinite(a) ? a : 0) - (Number.isFinite(b) ? b : 0)) > tolerance;
}

function salesDiffNormalizeInvoice(value) {
  const compact = String(value || "").replace(/[\s\-_\/\\.]+/g, "").toUpperCase();
  return compact.replace(/^0+/, "") || (compact ? "0" : "");
}

function salesDiffNormalizeDate(value) {
  return String(value || "").trim().replace(/\//g, "-");
}

function salesDiffFieldMismatches(row, tolerance = 1) {
  if (Array.isArray(row.field_mismatches) && row.field_mismatches.length) {
    return row.field_mismatches.slice();
  }
  const status = String(row.status || "");
  if (status === "Only in GSTR-1") return ["Only in GSTR-1"];
  if (status === "Only in Tally") return ["Only in Tally"];
  const mismatches = [];
  const g1Gstin = String(row.gstr1_gstin != null ? row.gstr1_gstin : row.gstin || "").trim().toUpperCase();
  const tGstin = String(row.tally_gstin != null ? row.tally_gstin : row.gstin || "").trim().toUpperCase();
  if (row.gstr1_gstin != null || row.tally_gstin != null) {
    if (g1Gstin !== tGstin && (g1Gstin || tGstin)) mismatches.push("GSTIN Mismatch");
  }
  const g1Inv = salesDiffNormalizeInvoice(row.gstr1_invoice_no != null ? row.gstr1_invoice_no : row.invoice_no);
  const tInv = salesDiffNormalizeInvoice(row.tally_invoice_no != null ? row.tally_invoice_no : row.invoice_no);
  if (row.gstr1_invoice_no != null || row.tally_invoice_no != null) {
    if (g1Inv !== tInv && (g1Inv || tInv)) mismatches.push("Invoice Number Mismatch");
  }
  const g1Date = salesDiffNormalizeDate(row.gstr1_invoice_date != null ? row.gstr1_invoice_date : row.invoice_date);
  const tDate = salesDiffNormalizeDate(row.tally_invoice_date != null ? row.tally_invoice_date : row.invoice_date);
  if (row.gstr1_invoice_date != null || row.tally_invoice_date != null) {
    if (g1Date !== tDate && (g1Date || tDate)) mismatches.push("Invoice Date Mismatch");
  }
  if (salesDiffMoneyMismatch(row.gstr1_taxable, row.tally_taxable, tolerance)) mismatches.push("Taxable Mismatch");
  if (salesDiffMoneyMismatch(row.gstr1_igst, row.tally_igst, tolerance)) mismatches.push("IGST Mismatch");
  if (salesDiffMoneyMismatch(row.gstr1_cgst, row.tally_cgst, tolerance)) mismatches.push("CGST Mismatch");
  if (salesDiffMoneyMismatch(row.gstr1_sgst, row.tally_sgst, tolerance)) mismatches.push("SGST Mismatch");
  if (salesDiffMoneyMismatch(row.gstr1_cess, row.tally_cess, tolerance)) mismatches.push("CESS Mismatch");
  return mismatches;
}

function salesDiffReasonLabel(row, tolerance = 1) {
  if (row.difference_reason) return String(row.difference_reason);
  const mismatches = salesDiffFieldMismatches(row, tolerance).filter(
    (label) => label !== "Only in GSTR-1" && label !== "Only in Tally"
  );
  if (mismatches.length > 1) return "Multiple Field Mismatch";
  if (mismatches.length === 1) return mismatches[0];
  return String(row.reason || row.status || "");
}

function salesDiffIdentityDisplay(g1Value, tallyValue, fallback) {
  const g1 = String(g1Value != null ? g1Value : "").trim();
  const tally = String(tallyValue != null ? tallyValue : "").trim();
  const shared = String(fallback || "").trim();
  if (g1 && tally && g1 !== tally) {
    return { text: `${g1} ≠ ${tally}`, mismatch: true };
  }
  return { text: g1 || tally || shared || "", mismatch: false };
}

function buildSalesDiffRowElement(row, money, blank) {
  const tr = document.createElement("tr");
  const tolerance = Number($("#salesReconTolerance")?.value || 1);
  const mismatches = new Set(salesDiffFieldMismatches(row, tolerance));
  const gstinCell = salesDiffIdentityDisplay(row.gstr1_gstin, row.tally_gstin, row.gstin);
  const invCell = salesDiffIdentityDisplay(row.gstr1_invoice_no, row.tally_invoice_no, row.invoice_no);
  const dateCell = salesDiffIdentityDisplay(row.gstr1_invoice_date, row.tally_invoice_date, row.invoice_date);
  const taxableMismatch = mismatches.has("Taxable Mismatch");
  const igstMismatch = mismatches.has("IGST Mismatch");
  const cgstMismatch = mismatches.has("CGST Mismatch");
  const sgstMismatch = mismatches.has("SGST Mismatch");
  const cessMismatch = mismatches.has("CESS Mismatch");
  const gstinMismatch = mismatches.has("GSTIN Mismatch") || gstinCell.mismatch;
  const invMismatch = mismatches.has("Invoice Number Mismatch") || invCell.mismatch;
  const dateMismatch = mismatches.has("Invoice Date Mismatch") || dateCell.mismatch;
  const reason = salesDiffReasonLabel(row, tolerance);
  const cells = [
    { html: `<span class="gst-status review">${escapeHtml(row.status || "")}</span>`, money: false, mismatch: false },
    { html: escapeHtml(gstinCell.text), money: false, mismatch: gstinMismatch },
    { html: escapeHtml(row.party_name || ""), money: false, mismatch: false },
    { html: escapeHtml(invCell.text), money: false, mismatch: invMismatch },
    { html: escapeHtml(dateCell.text), money: false, mismatch: dateMismatch },
    { html: escapeHtml(row.voucher_type || ""), money: false, mismatch: false },
    { html: blank(row.gstr1_taxable), money: true, mismatch: taxableMismatch },
    { html: blank(row.tally_taxable), money: true, mismatch: taxableMismatch },
    { html: money(row.difference), money: true, mismatch: false },
    { html: blank(row.gstr1_igst), money: true, mismatch: igstMismatch },
    { html: blank(row.tally_igst), money: true, mismatch: igstMismatch },
    { html: blank(row.gstr1_cgst), money: true, mismatch: cgstMismatch },
    { html: blank(row.tally_cgst), money: true, mismatch: cgstMismatch },
    { html: blank(row.gstr1_sgst), money: true, mismatch: sgstMismatch },
    { html: blank(row.tally_sgst), money: true, mismatch: sgstMismatch },
    { html: blank(row.gstr1_cess), money: true, mismatch: cessMismatch },
    { html: blank(row.tally_cess), money: true, mismatch: cessMismatch },
    { html: escapeHtml(reason), money: false, mismatch: false, reason: true },
  ];
  cells.forEach((cell) => {
    const td = document.createElement("td");
    const classes = [];
    if (cell.money) classes.push("money");
    if (cell.reason) classes.push("reason-cell");
    if (cell.mismatch) classes.push("value-mismatch");
    if (classes.length) td.className = classes.join(" ");
    td.innerHTML = cell.html;
    tr.appendChild(td);
  });
  return tr;
}

function bindSalesDiffTableRows() {
  const money = reconMoney;
  const blank = (value) => (value === null || value === undefined || value === "" ? "—" : money(value));
  const sourceRows = salesDiffSourceRows();
  const filtered = salesDiffExceptionRows(sourceRows);
  salesDiffFilteredRows = filtered;

  const tableBody = salesDiffQuery("salesDiffRows");
  console.info("[sales-diff] exception report", {
    total_reconciliation_records: sourceRows.length,
    exception_rows: filtered.length,
    table_body_found: Boolean(tableBody),
  });
  if (!tableBody) {
    console.error("[sales-diff] #salesDiffRows not found — cannot bind voucher grid");
    return filtered;
  }

  tableBody.replaceChildren();
  if (!filtered.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 18;
    td.textContent = sourceRows.length
      ? "No invoices contribute to Output GST Difference for this period."
      : "No vouchers to compare. Import GSTR-1 and sync Tally Sales.";
    tr.appendChild(td);
    tableBody.appendChild(tr);
  } else {
    const fragment = document.createDocumentFragment();
    filtered.forEach(row => fragment.appendChild(buildSalesDiffRowElement(row, money, blank)));
    tableBody.appendChild(fragment);
  }

  const wrap = salesDiffQuery("salesDiffTableWrap");
  if (wrap) {
    wrap.hidden = false;
    wrap.style.display = "block";
    wrap.scrollTop = 0;
  }
  return filtered;
}

function renderSalesDifferenceView(recon) {
  const money = reconMoney;
  const summary = recon?.summary || {};
  const sourceRows = Array.isArray(recon?.rows) ? recon.rows : [];
  salesDiffReconCache = recon;
  window.__salesDiffRecon = recon;
  console.info("[sales-diff] render", {
    total_reconciliation_records: sourceRows.length,
    counts: recon?.counts || {},
    excel_dataset_rows: sourceRows.length,
  });
  renderSalesDifferenceBadges(recon);
  const months = summary.contributing_period_labels || [];
  const monthText = months.length ? months.join(", ") : (summary.return_period === "ALL" ? "ALL / FY" : (summary.return_period || "—"));
  if (salesDiffQuery("salesDiffMonthLabel")) salesDiffQuery("salesDiffMonthLabel").textContent = monthText;
  if (salesDiffQuery("salesDiffPortalNet")) salesDiffQuery("salesDiffPortalNet").textContent = money(summary.gstr1_output_gst);
  if (salesDiffQuery("salesDiffTallyNet")) salesDiffQuery("salesDiffTallyNet").textContent = money(summary.tally_output_gst);
  if (salesDiffQuery("salesDiffAvailable")) salesDiffQuery("salesDiffAvailable").textContent = money(summary.output_gst_difference);
  if (salesDiffQuery("salesDiffPortalLabel")) {
    salesDiffQuery("salesDiffPortalLabel").textContent =
      `GST Portal Reported Difference · ${monthText} · ₹${money(summary.output_gst_difference)}`;
  }
  if (salesDiffQuery("salesDiffSummaryText")) {
    const exceptions = salesDiffExceptionRows(sourceRows);
    const top = exceptions.slice().sort((a, b) => Math.abs(Number(b.difference || 0)) - Math.abs(Number(a.difference || 0)))[0];
    const invNote = top
      ? ` Leading contributor: ${top.invoice_no || "—"} (${top.invoice_date || "—"}) · ₹${money(top.difference)}.`
      : "";
    salesDiffQuery("salesDiffSummaryText").textContent =
      `GST Portal Reported Difference for ${monthText}: ₹${money(summary.output_gst_difference)}. `
      + `Portal GSTR-1 values are preserved exactly (not auto-corrected). `
      + `${exceptions.length.toLocaleString("en-IN")} contributing invoice(s).`
      + invNote;
  }
  const bound = bindSalesDiffTableRows();
  const hasRows = sourceRows.length > 0;
  if ($("#reconSalesViewDiffBtn")) $("#reconSalesViewDiffBtn").disabled = !hasRows && !(gstReconGstr1Rows.length || gstReconTallySalesRows.length);
  if ($("#reconSalesDiffExportBtn")) $("#reconSalesDiffExportBtn").disabled = !hasRows;
  if ($("#reconSalesDiffExportCsvBtn")) $("#reconSalesDiffExportCsvBtn").disabled = !hasRows;
  return bound;
}

async function loadSalesDifferenceRecon(showError = true) {
  const tolerance = Number($("#salesReconTolerance")?.value || $("#reconTolerance")?.value || 1);
  try {
    // Let server load authoritative rows from SQLite (same source Excel rebuild uses).
    // Avoid posting huge client arrays that can truncate the response payload.
    const response = await fetch("/api/gst/recon/gstr1-difference", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        tolerance,
        returnPeriod: getGstReconPeriod(),
      }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "Could not build GSTR-1 difference reconciliation.");
    const rowCount = Array.isArray(result?.rows) ? result.rows.length : 0;
    console.info("[sales-diff] api result", {
      total_reconciliation_records: rowCount,
      counts: result?.counts || {},
      keys: result ? Object.keys(result) : [],
    });
    if (!rowCount && result?.counts && Object.values(result.counts).some(v => Number(v) > 0)) {
      console.error("[sales-diff] counts present but rows array empty — UI binding would show blank grid");
    }
    renderSalesDifferenceView(result);
    return result;
  } catch (failure) {
    if (showError) alert(failure.message || String(failure));
    throw failure;
  }
}

async function openSalesDifferenceDialog() {
  const dialog = salesDiffDialogRoot();
  if (!dialog) {
    console.error("[sales-diff] #salesDiffDialog missing");
    return;
  }
  const btn = $("#reconSalesViewDiffBtn");
  if (btn) {
    btn.disabled = true;
    btn.textContent = "Loading…";
  }
  try {
    if (typeof dialog.showModal === "function") dialog.showModal();
    else dialog.classList.remove("hidden");
    await loadSalesDifferenceRecon(true);
    bindSalesDiffTableRows();
  } catch (_failure) {
    /* alert already shown */
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = "View Difference";
    }
  }
}

async function exportSalesDifference(format = "xlsx") {
  try {
    if (!salesDiffReconCache || !Array.isArray(salesDiffReconCache.rows) || !salesDiffReconCache.rows.length) {
      await loadSalesDifferenceRecon(true);
    }
    const exceptionRows = salesDiffExceptionRows(salesDiffSourceRows());
    const exportRecon = {
      ...(salesDiffReconCache || {}),
      rows: exceptionRows,
    };
    const response = await fetch("/api/gst/recon/gstr1-difference-export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        recon: exportRecon,
        title: "Output GST Difference — Exceptions",
        tolerance: Number($("#salesReconTolerance")?.value || 1),
        returnPeriod: getGstReconPeriod(),
        format,
      }),
    });
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err.error || "Export failed.");
    }
    const blob = await response.blob();
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = format === "csv" ? "GSTR1_Output_GST_Exceptions.csv" : "GSTR1_Output_GST_Exceptions.xlsx";
    link.click();
    URL.revokeObjectURL(link.href);
  } catch (failure) {
    alert(failure.message || String(failure));
  }
}

async function saveGstReconSession(partial = {}) {
  const payload = {
    gstin: getGstPortalGstin(partial),
    financialYear: $("#margFinancialYear")?.value?.trim() || "2025-26",
  };
  if (partial.gstr2b) {
    payload.datasetKey = "GSTR-2B";
    const gstin = payload.gstin;
    payload.rows = (partial.gstr2b || []).map((row) => ({
      ...row,
      taxpayer_gstin: row.taxpayer_gstin || row.filing_gstin || gstin || "",
      filing_gstin: row.filing_gstin || row.taxpayer_gstin || gstin || "",
      financial_year: row.financial_year || "2025-26",
      return_type: "GSTR-2B",
    }));
  }
  if (partial.gstr2a) {
    payload.datasetKey = "GSTR-2A";
    const gstin = payload.gstin;
    payload.rows = (partial.gstr2a || []).map((row) => ({
      ...row,
      taxpayer_gstin: row.taxpayer_gstin || row.filing_gstin || gstin || "",
      filing_gstin: row.filing_gstin || row.taxpayer_gstin || gstin || "",
      financial_year: row.financial_year || "2025-26",
      return_type: "GSTR-2A",
    }));
  }
  if (partial.gstr3b !== undefined) payload.gstr3b = partial.gstr3b;
  if (partial.tally_sync !== undefined) payload.tally_sync = partial.tally_sync;
  if (partial.tally_sales_sync !== undefined) payload.tally_sales_sync = partial.tally_sales_sync;
  if (partial.results !== undefined) payload.results = partial.results;
  if (partial.gstr1_results !== undefined) payload.gstr1_results = partial.gstr1_results;
  if (partial.sales_dashboard !== undefined) payload.sales_dashboard = partial.sales_dashboard;
  if (partial.gstr3b_dashboard !== undefined) payload.gstr3b_dashboard = partial.gstr3b_dashboard;
  if (!Object.keys(payload).filter((key) => !["gstin", "financialYear"].includes(key)).length) return;
  await fetch("/api/gst/recon/save", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
}

let gstPortalContext = { gstin: "", financial_year: "2025-26" };

function inferTaxpayerGstinFromText(...values) {
  const pattern = /\b([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][A-Z0-9]Z[A-Z0-9])\b/i;
  for (const value of values) {
    const match = String(value || "").match(pattern);
    if (match) return match[1].toUpperCase();
  }
  return "";
}

function getGstPortalGstin(partial = {}) {
  return String(
    partial.gstin
    || gstPortalContext.gstin
    || inferTaxpayerGstinFromText(
      ...(partial.gstr2b || []).map((row) => row.taxpayer_gstin || row.filing_gstin || row.source_file || ""),
      ...(partial.gstr2a || []).map((row) => row.taxpayer_gstin || row.filing_gstin || row.source_file || ""),
      ...((gstReconRows || []).map((row) => row.taxpayer_gstin || row.filing_gstin || "")),
      ...((gstDatasets["GSTR-2A"] || []).map((row) => row.taxpayer_gstin || row.filing_gstin || "")),
      ...((gstReconGstr1Rows || []).map((row) => row.taxpayer_gstin || row.filing_gstin || "")),
      (gstDatasets["GSTR-3B"] || {}).gstin || ""
    )
    || ""
  ).trim().toUpperCase();
}

let gstReconPortalPageReady = false;

async function resetGstReconPortalSession(force = false) {
  /**
   * force=true (GST Refresh): clear portal imports in SQLite; keep Tally.
   * force=false (first module open): do NOT wipe DB — Overview/tabs share SQLite state.
   */
  if (gstReconPortalPageReady && !force) return;
  if (!force) {
    gstReconPortalPageReady = true;
    return { reset: false, restored: true };
  }
  const response = await fetch("/api/gst/recon/session-start", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      returnPeriod: getGstReconPeriod(),
      tolerance: Number($("#reconOverviewTolerance")?.value || 1),
    }),
  });
  const result = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(result.error || "Could not reset GST portal session.");
  gstReconPortalPageReady = true;
  gstDatasets["GSTR-2B"] = undefined;
  gstDatasets["GSTR-1"] = undefined;
  gstDatasets["GSTR-3B"] = undefined;
  return result;
}

async function fetchGstReconDatasets(include = []) {
  const keys = [...new Set((include || []).map((key) => String(key || "").trim().toLowerCase()).filter(Boolean))];
  if (!keys.length) return {};
  const missing = keys.filter((key) => !gstReconDatasetsLoaded.has(key));
  if (!missing.length) return {};
  const response = await fetch("/api/gst/recon/datasets", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      include: missing,
      gstin: getGstPortalGstin(),
      financialYear: $("#margFinancialYear")?.value?.trim() || gstPortalContext.financial_year || "2025-26",
    }),
  });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "Could not load GST datasets.");
  missing.forEach((key) => gstReconDatasetsLoaded.add(key));
  return data;
}

function applyGstReconDatasetBundle(data = {}, options = {}) {
  const portalAllowed = options.portalAllowed !== false;
  if (Object.prototype.hasOwnProperty.call(data, "gstr2b")) {
    gstReconRows = portalAllowed ? (data.gstr2b || []) : [];
    if (gstReconRows.length) {
      gstDatasets["GSTR-2B"] = gstReconRows;
      renderGstRecon2bSummary(gstReconRows);
    } else {
      gstDatasets["GSTR-2B"] = undefined;
      if ($("#recon2bSummaryPanel")) $("#recon2bSummaryPanel").classList.add("hidden");
    }
  }
  if (Object.prototype.hasOwnProperty.call(data, "gstr2a")) {
    const gstr2aRows = portalAllowed ? (data.gstr2a || []) : [];
    if (gstr2aRows.length) gstDatasets["GSTR-2A"] = gstr2aRows;
    else if (activeGstModule !== "reconciliation") gstDatasets["GSTR-2A"] = undefined;
    else if (!portalAllowed) gstDatasets["GSTR-2A"] = undefined;
  }
  if (Object.prototype.hasOwnProperty.call(data, "tally_purchase")) {
    gstReconTallyRows = data.tally_purchase || [];
  }
  if (Object.prototype.hasOwnProperty.call(data, "gstr1")) {
    gstReconGstr1Rows = portalAllowed ? (data.gstr1 || []) : [];
    if (gstReconGstr1Rows.length) gstDatasets["GSTR-1"] = gstReconGstr1Rows;
    else gstDatasets["GSTR-1"] = undefined;
  }
  if (Object.prototype.hasOwnProperty.call(data, "tally_sales")) {
    gstReconTallySalesRows = data.tally_sales || [];
  }
  if (Object.prototype.hasOwnProperty.call(data, "results")) {
    gstReconResults = portalAllowed ? (data.results || []) : [];
  }
  if (Object.prototype.hasOwnProperty.call(data, "gstr1_results")) {
    gstReconSalesResults = portalAllowed ? (data.gstr1_results || []) : [];
  }
  if (Object.prototype.hasOwnProperty.call(data, "gstr3b")) {
    const g3 = data.gstr3b || {};
    if (portalAllowed && g3 && Object.keys(g3).length && (g3.imported_periods || []).length) {
      gstDatasets["GSTR-3B"] = g3;
    } else if (!portalAllowed) {
      gstDatasets["GSTR-3B"] = undefined;
    }
  }
}

function datasetsNeededForGstModule(module = activeGstModule, tab = activeReconTab) {
  if (module === "reconciliation") return ["gstr2a", "gstr2b"];
  if (module !== "threeway") return [];
  if (tab === "purchase") return ["gstr2b", "tally_purchase", "results"];
  if (tab === "sales") return ["gstr1", "tally_sales", "gstr1_results"];
  if (tab === "gstr3b") return ["gstr3b"];
  if (tab === "payment") return [];
  return []; // overview: counts/meta only
}

async function ensureGstReconDatasetsForModule(module = activeGstModule, options = {}) {
  const include = datasetsNeededForGstModule(module, options.tab || activeReconTab);
  if (!include.length) return;
  const data = await fetchGstReconDatasets(include);
  if (data && Object.keys(data).length) {
    applyGstReconDatasetBundle(data, { portalAllowed: options.portalAllowed !== false });
  }
}

async function loadGstReconSession(options = {}) {
  const restorePortal = options.restorePortal !== false;
  const loadId = ++gstReconSessionLoadSeq;
  // First load after page open: drop previous-session portal files (F5 / reopen).
  if (!gstReconPortalPageReady) {
    try { await resetGstReconPortalSession(false); } catch (_) { gstReconPortalPageReady = true; }
  }
  if (loadId !== gstReconSessionLoadSeq) return;
  gstReconDatasetsLoaded = new Set();
  const response = await fetch("/api/gst/recon/load", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      mode: "summary",
      gstin: getGstPortalGstin(),
      financialYear: $("#margFinancialYear")?.value?.trim() || "2025-26",
    }),
  });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "Could not load saved GST reconciliation data.");
  // A newer import/open invalidated this load — do not wipe fresh in-memory data.
  if (loadId !== gstReconSessionLoadSeq) return;
  if (data.portal_context) {
    gstPortalContext = {
      gstin: String(data.portal_context.gstin || "").toUpperCase(),
      financial_year: data.portal_context.financial_year || "2025-26",
    };
  }
  gstReconDatasetCounts = data.counts || {};

  // After page/GST refresh, portal tables are empty. During the same browser page,
  // restorePortal keeps in-session imports when re-opening the module.
  const portalAllowed = restorePortal && gstReconPortalPageReady;
  gstReconRows = [];
  gstReconTallyRows = [];
  gstReconResults = [];
  gstReconGstr1Rows = [];
  gstReconTallySalesRows = [];
  gstReconSalesResults = [];
  const salesDash = data.sales_dashboard || null;
  const g3Dash = data.gstr3b_dashboard || null;
  // Empty {} meta leftovers must not revive dashboard cards after clear.
  salesReconDashboard = portalAllowed && salesDash && Object.keys(salesDash).length ? salesDash : null;
  gstr3bDashboard = portalAllowed && g3Dash && Object.keys(g3Dash).length ? g3Dash : null;

  if (portalAllowed && data.gstr3b && Object.keys(data.gstr3b).length && (data.gstr3b.imported_periods || []).length) {
    gstDatasets["GSTR-3B"] = data.gstr3b;
  } else {
    gstDatasets["GSTR-3B"] = undefined;
  }
  if (gstDatasets["GSTR-3B"]) {
    populateGstr3bPeriodSelector(
      gstDatasets["GSTR-3B"].imported_periods || Object.keys(gstDatasets["GSTR-3B"].periods || {}),
      getGstReconPeriod()
    );
  }

  // Paint Overview cards immediately from summary meta (fast), then refresh from API.
  if (activeGstModule === "threeway") {
    ensureGstReconPanelVisible();
    if (data.overview && Object.keys(data.overview).length) {
      renderGstReconOverview(data.overview);
    }
  }

  await ensureGstReconDatasetsForModule(activeGstModule, { portalAllowed, tab: activeReconTab });
  if (loadId !== gstReconSessionLoadSeq) return;

  // Keep one shared Period / Month across Overview / Sales / GSTR-3B (default ALL / FY).
  setGstReconPeriod(gstReconPeriodUserChosen ? getGstReconPeriod() : "ALL", { refresh: false, silent: true });
  if (gstReconGstr1Rows.length) gstDatasets["GSTR-1"] = gstReconGstr1Rows;
  else if (!gstReconDatasetsLoaded.has("gstr1")) gstDatasets["GSTR-1"] = undefined;

  if (gstReconResults.length) {
    const counts = gstReconResults.reduce((all, row) => {
      all[row.status] = (all[row.status] || 0) + 1;
      return all;
    }, {});
    renderGstReconMatchResults(gstReconResults, counts);
  } else if ($("#reconVoucherReportPlaceholder")) {
    $("#reconVoucherReportPlaceholder").classList.add("hidden");
    if ($("#reconVoucherMatchedCount")) $("#reconVoucherMatchedCount").textContent = "0";
    if ($("#reconVoucherMismatchCount")) $("#reconVoucherMismatchCount").textContent = "0";
  }
  if ($("#reconItcDashboardPanel")) $("#reconItcDashboardPanel").classList.add("hidden");

  renderGstReconStatus();
  renderSalesReconStatus();
  updateSalesReconReady();
  if (gstReconSalesResults.length || salesReconDashboard) {
    try {
      renderSalesReconDashboard(salesReconDashboard || { rows: gstReconSalesResults, counts: {}, cards: {}, output_summary: { gstr1: {}, tally: {}, difference: {} } });
      renderSalesReconTable();
    } catch (_) {}
  }
  if (gstr3bDashboard) {
    try { renderGstr3bDashboard(gstr3bDashboard); } catch (_) {}
  }
  if (gstReconRows.length && activeReconTab === "purchase") await refreshItcDashboard(false);
  else if (!gstReconRows.length) {
    // Empty ITC cards after refresh — do not keep stale period rows on screen.
    if ($("#reconItcGstr2b")) $("#reconItcGstr2b").textContent = "0.00";
    if ($("#reconItcGstr3b")) $("#reconItcGstr3b").textContent = "0.00";
    if ($("#reconItcTally")) $("#reconItcTally").textContent = "0.00";
    if ($("#reconItcAvailable")) $("#reconItcAvailable").textContent = "0.00";
    if ($("#reconItcRows")) $("#reconItcRows").innerHTML = "";
  }
  if (loadId !== gstReconSessionLoadSeq) return;
  if (activeGstModule === "threeway") {
    ensureGstReconPanelVisible();
    if (activeReconTab === "overview") {
      await refreshGstReconOverview(false);
    } else {
      await refreshActiveReconPeriodViews(false);
    }
    ensureGstReconPanelVisible();
  } else if (data.overview) {
    renderGstReconOverview(data.overview);
  }
}

async function syncGstReconTallyPurchase() {
  const button = $("#reconTallySyncBtn");
  const error = $("#gstError");
  button.disabled = true;
  button.textContent = "Syncing...";
  error.classList.add("hidden");
  try {
    const response = await fetch("/api/gst/tally/purchase-sync", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: "{}",
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "Tally purchase sync failed.");
    gstReconTallyRows = result.rows || [];
    await saveGstReconSession({ tally_sync: { company: result.company, count: result.count, synced_at: result.synced_at } });
    renderGstReconStatus();
    if (gstReconRows.length) {
      await refreshPurchaseDocumentTypeSummary(false);
      await refreshItcDashboard(false);
    }
  } catch (failure) {
    error.textContent = failure.message;
    error.classList.remove("hidden");
  } finally {
    button.disabled = false;
    button.textContent = "Sync Tally Purchase";
    updateGstReconReady();
  }
}

async function reconcileGstRecon2bTally() {
  const button = $("#recon2bTallyBtn");
  const error = $("#gstError");
  button.disabled = true;
  button.textContent = "Reconciling...";
  error.classList.add("hidden");
  try {
    const response = await fetch("/api/gst/recon/2b-tally", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        gstr2b: gstReconRows,
        tally_purchase: gstReconTallyRows,
        tolerance: Number($("#reconTolerance").value || 1),
      }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "GSTR-2B vs Tally reconciliation failed.");
    renderGstReconMatchResults(result.rows || [], result.counts || {}, result.document_summary || null);
    await saveGstReconSession({ results: result.rows || [] });
    await refreshItcDashboard(false);
  } catch (failure) {
    error.textContent = failure.message;
    error.classList.remove("hidden");
  } finally {
    button.disabled = false;
    button.textContent = "Reconcile 2B vs Tally";
    updateGstReconReady();
  }
}

async function refreshItcDashboard(showError = true) {
  const button = $("#reconItcBtn");
  const error = $("#gstError");
  if (button) {
    button.disabled = true;
    button.textContent = "Refreshing...";
  }
  if (showError) error.classList.add("hidden");
  try {
    const response = await fetch("/api/gst/recon/itc-dashboard", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        gstr2b: gstReconRows,
        tally_purchase: gstReconTallyRows,
        gstr3b: gstDatasets["GSTR-3B"] || {},
        tolerance: Number($("#reconTolerance").value || 1),
      }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "ITC dashboard refresh failed.");
    renderItcDashboardView(result);
  } catch (failure) {
    if (showError) {
      error.textContent = failure.message;
      error.classList.remove("hidden");
    }
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = "Refresh ITC Dashboard";
      updateGstReconReady();
    }
  }
}

function resolveGstClearReturnType(preferred = "") {
  /** Prefer explicit button return type, then active recon tab, then dropdown. */
  const forced = String(preferred || "").toUpperCase().trim();
  if (forced.includes("3B")) return "GSTR-3B";
  if (forced.includes("2B")) return "GSTR-2B";
  if (forced.includes("GSTR-1") || forced.includes("SALES")) return "GSTR-1";
  if (activeReconTab === "sales") return "GSTR-1";
  if (activeReconTab === "gstr3b") return "GSTR-3B";
  if (activeReconTab === "purchase") return "GSTR-2B";
  const rt = String($("#gstReturnType")?.value || "").toUpperCase();
  if (rt.includes("3B")) return "GSTR-3B";
  if (rt.includes("2B")) return "GSTR-2B";
  if (rt.includes("GSTR-1") || rt.includes("SALES")) return "GSTR-1";
  return "GSTR-2B";
}

function clearGstReconClientStateFor(returnType) {
  const kind = String(returnType || "").toUpperCase();
  if (kind === "GSTR-2B") {
    gstReconRows = [];
    gstReconResults = [];
    gstDatasets["GSTR-2B"] = undefined;
    gstReconDatasetsLoaded.delete("gstr2b");
    gstReconDatasetsLoaded.delete("results");
    if ($("#recon2bSummaryPanel")) $("#recon2bSummaryPanel").classList.add("hidden");
    if ($("#reconItcDashboardPanel")) $("#reconItcDashboardPanel").classList.add("hidden");
    if ($("#reconVoucherReportPlaceholder")) $("#reconVoucherReportPlaceholder").classList.add("hidden");
    if ($("#reconItcRows")) $("#reconItcRows").innerHTML = "";
    if ($("#recon2bGrossNetRows")) $("#recon2bGrossNetRows").innerHTML = "";
    ["reconItcGstr2b", "reconItcGstr3b", "reconItcTally", "reconItcAvailable"].forEach((id) => {
      if ($(`#${id}`)) $(`#${id}`).textContent = "0.00";
    });
  } else if (kind === "GSTR-1") {
    gstReconGstr1Rows = [];
    gstReconSalesResults = [];
    salesReconDashboard = null;
    gstDatasets["GSTR-1"] = undefined;
    gstReconDatasetsLoaded.delete("gstr1");
    gstReconDatasetsLoaded.delete("gstr1_results");
    if ($("#salesReconCards")) $("#salesReconCards").classList.add("hidden");
    if ($("#salesFyPeriodPanel")) $("#salesFyPeriodPanel").classList.add("hidden");
    if ($("#salesDocTypePanel")) $("#salesDocTypePanel").classList.add("hidden");
    if ($("#salesOutputSummaryPanel")) $("#salesOutputSummaryPanel").classList.add("hidden");
    if ($("#salesReconRows")) $("#salesReconRows").innerHTML = "";
    if ($("#salesFyPeriodRows")) $("#salesFyPeriodRows").innerHTML = "";
    if ($("#salesDocTypeRows")) $("#salesDocTypeRows").innerHTML = "";
  } else if (kind === "GSTR-3B") {
    gstr3bDashboard = null;
    gstDatasets["GSTR-3B"] = undefined;
    gstReconDatasetsLoaded.delete("gstr3b");
    populateGstr3bPeriodSelector([], getGstReconPeriod());
    ["gstr3bComparePanel", "gstr3bLiabilityPanel", "gstr3bItcPanel", "gstr3bUtilPanel", "gstr3bOutwardDrillPanel"].forEach((id) => {
      const node = document.getElementById(id);
      if (node) node.classList.add("hidden");
    });
    if ($("#gstr3bReconCards")) $("#gstr3bReconCards").classList.add("hidden");
    if ($("#gstr3bCompareRows")) $("#gstr3bCompareRows").innerHTML = "";
    if ($("#gstr3bLiabilityRows")) $("#gstr3bLiabilityRows").innerHTML = "";
    if ($("#gstr3bItcRows")) $("#gstr3bItcRows").innerHTML = "";
    if ($("#gstr3bUtilRows")) $("#gstr3bUtilRows").innerHTML = "";
    if ($("#gstr3bOutwardBooksRows")) $("#gstr3bOutwardBooksRows").innerHTML = "";
    ["g3bCardOutput", "g3bCardAvailable", "g3bCardClaimed", "g3bCardNet", "g3bCardInterest", "g3bCardLate", "g3bCardTotalCash", "g3bCardDiff"].forEach((id) => {
      if ($(`#${id}`)) $(`#${id}`).textContent = "Not Imported";
    });
    renderGstr3bCashPayableBreakdown(null, {
      panel: "#gstr3bCashBreakdownPanel",
      rows: "#gstr3bCashBreakdownRows",
      totals: "#gstr3bCashBreakdownTotals",
      showWhenReady: false,
    });
    renderGstr3bCashPayableBreakdown(null, {
      panel: "#overview3bCashBreakdown",
      rows: "#overview3bCashRows",
      totals: "#overview3bCashTotals",
      showWhenReady: false,
    });
    if ($("#gstr3bReconAdvice")) $("#gstr3bReconAdvice").textContent = "GSTR-3B not imported. Import portal GSTR-3B to reconcile liability, ITC and net payable.";
    if ($("#reconGstr3bPeriodStatus")) $("#reconGstr3bPeriodStatus").textContent = "GSTR-3B: Not loaded";
    if ($("#recon3bStatus")) $("#recon3bStatus").textContent = "GSTR-3B: Not loaded";
  }
  renderGstReconStatus();
  renderSalesReconStatus();
  updateGstReconReady();
  updateSalesReconReady();
}

async function clearGstReconSession(preferredReturnType = "") {
  const returnType = resolveGstClearReturnType(preferredReturnType);
  if (!confirm(`Clear imported ${returnType} data for FY 2025-26? Tally sync and other GST returns stay unchanged.`)) return;
  const response = await fetch("/api/gst/recon/clear", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      returnType,
      gstin: getGstPortalGstin(),
      financialYear: "2025-26",
      returnPeriod: getGstReconPeriod(),
      tolerance: Number($("#reconOverviewTolerance")?.value || $("#reconTolerance")?.value || 1),
    }),
  });
  const result = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = $("#gstError");
    if (error) {
      error.textContent = result.error || `Could not clear ${returnType}.`;
      error.classList.remove("hidden");
    }
    return;
  }
  if (result.gstin) gstPortalContext.gstin = String(result.gstin).toUpperCase();
  clearGstReconClientStateFor(result.return_type || returnType);
  if (result.overview) renderGstReconOverview(result.overview);
  else await refreshGstReconOverview(false);
}

function bindGstClearButtons() {
  document.querySelectorAll("[data-clear-return]").forEach((button) => {
    button.onclick = () => clearGstReconSession(button.getAttribute("data-clear-return") || "");
  });
  document.querySelectorAll("[data-purchase-clear]").forEach((button) => {
    button.onclick = () => clearPurchasePortalReturn(button.getAttribute("data-purchase-clear") || "");
  });
  document.querySelectorAll(".purchase-dash-card").forEach((button) => {
    button.onclick = () => {
      const key = button.dataset.purchaseFilter || "";
      purchaseDashFilter = purchaseDashFilter === key ? "" : key;
      renderPurchaseReconciliation();
    };
  });
  ["purchaseOrigMonthFilter", "purchase2bMonthFilter"].forEach((id) => {
    const node = document.getElementById(id);
    if (node) node.onchange = () => renderPurchaseReconciliation();
  });
  if ($("#purchaseClearPeriodFilters")) {
    $("#purchaseClearPeriodFilters").onclick = () => {
      if ($("#purchaseOrigMonthFilter")) $("#purchaseOrigMonthFilter").value = "";
      if ($("#purchase2bMonthFilter")) $("#purchase2bMonthFilter").value = "";
      renderPurchaseReconciliation();
    };
  }
  ["gstMismatchSearch"].forEach((id) => {
    const node = document.getElementById(id);
    if (node) node.oninput = debounce(() => {
      purchaseReconPages.mismatch = 1;
      renderPurchaseReconciliation();
    }, 250);
  });
}

function setReconTab(tab, options = {}) {
  activeReconTab = tab || "overview";
  document.querySelectorAll(".recon-tab").forEach(button => {
    button.classList.toggle("active", button.dataset.reconTab === activeReconTab);
  });
  if ($("#reconOverviewPane")) $("#reconOverviewPane").classList.toggle("hidden", activeReconTab !== "overview");
  if ($("#reconPurchasePane")) $("#reconPurchasePane").classList.toggle("hidden", activeReconTab !== "purchase");
  if ($("#reconSalesPane")) $("#reconSalesPane").classList.toggle("hidden", activeReconTab !== "sales");
  if ($("#reconGstr3bPane")) $("#reconGstr3bPane").classList.toggle("hidden", activeReconTab !== "gstr3b");
  if ($("#reconPaymentPane")) $("#reconPaymentPane").classList.toggle("hidden", activeReconTab !== "payment");
  ensureGstReconPanelVisible();
  if (options.refresh === false) return;
  if (activeReconTab === "payment") {
    setGstPaySubTab(activeGstPayTab || "summary");
    refreshGstPaymentLedgerStatus(false);
  }
  // Lazy-load heavy datasets only when the tab needs them.
  ensureGstReconDatasetsForModule("threeway", { portalAllowed: true, tab: activeReconTab })
    .then(() => {
      ensureGstReconPanelVisible();
      if (activeReconTab === "overview") return refreshGstReconOverview(false);
      if (activeReconTab === "sales") {
        renderSalesReconStatus();
        updateSalesReconReady();
        if (!gstReconSalesResults.length || !salesReconDashboard) return refreshSalesReconDashboard(false);
        renderSalesReconDashboard(salesReconDashboard);
        renderSalesReconTable();
        return null;
      }
      if (activeReconTab === "gstr3b") {
        if (!gstr3bDashboard) return refreshGstr3bDashboard(false);
        renderGstr3bDashboard(gstr3bDashboard);
        return null;
      }
      if (activeReconTab === "purchase") {
        renderGstReconStatus();
        if (gstReconRows.length) return refreshItcDashboard(false);
      }
      return null;
    })
    .catch(() => { ensureGstReconPanelVisible(); });
}

let activeGstPayTab = "summary";
let gstPaymentLedgerStatus = null;

function setGstPaySubTab(tab) {
  activeGstPayTab = tab || "summary";
  document.querySelectorAll(".recon-subtab").forEach((button) => {
    button.classList.toggle("active", button.dataset.payTab === activeGstPayTab);
  });
  const map = {
    summary: "#gstPaySummaryPane",
    cash: "#gstPayCashPane",
    itc: "#gstPayItcPane",
    recon: "#gstPayReconPane",
    adjust: "#gstPayAdjustPane",
  };
  Object.entries(map).forEach(([key, selector]) => {
    const node = $(selector);
    if (node) node.classList.toggle("hidden", activeGstPayTab !== key);
  });
}

function gstPayChip(el, ok, label) {
  if (!el) return;
  el.textContent = label;
  el.classList.toggle("ok", !!ok);
  el.classList.toggle("warn", !ok);
}

function gstPayMoney(value, blank = false) {
  if (value === null || value === undefined || value === "") return blank ? "—" : "0.00";
  return Number(value || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function gstPayDataset(datasets, ...keys) {
  for (const key of keys) {
    if (datasets && datasets[key]) return datasets[key];
  }
  return {};
}

function renderGstPaymentLedgerTables(status) {
  const money = gstPayMoney;
  const cashRows = ((status.cash_ledger || {}).rows) || [];
  const cashBody = $("#gstPayCashLedgerRows");
  if (cashBody) {
    cashBody.innerHTML = cashRows.length
      ? cashRows.map((row) => {
          const bad = row.balance_mismatch;
          const close = bad
            ? `<span class="gst-pay-mismatch">${money(row.closing_balance_calculated)} (portal ${money(row.closing_balance_portal)}; diff ${money(row.closing_difference)})</span>`
            : money(row.closing_balance_portal);
          return `<tr class="${bad ? "gst-pay-row-mismatch" : ""}">
            <td>${escapeHtml(row.period_label || row.period || "")}</td>
            <td class="money">${money(row.opening_balance)}</td>
            <td class="money">${money(row.cash_deposited)}</td>
            <td class="money">${money(row.other_credits)}</td>
            <td class="money">${money(row.tax_utilised)}</td>
            <td class="money">${money(row.interest_utilised)}</td>
            <td class="money">${money(row.late_fee_utilised)}</td>
            <td class="money">${money(row.penalty_utilised)}</td>
            <td class="money">${money(row.other_debit)}</td>
            <td class="money">${close}</td>
          </tr>`;
        }).join("")
      : `<tr><td colspan="10">Import Electronic Cash Ledger to begin.</td></tr>`;
  }
  const cashCompare = $("#gstPayCashCompareRows");
  if (cashCompare) {
    const detailRows = [];
    cashRows.forEach((row) => {
      const heads = row.head_detail || {};
      Object.entries(heads).forEach(([head, comps]) => {
        detailRows.push(`<tr>
          <td>${escapeHtml((row.period_label || row.period || "") + " · " + head.toUpperCase())}</td>
          <td class="money">${money((comps || {}).tax)}</td>
          <td class="money">${money((comps || {}).interest)}</td>
          <td class="money">${money((comps || {}).penalty)}</td>
          <td class="money">${money((comps || {}).fee)}</td>
          <td class="money">${money((comps || {}).others)}</td>
          <td class="money">${money((comps || {}).total)}</td>
        </tr>`);
      });
    });
    cashCompare.innerHTML = detailRows.length
      ? detailRows.join("")
      : `<tr><td colspan="7">Head-wise Tax / Interest / Penalty / Fee / Others appears after Cash Ledger import.</td></tr>`;
  }

  const itcRows = ((status.itc_ledger || {}).rows) || [];
  const itcBody = $("#gstPayItcLedgerRows");
  if (itcBody) {
    itcBody.innerHTML = itcRows.length
      ? itcRows.map((row) => `<tr class="${row.balance_mismatch ? "gst-pay-row-mismatch" : ""}">
          <td>${escapeHtml(row.period_label || row.period || "")}</td>
          <td class="money">${money(row.opening_igst)}</td>
          <td class="money">${money(row.opening_cgst)}</td>
          <td class="money">${money(row.opening_sgst)}</td>
          <td class="money">${money(row.opening_cess)}</td>
          <td class="money">${money(row.itc_credit_igst)}</td>
          <td class="money">${money(row.itc_credit_cgst)}</td>
          <td class="money">${money(row.itc_credit_sgst)}</td>
          <td class="money">${money(row.itc_credit_cess)}</td>
          <td class="money">${money(row.itc_utilised_igst)}</td>
          <td class="money">${money(row.itc_utilised_cgst)}</td>
          <td class="money">${money(row.itc_utilised_sgst)}</td>
          <td class="money">${money(row.itc_utilised_cess)}</td>
          <td class="money">${money(row.closing_igst)}</td>
          <td class="money">${money(row.closing_cgst)}</td>
          <td class="money">${money(row.closing_sgst)}</td>
          <td class="money">${money(row.closing_cess)}</td>
          <td class="money">${money(row.closing_total)}</td>
        </tr>`).join("")
      : `<tr><td colspan="18">Import Electronic Credit / ITC Ledger to begin.</td></tr>`;
  }
  const matrixNote = status.itc_cross_head_note
    || ((status.itc_ledger || {}).matrix_note)
    || "Cross-head utilisation breakup not available in Electronic Credit Ledger source";
  const matrixBody = $("#gstPayItcMatrixRows");
  if (matrixBody) {
    matrixBody.innerHTML = `<tr><td colspan="5">${escapeHtml(matrixNote)}</td></tr>`;
  }

  const reconRows = ((status.payment_recon || {}).rows) || [];
  const reconBody = $("#gstPayReconRows");
  if (reconBody) {
    reconBody.innerHTML = reconRows.length
      ? reconRows.map((row) => `<tr>
          <td>${escapeHtml(row.tax_period || row.period_key || "")}</td>
          <td>${escapeHtml(row.deposit_date || row.matched_cash_date || "")}</td>
          <td>${escapeHtml(row.cpin || row.reference_no || "")}</td>
          <td class="money">—</td>
          <td class="money">—</td>
          <td class="money">—</td>
          <td class="money">—</td>
          <td class="money">—</td>
          <td class="money">—</td>
          <td class="money">${money(row.amount ?? row.matched_cash_total, true)}</td>
          <td class="money">—</td>
          <td class="money">${money(row.difference, true)}</td>
          <td>${escapeHtml(row.match_status || row.status || "")}</td>
        </tr>`).join("")
      : `<tr><td colspan="13">No payment reconciliation rows yet. Import Challan History and Cash Ledger.</td></tr>`;
  }

  const g3Rows = ((status.gstr3b_link || {}).rows) || [];
  const g3Body = $("#gstPayGstr3bLinkRows");
  if (g3Body) {
    const visible = g3Rows.filter((row) => row.has_gstr3b || row.cash_tax_debit || row.credit_itc_debit);
    g3Body.innerHTML = visible.length
      ? visible.map((row) => `<tr>
          <td>${escapeHtml(row.period_label || row.period || "")}</td>
          <td class="money">${money(row.gstr3b_tax_payable)}</td>
          <td class="money">${money(row.cash_tax_debit)}</td>
          <td class="money">${money(row.credit_itc_debit)}</td>
          <td class="money">${money(row.tax_difference)}</td>
          <td class="money">${money(row.gstr3b_interest)}</td>
          <td class="money">${money(row.cash_interest_debit)}</td>
          <td class="money">${money(row.interest_difference)}</td>
          <td class="money">${money(row.gstr3b_late_fee)}</td>
          <td class="money">${money(row.cash_fee_debit)}</td>
          <td class="money">${money(row.late_fee_difference)}</td>
        </tr>`).join("")
      : `<tr><td colspan="11">Import Cash/Credit ledgers and GSTR-3B to compare liability vs utilisation by component.</td></tr>`;
  }

  const adjustRows = ((status.tally_adjustments || {}).rows) || [];
  const tallyRecon = ((status.tally_adjustments || {}).tally_recon) || [];
  const adjustBody = $("#gstPayAdjustRows");
  if (adjustBody) {
    const lines = [
      ...adjustRows.map((row) => `<tr>
        <td>${escapeHtml(row.issue || "")}</td>
        <td>${escapeHtml(row.period_ref || "")}</td>
        <td class="money">${money(row.portal_amount, true)}</td>
        <td class="money">${money(row.tally_amount, true)}</td>
        <td class="money">${money(row.difference, true)}</td>
        <td>${escapeHtml(row.preview_hint || "preview only")}</td>
      </tr>`),
      ...tallyRecon.map((row) => `<tr>
        <td>${escapeHtml(row.particulars || "")}</td>
        <td>—</td>
        <td class="money">${money(row.portal, true)}</td>
        <td class="money">${money(row.tally, true)}</td>
        <td class="money">${money(row.difference, true)}</td>
        <td>${escapeHtml(row.status || "")}</td>
      </tr>`),
    ];
    adjustBody.innerHTML = lines.length
      ? lines.join("")
      : `<tr><td colspan="6">No adjustment candidates yet.</td></tr>`;
  }
}

function renderGstPaymentLedgerStatus(status) {
  gstPaymentLedgerStatus = status || gstPaymentLedgerStatus || {};
  const datasets = gstPaymentLedgerStatus.datasets || {};
  const payment = gstPayDataset(datasets, "challan_history", "GST_PAYMENT_LIST");
  const cash = gstPayDataset(datasets, "cash_ledger", "GST_CASH_LEDGER");
  const itc = gstPayDataset(datasets, "credit_ledger", "GST_ITC_LEDGER");
  gstPayChip(
    $("#gstPayStatusPayment"),
    payment.imported,
    payment.imported
      ? `Challan History: Imported (${Number(payment.record_count || payment.row_count || 0).toLocaleString("en-IN")} rows)`
      : "Challan History: Not Imported"
  );
  gstPayChip(
    $("#gstPayStatusCash"),
    cash.imported,
    cash.imported
      ? `Cash Ledger: Imported (${Number(cash.record_count || cash.row_count || 0).toLocaleString("en-IN")} rows)`
      : "Cash Ledger: Not Imported"
  );
  gstPayChip(
    $("#gstPayStatusItc"),
    itc.imported,
    itc.imported
      ? `ITC/Credit Ledger: Imported (${Number(itc.record_count || itc.row_count || 0).toLocaleString("en-IN")} rows)`
      : "ITC/Credit Ledger: Not Imported"
  );
  gstPayChip(
    $("#gstPayStatus3b"),
    gstPaymentLedgerStatus.gstr3b_imported,
    gstPaymentLedgerStatus.gstr3b_imported ? "GSTR-3B: Imported" : "GSTR-3B: Not Imported"
  );
  gstPayChip(
    $("#gstPayStatusTally"),
    gstPaymentLedgerStatus.tally_connected,
    gstPaymentLedgerStatus.tally_connected ? "Tally: Connected" : "Tally: Not Connected"
  );
  const noteFor = (keys, noteId) => {
    const item = gstPayDataset(datasets, ...keys);
    const note = $(noteId);
    if (!note) return;
    if (!item.imported) {
      note.textContent = "Not imported";
      return;
    }
    const validation = item.validation || {};
    const stats = item.stats || {};
    const parts = [
      item.file_name ? `Imported: ${item.file_name}` : "Imported",
      `${Number(item.record_count || item.row_count || 0).toLocaleString("en-IN")} rows`,
    ];
    if (validation.detected_gstin || validation.from_date || validation.to_date) {
      parts.push(`GSTIN ${validation.detected_gstin || "—"}`);
      parts.push(`From ${validation.from_date || "—"}`);
      parts.push(`To ${validation.to_date || "—"}`);
    }
    if (stats.credit_rows != null || stats.debit_rows != null) {
      parts.push(`CR ${stats.credit_rows ?? "—"} / DR ${stats.debit_rows ?? "—"}`);
    }
    if (validation.status) parts.push(`Validation ${validation.status}`);
    note.textContent = parts.join(" · ");
  };
  noteFor(["challan_history", "GST_PAYMENT_LIST"], "#gstPayPaymentImportNote");
  noteFor(["cash_ledger", "GST_CASH_LEDGER"], "#gstPayCashImportNote");
  noteFor(["credit_ledger", "GST_ITC_LEDGER"], "#gstPayItcImportNote");

  // Prefer detailed challan merge stats when available.
  const paymentMeta = payment.meta || {};
  const mergeStats = paymentMeta.merge_stats || payment.stats || {};
  if (payment.imported && $("#gstPayPaymentImportNote") && (mergeStats.unique_challans != null || mergeStats.files_processed != null)) {
    $("#gstPayPaymentImportNote").textContent =
      `Files ${mergeStats.files_processed || 1} · Rows read ${mergeStats.rows_read || payment.record_count || 0} · Unique ${mergeStats.unique_challans || payment.record_count || 0} · Duplicates skipped ${mergeStats.duplicates_skipped || 0} · PAID ${mergeStats.paid_challans ?? "—"} · FAILED ${mergeStats.failed_challans ?? "—"} · PAID amount ${gstPayMoney(mergeStats.total_paid_amount)}`;
  }

  const cards = gstPaymentLedgerStatus.summary_cards || {};
  const setCard = (id, value) => {
    if ($(`#${id}`)) $(`#${id}`).textContent = cards.calculated ? gstPayMoney(value, value == null) : "—";
  };
  setCard("gstPayCardLiability", cards.gst_liability);
  setCard("gstPayCardItcUtil", cards.itc_utilised);
  setCard("gstPayCardCashUtil", cards.cash_utilised);
  setCard("gstPayCardTaxCash", cards.gst_tax_paid_cash);
  setCard("gstPayCardInterest", cards.interest_paid);
  setCard("gstPayCardLate", cards.late_fee_paid);
  setCard("gstPayCardTotalCash", cards.total_cash_paid);
  setCard("gstPayCardCashClose", cards.closing_cash_ledger_balance);
  setCard("gstPayCardItcClose", cards.closing_itc_balance);

  const challanNote = $("#gstPayChallanSummaryNote");
  if (challanNote) {
    const cs = gstPaymentLedgerStatus.challan_summary || {};
    challanNote.textContent = payment.imported
      ? `Challans ${cs.total_challans || 0} · Paid ${cs.paid_challans || 0} · Failed ${cs.failed_challans || 0} · Paid amount ${gstPayMoney(cs.total_paid_amount)} (deposits only, not utilisation)`
      : "Challan summary appears after Challan History import.";
  }
  renderGstPaymentLedgerTables(gstPaymentLedgerStatus);
}

async function refreshGstPaymentLedgerStatus(showError = false) {
  const error = $("#gstError");
  try {
    const response = await fetch("/api/gst/recon/payment/status", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        gstin: getGstPortalGstin(),
        financialYear: $("#gstPayFinancialYear")?.value || "2025-26",
      }),
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error || "Could not load GST Payment & Ledger status.");
    renderGstPaymentLedgerStatus(data);
  } catch (failure) {
    if (showError && error) {
      error.textContent = failure.message || String(failure);
      error.classList.remove("hidden");
    }
  }
}

async function fileToBase64Payload(fileInput) {
  const file = fileInput?.files?.[0];
  if (!file) throw new Error("Select a file first.");
  return fileObjectToBase64Payload(file);
}

async function fileObjectToBase64Payload(file) {
  const buffer = await file.arrayBuffer();
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return { name: file.name, data: btoa(binary) };
}

async function filesToBase64Payloads(fileInput) {
  const files = Array.from(fileInput?.files || []);
  if (!files.length) throw new Error("Select a file first.");
  const packed = [];
  for (const file of files) {
    packed.push(await fileObjectToBase64Payload(file));
  }
  return packed;
}

function updateGstPayPaymentFileSelection() {
  const input = $("#gstPayPaymentFile");
  const note = $("#gstPayPaymentFileSelection");
  if (!note) return;
  const files = Array.from(input?.files || []);
  if (!files.length) {
    note.textContent = "No files selected";
    return;
  }
  if (files.length === 1) {
    note.textContent = `1 file selected: ${files[0].name}`;
    return;
  }
  const names = files.map((f) => f.name).join(", ");
  note.textContent = `${files.length} files selected: ${names}`;
}

async function importGstPaymentLedgerDataset(dataType, fileInputId, showError = true) {
  const error = $("#gstError");
  if (showError && error) error.classList.add("hidden");
  try {
    const input = $(fileInputId);
    const isChallan = ["GST_PAYMENT_LIST", "challan_history", "payment_list"].includes(String(dataType || ""));
    const packedFiles = isChallan
      ? await filesToBase64Payloads(input)
      : [await fileToBase64Payload(input)];
    const response = await fetch("/api/gst/recon/payment/import", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        dataType,
        gstin: getGstPortalGstin(),
        financialYear: $("#gstPayFinancialYear")?.value || "2025-26",
        files: packedFiles,
      }),
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok) {
      const detail = result.detail ? ` — ${result.detail}` : "";
      throw new Error((result.error || `Could not import ${dataType}.`) + detail);
    }
    if (result.gstin) gstPortalContext.gstin = String(result.gstin).toUpperCase();
    if (result.status) renderGstPaymentLedgerStatus(result.status);
    else await refreshGstPaymentLedgerStatus(false);
    if (result.merge_stats && $("#gstPayPaymentImportNote")) {
      const ms = result.merge_stats;
      $("#gstPayPaymentImportNote").textContent =
        `Files ${ms.files_processed || packedFiles.length} · Rows read ${ms.rows_read || 0} · Unique ${ms.unique_challans || 0} · Duplicates skipped ${ms.duplicates_skipped || 0} · PAID ${ms.paid_challans || 0} · FAILED ${ms.failed_challans || 0} · PAID amount ${gstPayMoney(ms.total_paid_amount)}`;
    } else if (result.message && error) {
      error.textContent = result.message + (result.parse_error ? ` (${result.parse_error})` : "");
      error.classList.toggle("hidden", !result.format_pending && !result.parse_error);
    }
    return result;
  } catch (failure) {
    if (showError && error) {
      error.textContent = failure.message || String(failure);
      error.classList.remove("hidden");
    }
    throw failure;
  }
}

async function clearGstPaymentLedgerDataset(dataType) {
  if (!confirm(`Clear imported ${dataType.replace(/_/g, " ")} for FY 2025-26? GSTR-1/2B/3B and Tally data stay unchanged.`)) return;
  const error = $("#gstError");
  const response = await fetch("/api/gst/recon/payment/clear", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      dataType,
      gstin: getGstPortalGstin(),
      financialYear: $("#gstPayFinancialYear")?.value || "2025-26",
    }),
  });
  const result = await response.json().catch(() => ({}));
  if (!response.ok) {
    if (error) {
      error.textContent = result.error || `Could not clear ${dataType}.`;
      error.classList.remove("hidden");
    }
    return;
  }
  if (result.status) renderGstPaymentLedgerStatus(result.status);
  else await refreshGstPaymentLedgerStatus(false);
}

function bindGstPaymentLedgerControls() {
  document.querySelectorAll(".recon-subtab").forEach((button) => {
    button.onclick = () => setGstPaySubTab(button.dataset.payTab || "summary");
  });
  if ($("#gstPayImportPaymentBtn")) {
    $("#gstPayImportPaymentBtn").onclick = () => importGstPaymentLedgerDataset("GST_PAYMENT_LIST", "#gstPayPaymentFile");
  }
  if ($("#gstPayPaymentFile")) {
    $("#gstPayPaymentFile").onchange = updateGstPayPaymentFileSelection;
    updateGstPayPaymentFileSelection();
  }
  if ($("#gstPayImportCashBtn")) {
    $("#gstPayImportCashBtn").onclick = () => importGstPaymentLedgerDataset("GST_CASH_LEDGER", "#gstPayCashFile");
  }
  if ($("#gstPayImportItcBtn")) {
    $("#gstPayImportItcBtn").onclick = () => importGstPaymentLedgerDataset("GST_ITC_LEDGER", "#gstPayItcFile");
  }
  document.querySelectorAll("[data-pay-clear]").forEach((button) => {
    button.onclick = () => clearGstPaymentLedgerDataset(button.getAttribute("data-pay-clear") || "");
  });
  if ($("#gstPayFinancialYear")) {
    $("#gstPayFinancialYear").onchange = () => refreshGstPaymentLedgerStatus(false);
  }
  if ($("#gstPayPrepareEntryBtn")) {
    $("#gstPayPrepareEntryBtn").onclick = () => {
      const panel = $("#gstPayAdjustPreviewPanel");
      if (panel) panel.classList.remove("hidden");
      const cards = (gstPaymentLedgerStatus && gstPaymentLedgerStatus.summary_cards) || {};
      const setVal = (id, value) => {
        if ($(id)) $(id).value = value == null || value === "" ? "" : String(value);
      };
      setVal("#gstPayPreviewVoucherType", "Payment");
      setVal("#gstPayPreviewIgst", cards.gst_tax_paid_cash != null ? "" : "");
      setVal("#gstPayPreviewInterest", cards.interest_paid);
      setVal("#gstPayPreviewLate", cards.late_fee_paid);
      setVal("#gstPayPreviewTotal", cards.total_cash_paid);
      if ($("#gstPayPreviewNarration")) {
        $("#gstPayPreviewNarration").value =
          "PREVIEW ONLY — Portal cash util Tax/Interest/Late Fee/Penalty/Others. No Tally voucher is posted without explicit confirmation.";
      }
    };
  }
}

function renderSalesReconStatus() {
  const g1 = gstReconGstr1Rows.length || gstDatasets["GSTR-1"]?.length || 0;
  const tally = gstReconTallySalesRows.length;
  if ($("#reconGstr1Status")) {
    $("#reconGstr1Status").textContent = g1
      ? `GSTR-1: ${g1.toLocaleString("en-IN")} rows loaded`
      : "GSTR-1: Not loaded";
  }
  if ($("#reconTallySalesStatus")) {
    $("#reconTallySalesStatus").textContent = tally
      ? `Tally Sales: ${tally.toLocaleString("en-IN")} vouchers synced`
      : "Tally Sales: Not synced";
  }
}

function updateSalesReconReady() {
  const ready = Boolean(gstReconGstr1Rows.length && gstReconTallySalesRows.length);
  if ($("#reconGstr1TallyBtn")) $("#reconGstr1TallyBtn").disabled = !ready;
  if ($("#reconSalesDashBtn")) $("#reconSalesDashBtn").disabled = !gstReconGstr1Rows.length;
}

function renderSalesFyPeriodBreakdown(dashboard) {
  const panel = $("#salesFyPeriodPanel");
  const rowsEl = $("#salesFyPeriodRows");
  const totalsEl = $("#salesFyPeriodTotals");
  const netEl = $("#salesFyNetOutputGst");
  if (!panel || !rowsEl) return;
  const breakdown = dashboard?.period_breakdown;
  const isFy = (dashboard?.period_mode === "fy_all") || String(dashboard?.return_period || "").toUpperCase() === "ALL";
  if (!isFy || !breakdown) {
    panel.classList.add("hidden");
    rowsEl.innerHTML = "";
    if (totalsEl) totalsEl.innerHTML = "";
    return;
  }
  const money = reconMoney;
  const months = breakdown.months || [];
  rowsEl.innerHTML = months.map(row => {
    const missing = !row.present;
    return `<tr class="${missing ? "fy-period-missing" : ""}">
      <td>${escapeHtml(row.period_label || row.period || "")}</td>
      <td>${Number(row.count || 0).toLocaleString("en-IN")}</td>
      <td class="money">${money(row.taxable_value)}</td>
      <td class="money">${money(row.igst)}</td>
      <td class="money">${money(row.cgst)}</td>
      <td class="money">${money(row.sgst)}</td>
      <td class="money">${money(row.cess)}</td>
      <td class="money">${money(row.output_gst)}</td>
    </tr>`;
  }).join("");
  const tot = breakdown.totals || {};
  if (totalsEl) {
    totalsEl.innerHTML = `<tr>
      <th>FY Total</th>
      <th>${Number(tot.count || 0).toLocaleString("en-IN")}</th>
      <th class="money">${money(tot.taxable_value)}</th>
      <th class="money">${money(tot.igst)}</th>
      <th class="money">${money(tot.cgst)}</th>
      <th class="money">${money(tot.sgst)}</th>
      <th class="money">${money(tot.cess)}</th>
      <th class="money">${money(tot.output_gst)}</th>
    </tr>`;
  }
  if (netEl) netEl.textContent = money(tot.output_gst);
  panel.classList.remove("hidden");
}

function renderSalesReconDashboard(dashboard) {
  salesReconDashboard = dashboard || salesReconDashboard;
  if (!salesReconDashboard) return;
  const cards = salesReconDashboard.cards || {};
  const money = reconMoney;
  if ($("#salesReconCards")) $("#salesReconCards").classList.remove("hidden");
  if ($("#salesCardGstr1Count")) $("#salesCardGstr1Count").textContent = Number(cards.gstr1_invoice_count || 0).toLocaleString("en-IN");
  if ($("#salesCardTallyCount")) $("#salesCardTallyCount").textContent = Number(cards.tally_sales_count || 0).toLocaleString("en-IN");
  if ($("#salesCardExact")) $("#salesCardExact").textContent = Number(cards.exact_match || 0).toLocaleString("en-IN");
  if ($("#salesCardMismatch")) $("#salesCardMismatch").textContent = Number(cards.mismatch || 0).toLocaleString("en-IN");
  if ($("#salesCardMissingTally")) $("#salesCardMissingTally").textContent = Number(cards.missing_in_tally || 0).toLocaleString("en-IN");
  if ($("#salesCardMissingGstr1")) $("#salesCardMissingGstr1").textContent = Number(cards.missing_in_gstr1 || 0).toLocaleString("en-IN");
  if ($("#salesCardGstr1Taxable")) $("#salesCardGstr1Taxable").textContent = money(cards.gstr1_taxable);
  if ($("#salesCardTallyTaxable")) $("#salesCardTallyTaxable").textContent = money(cards.tally_taxable);
  if ($("#salesCardOutputDiff")) $("#salesCardOutputDiff").textContent = money(cards.output_gst_difference);
  if ($("#salesCardPortalNetGst")) $("#salesCardPortalNetGst").textContent = money(cards.portal_net_gst);
  if ($("#salesCardTallyNetGst")) $("#salesCardTallyNetGst").textContent = money(cards.tally_net_gst);
  renderSalesFyPeriodBreakdown(salesReconDashboard);
  const docSummary = salesReconDashboard.document_summary;
  if (docSummary) {
    renderSignedDocumentTypeSummary(docSummary, {
      panelId: "salesDocTypePanel",
      rowsId: "salesDocTypeRows",
      netLabelId: "salesDocTypeNetLabel",
      portalNetId: "salesPortalNetGst",
      tallyNetId: "salesTallyNetGst",
      diffId: "salesNetGstDiff",
      statusId: "salesNetGstStatus",
    });
  }
  const summary = salesReconDashboard.output_summary || {};
  const g1 = summary.gstr1 || {};
  const tally = summary.tally || {};
  const diff = summary.difference || {};
  const lines = [
    ["Taxable Value", g1.taxable_value, tally.taxable_value, diff.taxable_value],
    ["Output IGST", g1.igst, tally.igst, diff.igst],
    ["Output CGST", g1.cgst, tally.cgst, diff.cgst],
    ["Output SGST", g1.sgst, tally.sgst, diff.sgst],
    ["Output CESS", g1.cess, tally.cess, diff.cess],
    ["Total Output GST", g1.output_gst, tally.output_gst, diff.output_gst],
  ];
  if ($("#salesOutputSummaryRows")) {
    $("#salesOutputSummaryRows").innerHTML = lines.map(([label, a, b, c]) =>
      `<tr><td>${label}</td><td class="money">${money(a)}</td><td class="money">${money(b)}</td><td class="money">${money(c)}</td></tr>`
    ).join("");
  }
  if ($("#salesOutputSummaryPanel")) $("#salesOutputSummaryPanel").classList.remove("hidden");
  if ($("#reconSalesViewDiffBtn")) {
    $("#reconSalesViewDiffBtn").disabled = !(Number(cards.gstr1_invoice_count || 0) || Number(cards.tally_sales_count || 0));
  }
  if ($("#salesReconMatchCounts")) {
    $("#salesReconMatchCounts").innerHTML = Object.entries(salesReconDashboard.counts || {}).map(
      ([name, count]) => `<span>${escapeHtml(name)}<strong>${Number(count).toLocaleString("en-IN")}</strong></span>`
    ).join("");
  }
  const advice = [];
  if (salesReconDashboard.period_mode === "fy_all") {
    const fyOut = salesReconDashboard.period_breakdown?.totals?.output_gst;
    if (fyOut != null) advice.push(`FY 2025-26 GSTR-1 Net Output GST ₹${money(fyOut)}.`);
  }
  if (cards.missing_in_tally) advice.push(`${cards.missing_in_tally} portal invoice(s) missing in Tally.`);
  if (cards.missing_in_gstr1) advice.push(`${cards.missing_in_gstr1} Tally sales voucher(s) missing in GSTR-1.`);
  if (Math.abs(Number(cards.output_gst_difference || 0)) > 1) advice.push(`Output GST / Net GST difference ₹${money(cards.output_gst_difference)}.`);
  if (cards.net_gst_matched) advice.push("Net GST matched with Tally.");
  if ($("#salesReconAdvice")) $("#salesReconAdvice").textContent = advice.join(" ") || "Review Exact Match and mismatch rows before filing.";
}

function filteredSalesReconRows() {
  const search = ($("#salesReconSearch")?.value || "").trim().toLowerCase();
  const status = $("#salesReconStatusFilter")?.value || "";
  const dateFilter = ($("#salesReconDateFilter")?.value || "").trim().toLowerCase();
  const gstinFilter = ($("#salesReconGstinFilter")?.value || "").trim().toLowerCase();
  const voucherFilter = $("#salesReconVoucherFilter")?.value || "";
  const sortBy = $("#salesReconSort")?.value || "status";
  const mismatch = new Set([
    "Value Difference", "Tax Difference", "Date Difference", "GSTIN Difference",
    "Invoice Number Difference", "Possible Match", "Duplicate",
  ]);
  let rows = [...(gstReconSalesResults || [])];
  rows = rows.filter(row => {
    if (status === "Mismatch" && !mismatch.has(row.status)) return false;
    if (status && status !== "Mismatch" && row.status !== status) return false;
    if (dateFilter && !String(row.invoice_date || "").toLowerCase().includes(dateFilter)) return false;
    if (gstinFilter && !String(row.gstin || "").toLowerCase().includes(gstinFilter)) return false;
    if (voucherFilter && !String(row.voucher_type || row.document_type || "").includes(voucherFilter)) return false;
    if (search) {
      const blob = `${row.gstin || ""} ${row.party_name || ""} ${row.invoice_no || ""} ${row.status || ""}`.toLowerCase();
      if (!blob.includes(search)) return false;
    }
    return true;
  });
  rows.sort((a, b) => {
    if (sortBy === "date") return String(a.invoice_date || "").localeCompare(String(b.invoice_date || ""));
    if (sortBy === "diff") return Number(b.total_difference || 0) - Number(a.total_difference || 0);
    if (sortBy === "party") return String(a.party_name || "").localeCompare(String(b.party_name || ""));
    return String(a.status || "").localeCompare(String(b.status || ""));
  });
  return rows;
}

function renderSalesReconTable() {
  if (!$("#salesReconRows")) return;
  const money = reconMoney;
  const rows = filteredSalesReconRows();
  const pages = Math.max(1, Math.ceil(rows.length / SALES_RECON_PAGE_SIZE));
  salesReconPage = Math.min(Math.max(1, salesReconPage), pages);
  const start = (salesReconPage - 1) * SALES_RECON_PAGE_SIZE;
  const pageRows = rows.slice(start, start + SALES_RECON_PAGE_SIZE);
  $("#salesReconRows").innerHTML = pageRows.map((row, index) => {
    const absIndex = start + index;
    const cls = row.status === "Exact Match" ? "matched" : "review";
    const taxableDiff = Number(row.tally_taxable || 0) - Number(row.gstr1_taxable || 0);
    return `<tr class="${cls}">
      <td><span class="gst-status ${cls}">${escapeHtml(row.status || "")}</span></td>
      <td>${escapeHtml(row.gstin || "")}</td>
      <td>${escapeHtml(row.party_name || "")}</td>
      <td>${escapeHtml(row.invoice_no || "")}</td>
      <td>${escapeHtml(row.invoice_date || "")}</td>
      <td>${escapeHtml(row.voucher_type || row.document_type || "")}</td>
      <td class="money">${money(row.gstr1_taxable)}</td>
      <td class="money">${money(row.tally_taxable)}</td>
      <td class="money">${money(taxableDiff)}</td>
      <td class="money">${money(row.gstr1_igst)}</td><td class="money">${money(row.tally_igst)}</td>
      <td class="money">${money(row.gstr1_cgst)}</td><td class="money">${money(row.tally_cgst)}</td>
      <td class="money">${money(row.gstr1_sgst)}</td><td class="money">${money(row.tally_sgst)}</td>
      <td class="money">${money(row.gstr1_cess)}</td><td class="money">${money(row.tally_cess)}</td>
      <td class="money">${money(row.total_difference)}</td>
      <td><select class="sales-review-action" data-index="${absIndex}">
        <option value="">Action</option>
        <option ${row.review_action === "Review" ? "selected" : ""}>Review</option>
        <option ${row.review_action === "Accept Match" ? "selected" : ""}>Accept Match</option>
        <option ${row.review_action === "Mark Corrected" ? "selected" : ""}>Mark Corrected</option>
        <option ${row.review_action === "Ignore" ? "selected" : ""}>Ignore</option>
        <option ${row.review_action === "View Tally Voucher" ? "selected" : ""}>View Tally Voucher</option>
        <option ${row.review_action === "View Portal Record" ? "selected" : ""}>View Portal Record</option>
      </select></td>
    </tr>`;
  }).join("") || `<tr><td colspan="19">Import GSTR-1 and sync Tally Sales, then reconcile.</td></tr>`;
  if ($("#salesReconPageLabel")) $("#salesReconPageLabel").textContent = `Page ${salesReconPage} / ${pages} (${rows.length} rows)`;
  document.querySelectorAll(".sales-review-action").forEach(select => {
    select.onchange = async () => {
      const filtered = filteredSalesReconRows();
      const row = filtered[Number(select.dataset.index)];
      if (!row || !select.value) return;
      if (select.value === "View Tally Voucher") {
        alert(row.tally ? JSON.stringify(row.tally, null, 2).slice(0, 1200) : "No Tally voucher linked.");
        return;
      }
      if (select.value === "View Portal Record") {
        alert(row.gstr1 ? JSON.stringify(row.gstr1, null, 2).slice(0, 1200) : "No portal record linked.");
        return;
      }
      try {
        const response = await fetch("/api/gst/recon/review-action", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            results: gstReconSalesResults,
            invoiceNo: row.invoice_no,
            gstin: row.gstin || "",
            action: select.value,
            returnPeriod: getGstReconPeriod(),
          }),
        });
        const result = await response.json();
        if (!response.ok) throw new Error(result.error || "Review action failed.");
        gstReconSalesResults = result.rows || [];
        renderSalesReconTable();
      } catch (failure) {
        alert(failure.message);
      }
    };
  });
}

async function syncGstReconTallySales() {
  const button = $("#reconTallySalesSyncBtn");
  const error = $("#gstError");
  button.disabled = true;
  button.textContent = "Syncing Sales...";
  error.classList.add("hidden");
  try {
    const response = await fetch("/api/gst/tally/sales-sync", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: "{}",
    });
    const result = await response.json();
    if (!response.ok || result.ok === false) {
      // Failed sync must not clear GSTR-1 or force zero Tally cards via recon.
      if (Array.isArray(result.rows) && result.preserved_existing) {
        gstReconTallySalesRows = result.rows;
      }
      throw new Error(result.error || "Tally Sales sync failed.");
    }
    gstReconTallySalesRows = result.rows || [];
    renderSalesReconStatus();
    updateSalesReconReady();
    if (gstReconGstr1Rows.length && gstReconTallySalesRows.length) {
      await refreshSalesReconDashboard(false);
    }
  } catch (failure) {
    error.textContent = failure.message;
    error.classList.remove("hidden");
    updateSalesReconReady();
  } finally {
    button.disabled = false;
    button.textContent = "Sync Tally Sales";
    updateSalesReconReady();
  }
}

async function reconcileGstReconGstr1Tally() {
  const button = $("#reconGstr1TallyBtn");
  const error = $("#gstError");
  if (!gstReconTallySalesRows.length) {
    error.textContent = "Sync Tally Sales successfully before reconciliation. GSTR-1 portal totals are unchanged.";
    error.classList.remove("hidden");
    return;
  }
  button.disabled = true;
  button.textContent = "Reconciling...";
  error.classList.add("hidden");
  try {
    const response = await fetch("/api/gst/recon/gstr1-tally", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        gstr1: gstReconGstr1Rows,
        tally_sales: gstReconTallySalesRows,
        tolerance: Number($("#salesReconTolerance")?.value || $("#reconTolerance")?.value || 1),
        returnPeriod: getGstReconPeriod(),
      }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "GSTR-1 vs Tally reconciliation failed.");
    gstReconSalesResults = result.rows || [];
    renderSalesReconDashboard(result.dashboard || null);
    renderSalesReconTable();
    loadSalesDifferenceRecon(false).catch(() => {});
    await refreshGstReconOverview(false);
  } catch (failure) {
    error.textContent = failure.message;
    error.classList.remove("hidden");
  } finally {
    button.disabled = false;
    button.textContent = "Reconcile GSTR-1 vs Tally";
    updateSalesReconReady();
  }
}

async function refreshSalesReconDashboard(showError = true) {
  const button = $("#reconSalesDashBtn");
  const error = $("#gstError");
  if (button) {
    button.disabled = true;
    button.textContent = "Refreshing...";
  }
  if (showError) error.classList.add("hidden");
  try {
    await ensureGstReconDatasetsForModule("threeway", { portalAllowed: true, tab: "sales" });
    const needRows = !gstReconSalesResults.length;
    const response = await fetch("/api/gst/recon/sales-dashboard", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        tolerance: Number($("#salesReconTolerance")?.value || 1),
        returnPeriod: getGstReconPeriod(),
        includeRows: needRows,
      }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "Sales dashboard refresh failed.");
    if (result.rows) gstReconSalesResults = result.rows;
    else if (needRows) {
      gstReconDatasetsLoaded.delete("gstr1_results");
      const bundle = await fetchGstReconDatasets(["gstr1_results"]);
      applyGstReconDatasetBundle(bundle, { portalAllowed: true });
    }
    renderSalesReconDashboard(result);
    renderSalesReconTable();
  } catch (failure) {
    if (showError) {
      error.textContent = failure.message;
      error.classList.remove("hidden");
    }
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = "Refresh Sales Dashboard";
      updateSalesReconReady();
    }
  }
}

function overviewValue(imported, value, asMoney = false) {
  if (!imported) return "Not Imported";
  if (value === null || value === undefined || value === "") return "Not Imported";
  return asMoney ? reconMoney(value) : Number(value || 0).toLocaleString("en-IN");
}

function setReconImportChip(node, ok, labelOk, labelBad) {
  if (!node) return;
  node.textContent = ok ? labelOk : labelBad;
  node.classList.toggle("ok", Boolean(ok));
  node.classList.toggle("warn", !ok);
}

function renderGstReconOverview(overview) {
  if (!overview) return;
  const purchase = overview.purchase || {};
  const sales = overview.sales || {};
  const g3 = overview.gstr3b || {};
  const status = overview.import_status || {};
  const g2bOk = Boolean(status.gstr2b_imported ?? purchase.imported);
  const g1Ok = Boolean(status.gstr1_imported ?? sales.imported);
  const g3Ok = Boolean(status.gstr3b_imported ?? g3.imported);
  const tallyOk = Boolean(status.tally_connected);

  setReconImportChip($("#reconStatusTally"), tallyOk, "✅ Tally Connected", "❌ Tally Not Connected");
  setReconImportChip($("#reconStatus2b"), g2bOk, "✅ GSTR-2B Imported", "❌ GSTR-2B Not Imported");
  setReconImportChip($("#reconStatus1"), g1Ok, "✅ GSTR-1 Imported", "❌ GSTR-1 Not Imported");
  setReconImportChip($("#reconStatus3b"), g3Ok, "✅ GSTR-3B Imported", "❌ GSTR-3B Not Imported");

  if ($("#overview2bStatusLabel")) $("#overview2bStatusLabel").textContent = g2bOk ? "GSTR-2B: Imported" : "GSTR-2B: Not Imported";
  if ($("#overviewGstr1StatusLabel")) $("#overviewGstr1StatusLabel").textContent = g1Ok ? "GSTR-1: Imported" : "GSTR-1: Not Imported";
  if ($("#overview3bStatusLabel")) $("#overview3bStatusLabel").textContent = g3Ok ? "GSTR-3B: Imported" : "GSTR-3B: Not Imported";

  if ($("#overview2bItc")) $("#overview2bItc").textContent = overviewValue(g2bOk, purchase.gstr2b_itc ?? purchase.gstr2b_available_itc, true);
  if ($("#overviewPurchaseMismatch")) $("#overviewPurchaseMismatch").textContent = overviewValue(g2bOk && purchase.ready_to_reconcile, purchase.purchase_mismatch);
  if ($("#overviewMissingPurchase")) $("#overviewMissingPurchase").textContent = overviewValue(g2bOk && purchase.ready_to_reconcile, purchase.missing_purchase);

  if ($("#overviewOutputGst")) $("#overviewOutputGst").textContent = overviewValue(g1Ok, sales.output_gst, true);
  if ($("#overviewGstr1Match")) $("#overviewGstr1Match").textContent = overviewValue(g1Ok && sales.ready_to_reconcile, sales.gstr1_match);
  if ($("#overviewGstr1Mismatch")) $("#overviewGstr1Mismatch").textContent = overviewValue(g1Ok && sales.ready_to_reconcile, sales.gstr1_mismatch);
  if ($("#overviewMissingSales")) $("#overviewMissingSales").textContent = overviewValue(g1Ok && sales.ready_to_reconcile, sales.missing_sales);

  if ($("#overview3bOutput")) $("#overview3bOutput").textContent = overviewValue(g3Ok && g3.ready_to_reconcile, g3.output_gst, true);
  if ($("#overview3bAvailable")) $("#overview3bAvailable").textContent = overviewValue(g3Ok && g3.ready_to_reconcile, g3.available_itc, true);
  if ($("#overview3bClaimed")) $("#overview3bClaimed").textContent = overviewValue(g3Ok && g3.ready_to_reconcile, g3.claimed_itc, true);
  if ($("#overview3bNet")) $("#overview3bNet").textContent = overviewValue(g3Ok && g3.ready_to_reconcile, g3.net_gst_payable, true);
  if ($("#overview3bInterest")) $("#overview3bInterest").textContent = overviewValue(g3Ok && g3.ready_to_reconcile, g3.net_gst_interest_payable ?? g3.interest, true);
  if ($("#overview3bLateFee")) $("#overview3bLateFee").textContent = overviewValue(g3Ok && g3.ready_to_reconcile, g3.net_gst_late_fee_payable ?? g3.late_fee, true);
  if ($("#overview3bTotalCash")) $("#overview3bTotalCash").textContent = overviewValue(g3Ok && g3.ready_to_reconcile, g3.total_cash_payable, true);
  if ($("#overview3bDiff")) $("#overview3bDiff").textContent = overviewValue(g3Ok && g3.ready_to_reconcile, g3.books_vs_3b_difference, true);
  renderGstr3bCashPayableBreakdown(
    g3.cash_payable_breakdown || (overview.gstr3b_dashboard || {}).cash_payable_breakdown,
    {
      panel: "#overview3bCashBreakdown",
      rows: "#overview3bCashRows",
      totals: "#overview3bCashTotals",
      showWhenReady: g3Ok && g3.ready_to_reconcile,
    }
  );
}

function renderGstr3bCashPayableBreakdown(breakdown, targets = {}) {
  const panel = $(targets.panel || "#gstr3bCashBreakdownPanel");
  const rowsEl = $(targets.rows || "#gstr3bCashBreakdownRows");
  const totalsEl = $(targets.totals || "#gstr3bCashBreakdownTotals");
  if (!panel || !rowsEl) return;
  const months = (breakdown && breakdown.months) || [];
  const anyPresent = months.some((row) => row && row.present);
  const show = targets.showWhenReady !== false && breakdown && breakdown.imported !== false && anyPresent;
  if (!show) {
    panel.classList.add("hidden");
    rowsEl.innerHTML = "";
    if (totalsEl) totalsEl.innerHTML = "";
    return;
  }
  const money = reconMoney;
  rowsEl.innerHTML = months.map((row) => {
    const missing = !row.present;
    const hasExtra = Number(row.interest_payable || 0) || Number(row.late_fee_payable || 0);
    return `<tr class="${missing ? "fy-period-missing" : ""}${hasExtra ? " gstr3b-cash-has-extra" : ""}">
      <td>${escapeHtml(row.period_label || row.period || "")}${missing ? " (missing)" : ""}</td>
      <td class="money">${money(row.gst_tax_payable)}</td>
      <td class="money">${money(row.interest_payable)}</td>
      <td class="money">${money(row.late_fee_payable)}</td>
      <td class="money">${money(row.total_cash_payable)}</td>
    </tr>`;
  }).join("");
  const totals = (breakdown && breakdown.totals) || {};
  if (totalsEl) {
    totalsEl.innerHTML = `<tr>
      <th>Total</th>
      <th class="money">${money(totals.gst_tax_payable)}</th>
      <th class="money">${money(totals.interest_payable)}</th>
      <th class="money">${money(totals.late_fee_payable)}</th>
      <th class="money">${money(totals.total_cash_payable)}</th>
    </tr>`;
  }
  panel.classList.remove("hidden");
}

function gstr3bTaxCells(bucket) {
  const b = bucket || {};
  return `<td class="money">${reconMoney(b.igst)}</td><td class="money">${reconMoney(b.cgst)}</td><td class="money">${reconMoney(b.sgst)}</td><td class="money">${reconMoney(b.cess)}</td><td class="money">${reconMoney(b.output_gst)}</td>`;
}

function populateGstr3bPeriodSelector(importedPeriods, selectedPeriod) {
  const select = $("#gstr3bReconPeriod");
  if (!select || select.tagName !== "SELECT") return;
  const labels = {
    ALL: "ALL / FY 2025-26",
    "042025": "Apr-25 (042025)",
    "052025": "May-25 (052025)",
    "062025": "Jun-25 (062025)",
    "072025": "Jul-25 (072025)",
    "082025": "Aug-25 (082025)",
    "092025": "Sep-25 (092025)",
    "102025": "Oct-25 (102025)",
    "112025": "Nov-25 (112025)",
    "122025": "Dec-25 (122025)",
    "012026": "Jan-26 (012026)",
    "022026": "Feb-26 (022026)",
    "032026": "Mar-26 (032026)",
  };
  const imported = new Set(
    (importedPeriods || [])
      .map((period) => String(period || "").replace(/\D/g, "").slice(0, 6))
      .filter(Boolean)
  );
  Array.from(select.options).forEach((option) => {
    const value = String(option.value || "");
    if (value === "ALL") {
      option.disabled = false;
      option.textContent = imported.size
        ? `ALL / FY 2025-26 (${imported.size} months)`
        : labels.ALL;
      return;
    }
    const has = imported.has(value);
    option.disabled = imported.size > 0 ? !has : false;
    option.textContent = has ? `${labels[value] || value} ✓` : (labels[value] || value);
  });
  const preferredRaw = String(selectedPeriod || select.value || "ALL").trim() || "ALL";
  const preferred = preferredRaw.toUpperCase() === "ALL" ? "ALL" : preferredRaw.replace(/\D/g, "").slice(0, 6);
  if (preferred === "ALL" || !imported.size || imported.has(preferred)) {
    select.value = preferred;
  } else {
    select.value = "ALL";
  }
}

async function ensureGstr3bOutwardDrilldown() {
  if (!gstr3bDashboard) return null;
  if (gstr3bDashboard.outward_classification_drilldown) {
    return gstr3bDashboard.outward_classification_drilldown;
  }
  const response = await fetch("/api/gst/recon/gstr3b-drilldown", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      returnPeriod: getGstReconPeriod(),
      tolerance: Number($("#gstr3bReconTolerance")?.value || $("#reconOverviewTolerance")?.value || 1),
    }),
  });
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || "Could not load GSTR-3B drilldown.");
  gstr3bDashboard.outward_classification_drilldown = result.outward_classification_drilldown || {};
  return gstr3bDashboard.outward_classification_drilldown;
}

async function renderGstr3bOutwardDrilldown(kind) {
  const panel = $("#gstr3bOutwardDrillPanel");
  if (!panel || !gstr3bDashboard) return;
  try {
    await ensureGstr3bOutwardDrilldown();
  } catch (_) {
    return;
  }
  const drill = gstr3bDashboard.outward_classification_drilldown || {};
  const key = kind === "nil_exempt" ? "nil_exempt" : "taxable";
  const block = drill[key] || {};
  const money = reconMoney;
  panel.classList.remove("hidden");
  if ($("#gstr3bOutwardDrillTitle")) {
    $("#gstr3bOutwardDrillTitle").textContent = `${block.label || key} — classification drill-down`;
  }
  const findings = (drill.findings || {}).summary || "";
  const notes = (drill.anomaly_notes || []).join(" ");
  if ($("#gstr3bOutwardDrillFinding")) {
    $("#gstr3bOutwardDrillFinding").textContent = [
      findings,
      notes && notes !== findings ? notes : "",
      block.portal_field ? `Portal field: ${block.portal_field}.` : "",
      `Books source: ${drill.books_source || "—"}.`,
      `Books ${key === "nil_exempt" ? "nil" : "taxable"} vouchers: ${Number(block.books_count || 0).toLocaleString("en-IN")} / ₹${money(block.books_total)}.`,
      `Portal total: ₹${money(block.portal_total)}.`,
    ].filter(Boolean).join(" ");
  }
  if ($("#gstr3bOutwardPortalRows")) {
    $("#gstr3bOutwardPortalRows").innerHTML = (drill.portal_period_breakdown || []).map((row) => {
      const anomalous = Boolean(row.anomaly_note);
      return `<tr class="${anomalous ? "difference-warning" : ""}">
        <td>${escapeHtml(row.period_label || row.period || "")}</td>
        <td>${escapeHtml(row.source_file || (row.present ? "—" : "missing"))}</td>
        <td class="money">${money(row.table_31a_taxable_value)}</td>
        <td class="money">${money(row.books_taxable_value)}</td>
        <td class="money">${money(row.taxable_difference_books_minus_31a)}</td>
        <td class="money">${money(row.table_31c_nil_exempt_taxable_value)}</td>
        <td class="money">${money(row.books_nil_exempt_taxable_value)}</td>
        <td class="money">${money(row.nil_difference_books_minus_31c)}</td>
      </tr>`;
    }).join("") || `<tr><td colspan="8">No GSTR-3B period rows.</td></tr>`;
  }
  if ($("#gstr3bOutwardBooksHeading")) {
    $("#gstr3bOutwardBooksHeading").textContent = key === "nil_exempt"
      ? `Books Nil/Exempt vouchers (${Number(block.books_count || 0)})`
      : `Books Taxable vouchers (${Number(block.books_count || 0)})`;
  }
  if ($("#gstr3bOutwardBooksRows")) {
    const rows = block.books_rows || [];
    const capped = rows.slice(0, 500);
    $("#gstr3bOutwardBooksRows").innerHTML = capped.map((row) => `<tr>
      <td>${escapeHtml(row.invoice_no || "")}</td>
      <td>${escapeHtml(row.invoice_date || "")}</td>
      <td>${escapeHtml(row.source_period || "")}</td>
      <td>${escapeHtml(row.party_name || "")}</td>
      <td>${escapeHtml(row.voucher_type || "")}</td>
      <td class="money">${money(row.gst_rate)}</td>
      <td class="money">${money(row.taxable_value)}</td>
      <td class="money">${money(row.output_gst)}</td>
      <td>${escapeHtml(row.classification || "")}</td>
    </tr>`).join("") || `<tr><td colspan="9">No Books vouchers in this class for the selected period.</td></tr>`;
    if (rows.length > capped.length) {
      $("#gstr3bOutwardBooksRows").insertAdjacentHTML(
        "beforeend",
        `<tr><td colspan="9">Showing first ${capped.length} of ${rows.length} vouchers.</td></tr>`
      );
    }
  }
}

function renderGstr3bDashboard(dashboard) {
  if (!dashboard || dashboard.imported === false || dashboard.status === "Not Imported") {
    clearGstReconClientStateFor("GSTR-3B");
    return;
  }
  gstr3bDashboard = dashboard;
  const cards = gstr3bDashboard.cards || {};
  const money = reconMoney;
  populateGstr3bPeriodSelector(
    gstr3bDashboard.imported_periods || (gstDatasets["GSTR-3B"] || {}).imported_periods || Object.keys((gstDatasets["GSTR-3B"] || {}).periods || {}),
    getGstReconPeriod()
  );
  if ($("#gstr3bReconCards")) $("#gstr3bReconCards").classList.remove("hidden");
  if ($("#g3bCardOutput")) $("#g3bCardOutput").textContent = money(cards.output_gst);
  if ($("#g3bCardAvailable")) $("#g3bCardAvailable").textContent = money(cards.available_itc);
  if ($("#g3bCardClaimed")) $("#g3bCardClaimed").textContent = money(cards.claimed_itc);
  if ($("#g3bCardNet")) $("#g3bCardNet").textContent = money(cards.net_gst_payable);
  if ($("#g3bCardInterest")) $("#g3bCardInterest").textContent = money(cards.net_gst_interest_payable ?? cards.interest);
  if ($("#g3bCardLate")) $("#g3bCardLate").textContent = money(cards.net_gst_late_fee_payable ?? cards.late_fee);
  if ($("#g3bCardTotalCash")) $("#g3bCardTotalCash").textContent = money(cards.total_cash_payable);
  if ($("#g3bCardDiff")) $("#g3bCardDiff").textContent = money(cards.books_vs_3b_difference);
  renderGstr3bCashPayableBreakdown(gstr3bDashboard.cash_payable_breakdown, {
    panel: "#gstr3bCashBreakdownPanel",
    rows: "#gstr3bCashBreakdownRows",
    totals: "#gstr3bCashBreakdownTotals",
    showWhenReady: true,
  });
  if ($("#gstr3bCompareRows")) {
    $("#gstr3bCompareRows").innerHTML = (gstr3bDashboard.rows || []).map(row => {
      const key = row.drilldown_key || "";
      const clickable = key === "taxable" || key === "nil_exempt";
      const label = clickable
        ? `<button type="button" class="linkish gstr3b-drill-link" data-drill="${escapeHtml(key)}">${escapeHtml(row.particulars || "")}</button>`
        : escapeHtml(row.particulars || "");
      return `<tr class="${Math.abs(Number(row.difference || 0)) > 1 ? "difference-warning" : "difference-ok"}"><td>${label}</td><td class="money">${money(row.books)}</td><td class="money">${money(row.gstr3b)}</td><td class="money">${money(row.difference)}</td></tr>`;
    }).join("");
    $("#gstr3bCompareRows").querySelectorAll(".gstr3b-drill-link").forEach((btn) => {
      btn.onclick = () => renderGstr3bOutwardDrilldown(btn.getAttribute("data-drill"));
    });
  }
  if ($("#gstr3bComparePanel")) $("#gstr3bComparePanel").classList.remove("hidden");
  const drill = gstr3bDashboard.outward_classification_drilldown || {};
  if ((drill.anomaly_notes || []).length) renderGstr3bOutwardDrilldown("nil_exempt");
  else if ($("#gstr3bOutwardDrillPanel")) $("#gstr3bOutwardDrillPanel").classList.add("hidden");
  const booksOut = ((gstr3bDashboard.books_liability || {}).books_output) || {};
  const portalOut = gstr3bDashboard.portal_outward || {};
  if ($("#gstr3bLiabilityRows")) {
    $("#gstr3bLiabilityRows").innerHTML = [
      ["Books Output (Tally)", booksOut],
      ["GSTR-3B / Portal Outward", portalOut],
    ].map(([label, bucket]) => `<tr><td>${label}</td>${gstr3bTaxCells(bucket)}</tr>`).join("");
  }
  if ($("#gstr3bLiabilityPanel")) $("#gstr3bLiabilityPanel").classList.remove("hidden");
  const itc = gstr3bDashboard.itc || {};
  if ($("#gstr3bItcRows")) {
    $("#gstr3bItcRows").innerHTML = [
      ["Available ITC", itc.available_itc],
      ["Claimed ITC", itc.claimed_itc],
      ["Eligible ITC", itc.eligible_itc],
      ["Ineligible ITC", itc.ineligible_itc],
      ["Pending ITC", itc.pending_itc],
      ["Reversed ITC", itc.reversed_itc],
      ["Unused ITC", itc.unused_itc],
    ].map(([label, bucket]) => `<tr><td>${label}</td>${gstr3bTaxCells(bucket)}</tr>`).join("");
  }
  if ($("#gstr3bItcPanel")) $("#gstr3bItcPanel").classList.remove("hidden");
  const util = gstr3bDashboard.utilisation || {};
  const payable = (gstr3bDashboard.payable || {}).net_gst_payable || {};
  const interestBucket = (gstr3bDashboard.payable || {}).interest || {};
  const lateBucket = (gstr3bDashboard.payable || {}).late_fee || {};
  const totalCashBucket = (gstr3bDashboard.payable || {}).total_cash_payable_bucket || {};
  if ($("#gstr3bUtilRows")) {
    $("#gstr3bUtilRows").innerHTML = [
      ["Liability", util.liability],
      ["ITC Utilised", util.itc_utilised],
      ["Cash Required (Tax)", util.cash_required],
      ["Remaining ITC", util.remaining_itc],
      ["Net GST Payable (Tax only)", payable],
      ["Net GST Interest Payable", interestBucket],
      ["Net GST Late Fee Payable", lateBucket],
      ["Total Cash Payable", totalCashBucket],
    ].map(([label, bucket]) => `<tr><td>${label}</td>${gstr3bTaxCells(bucket)}</tr>`).join("");
  }
  if ($("#gstr3bUtilPanel")) $("#gstr3bUtilPanel").classList.remove("hidden");
  if ($("#reconGstr3bPeriodStatus")) {
    const period = gstr3bDashboard.return_period || (gstDatasets["GSTR-3B"] || {}).return_period || "";
    const count = (gstr3bDashboard.imported_periods || (gstDatasets["GSTR-3B"] || {}).imported_periods || []).length;
    if (period === "ALL") {
      $("#reconGstr3bPeriodStatus").textContent = count
        ? `GSTR-3B: ALL / FY (${count} months)`
        : "GSTR-3B: ALL / FY";
    } else {
      $("#reconGstr3bPeriodStatus").textContent = period ? `GSTR-3B: ${period}` : "GSTR-3B: Loaded";
    }
  }
  if ($("#gstr3bReconAdvice")) {
    const advice = [];
    if (gstr3bDashboard.books_warning) advice.push(String(gstr3bDashboard.books_warning));
    if (gstr3bDashboard.books_source) advice.push(`Books source: ${gstr3bDashboard.books_source}.`);
    if (gstr3bDashboard.books_itc_source || gstr3bDashboard.books_itc_field) {
      const itcBits = [
        gstr3bDashboard.books_itc_source,
        gstr3bDashboard.books_itc_field ? `field ${gstr3bDashboard.books_itc_field}` : "",
        gstr3bDashboard.books_itc_formula ? `formula ${gstr3bDashboard.books_itc_formula}` : "",
        gstr3bDashboard.books_itc_uncertain_excluded
          ? `uncertain vouchers excluded ${gstr3bDashboard.books_itc_uncertain_excluded}`
          : "",
      ].filter(Boolean);
      advice.push(`Books ITC Claimed: ${itcBits.join("; ")}.`);
    }
    if (Math.abs(Number(cards.books_vs_3b_difference || 0)) > 1) advice.push(`Books vs GSTR-3B difference ₹${money(cards.books_vs_3b_difference)}.`);
    if (Number(cards.net_gst_payable || 0) > 0) advice.push(`Net GST tax payable ₹${money(cards.net_gst_payable)}.`);
    const interestAdvice = Number(cards.net_gst_interest_payable != null ? cards.net_gst_interest_payable : (cards.interest || 0));
    const lateAdvice = Number(cards.net_gst_late_fee_payable != null ? cards.net_gst_late_fee_payable : (cards.late_fee || 0));
    if (interestAdvice || lateAdvice) {
      advice.push(`Interest ₹${money(interestAdvice)}, Late fee ₹${money(lateAdvice)}.`);
    }
    if (Number(cards.total_cash_payable || 0) > 0) advice.push(`Total cash payable ₹${money(cards.total_cash_payable)}.`);
    $("#gstr3bReconAdvice").textContent = advice.join(" ") || "Books liability, ITC utilisation and net payable are aligned for review.";
  }
}

async function refreshGstr3bDashboard(showError = true) {
  const button = $("#reconGstr3bDashBtn");
  const error = $("#gstError");
  if (button) {
    button.disabled = true;
    button.textContent = "Reconciling...";
  }
  if (showError) error.classList.add("hidden");
  try {
    const response = await fetch("/api/gst/recon/gstr3b-dashboard", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        returnPeriod: getGstReconPeriod(),
        tolerance: Number($("#gstr3bReconTolerance")?.value || $("#reconOverviewTolerance")?.value || 1),
        includeDrilldown: false,
      }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "GSTR-3B vs Books reconciliation failed.");
    if (result.imported === false || result.status === "Not Imported") {
      clearGstReconClientStateFor("GSTR-3B");
      await saveGstReconSession({ gstr3b_dashboard: {} });
      if (activeReconTab === "overview") await refreshGstReconOverview(false);
      return;
    }
    renderGstr3bDashboard(result);
    await saveGstReconSession({ gstr3b_dashboard: result });
    if (activeReconTab === "overview") await refreshGstReconOverview(false);
  } catch (failure) {
    if (showError) {
      error.textContent = failure.message;
      error.classList.remove("hidden");
    }
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = "Reconcile Books vs GSTR-3B";
    }
  }
}

async function exportGstr3bReport(report, format = "xlsx") {
  const response = await fetch("/api/gst/recon/gstr3b-export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      report,
      format,
      dashboard: gstr3bDashboard || undefined,
      returnPeriod: getGstReconPeriod(),
      tolerance: Number($("#gstr3bReconTolerance")?.value || 1),
      title: `GSTR3B_${report}`,
    }),
  });
  if (!response.ok) {
    const result = await response.json().catch(() => ({}));
    throw new Error(result.error || "GSTR-3B export failed.");
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `GSTR3B_${report}.${format === "csv" ? "csv" : "xlsx"}`;
  link.click();
  URL.revokeObjectURL(url);
}

async function refreshGstReconOverview(showError = true) {
  try {
    const response = await fetch("/api/gst/recon/overview", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        returnPeriod: getGstReconPeriod(),
        tolerance: Number($("#reconOverviewTolerance")?.value || $("#reconTolerance")?.value || 1),
      }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "Overview refresh failed.");
    renderGstReconOverview(result);
  } catch (failure) {
    if (showError) {
      const error = $("#gstError");
      error.textContent = failure.message;
      error.classList.remove("hidden");
    }
  }
}

async function runOneClickGstSync() {
  const button = $("#reconOneClickSyncBtn");
  const error = $("#gstError");
  const progress = $("#reconOneClickProgress");
  button.disabled = true;
  error.classList.add("hidden");
  progress.classList.remove("hidden");
  $("#reconOneClickPercent").textContent = "5%";
  $("#reconOneClickLabel").textContent = "Starting combined GST sync...";
  $("#reconOneClickBar").style.width = "5%";
  try {
    const response = await fetch("/api/gst/recon/one-click-sync", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        returnPeriod: getGstReconPeriod(),
        tolerance: Number($("#reconOverviewTolerance")?.value || $("#reconTolerance")?.value || 1),
      }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "One Click GST Sync failed.");
    (result.steps || []).forEach(step => {
      $("#reconOneClickPercent").textContent = `${step.percent}%`;
      $("#reconOneClickLabel").textContent = step.label;
      $("#reconOneClickBar").style.width = `${step.percent}%`;
    });
    if (result.purchase?.synced?.rows) gstReconTallyRows = result.purchase.synced.rows;
    if (result.purchase?.rows) {
      gstReconResults = result.purchase.rows;
      renderGstReconMatchResults(
        result.purchase.rows,
        result.purchase.counts || {},
        result.purchase.document_summary || null
      );
    }
    if (result.sales?.synced?.rows) gstReconTallySalesRows = result.sales.synced.rows;
    if (result.sales?.rows) gstReconSalesResults = result.sales.rows;
    if (result.sales?.dashboard) renderSalesReconDashboard(result.sales.dashboard);
    else if (result.sales?.document_summary) {
      renderSignedDocumentTypeSummary(result.sales.document_summary, {
        panelId: "salesDocTypePanel",
        rowsId: "salesDocTypeRows",
        netLabelId: "salesDocTypeNetLabel",
        portalNetId: "salesPortalNetGst",
        tallyNetId: "salesTallyNetGst",
        diffId: "salesNetGstDiff",
        statusId: "salesNetGstStatus",
      });
    }
    if (result.itc_dashboard) renderItcDashboardView(result.itc_dashboard);
    if (result.gstr3b_dashboard) renderGstr3bDashboard(result.gstr3b_dashboard);
    if (result.overview) renderGstReconOverview(result.overview);
    renderGstReconStatus();
    renderSalesReconStatus();
    updateGstReconReady();
    updateSalesReconReady();
    renderSalesReconTable();
    $("#reconOneClickPercent").textContent = "100%";
    $("#reconOneClickLabel").textContent = "Completed";
    $("#reconOneClickBar").style.width = "100%";
  } catch (failure) {
    error.textContent = failure.message;
    error.classList.remove("hidden");
    $("#reconOneClickLabel").textContent = failure.message;
  } finally {
    button.disabled = false;
  }
}

async function exportSalesRecon(format) {
  const statusFilter = $("#salesReconStatusFilter")?.value || "";
  const response = await fetch("/api/gst/recon/sales-export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      rows: filteredSalesReconRows(),
      format,
      statusFilter,
      title: statusFilter ? `GSTR1_Tally_${statusFilter}` : "GSTR1_vs_Tally_Full",
      returnPeriod: getGstReconPeriod(),
    }),
  });
  if (!response.ok) {
    const result = await response.json().catch(() => ({}));
    throw new Error(result.error || "Export failed.");
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `GSTR1_Tally.${format === "csv" ? "csv" : "xlsx"}`;
  link.click();
  URL.revokeObjectURL(url);
}

$("#reconTallySyncBtn").onclick = syncGstReconTallyPurchase;
$("#recon2bTallyBtn").onclick = reconcileGstRecon2bTally;
$("#reconItcBtn").onclick = () => refreshItcDashboard(true);
if ($("#reconItcViewDiffBtn")) $("#reconItcViewDiffBtn").onclick = () => openItcDifferenceDialog();
if ($("#reconItcDiffExportBtn")) $("#reconItcDiffExportBtn").onclick = () => exportItcDifferenceExcel();
if ($("#itcDiffExportBtn")) $("#itcDiffExportBtn").onclick = () => exportItcDifferenceExcel();
if ($("#reconSalesViewDiffBtn")) $("#reconSalesViewDiffBtn").onclick = () => openSalesDifferenceDialog();
if ($("#reconSalesDiffExportBtn")) $("#reconSalesDiffExportBtn").onclick = () => exportSalesDifference("xlsx");
if ($("#reconSalesDiffExportCsvBtn")) $("#reconSalesDiffExportCsvBtn").onclick = () => exportSalesDifference("csv");
if ($("#salesDiffExportXlsxBtn")) $("#salesDiffExportXlsxBtn").onclick = () => exportSalesDifference("xlsx");
if ($("#salesDiffExportCsvBtn")) $("#salesDiffExportCsvBtn").onclick = () => exportSalesDifference("csv");
if (salesDiffQuery("salesDiffCloseBtn")) {
  salesDiffQuery("salesDiffCloseBtn").onclick = () => {
    const dialog = salesDiffDialogRoot();
    if (dialog && typeof dialog.close === "function") dialog.close();
    else if (dialog) dialog.classList.add("hidden");
  };
}
bindGstClearButtons();
bindGstPaymentLedgerControls();
document.querySelectorAll(".recon-tab").forEach(button => {
  button.onclick = () => setReconTab(button.dataset.reconTab);
});
if ($("#reconTallySalesSyncBtn")) $("#reconTallySalesSyncBtn").onclick = syncGstReconTallySales;
if ($("#reconGstr1TallyBtn")) $("#reconGstr1TallyBtn").onclick = reconcileGstReconGstr1Tally;
if ($("#reconSalesDashBtn")) $("#reconSalesDashBtn").onclick = () => refreshSalesReconDashboard(true);
if ($("#salesReconPeriod")) {
  $("#salesReconPeriod").onchange = () => {
    setGstReconPeriod($("#salesReconPeriod").value, { userChosen: true, refresh: false, silent: true });
    refreshActiveReconPeriodViews(false);
  };
}
if ($("#gstr3bReconPeriod")) {
  $("#gstr3bReconPeriod").onchange = () => {
    setGstReconPeriod($("#gstr3bReconPeriod").value, { userChosen: true, refresh: false, silent: true });
    refreshActiveReconPeriodViews(false);
  };
}
if ($("#reconOneClickSyncBtn")) $("#reconOneClickSyncBtn").onclick = runOneClickGstSync;
if ($("#reconOverviewRefreshBtn")) $("#reconOverviewRefreshBtn").onclick = () => refreshGstReconOverview(true);
if ($("#reconPeriodFilter")) {
  $("#reconPeriodFilter").onchange = () => {
    setGstReconPeriod($("#reconPeriodFilter").value, { userChosen: true, refresh: false, silent: true });
    refreshActiveReconPeriodViews(false);
  };
}
const debouncedSalesReconTable = debounce(() => { salesReconPage = 1; renderSalesReconTable(); }, 250);
["salesReconSearch", "salesReconStatusFilter", "salesReconDateFilter", "salesReconGstinFilter", "salesReconVoucherFilter", "salesReconSort"].forEach(id => {
  const node = document.getElementById(id);
  if (!node) return;
  if (node.tagName === "SELECT") {
    node.onchange = () => { salesReconPage = 1; renderSalesReconTable(); };
  } else {
    node.oninput = debouncedSalesReconTable;
    node.onchange = () => { salesReconPage = 1; renderSalesReconTable(); };
  }
});
if ($("#salesFiltersToggle") && $("#salesFiltersPanel")) {
  $("#salesFiltersToggle").onclick = () => {
    const panel = $("#salesFiltersPanel");
    const button = $("#salesFiltersToggle");
    const open = panel.hasAttribute("hidden");
    if (open) {
      panel.removeAttribute("hidden");
      panel.classList.remove("collapsed");
      button.classList.add("is-open");
      button.setAttribute("aria-expanded", "true");
    } else {
      panel.setAttribute("hidden", "");
      panel.classList.add("collapsed");
      button.classList.remove("is-open");
      button.setAttribute("aria-expanded", "false");
    }
  };
}
if ($("#salesReconPrevPage")) $("#salesReconPrevPage").onclick = () => { salesReconPage -= 1; renderSalesReconTable(); };
if ($("#salesReconNextPage")) $("#salesReconNextPage").onclick = () => { salesReconPage += 1; renderSalesReconTable(); };
if ($("#salesReconExportXlsx")) $("#salesReconExportXlsx").onclick = async () => {
  try { await exportSalesRecon("xlsx"); } catch (failure) { alert(failure.message); }
};
if ($("#salesReconExportCsv")) $("#salesReconExportCsv").onclick = async () => {
  try { await exportSalesRecon("csv"); } catch (failure) { alert(failure.message); }
};
if ($("#reconGstr3bDashBtn")) $("#reconGstr3bDashBtn").onclick = () => refreshGstr3bDashboard(true);
if ($("#reconGstr3bLiabilityBtn")) $("#reconGstr3bLiabilityBtn").onclick = async () => {
  try {
    await refreshGstr3bDashboard(true);
    if ($("#gstr3bLiabilityPanel")) $("#gstr3bLiabilityPanel").scrollIntoView({ behavior: "smooth", block: "nearest" });
  } catch (failure) { alert(failure.message); }
};
if ($("#reconGstr3bItcBtn")) $("#reconGstr3bItcBtn").onclick = async () => {
  try {
    await refreshGstr3bDashboard(true);
    if ($("#gstr3bItcPanel")) $("#gstr3bItcPanel").scrollIntoView({ behavior: "smooth", block: "nearest" });
  } catch (failure) { alert(failure.message); }
};
if ($("#reconGstr3bUtilBtn")) $("#reconGstr3bUtilBtn").onclick = async () => {
  try {
    await refreshGstr3bDashboard(true);
    if ($("#gstr3bUtilPanel")) $("#gstr3bUtilPanel").scrollIntoView({ behavior: "smooth", block: "nearest" });
  } catch (failure) { alert(failure.message); }
};
if ($("#gstr3bExportBooks")) $("#gstr3bExportBooks").onclick = async () => {
  try { await exportGstr3bReport("books_vs_3b", "xlsx"); } catch (failure) { alert(failure.message); }
};
if ($("#gstr3bExportLiability")) $("#gstr3bExportLiability").onclick = async () => {
  try { await exportGstr3bReport("liability", "xlsx"); } catch (failure) { alert(failure.message); }
};
if ($("#gstr3bExportItc")) $("#gstr3bExportItc").onclick = async () => {
  try { await exportGstr3bReport("itc", "xlsx"); } catch (failure) { alert(failure.message); }
};
if ($("#gstr3bExportCash")) $("#gstr3bExportCash").onclick = async () => {
  try { await exportGstr3bReport("cash", "xlsx"); } catch (failure) { alert(failure.message); }
};
if ($("#gstr3bExportDiff")) $("#gstr3bExportDiff").onclick = async () => {
  try { await exportGstr3bReport("tax_diff", "xlsx"); } catch (failure) { alert(failure.message); }
};
if ($("#gstr3bExportCsv")) $("#gstr3bExportCsv").onclick = async () => {
  try { await exportGstr3bReport("books_vs_3b", "csv"); } catch (failure) { alert(failure.message); }
};
function renderGstPaymentReview(){const payment=gstDatasets["GST Payment List"]||{},cash=gstDatasets["GST Cash Ledger"]||{},g3=gstDatasets["GSTR-3B"]||{},records=[...(payment.records||[]),...(cash.records||[])],money=v=>Number(v||0).toLocaleString("en-IN",{minimumFractionDigits:2,maximumFractionDigits:2}),itc=Number(g3.igst||0)+Number(g3.cgst||0)+Number(g3.sgst||0);$("#gstPaymentReviewRows").innerHTML=records.map((row,index)=>`<tr><td>${escapeHtml(row.date||row["Deposit Date"]||row.tax_period||"")}</td><td>${escapeHtml(row.description||row["Mode of Payment"]||row.CPIN||"GST Payment")}</td><td>${money(row.Amount||0)}</td><td>${index===records.length-1?money(g3.igst):"—"}</td><td>${index===records.length-1?money(g3.cgst):"—"}</td><td>${index===records.length-1?money(g3.sgst):"—"}</td><td>${index===records.length-1?money(itc):"—"}</td><td>${$("#gstPaymentBankLedger").value.trim()?"Ready for review":"Select Cash / Bank Ledger"}</td></tr>`).join("")||`<tr><td colspan="8">Upload GST Payment List/Cash Ledger and GSTR-3B to prepare the review.</td></tr>`;$("#gstPaymentReviewStatus").textContent=`Payment rows: ${records.length.toLocaleString("en-IN")} · GSTR-3B ITC available: ₹${money(itc)}. Review the month/date and select the Tally Cash/Bank Ledger before entry preparation.`;}
$("#gstPaymentBankLedger").onchange=renderGstPaymentReview;
function gstMonthKey(row){const p=String(row.source_period||"").replace(/\D/g,"");if(p.length===6)return p.slice(2)+p.slice(0,2);const d=String(row.invoice_date||row.date||"");let m=d.match(/(\d{4})[-\/]?(\d{2})/);if(m)return m[1]+m[2];m=d.match(/\d{1,2}[-\/](\d{1,2})[-\/](\d{4})/);return m?m[2]+m[1].padStart(2,"0"):"";}
function gstMonthlyTotals(rows){return (rows||[]).reduce((all,row)=>{const key=gstMonthKey(row);if(!key)return all;const t=all[key]||(all[key]={taxable_value:0,igst:0,cgst:0,sgst:0});["taxable_value","igst","cgst","sgst"].forEach(k=>t[k]+=Number(row[k]||0));return all;},{});}
function normalizeGstr3bPeriods(periods){const normalized={};Object.entries(periods||{}).forEach(([period,values])=>{const digits=String(period).replace(/\D/g,"");const key=digits.length===6?digits.slice(2)+digits.slice(0,2):period;normalized[key]=values||{};});return normalized;}
async function reconcileGst() {
  const button = $("#gstReconcileBtn");
  const error = $("#gstError");
  button.disabled = true;
  button.textContent = "Reconciling...";
  error.classList.add("hidden");
  try {
    const response = await fetch("/api/gst/reconcile", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        gstr2a: gstDatasets["GSTR-2A"] || [], gstr2b: gstDatasets["GSTR-2B"] || [],
        tolerance: $("#gstTolerance").value
      })
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "GST reconciliation failed.");
    gstRows = result.rows || [];
    const money = value => Number(value || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    restorePurchaseItemMappings(gstRows);
    gstRows.forEach(row => {
      row.selected = false;
      row.party_ledger = purchaseLedgerMatch(row);
      const rate = purchaseGstRate(row);
      const mapped = resolvePurchaseLedgerForRow(row, rate);
      row.purchase_ledger = mapped.ledger;
      if (mapped.required) {
        row.purchase_ledger_required = true;
        if (!row.review_required && row.ready_for_tally) {
          /* keep ready but preview will flag ledger */
        }
      }
      row.ready_for_purchase_note = isPurchaseNote(row) && !row.purchase_booked;
      if (!isPurchaseNote(row) && !row.purchase_booked) row.ready_for_tally = true;
    });
    purchase2a2bDashboard = result.dashboard || null;
    $("#gstRows").innerHTML = gstRows.map((row, index) => `<tr>
      <td>${row.ready_for_tally ? `<input class="gst-row-select" data-index="${index}" type="checkbox">` : "—"}</td>
      <td>${escapeHtml(row.gstin)}</td><td>${row.ready_for_tally ? `<input class="gst-party-ledger" data-index="${index}" value="${escapeHtml(row.party_ledger)}">` : escapeHtml(row.party_name || "")}</td>
      <td>${escapeHtml(row.invoice_no)}</td><td>${escapeHtml(row.original_invoice_date || row.invoice_date)}</td>
      <td class="money">${money(row.invoice_value)}</td><td class="money">${money(row.taxable_value)}</td>
      <td class="money">${money(row.igst)}</td><td class="money">${money(row.cgst)}</td><td class="money">${money(row.sgst)}</td>
      <td><span class="gst-status ${String(row.status || "").includes("Matched") ? "matched" : "review"}">${escapeHtml(row.status)}</span></td></tr>`).join("");
    $("#gstMatchCounts").innerHTML = Object.entries(result.counts || {}).map(([name, count]) => `<span>${escapeHtml(name)}<strong>${Number(count).toLocaleString("en-IN")}</strong></span>`).join("");
    $("#gstMatchCounts").classList.remove("hidden");
    $("#gstTallyPanel").classList.remove("hidden");
    document.querySelectorAll(".gst-row-select").forEach(input => input.onchange = () => {
      gstRows[Number(input.dataset.index)].selected = input.checked;
    });
    document.querySelectorAll(".gst-party-ledger").forEach(input => input.oninput = () => {
      gstRows[Number(input.dataset.index)].party_ledger = input.value.trim();
    });
    renderPurchaseReconciliation();
    setPurchaseSheetView("match");
    updatePurchaseImportStatus();
    $(".gst-next-note").textContent = "Reconciliation complete. Review categories, then preview before sending Purchase vouchers to Tally.";
    return true;
  } catch (failure) {
    error.textContent = failure.message;
    error.classList.remove("hidden");
    if ($("#purchase2aWorkspace")) {
      $("#purchase2aWorkspace").classList.remove("hidden");
      $("#purchase2aWorkspace").parentElement?.classList.remove("hidden");
    }
    return false;
  } finally {
    button.disabled = false;
    button.textContent = "Match GSTR-2A & 2B";
  }
}
$("#gstReconcileBtn").onclick = reconcileGst;
async function sendSelectedPurchaseRows(predicate, button, idleText) {
  const selectedCandidates = gstRows.filter(row => row.selected && row.ready_for_tally && predicate(row));
  const problematic = selectedCandidates.filter(row => !purchaseAllowedForTally(row));
  const selected = selectedCandidates.filter(purchaseAllowedForTally);
  if (!selectedCandidates.length) return alert("No reviewed Purchase Invoice is selected to send.");
  const money = (v) => Number(v || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const issues = [];
  const preparePreviewRow = (row) => {
    if (!row.expense_ledger && !row.sales_allocations?.length) {
      setRowItemName(-1, salesItemLabel(row) || `${purchaseGstRate(row)}% Items`, row);
    }
    const rate = purchaseGstRate(row);
    const mapped = resolvePurchaseLedgerForRow(row, rate);
    row.purchase_ledger = mapped.ledger || row.purchase_ledger || "";
    if (mapped.required || !row.purchase_ledger) {
      row.purchase_ledger_required = true;
      issues.push(`${row.invoice_no || "?"}: Purchase Ledger Required`);
    }
    if (!String(row.party_ledger || row.party_name || "").trim()) {
      issues.push(`${row.invoice_no || "?"}: Party ledger required`);
    }
    return row;
  };
  const previewRows = selected.map(preparePreviewRow);
  const problematicPreviewRows = problematic.map(preparePreviewRow);
  const dialog = $("#purchaseTallyPreviewDialog");
  const body = $("#purchaseTallyPreviewRows");
  if (body) {
    body.innerHTML = previewRows.map((row) => `<tr>
      <td>${escapeHtml(row.tally_entry_date || "")}</td>
      <td>${escapeHtml(row.original_invoice_no || row.invoice_no || "")}</td>
      <td>${escapeHtml(row.original_invoice_date || row.invoice_date || "")}</td>
      <td>${escapeHtml(row.party_ledger || row.party_name || "")}</td>
      <td>${escapeHtml(row.gstin || "")}</td>
      <td>${escapeHtml(row.purchase_ledger_required ? "Purchase Ledger Required" : (row.purchase_ledger || row.expense_ledger || "—"))}</td>
      <td class="money">${money(row.taxable_value)}</td>
      <td class="money">${money(row.igst)}</td>
      <td class="money">${money(row.cgst)}</td>
      <td class="money">${money(row.sgst)}</td>
      <td class="money">${money(row.cess)}</td>
      <td class="money">${money(row.invoice_value)}</td>
      <td>${row.available_in_gstr2a ? "Yes" : "No"}</td>
      <td>${row.available_in_gstr2b ? "Yes" : "No"}</td>
      <td>${escapeHtml(row.tally_status || "Missing in Tally")}</td>
      <td>${escapeHtml(row.itc_status || "")}</td>
    </tr>`).join("");
  }
  if ($("#purchaseTallyPreviewIssues")) {
    $("#purchaseTallyPreviewIssues").textContent = issues.length
      ? `Review required before send: ${issues.join("; ")}`
      : `${previewRows.length} voucher(s) ready. ${problematic.length} problematic invoice(s) will be skipped individually.`;
  }
  const skippedPanel = $("#purchaseTallySkippedPanel");
  const skippedBody = $("#purchaseTallySkippedRows");
  const differenceReasons = new Map(itcDifferenceInvoices().map(({row, reason}) => [purchaseDocumentKey(row), reason]));
  if (skippedPanel) skippedPanel.classList.toggle("hidden", !problematic.length);
  if ($("#purchaseTallySkippedCount")) $("#purchaseTallySkippedCount").textContent = problematic.length.toLocaleString("en-IN");
  if (skippedBody) {
    skippedBody.innerHTML = problematic.map((entry) => {
      const row = entry.gstr2b || entry.gstr2a || entry;
      const reason = differenceReasons.get(purchaseDocumentKey(entry)) || entry.itc_status || entry.status || "ITC review required";
      return `<tr><td>${escapeHtml(row.invoice_no || "")}</td><td>${escapeHtml(entry.party_ledger || row.party_name || "")}</td><td>${escapeHtml(row.gstin || "")}</td><td>${escapeHtml(row.original_invoice_date || row.invoice_date || "")}</td><td>${escapeHtml(row.gstr2b_period || row.source_period || "")}</td><td class="itc-reason">${escapeHtml(reason)}</td></tr>`;
    }).join("");
  }
  const readyButton = $("#purchaseTallyPreviewConfirm");
  const problematicButton = $("#purchaseTallyPreviewProblematicConfirm");
  if (readyButton) {
    readyButton.textContent = `Send Ready ${previewRows.length.toLocaleString("en-IN")} to Tally`;
    readyButton.classList.toggle("hidden", !previewRows.length);
  }
  if (problematicButton) {
    problematicButton.textContent = `Send ITC Difference ${problematicPreviewRows.length.toLocaleString("en-IN")} to Tally`;
    problematicButton.classList.toggle("hidden", !problematicPreviewRows.length);
  }
  const sendChoice = await new Promise((resolve) => {
    if (!dialog || typeof dialog.showModal !== "function") {
      resolve(confirm(`Create ${selected.length} ready Purchase Voucher(s) in the currently open Tally company? Take a Tally backup first.`) ? "default" : "cancel");
      return;
    }
    const onClose = () => {
      dialog.removeEventListener("close", onClose);
      resolve(dialog.returnValue || "cancel");
    };
    dialog.addEventListener("close", onClose);
    dialog.returnValue = "cancel";
    dialog.showModal();
  });
  if (sendChoice === "cancel") return;
  const chosenRows = sendChoice === "problematic" ? problematicPreviewRows : previewRows;
  const groupLabel = sendChoice === "problematic" ? "ITC Difference" : "Ready";
  const sendable = chosenRows.filter((row) => !row.purchase_ledger_required && String(row.party_ledger || row.party_name || "").trim());
  if (!sendable.length) return alert("No invoice passed Purchase Ledger / Party validation. Fix Review Required rows and try again.");
  button.disabled = true;
  button.textContent = "Sending...";
  startTallySendProgress(`Sending ${groupLabel} Purchase vouchers`,()=>sendSelectedPurchaseRows(predicate,button,idleText));
  try {
    const ledgerConfig = {
      partyParent:"Sundry Creditors", purchaseLedger:$("#gstPurchaseLedger").value,
      purchaseLedger0:$("#gstPurchaseLedger0").value, purchaseLedger5:$("#gstPurchaseLedger5").value,
      purchaseLedger12:$("#gstPurchaseLedger12").value, purchaseLedger18:$("#gstPurchaseLedger18").value,
      purchaseLedger28:$("#gstPurchaseLedger28").value,
      igstLedger:$("#gstIgstLedger").value, cgstLedger:$("#gstCgstLedger").value,
      sgstLedger:$("#gstSgstLedger").value, cessLedger: ($("#gstCessLedger") || {}).value || "Input Cess",
      roundLedger:$("#gstRoundLedger").value
    };
    $("#tallySendMessage").textContent=`Checking ledgers and items for ${sendable.length.toLocaleString("en-IN")} ${groupLabel} voucher(s)...`;updateTallySendProgress(Math.max(tallySendValue,15));
    const ensureResponse = await fetch("/api/gst/party-ledgers/ensure", {method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({rows:sendable,ledgers:ledgerConfig})});
    const ensureResult = await ensureResponse.json();
    if(!ensureResponse.ok) throw new Error(ensureResult.error||"Purchase party ledger creation failed.");
    const mappings=ensureResult.mappings||{};
    const purchaseLedgers=ensureResult.purchaseLedgers||{};
    [0,5,12,18,28].forEach(rate=>{const field=`purchaseLedger${rate}`;if(purchaseLedgers[field]){ledgerConfig[field]=purchaseLedgers[field];$(`#gstPurchaseLedger${rate}`).value=purchaseLedgers[field];}});
    gstRows.forEach(row=>{const key=String(row.party_ledger||row.party_name||"").trim().toLowerCase();if(mappings[key])row.party_ledger=mappings[key];});
    const expenseMappings=ensureResult.expenseMappings||{};
    gstRows.forEach(row=>{const key=String(row.expense_ledger||"").trim().toLowerCase();if(key&&expenseMappings[key])row.expense_ledger=expenseMappings[key];});
    $("#tallySendMessage").textContent=`Sending ${groupLabel} vouchers to Tally...`;updateTallySendProgress(Math.max(tallySendValue,50));
    const response = await fetch("/api/gst/tally/send", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ rows: sendable.map(row => ({...row, selected:true, ready_for_tally:true})), ledgers: ledgerConfig, tolerance: $("#gstTolerance")?.value || 1 })
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "Tally import failed.");
    const skipped = result.skipped || [];
    const failed = result.failed || [];
    skipped.forEach((item) => {
      const row = gstRows.find((r) => String(r.invoice_no) === String(item.invoice_no) && String(r.gstin) === String(item.gstin));
      if (row) {
        row.tally_status = "Already in Tally";
        row.purchase_booked = true;
        row.ready_for_tally = false;
        row.tally_voucher_no = item.tally_voucher_no || "";
        if (row.category === "only_2a") {
          row.status = "Booked in Tally / Not Available in GSTR-2B";
          row.itc_status = "ITC Pending / Not Available in GSTR-2B";
        }
      }
    });
    failed.forEach((item) => {
      const row = gstRows.find((r) => String(r.invoice_no) === String(item.invoice_no) && String(r.gstin) === String(item.gstin));
      if (row) {
        row.tally_status = item.status || "Tally Entry Failed";
        row.tally_error = item.error || "";
        row.review_required = true;
        if (String(item.status || "").includes("Failed")) row.status = "Tally Entry Failed";
      }
    });
    if (result.created) {
      sendable.forEach((row) => {
        const failedHit = failed.some((item) => String(item.invoice_no) === String(row.invoice_no) && String(item.gstin) === String(row.gstin));
        const skippedHit = skipped.some((item) => String(item.invoice_no) === String(row.invoice_no) && String(item.gstin) === String(row.gstin));
        if (failedHit || skippedHit) return;
        row.tally_status = "Sent to Tally";
        row.purchase_booked = true;
        row.ready_for_tally = false;
        if (row.category === "only_2a") {
          row.status = "Booked in Tally / Not Available in GSTR-2B";
          row.itc_status = "ITC Pending / Not Available in GSTR-2B";
        } else if (String(row.status || "").includes("2A + 2B")) {
          row.status = "Sent to Tally";
        }
      });
    }
    purchase2a2bDashboard = null;
    renderPurchaseReconciliation();
    const parts = [
      result.message || `${result.created || 0} voucher(s) created.`,
    ];
    if (skipped.length) parts.push(`Already in Tally: ${skipped.map((x) => x.invoice_no).join(", ")}`);
    if (failed.length) parts.push(`Failed/Review: ${failed.map((x) => `${x.invoice_no} (${x.error || x.status})`).join("; ")}`);
    const accepted=Number(result.created||0);const success=failed.length===0&&(accepted>0||skipped.length>0);
    finishTallySendProgress(success,`Accepted ${accepted.toLocaleString("en-IN")} voucher(s).\n${parts.join("\n")}`);
  } catch (failure) {
    $("#gstError").textContent = failure.message;
    $("#gstError").classList.remove("hidden");
    finishTallySendProgress(false,failure.message||"Purchase voucher import failed.");
  } finally {
    button.disabled = false;
    button.textContent = idleText;
  }
}
const purchaseInvoice = row => !isPurchaseNote(row);
const matchedPurchase = row => purchaseInvoice(row) && Boolean(row.gstr2a && row.gstr2b) && row.category !== "mismatch" && !String(row.status || "").includes("Mismatch");
const only2aPurchase = row => purchaseInvoice(row) && Boolean(row.gstr2a && !row.gstr2b);
const only2bPurchase = row => purchaseInvoice(row) && Boolean(row.gstr2b && !row.gstr2a);
const mismatchPurchase = row => purchaseInvoice(row) && row.category === "mismatch";
function selectAllPurchaseRows(predicate) {
  const targets=gstRows.filter(predicate),select=targets.some(row=>!row.selected);
  targets.forEach(row=>row.selected=select);
  renderPurchaseReconciliation();
}
async function exportPurchasePreview(title,predicate) {
  const rows=gstRows.filter(predicate);
  if(!rows.length)return alert("No rows are available in this box.");
  const response=await fetch("/api/gst/preview-xlsx",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({title,rows})});
  if(!response.ok){const failure=await response.json().catch(()=>({}));return alert(failure.error||"Excel export failed.");}
  const blob=await response.blob(),link=document.createElement("a");link.href=URL.createObjectURL(blob);link.download=`${title.replace(/[^A-Za-z0-9_-]+/g,"_")}.xlsx`;link.click();URL.revokeObjectURL(link.href);
}
async function exportPurchaseSummary(title,summaryRows) {
  const response=await fetch("/api/gst/preview-xlsx",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({title,summaryRows})});
  if(!response.ok)return alert("Excel export failed.");const blob=await response.blob(),link=document.createElement("a");link.href=URL.createObjectURL(blob);link.download=`${title.replace(/[^A-Za-z0-9_-]+/g,"_")}.xlsx`;link.click();URL.revokeObjectURL(link.href);
}
function purchaseSummaryExportRows(){
  const source=gstDatasets["GSTR-2B"]||[];
  const portal=purchaseGstr2Breakdown(source);
  const eligible=purchaseGstr2Breakdown(source.filter(purchaseAllowedForTally));
  const format=(Particulars,t)=>({
    Particulars,
    Count:t.invoices,
    Taxable:t.taxable_value,
    IGST:t.igst,
    CGST:t.cgst,
    SGST:t.sgst,
    CESS:t.cess||0,
    "Total ITC": t.itc!=null ? Number(t.itc) : (Number(t.igst||0)+Number(t.cgst||0)+Number(t.sgst||0)+Number(t.cess||0)),
  });
  const amendment={
    invoices:Number(portal.increase.invoices||0)+Number(portal.decrease.invoices||0),
    taxable_value:Number(portal.increase.taxable_value||0)-Number(portal.decrease.taxable_value||0),
    igst:Number(portal.increase.igst||0)-Number(portal.decrease.igst||0),
    cgst:Number(portal.increase.cgst||0)-Number(portal.decrease.cgst||0),
    sgst:Number(portal.increase.sgst||0)-Number(portal.decrease.sgst||0),
    cess:Number(portal.increase.cess||0)-Number(portal.decrease.cess||0),
    itc:Number(portal.summary?.amendment_itc||0),
  };
  return [
    ["Gross Invoice ITC", portal.total],
    ["(+) Debit Note ITC", portal.debit],
    ["(−) Credit Note ITC", portal.credit],
    ["(±) Amendment Adjustment", amendment],
    ["Net GSTR-2B ITC", portal.net],
    ["Tally Eligible Net ITC", eligible.net],
  ].map(([label,total])=>format(label,total));
}
function purchaseBreakdownExportRows(source,label){return purchaseReturnBreakdown(source,label).map(([Particulars,t])=>({Particulars,Invoices:t.invoices,Taxable:t.taxable_value,IGST:t.igst,CGST:t.cgst,SGST:t.sgst}));}
$("#gstSelectAllMatched").onclick=()=>selectAllPurchaseRows(matchedPurchase);
$("#gstSelectAllOnly2a").onclick=()=>selectAllPurchaseRows(only2aPurchase);
$("#gstSelectAllOnly2b").onclick=()=>selectAllPurchaseRows(only2bPurchase);
$("#gstSelectAllMismatch").onclick=()=>selectAllPurchaseRows(mismatchPurchase);
$("#gstSelectAllNotes").onclick=()=>selectAllPurchaseRows(isPurchaseNote);
$("#gstExportMatched").onclick=()=>exportPurchasePreview("Matched_GSTR2A_GSTR2B",matchedPurchase);
$("#gstExportOnly2a").onclick=()=>exportPurchasePreview("Only_GSTR2A",only2aPurchase);
$("#gstExportOnly2b").onclick=()=>exportPurchasePreview("Only_GSTR2B",only2bPurchase);
$("#gstExportMismatch").onclick=()=>exportPurchasePreview("GSTR2A_GSTR2B_Mismatch",mismatchPurchase);
$("#gstExportNotes").onclick=()=>exportPurchasePreview("Credit_Debit_Note_Amendment",isPurchaseNote);
$("#gstExportSummaryBtn").onclick=event=>{event.preventDefault();event.stopPropagation();exportPurchaseSummary("GSTR2_Summary",purchaseSummaryExportRows());};
$("#gstExport2aSummaryBtn").onclick=event=>{event.preventDefault();event.stopPropagation();exportPurchaseSummary("GSTR2A_Summary",purchaseBreakdownExportRows(gstDatasets["GSTR-2A"]||[],"GSTR-2A"));};
$("#gstExport3bCompareBtn").onclick=event=>{event.preventDefault();event.stopPropagation();const report=gstDatasets["GSTR-3B"]||{},rcm=report.reverse_charge||{},gstr2=purchaseNetGstr2bTotals(),diff={igst:Number(report.igst||0)-Number(gstr2.igst||0),cgst:Number(report.cgst||0)-Number(gstr2.cgst||0),sgst:Number(report.sgst||0)-Number(gstr2.sgst||0),cess:Number(report.cess||0)};const format=(Particulars,t)=>({Particulars,Taxable:Number(t.taxable_value||0),IGST:Number(t.igst||0),CGST:Number(t.cgst||0),SGST:Number(t.sgst||0),CESS:Number(t.cess||0)});exportPurchaseSummary("GSTR3B_ITC_RCM_Comparison",[format("Net/Gross GSTR-2B ITC",gstr2),format("GSTR-3B Claimed ITC",report),format("Difference: GSTR-3B - GSTR-2B",diff),format("GSTR-3B Reverse Charge",rcm)]);};
$("#gstSendTallyBtn").onclick = () => sendSelectedPurchaseRows(purchaseInvoice,$("#gstSendTallyBtn"),"Send Selected to Tally");
$("#gstSendAllPurchaseBtn").onclick = event => { event.preventDefault(); event.stopPropagation(); return sendSelectedPurchaseRows(purchaseInvoice,$("#gstSendAllPurchaseBtn"),"Send All Selected to Tally"); };
$("#gstSendMatchedPurchaseBtn").onclick = () => sendSelectedPurchaseRows(matchedPurchase,$("#gstSendMatchedPurchaseBtn"),"Send Selected Matched to Tally");
$("#gstSendOnly2aPurchaseBtn").onclick = () => sendSelectedPurchaseRows(only2aPurchase,$("#gstSendOnly2aPurchaseBtn"),"Send Selected 2A-only to Tally");
$("#gstSendOnly2bPurchaseBtn").onclick = () => sendSelectedPurchaseRows(only2bPurchase,$("#gstSendOnly2bPurchaseBtn"),"Send Selected 2B-only to Tally");
$("#gstSendMismatchPurchaseBtn").onclick = () => sendSelectedPurchaseRows(mismatchPurchase,$("#gstSendMismatchPurchaseBtn"),"Send Selected Reviewed Mismatch to Tally");

async function applyPurchaseItem(inputId,predicate) {
  await applyItemToSelectedRows({
    itemInputId: inputId,
    predicate,
    refresh: renderPurchaseReconciliation,
    emptySelectMessage: "Select at least one row from this box.",
  });
}
$("#gstMatchedItem").onfocus = () => {
  const selected = gstRows.filter((row) => row.selected && matchedPurchase(row));
  if (selected.length) showItemSuggestions(selected[0]);
};
$("#gstOnly2aItem").onfocus = () => {
  const selected = gstRows.filter((row) => row.selected && only2aPurchase(row));
  if (selected.length) showItemSuggestions(selected[0]);
};
$("#gstOnly2bItem").onfocus = () => {
  const selected = gstRows.filter((row) => row.selected && only2bPurchase(row));
  if (selected.length) showItemSuggestions(selected[0]);
};
$("#gstPurchaseNoteItem").onfocus = () => {
  const selected = gstRows.filter((row) => row.selected && isPurchaseNote(row));
  if (selected.length) showItemSuggestions(selected[0]);
};
$("#gstApplyMatchedItem").onclick=()=>applyPurchaseItem("gstMatchedItem",matchedPurchase);
$("#gstApplyOnly2aItem").onclick=()=>applyPurchaseItem("gstOnly2aItem",only2aPurchase);
$("#gstApplyOnly2bItem").onclick=()=>applyPurchaseItem("gstOnly2bItem",only2bPurchase);
$("#gstApplyMismatchItem").onclick=()=>applyPurchaseItem("gstMismatchItem",mismatchPurchase);
$("#gstApplyPurchaseNoteItem").onclick=()=>applyPurchaseItem("gstPurchaseNoteItem",isPurchaseNote);
async function sendSelectedPurchaseNotes(button=$("#gstSendPurchaseNotesBtn"),idleText="Send Selected Notes to Tally",predicate=isPurchaseNote) {
  const selected=gstRows.filter(row=>row.selected&&row.ready_for_purchase_note&&predicate(row)&&purchaseAllowedForTally(row));
  if(!selected.length)return alert("Select at least one reviewed Credit Note, Debit Note or Invoice Amendment.");
  if(!confirm(`Take a Tally backup first. Send ${selected.length} selected Purchase Note/Amendment entry(s) to Tally?`))return;
  button.disabled=true;button.textContent="Sending...";
  startTallySendProgress(`Sending ${selected.length} Purchase Note/Amendment voucher(s)`,()=>sendSelectedPurchaseNotes(button,idleText,predicate));
  const ledgers={partyParent:"Sundry Creditors",purchaseLedger:$("#gstPurchaseLedger").value,purchaseLedger0:$("#gstPurchaseLedger0").value,purchaseLedger5:$("#gstPurchaseLedger5").value,purchaseLedger12:$("#gstPurchaseLedger12").value,purchaseLedger18:$("#gstPurchaseLedger18").value,purchaseLedger28:$("#gstPurchaseLedger28").value,igstLedger:$("#gstIgstLedger").value,cgstLedger:$("#gstCgstLedger").value,sgstLedger:$("#gstSgstLedger").value,roundLedger:$("#gstRoundLedger").value};
  try{
    $("#tallySendMessage").textContent="Checking supplier, item and tax ledgers...";updateTallySendProgress(Math.max(tallySendValue,15));
    const ensureResponse=await fetch("/api/gst/party-ledgers/ensure",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({rows:selected,ledgers})});const ensured=await ensureResponse.json();if(!ensureResponse.ok)throw new Error(ensured.error||"Supplier ledger creation failed.");
    const mappings=ensured.mappings||{};selected.forEach(row=>{const key=String(row.party_ledger||row.party_name||"").trim().toLowerCase();if(mappings[key])row.party_ledger=mappings[key];});const purchaseLedgers=ensured.purchaseLedgers||{};[0,5,12,18,28].forEach(rate=>{const field=`purchaseLedger${rate}`;if(purchaseLedgers[field])ledgers[field]=purchaseLedgers[field];});
    $("#tallySendMessage").textContent=`Sending ${selected.length} voucher(s) one by one so one failure cannot stop the others...`;updateTallySendProgress(Math.max(tallySendValue,45));
    const response=await fetch("/api/gst/purchase-notes/tally/send",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({rows:selected,ledgers})});const result=await response.json();if(!response.ok)throw new Error(result.error||"Purchase Note import failed.");
    const created=Number(result.created||0),altered=Number(result.altered||0),failed=(result.failed||[]),ignored=Number(result.ignored||0);
    const lines=[`Created: ${created}`,`Already Exists / Altered: ${altered}`,`Ignored: ${ignored}`,`Failed: ${failed.length}`];
    if(failed.length)lines.push("",...failed.slice(0,20).map(item=>`${item.invoice_no||"?"}: ${item.error||"Failed"}`));
    finishTallySendProgress(failed.length===0&&created+altered>0,lines.join("\n"),{counts:{created,already_exists_count:altered,errors:failed.length,exceptions:Number(result.exceptions||0)}});
  }catch(failure){$("#gstError").textContent=failure.message;$("#gstError").classList.remove("hidden");finishTallySendProgress(false,failure.message||"Purchase Note import failed.");}finally{button.disabled=false;button.textContent=idleText;}
}
$("#gstSendPurchaseNotesBtn").onclick = () => sendSelectedPurchaseNotes();
$("#itcExcludeAllBtn").onclick=()=>{itcDifferenceInvoices().forEach(({row})=>itcTallyExcluded.add(purchaseDocumentKey(row)));renderGstr2Summary([],[],[],[]);};
$("#itcIncludeAllBtn").onclick=()=>{itcDifferenceInvoices().forEach(({row})=>itcTallyExcluded.delete(purchaseDocumentKey(row)));renderGstr2Summary([],[],[],[]);};
$("#itcDifferenceSearch").oninput=renderItcDifferenceInvoices;
["itcFilterPeriod","itcFilterGstin","itcFilterParty","itcFilterInvoice","itcFilterDate","itcFilterType","itcFilterTaxable","itcFilterIgst","itcFilterCgst","itcFilterSgst","itcFilterReason"].forEach(id=>{$(`#${id}`).oninput=renderItcDifferenceInvoices;});
$("#itcFilterAction").onchange=renderItcDifferenceInvoices;
$("#itcClearFilters").onclick=()=>{["itcDifferenceSearch","itcFilterPeriod","itcFilterGstin","itcFilterParty","itcFilterInvoice","itcFilterDate","itcFilterType","itcFilterTaxable","itcFilterIgst","itcFilterCgst","itcFilterSgst","itcFilterReason","itcFilterAction"].forEach(id=>{const field=$(`#${id}`);if(field)field.value="";});renderItcDifferenceInvoices();};
const debouncedPurchaseReconSearch = debounce(() => {
  purchaseReconPages = { matched: 1, only2a: 1, only2b: 1, notes: 1, mismatch: 1 };
  renderPurchaseReconciliation();
}, 250);
[$("#gstMatchedSearch"),$("#gstOnly2aSearch"),$("#gstOnly2bSearch"),$("#gstMismatchSearch"),$("#gstNotesSearch")].forEach(input=>{
  if(input)input.oninput=debouncedPurchaseReconSearch;
});
document.querySelectorAll("details.purchase-always-open").forEach(box=>{box.open=true;const summary=box.querySelector(":scope > summary");if(summary)summary.onclick=event=>event.preventDefault();box.addEventListener("toggle",()=>{if(!box.open)box.open=true;});});

$("#chooseBtn").onclick = () => input.click();
input.onchange = () => upload([...input.files]);
["dragenter", "dragover"].forEach(name => card.addEventListener(name, event => { event.preventDefault(); card.classList.add("drag"); }));
["dragleave", "drop"].forEach(name => card.addEventListener(name, event => { event.preventDefault(); card.classList.remove("drag"); }));
card.addEventListener("drop", event => upload([...event.dataTransfer.files]));
bankLedgerInput.addEventListener("change", () => applyBankLedgerToRows(true));
bankLedgerInput.addEventListener("blur", () => applyBankLedgerToRows(true));
document.querySelectorAll(".chip").forEach(button => button.onclick = () => {
  document.querySelectorAll(".chip").forEach(item => item.classList.remove("active"));
  button.classList.add("active");
  filter = button.dataset.filter;
  render();
});

function showMessage(text, error = false) {
  message.textContent = text;
  message.classList.remove("hidden");
  message.style.background = error ? "#fff0f0" : "#edf9f3";
  message.style.color = error ? "#a11" : "#176a49";
}

function setProgress(percent, text = "Processing statement...") {
  const safe = Math.max(0, Math.min(100, Math.round(percent)));
  processingPercent.textContent = `${safe}%`;
  processingText.textContent = text;
  progressBar.style.width = `${safe}%`;
  processingStatus.classList.toggle("active", safe > 0 && safe < 100);
  processingStatus.classList.toggle("complete", safe === 100);
}

function beginProcessingProgress() {
  clearInterval(processingTimer);
  let value = 55;
  setProgress(value);
  processingTimer = setInterval(() => {
    value = Math.min(94, value + (value < 80 ? 2 : 1));
    setProgress(value);
  }, 700);
}

function finishProcessingProgress(state = "complete") {
  clearInterval(processingTimer);
  processingTimer = null;
  if (state === "complete") setProgress(100, "Processing complete");
  else if (state === "cancelled") setProgress(0, "Upload cancelled");
  else setProgress(0, "Processing failed");
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onprogress = event => {
      if (event.lengthComputable) setProgress((event.loaded / event.total) * 50, "Reading file...");
    };
    reader.onerror = () => reject(reader.error || new Error("Could not read the file."));
    reader.onload = () => {
      const bytes = new Uint8Array(reader.result);
      let binary = "";
      for (let index = 0; index < bytes.length; index += 0x8000) {
        binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
      }
      setProgress(50, "Uploading statement...");
      resolve(btoa(binary));
    };
    reader.readAsArrayBuffer(file);
  });
}

async function upload(files) {
  if (!files.length) return;
  if (!bankLedgerInput.value.trim()) {
    showMessage("Enter the Tally bank ledger before uploading a statement.", true);
    bankLedgerInput.focus();
    return;
  }
  showMessage("Processing statement...");
  setProgress(0, "Starting upload...");
  for (const file of files) {
    try {
      const packed = { name: file.name, data: await fileToBase64(file), password: "" };
      beginProcessingProgress();
      await processPackedFile(packed);
    } catch (error) {
      finishProcessingProgress("failed");
      showMessage(`${file.name}: ${error.message || "Processing failed."}`, true);
    }
  }
  input.value = "";
}

async function processPackedFile(packed) {
  const response = await fetch("/api/parse", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ bankLedger: bankLedgerInput.value.trim(), files: [packed] })
  });
  const data = await response.json();
  if (!response.ok) {
    if (String(data.error || "").includes("Incorrect PDF password")) {
      showMessage("Incorrect PDF password. Please try again.", true);
      const retryPassword = await askForPassword(packed.name);
      if (retryPassword === null) {
        finishProcessingProgress("cancelled");
        showMessage(`${packed.name}: upload cancelled.`);
        return;
      }
      packed.password = retryPassword;
      beginProcessingProgress();
      return processPackedFile(packed);
    }
    throw new Error(data.error || "Statement processing failed.");
  }
  const fileResult = data.files[0] || {};
  if (fileResult.password_required) {
    clearInterval(processingTimer);
    setProgress(55, "PDF password required");
    const password = await askForPassword(packed.name);
    if (password === null) {
      finishProcessingProgress("cancelled");
      showMessage(`${packed.name}: upload cancelled.`);
      return;
    }
    packed.password = password;
    beginProcessingProgress();
    return processPackedFile(packed);
  }
  if (fileResult.mapping_required) {
    clearInterval(processingTimer);
    setProgress(75, "Column mapping required");
    const mappedRows = await askForMapping(fileResult);
    if (!mappedRows) {
      finishProcessingProgress("cancelled");
      showMessage(`${packed.name}: mapping cancelled.`);
      return;
    }
    prepareStatementRows(mappedRows, packed.name);
    const duplicates = addUniqueRows(mappedRows);
    rememberStatementBalance(mappedRows, fileResult);
    recalculateAllStatementAmounts();
    rebalanceStatementBalances();
    sortRowsByDate();
    $("#workspace").classList.remove("hidden");
    finishProcessingProgress("complete");
    showMessage(`${packed.name}: ${mappedRows.length - duplicates} transactions added${duplicates ? `; ${duplicates} duplicate${duplicates === 1 ? "" : "s"} skipped` : ""}.`);
    render();
    if (tallyMasters.connected) await autoMatchBankLedgers();
    await refreshLicense();
    return;
  }
  prepareStatementRows(data.rows, packed.name);
  const duplicates = addUniqueRows(data.rows);
  rememberStatementBalance(data.rows, fileResult);
  recalculateAllStatementAmounts();
  rebalanceStatementBalances();
  sortRowsByDate();
  $("#workspace").classList.remove("hidden");
  finishProcessingProgress("complete");
  showMessage(`${packed.name}: ${data.rows.length - duplicates} transactions added${duplicates ? `; ${duplicates} duplicate${duplicates === 1 ? "" : "s"} skipped` : ""} - ${fileResult.format || "automatic"}.`);
  render();
  if (tallyMasters.connected) await autoMatchBankLedgers();
  await refreshLicense();
}

function askForPassword(filename) {
  return new Promise(resolve => {
    const dialog = $("#passwordDialog");
    let settled = false;
    const finish = value => {
      if (settled) return;
      settled = true;
      resolve(value);
    };
    $("#passwordFile").textContent = filename;
    $("#pdfPassword").value = "";
    $("#passwordSubmit").onclick = event => {
      event.preventDefault();
      const value = $("#pdfPassword").value;
      if (!value) return;
      dialog.close();
      finish(value);
    };
    const cancel = dialog.querySelector("button[value='cancel']");
    cancel.onclick = event => { event.preventDefault(); dialog.close(); finish(null); };
    dialog.oncancel = event => { event.preventDefault(); dialog.close(); finish(null); };
    dialog.showModal();
    $("#pdfPassword").focus();
  });
}

const mappingDefinitions = [
  ["date", "Date *"], ["value_date", "Value Date"], ["particulars", "Particulars *"],
  ["debit", "Debit"], ["credit", "Credit"], ["balance", "Balance"],
  ["reference", "Reference / UTR"], ["instrument", "Instrument / Cheque"],
  ["amount", "Single Amount"], ["direction", "Dr / Cr"]
];

function askForMapping(meta) {
  return new Promise(resolve => {
    const dialog = $("#mappingDialog");
    let settled = false;
    const finish = value => {
      if (settled) return;
      settled = true;
      resolve(value);
    };
    const mappingError = $("#mappingError");
    const submitButton = $("#mappingSubmit");
    const grid = meta.grid || [];
    const showMappingError = text => {
      mappingError.textContent = text || "Mapping could not be applied.";
      mappingError.classList.remove("hidden");
    };
    mappingError.textContent = "";
    mappingError.classList.add("hidden");
    submitButton.disabled = false;
    submitButton.textContent = "Apply & Save Mapping";
    $("#mappingFile").textContent = meta.filename;
    $("#profileName").value = meta.filename.replace(/\.[^.]+$/, "");
    const headerSelect = $("#headerRow");
    headerSelect.innerHTML = grid.slice(0, 30).map((row, index) => `<option value="${index}" ${index === Number(meta.header_row) ? "selected" : ""}>Row ${index + 1}: ${esc(row.slice(0, 5).join(" | ")).slice(0, 100)}</option>`).join("");
    const fields = $("#mappingFields");
    fields.innerHTML = mappingDefinitions.map(([key, label]) => `<label>${label}<select data-map="${key}"></select></label>`).join("");

    function refreshColumns() {
      const headerRow = Number(headerSelect.value || 0);
      const headers = grid[headerRow] || [];
      fields.querySelectorAll("select[data-map]").forEach(select => {
        const key = select.dataset.map;
        select.innerHTML = `<option value="-1">(Not used)</option>` + headers.map((value, index) => `<option value="${index}">${columnName(index)} - ${esc(String(value || "Column " + (index + 1)))}</option>`).join("");
        const aliases = {
          date: ["date", "txn date", "transaction date", "booking date", "booking dt", "posting date", "posting dt"], value_date: ["value date"],
          particulars: ["particulars", "description", "narration", "remarks", "details", "memo", "memo text"],
          debit: ["debit", "withdrawal", "withdrawal amount", "dr amount", "paid out", "payment"], credit: ["credit", "deposit", "deposit amount", "cr amount", "money in", "receipt"],
          balance: ["balance", "closing balance", "closing total", "running balance"], reference: ["reference", "ref no", "utr", "cheque no"],
          instrument: ["instrument"], amount: ["amount"], direction: ["dr/cr", "type", "direction"]
        };
        const found = headers.findIndex(value => (aliases[key] || []).includes(String(value || "").trim().toLowerCase()));
        select.value = found >= 0 ? String(found) : "-1";
      });
      renderMappingPreview(grid, headerRow);
    }
    headerSelect.onchange = refreshColumns;
    refreshColumns();

    $("#mappingSubmit").onclick = async event => {
      event.preventDefault();
      const mapping = {};
      fields.querySelectorAll("select[data-map]").forEach(select => mapping[select.dataset.map] = Number(select.value));
      if (mapping.date < 0 || mapping.particulars < 0) {
        showMappingError("Mapping requires Date and Particulars columns.");
        return;
      }
      if (mapping.debit < 0 && mapping.credit < 0 && mapping.amount < 0 && mapping.balance < 0) {
        showMappingError("Select Debit/Credit, Single Amount, or Balance.");
        return;
      }
      mappingError.classList.add("hidden");
      submitButton.disabled = true;
      submitButton.textContent = "Applying...";
      try {
        const response = await fetch("/api/map", {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ filename: meta.filename, grid, headerRow: Number(headerSelect.value), mapping, bankLedger: bankLedgerInput.value.trim(), profileName: $("#profileName").value, save: true, chargeToken: meta.charge_token })
        });
        const result = await response.json();
        if (!response.ok) {
          showMappingError(result.error);
          return;
        }
        await refreshLicense();
        dialog.close();
        finish(result.rows);
      } catch (error) {
        showMappingError(error.message);
      } finally {
        submitButton.disabled = false;
        submitButton.textContent = "Apply & Save Mapping";
      }
    };
    const cancel = dialog.querySelector("button[value='cancel']");
    cancel.onclick = event => { event.preventDefault(); dialog.close(); finish(null); };
    dialog.oncancel = event => { event.preventDefault(); dialog.close(); finish(null); };
    dialog.showModal();
  });
}

function columnName(index) {
  let name = "";
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) name = String.fromCharCode(65 + ((value - 1) % 26)) + name;
  return name;
}

function renderMappingPreview(grid, headerRow) {
  const start = Math.max(0, headerRow);
  const sample = grid.slice(start, start + 8);
  $("#mappingPreview").innerHTML = sample.map((row, rowIndex) => `<tr>${row.map(value => rowIndex === 0 ? `<th>${esc(value)}</th>` : `<td>${esc(value)}</td>`).join("")}</tr>`).join("");
}

function money(value, showZero = false) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || (!showZero && !numeric)) return "";
  return numeric.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function esc(value) { return String(value ?? "").replace(/[&<>"']/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char])); }
function cleanLedgerName(value) {
  return String(value || "")
    .replace(/&#13;&#10;|&#10;|&#13;/gi, " ")
    .replace(/[\r\n\t\u0004\u000b\f]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function sanitizeAllLedgerNames() {
  rows.forEach(row => {
    row.bank_ledger = cleanLedgerName(row.bank_ledger);
    row.counter_ledger = cleanLedgerName(row.counter_ledger);
  });
}

function update(index, key, value) {
  if (key === "counter_ledger" || key === "bank_ledger") value = cleanLedgerName(value);
  rows[index][key] = value;
  if (key === "approval") render();
}

function applyBankLedgerToRows(showNotice = false) {
  const ledger = bankLedgerInput.value.trim();
  if (!ledger) {
    showMessage("Tally bank ledger cannot be blank.", true);
    bankLedgerInput.focus();
    return false;
  }
  let changed = 0;
  rows.forEach(row => {
    const previous = String(row.bank_ledger || "");
    if (previous === ledger) return;
    row.bank_ledger = ledger;
    changed++;
  });
  if (changed) render();
  if (showNotice && changed) showMessage(`Bank ledger updated in ${changed} transaction${changed === 1 ? "" : "s"}.`);
  return true;
}

async function autoMatchBankLedgers() {
  if (!rows.length || !tallyMasters.connected) return 0;
  const response = await fetch("/api/tally/match-ledgers", {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ particulars: rows.map(row => row.particulars || "") })
  });
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || "Tally ledger matching failed.");
  let matched = 0;
  (result.matches || []).forEach(match => {
    const row = rows[Number(match.index)];
    if (!row || !match.matched || !match.ledger) return;
    if (row.counter_ledger === "Suspense" || !row.counter_ledger) {
      row.counter_ledger = cleanLedgerName(match.ledger);
      row.tally_match_score = match.score;
      row.confidence = match.score >= 0.8 ? "High" : "Review";
      matched += 1;
    }
  });
  if (matched) {
    render();
    showMessage(`${matched} bank transaction(s) matched with synced Tally ledgers. Please review before import.`);
  }
  return matched;
}

function transactionFingerprint(row) {
  const text = value => String(value || "").trim().toLowerCase().replace(/\s+/g, " ");
  const numeric = value => Number(value || 0).toFixed(2);
  return [
    text(row.date),
    numeric(row.debit),
    numeric(row.credit),
    numeric(row.balance),
    text(row.reference),
    text(row.instrument),
    text(row.particulars)
  ].join("|");
}

// Boundary transactions can repeat across consecutive statement PDFs with
// slightly different narration. Match date + running balance + amount so that
// different rows that land on the same balance (e.g. several 0.00 closures and
// a 217.12 folio charge on one day) are not treated as duplicates.
function transactionBalanceFingerprint(row) {
  const hasBalance = row.balance_available === true ||
    (row.balance !== null && row.balance !== undefined && row.balance !== "");
  if (!hasBalance || !row.date) return "";
  return [
    String(row.date).trim(),
    Number(row.balance || 0).toFixed(2),
    Number(row.debit || 0).toFixed(2),
    Number(row.credit || 0).toFixed(2)
  ].join("|");
}

function statementDateKey(value) {
  const text = String(value || "").trim();
  let match = text.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (match) return `${match[3]}-${match[2]}-${match[1]}`;
  match = text.match(/^(\d{2})[/.](\d{2})[/.](\d{4})$/);
  if (match) return `${match[3]}-${match[2]}-${match[1]}`;
  return text;
}

function chronologicalRows(sourceRows) {
  const hasBalance = row => row.balance_available === true ||
    (row.balance !== null && row.balance !== undefined && row.balance !== "");
  const balanceRows = sourceRows.filter(hasBalance);
  const newestFirst = balanceRows.length > 1 &&
    statementDateKey(balanceRows[0].date) > statementDateKey(balanceRows[balanceRows.length - 1].date);
  return balanceRows.map((row, index) => ({ row, index })).sort((a, b) => {
    const dateOrder = statementDateKey(a.row.date).localeCompare(statementDateKey(b.row.date));
    if (dateOrder) return dateOrder;
    const orderA = Number(a.row._statementOrder);
    const orderB = Number(b.row._statementOrder);
    if (!Number.isNaN(orderA) && !Number.isNaN(orderB) && orderA !== orderB) return orderA - orderB;
    return newestFirst ? b.index - a.index : a.index - b.index;
  }).map(item => item.row);
}

function sourceBalance(row) {
  return Number(row._sourceBalance ?? row.balance ?? 0);
}

function recalculateAllStatementAmounts() {
  for (const summary of statementSummaries) {
    const ordered = rows
      .filter(row => row._statementId === summary.statementId)
      .sort((a, b) => Number(a._statementOrder || 0) - Number(b._statementOrder || 0));
    if (!ordered.length) continue;
    const first = ordered[0];
    let previous;
    if (Number(first._statementOrder || 0) === 0) {
      previous = Number((summary.sourceOpening ?? summary.opening ??
        (sourceBalance(first) + Number(first.debit || 0) - Number(first.credit || 0))).toFixed(2));
    } else {
      previous = Number((sourceBalance(first) + Number(first.debit || 0) - Number(first.credit || 0)).toFixed(2));
    }
    for (const row of ordered) {
      const balance = sourceBalance(row);
      const change = Number((balance - previous).toFixed(2));
      row.debit = change < 0 ? Math.abs(change) : 0;
      row.credit = change > 0 ? change : 0;
      previous = balance;
    }
  }
}

function prepareStatementRows(fileRows, sourceName = "") {
  const normalizedName = String(sourceName || "").trim().toLowerCase();
  if (normalizedName) {
    const replacedIds = new Set(rows
      .filter(row => row._statementFile === normalizedName)
      .map(row => row._statementId));
    if (replacedIds.size) {
      rows = rows.filter(row => !replacedIds.has(row._statementId));
      statementSummaries = statementSummaries.filter(item => !replacedIds.has(item.statementId));
    }
  }
  const statementId = ++statementSequence;
  chronologicalRows(fileRows).forEach((row, order) => {
    row._statementId = statementId;
    row._statementFile = normalizedName;
    row._statementOrder = order;
    const hasBalance = row.balance_available === true ||
      (row.balance !== null && row.balance !== undefined && row.balance !== "");
    if (hasBalance) {
      row._sourceBalance = Number(row.balance || 0);
    }
  });
}

function addUniqueRows(newRows) {
  const known = new Set(rows.map(transactionFingerprint));
  const knownBalances = new Set(rows.map(transactionBalanceFingerprint).filter(Boolean));
  let duplicates = 0;
  newRows.forEach(row => {
    const fingerprint = transactionFingerprint(row);
    const balanceFingerprint = transactionBalanceFingerprint(row);
    if (known.has(fingerprint) || (balanceFingerprint && knownBalances.has(balanceFingerprint))) {
      duplicates++;
      return;
    }
    known.add(fingerprint);
    if (balanceFingerprint) knownBalances.add(balanceFingerprint);
    rows.push(row);
  });
  return duplicates;
}

function rebalanceStatementBalances() {
  const groups = new Map();
  rows.forEach(row => {
    if (!row._statementId) return;
    if (!groups.has(row._statementId)) groups.set(row._statementId, []);
    groups.get(row._statementId).push(row);
  });
  if (!groups.size || !statementSummaries.length) return;

  const ordered = [...statementSummaries].sort((a, b) =>
    statementDateKey(a.startDate).localeCompare(statementDateKey(b.startDate)) ||
    statementDateKey(a.endDate).localeCompare(statementDateKey(b.endDate))
  );
  let previousClosing = null;
  for (const summary of ordered) {
    const group = groups.get(summary.statementId) || [];
    const sourceOpening = summary.sourceOpening ?? summary.opening ?? 0;
    const sourceClosing = summary.sourceClosing ?? summary.closing ?? sourceOpening;
    const offset = previousClosing === null ? 0 : Number((previousClosing - sourceOpening).toFixed(2));
    group.forEach(row => {
      if (row._sourceBalance === undefined && (row.balance_available === true || Number(row.balance))) {
        row._sourceBalance = Number((Number(row.balance || 0) - Number(summary.balanceOffset || 0)).toFixed(2));
      }
      if (row._sourceBalance !== undefined) {
        row.balance = Number((row._sourceBalance + offset).toFixed(2));
      }
    });
    summary.opening = Number((sourceOpening + offset).toFixed(2));
    summary.closing = Number((sourceClosing + offset).toFixed(2));
    summary.balanceOffset = offset;
    previousClosing = summary.closing;
  }
}

function sortRowsByDate() {
  rows.sort((a, b) =>
    statementDateKey(a.date).localeCompare(statementDateKey(b.date)) ||
    Number(a._statementId || 0) - Number(b._statementId || 0) ||
    Number(a._statementOrder || 0) - Number(b._statementOrder || 0)
  );
}

function balanceSummaryForRows(sourceRows, meta = {}) {
  const balanceRows = sourceRows.filter(row => row.balance_available === true || Number(row.balance));
  if (!balanceRows.length) return { opening: null, closing: null };
  // Statements may list transactions oldest-first or newest-first. Work on a
  // chronological copy so a newest-first PDF does not show its closing balance
  // as the opening balance.
  const chronological = chronologicalRows(balanceRows);
  const first = chronological[0];
  const last = chronological[chronological.length - 1];
  return {
    startDate: String(first.date || ""),
    endDate: String(last.date || ""),
    opening: meta.opening !== undefined ? Number(meta.opening) :
      Number(first.balance || 0) + Number(first.debit || 0) - Number(first.credit || 0),
    closing: meta.closing !== undefined ? Number(meta.closing) : Number(last.balance || 0)
  };
}

function rememberStatementBalance(fileRows, meta = {}) {
  const summary = balanceSummaryForRows(fileRows, meta);
  const statementId = fileRows[0]?._statementId;
  if (summary.opening === null || !statementId) return;
  const entry = {
    ...summary,
    statementId,
    sourceOpening: meta.opening !== undefined ? Number(meta.opening) : summary.opening,
    sourceClosing: meta.closing !== undefined ? Number(meta.closing) : summary.closing
  };
  const existing = statementSummaries.findIndex(item => item.statementId === statementId);
  if (existing >= 0) statementSummaries[existing] = entry;
  else statementSummaries.push(entry);
}

function statementBalances() {
  if (!statementSummaries.length) return balanceSummaryForRows(rows);
  const ordered = [...statementSummaries].sort((a, b) =>
    statementDateKey(a.startDate).localeCompare(statementDateKey(b.startDate)) ||
    statementDateKey(a.endDate).localeCompare(statementDateKey(b.endDate))
  );
  const latest = [...statementSummaries].sort((a, b) =>
    statementDateKey(a.endDate).localeCompare(statementDateKey(b.endDate)) ||
    statementDateKey(a.startDate).localeCompare(statementDateKey(b.startDate))
  ).pop();
  return { opening: ordered[0].opening, closing: latest.closing };
}

function render() {
  sanitizeAllLedgerNames();
  recalculateAllStatementAmounts();
  rebalanceStatementBalances();
  $("#mTotal").textContent = rows.length.toLocaleString("en-IN");
  for (const value of ["Receipt", "Payment", "Contra"]) $("#m" + value).textContent = rows.filter(row => row.voucher_type === value).length.toLocaleString("en-IN");
  const readyBox = $("#mReady") || $("#mPending");
  if (readyBox) readyBox.textContent = rows.filter(row => row.approval === "Ready").length.toLocaleString("en-IN");
  const balances = statementBalances();
  $("#totalDebit").textContent = money(rows.reduce((total, row) => total + Number(row.debit || 0), 0), true);
  $("#totalCredit").textContent = money(rows.reduce((total, row) => total + Number(row.credit || 0), 0), true);
  $("#openingBalance").textContent = balances.opening === null ? "—" : money(balances.opening, true);
  $("#closingBalance").textContent = balances.closing === null ? "—" : money(balances.closing, true);
  const visible = rows.map((row, index) => ({ row, index })).filter(({ row }) => filter === "All" || row.approval === filter);
  $("#rows").innerHTML = visible.map(({ row, index }) => `<tr><td>${esc(row.date)}</td><td>${esc(row.particulars)}</td><td class="money">${money(row.debit)}</td><td class="money">${money(row.credit)}</td><td class="money">${money(row.balance, row.balance_available === true)}</td><td><select onchange="update(${index},'voucher_type',this.value)">${["Receipt", "Payment", "Contra"].map(value => `<option ${row.voucher_type === value ? "selected" : ""}>${value}</option>`).join("")}</select></td><td><input value="${esc(row.counter_ledger)}" onchange="update(${index},'counter_ledger',this.value)"></td><td><textarea onchange="update(${index},'narration',this.value)">${esc(row.narration)}</textarea></td><td><select class="status ${row.approval.toLowerCase()}" onchange="update(${index},'approval',this.value)"><option ${row.approval === "Ready" ? "selected" : ""}>Ready</option><option ${row.approval === "Hold" ? "selected" : ""}>Hold</option></select></td></tr>`).join("");
}

async function download(path) {
  if (!applyBankLedgerToRows()) return;
  if (path.endsWith("/xml") && !balancesReconciled()) return;
  const response = await fetch(path, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ rows }) });
  if (!response.ok) { const error = await response.json(); showMessage(error.error, true); return; }
  const blob = await response.blob();
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = path.endsWith("xlsx") ? "Bank2Tally_Review.xlsx" : "Bank2Tally_Import.xml";
  link.click();
  URL.revokeObjectURL(link.href);
}

function statementChainCheck(groupRows) {
  const ordered = [...groupRows].sort((a, b) => Number(a._statementOrder || 0) - Number(b._statementOrder || 0));
  if (!ordered.length) return null;
  const first = ordered[0];
  const last = ordered[ordered.length - 1];
  const debit = groupRows.reduce((total, row) => total + Number(row.debit || 0), 0);
  const credit = groupRows.reduce((total, row) => total + Number(row.credit || 0), 0);
  const rowOpening = Number((sourceBalance(first) + Number(first.debit || 0) - Number(first.credit || 0)).toFixed(2));
  const rowClosing = sourceBalance(last);
  const calculatedClosing = Number((rowOpening + credit - debit).toFixed(2));
  return { rowOpening, rowClosing, calculatedClosing, startDate: first.date, endDate: last.date };
}

function balancesReconciled() {
  recalculateAllStatementAmounts();
  rebalanceStatementBalances();
  const balances = statementBalances();
  if (balances.opening === null || balances.closing === null) {
    showMessage("Opening and Closing Balance are required before Tally export.", true);
    return false;
  }

  if (statementSummaries.length > 1) {
    for (const summary of statementSummaries) {
      const groupRows = rows.filter(row => row._statementId === summary.statementId);
      const check = statementChainCheck(groupRows);
      if (!check) continue;
      if (Math.abs(check.calculatedClosing - check.rowClosing) > 0.05) {
        const label = summary.startDate && summary.endDate ? `${summary.startDate} to ${summary.endDate}` : `${check.startDate} to ${check.endDate}`;
        showMessage(`Tally export stopped: ${label} has inconsistent debit/credit amounts. Calculated closing ${money(check.calculatedClosing, true)} does not match running balance ${money(check.rowClosing, true)}. Export Review Excel for checking.`, true);
        return false;
      }
    }
    return true;
  }

  const check = statementChainCheck(rows);
  if (!check) {
    showMessage("Opening and Closing Balance are required before Tally export.", true);
    return false;
  }
  if (Math.abs(check.calculatedClosing - check.rowClosing) > 0.05) {
    showMessage(`Tally export stopped: calculated closing ${money(check.calculatedClosing, true)} does not match running balance ${money(check.rowClosing, true)}. Export Review Excel for checking.`, true);
    return false;
  }
  if (Math.abs(check.rowClosing - Number(balances.closing)) > 0.05) {
    showMessage(`Tally export stopped: last running balance ${money(check.rowClosing, true)} does not match statement closing ${money(balances.closing, true)}. Export Review Excel for checking.`, true);
    return false;
  }
  return true;
}

async function resolveTallyLedgerNames() {
  if (!tallyMasters.connected || !rows.length) return 0;
  const names = [...new Set(rows.flatMap(row => [row.bank_ledger, row.counter_ledger]).filter(Boolean))];
  if (!names.length) return 0;
  const response = await fetch("/api/tally/resolve-ledgers", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ names })
  });
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || "Tally ledger resolve failed.");
  const lookup = new Map((result.resolved || []).map(item => [item.source, item.ledger]));
  let changed = 0;
  rows.forEach(row => {
    for (const key of ["bank_ledger", "counter_ledger"]) {
      const current = cleanLedgerName(row[key]);
      const resolved = lookup.get(current);
      if (resolved && resolved !== current) {
        row[key] = resolved;
        changed++;
      }
    }
  });
  if (changed) render();
  return changed;
}

function markPartialTallyRows(missingLedgers) {
  const missing = new Set((missingLedgers || []).map(name => cleanLedgerName(name).toLowerCase()).filter(Boolean));
  if (!missing.size) return 0;
  let retryCount = 0;
  rows.forEach(row => {
    if (row.approval !== "Ready") return;
    const counter = cleanLedgerName(row.counter_ledger).toLowerCase();
    if (missing.has(counter)) {
      retryCount += 1;
      return;
    }
    row.approval = "Hold";
  });
  if (retryCount) render();
  return retryCount;
}

async function sendDirectlyToTally() {
  if (!applyBankLedgerToRows() || !balancesReconciled()) return;
  sanitizeAllLedgerNames();
  if (tallyMasters.connected) await resolveTallyLedgerNames();
  const readyCount = rows.filter(row => row.approval === "Ready").length;
  if (!readyCount) {
    showMessage("No Ready rows left to send. Change Hold rows back to Ready if you need to retry.", true);
    return;
  }
  if (!window.confirm("Take a Tally company backup first. Send all Ready transactions directly to the currently open Tally company?")) return;
  const button = $("#tallyDirectBtn");
  button.disabled = true;
  button.textContent = "Sending to Tally...";
  startTallySendProgress(`Sending ${readyCount} bank voucher(s)`, sendDirectlyToTally);
  try {
    const response = await fetch("/api/tally/send", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ rows })
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "Tally import failed.");
    const created = Number(result.created || 0);
    const altered = Number(result.altered || 0);
    const ignored = Number(result.ignored || 0);
    const errors = Number(result.errors || 0) + Number(result.exceptions || 0);
    if (errors > 0) {
      const details = Array.isArray(result.details) && result.details.length
        ? ` First reason: ${result.details[0]}`
        : "";
      const retryCount = markPartialTallyRows(result.missing_ledgers || []);
      const errorMessage = `Tally accepted ${created + altered} voucher(s), but reported ${errors} error(s). Ignored: ${ignored}.${details}`;
      showMessage(errorMessage, true);
      finishTallySendProgress(false, errorMessage, {
        retryCount,
        retryAction: retryCount ? sendDirectlyToTally : null,
      });
      return;
    }
    rows.forEach(row => {
      if (row.approval === "Ready") row.approval = "Hold";
    });
    render();
    const successMessage = `${created + altered} bank voucher(s) successfully accepted by Tally. Created: ${created}, altered: ${altered}, ignored: ${ignored}.`;
    showMessage(successMessage);
    finishTallySendProgress(true, successMessage);
  } catch (error) {
    showMessage(error.message, true);
    finishTallySendProgress(false, error.message);
  } finally {
    button.disabled = false;
    button.textContent = "Send Directly to Tally";
  }
}

function displayTallyDate(value) {
  const text = String(value || "");
  return text.length === 8 ? `${text.slice(6, 8)}-${text.slice(4, 6)}-${text.slice(0, 4)}` : text;
}

async function openUndoImport() {
  const response = await fetch("/api/tally/history", {
    method: "POST", headers: { "Content-Type": "application/json" }, body: "{}"
  });
  const result = await response.json();
  if (!response.ok) {
    showMessage(result.error || "Could not load Tally import history.", true);
    return;
  }
  const select = $("#undoBatch");
  select.innerHTML = (result.batches || []).map(batch =>
    `<option value="${esc(batch.batch_id)}">${esc(batch.bank_ledger)} | ${displayTallyDate(batch.from_date)} to ${displayTallyDate(batch.to_date)} | ${Number(batch.count).toLocaleString("en-IN")} vouchers</option>`
  ).join("");
  if (!select.options.length) {
    showMessage("No undoable Bank2Tally direct-import batch was found. Older imports without a Batch ID cannot be deleted automatically.", true);
    return;
  }
  $("#undoPin").value = "";
  $("#undoError").classList.add("hidden");
  $("#undoDialog").showModal();
}

async function confirmUndoImport() {
  const batchId = $("#undoBatch").value;
  const pin = $("#undoPin").value.trim();
  const errorBox = $("#undoError");
  if (!/^(?:\d{4}|\d{6})$/.test(pin)) {
    errorBox.textContent = "Enter your 4 or 6 digit Login PIN.";
    errorBox.classList.remove("hidden");
    return;
  }
  const label = $("#undoBatch").selectedOptions[0]?.textContent || "selected batch";
  if (!window.confirm(`Take a Tally company backup first. Permanently delete ${label}?`)) return;
  const button = $("#undoConfirmBtn");
  button.disabled = true;
  button.textContent = "Deleting from Tally...";
  try {
    const response = await fetch("/api/tally/undo", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ batchId, pin })
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "Undo failed.");
    $("#undoDialog").close();
    showMessage(`Undo complete: ${Number(result.deleted).toLocaleString("en-IN")} Bank2Tally vouchers deleted from Tally.`);
  } catch (error) {
    errorBox.textContent = error.message;
    errorBox.classList.remove("hidden");
  } finally {
    button.disabled = false;
    button.textContent = "Delete This Batch";
  }
}
$("#excelBtn").onclick = () => download("/api/export/xlsx");
$("#xmlBtn").onclick = () => download("/api/export/xml");
$("#tallyDirectBtn").onclick = sendDirectlyToTally;
$("#undoTallyBtn").onclick = openUndoImport;
$("#undoConfirmBtn").onclick = confirmUndoImport;

// Sales / Credit / Debit bulk voucher builder
let bulkVoucherRows = [];
const bulkAutoParties = ["Aarav Sharma","Ananya Das","Arjun Roy","Diya Singh","Ishaan Gupta","Kavya Paul","Neha Verma","Rahul Dey","Riya Traders","Saanvi Enterprise","Siddharth Store","Vikram Agency"];

function bulkMoney(value) { return Number(value || 0).toLocaleString("en-IN", {minimumFractionDigits:2, maximumFractionDigits:2}); }
function bulkNumber(id) { return Number($(id).value || 0); }
function bulkInvoiceNumber(start, index) {
  const match = String(start || "INV0001").match(/^(.*?)(\d+)$/);
  if (!match) return `${start || "INV"}${index + 1}`;
  return `${match[1]}${String(Number(match[2]) + index).padStart(match[2].length, "0")}`;
}
function bulkDateAt(start, end, index, count) {
  const first = new Date(`${start}T12:00:00`), last = new Date(`${end}T12:00:00`);
  const days = Math.max(0, Math.round((last - first) / 86400000));
  const offset = count <= 1 ? 0 : Math.round(days * index / (count - 1));
  const date = new Date(first.getTime() + offset * 86400000);
  return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,"0")}-${String(date.getDate()).padStart(2,"0")}`;
}
function distributeBulkTotal(groups, total, weights) {
  const low = groups.map(g => Math.round(g.min * 100));
  const high = groups.map(g => Math.round(g.max * 100));
  const wanted = Math.round(total * 100);
  const minimum = low.reduce((a,b)=>a+b,0), maximum = high.reduce((a,b)=>a+b,0);
  if (wanted < minimum || wanted > maximum) throw new Error(`Total Amount must be between ₹${bulkMoney(minimum/100)} and ₹${bulkMoney(maximum/100)} for these rules.`);
  const result = [...low]; let left = wanted - minimum;
  const active = new Set(result.map((_,i)=>i));
  while (left > 0 && active.size) {
    const weightTotal = [...active].reduce((sum,i)=>sum + Math.max(0.01, Number(weights[i] || 1)),0);
    let used = 0;
    for (const i of [...active]) {
      const capacity = high[i] - result[i];
      const add = Math.min(capacity, Math.max(1, Math.floor(left * Math.max(0.01, Number(weights[i] || 1)) / weightTotal)));
      result[i] += add; left -= add; used += add;
      if (result[i] >= high[i]) active.delete(i);
      if (left <= 0) break;
    }
    if (!used) break;
  }
  let cursor = 0;
  while (left > 0) { const i = cursor++ % result.length; if (result[i] < high[i]) { result[i]++; left--; } }
  return result.map(value => value / 100);
}
function bulkTaxParts(amount, rate, taxType) {
  const taxable = Number((amount / (1 + rate / 100)).toFixed(2));
  const tax = Number((amount - taxable).toFixed(2));
  if (taxType === "igst") return {taxable, igst:tax, cgst:0, sgst:0};
  const cgst = Number((tax / 2).toFixed(2));
  return {taxable, igst:0, cgst, sgst:Number((tax-cgst).toFixed(2))};
}
async function readBulkImport() {
  const file = $("#bulkExcelFile").files[0];
  if (!file) return [];
  const response = await fetch("/api/bulk-vouchers/parse", {method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({files:[{name:file.name,data:await fileToBase64(file)}]})});
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || "Could not read the Excel file.");
  return result.rows || [];
}
function renderBulkPreview() {
  const body = $("#bulkPreviewRows"); body.innerHTML = "";
  bulkVoucherRows.forEach((row,index) => {
    const tr = document.createElement("tr");
    const itemList = row.use_ledger_entry ? "tallyLedgerList" : "tallyItemList";
    tr.innerHTML = `<td><input type="checkbox" class="bulk-row-select" data-index="${index}" ${row.selected?"checked":""}></td><td>${row.bulk_type}</td><td><input class="bulk-edit" data-field="invoice_no" data-index="${index}" value="${escapeHtml(row.invoice_no)}"></td><td><input class="bulk-edit" data-field="invoice_date" data-index="${index}" type="date" value="${row.invoice_date}"></td><td><input class="bulk-edit bulk-party-edit" data-field="party_ledger" data-index="${index}" list="tallyLedgerList" value="${escapeHtml(row.party_ledger)}"></td><td><input class="bulk-edit" data-field="item_name" data-index="${index}" list="${itemList}" value="${escapeHtml(row.item_name)}"></td><td>${bulkMoney(row.invoice_value)}</td><td>${bulkMoney(row.taxable_value)}</td><td>${row.gst_rate}%</td><td>${bulkMoney(row.igst)}</td><td>${bulkMoney(row.cgst)}</td><td>${bulkMoney(row.sgst)}</td>`;
    body.appendChild(tr);
  });
  body.querySelectorAll(".bulk-row-select").forEach(input => input.onchange = () => { bulkVoucherRows[Number(input.dataset.index)].selected = input.checked; updateBulkSummary(); });
  body.querySelectorAll(".bulk-edit").forEach(input => input.onchange = () => {
    const row = bulkVoucherRows[Number(input.dataset.index)]; row[input.dataset.field] = input.value;
    if (input.dataset.field === "item_name") row.sales_allocations[0].item_name = input.value;
  });
  $("#bulkPreviewWrap").classList.remove("hidden"); $("#bulkSelectAllBtn").classList.remove("hidden"); $("#bulkSendBtn").classList.remove("hidden");
  updateBulkSummary();
}
function updateBulkSummary() {
  const selected = bulkVoucherRows.filter(row=>row.selected), total = selected.reduce((sum,row)=>sum+row.invoice_value,0);
  $("#bulkSummary").innerHTML = `<strong>${selected.length} of ${bulkVoucherRows.length} selected</strong><span>Selected Total ₹${bulkMoney(total)}</span>`;
  $("#bulkSummary").classList.remove("hidden");
  $("#bulkSelectAllBtn").textContent = selected.length === bulkVoucherRows.length ? "Unselect All" : "Select All";
}
async function generateBulkPreview() {
  const error = $("#bulkError"); error.classList.add("hidden");
  const button = $("#bulkGenerateBtn"); button.disabled = true; button.textContent = "Generating...";
  try {
    const imported = await readBulkImport();
    const count = imported.length || Math.floor(bulkNumber("#bulkEntryCount"));
    if (!count) throw new Error("Enter Number of Entries or upload an Excel file.");
    const cashCount = Math.min(count, Math.max(0, Math.floor(bulkNumber("#bulkCashCount"))));
    const total = bulkNumber("#bulkTotalAmount") || imported.reduce((sum,row)=>sum+Number(row.amount||0),0);
    if (total <= 0) throw new Error("Enter the Total Amount.");
    const personMin=bulkNumber("#bulkPersonMin"), personMax=bulkNumber("#bulkPersonMax"), cashMin=bulkNumber("#bulkCashMin"), cashMax=bulkNumber("#bulkCashMax");
    if (personMax < personMin || cashMax < cashMin) throw new Error("Amount To cannot be less than Amount From.");
    const groups = Array.from({length:count},(_,i)=>i<cashCount?{min:cashMin,max:cashMax}:{min:personMin,max:personMax});
    const amounts = distributeBulkTotal(groups,total,Array.from({length:count},(_,i)=>Number(imported[i]?.amount||((i%7)+1))));
    const names = $("#bulkPartyNames").value.split(/\r?\n|,/).map(v=>v.trim()).filter(Boolean);
    const startDate=$("#bulkStartDate").value, endDate=$("#bulkEndDate").value;
    if (!startDate || !endDate || endDate < startDate) throw new Error("Choose a valid Starting Date and Ending Date.");
    const type=$("#bulkVoucherType").value, rate=bulkNumber("#bulkGstRate"), taxType=$("#bulkTaxType").value;
    bulkVoucherRows = amounts.map((amount,i)=>{
      const source=imported[i]||{}, sourceAllocations=Array.isArray(source.sales_allocations)?source.sales_allocations:[], rowRate=Number(source.gst_rate||source.rate||rate), parts=sourceAllocations.length?{
        taxable:sourceAllocations.reduce((s,a)=>s+Number(a.taxable_value||0),0),
        igst:sourceAllocations.reduce((s,a)=>s+Number(a.igst||0),0),
        cgst:sourceAllocations.reduce((s,a)=>s+Number(a.cgst||0),0),
        sgst:sourceAllocations.reduce((s,a)=>s+Number(a.sgst||0),0)
      }:bulkTaxParts(amount,rowRate,taxType);
      const party = String(source.party||source.party_name||"").trim() || (i<cashCount?"Cash":(names.length?names[(i-cashCount)%names.length]:bulkAutoParties[(i-cashCount)%bulkAutoParties.length]));
      const invoiceNo=String(source.invoice_no||bulkInvoiceNumber($("#bulkStartInvoice").value,i));
      const invoiceDate=String(source.date||bulkDateAt(startDate,endDate,i,count)).slice(0,10);
      const item=String(source.item||$("#bulkItemName").value||`${rate}% Items`), quantity=Math.max(.001,Number(source.quantity||bulkNumber("#bulkQuantity")||1));
      const useLedgerEntry = type === "Credit Note";
      const allocations=sourceAllocations.length?sourceAllocations.map(a=>({...a,item_name:item,use_ledger_entry:useLedgerEntry,preserve_item_name:true,unit:String(a.unit||"Pcs")})):[{taxable_value:parts.taxable,rate:rowRate,item_name:item,use_ledger_entry:useLedgerEntry,preserve_item_name:true,quantity,unit:"Pcs",hsn:String(source.hsn||""),igst:parts.igst,cgst:parts.cgst,sgst:parts.sgst,cess:0}];
      return {selected:true,bulk_type:type,invoice_no:invoiceNo,invoice_date:invoiceDate,party_name:party,party_ledger:party,gstin:String(source.gstin||""),invoice_value:Number(source.amount||amount),taxable_value:parts.taxable,igst:parts.igst,cgst:parts.cgst,sgst:parts.sgst,cess:allocations.reduce((s,a)=>s+Number(a.cess||0),0),gst_rate:rowRate,item_name:item,use_ledger_entry:useLedgerEntry,ready_for_sales_tally:type==="Sales Invoice",ready_for_note_tally:type!=="Sales Invoice",sales_allocations:allocations};
    });
    renderBulkPreview();
  } catch (failure) { error.textContent=failure.message; error.classList.remove("hidden"); }
  finally { button.disabled=false; button.textContent="Generate Preview"; }
}
async function sendBulkVouchers() {
  const rows=bulkVoucherRows.filter(row=>row.selected); if(!rows.length)return alert("Select at least one voucher.");
  if(!confirm(`Take a Tally backup first. Send ${rows.length} voucher(s) to Tally?`))return;
  const type=$("#bulkVoucherType").value, ledgerName=$("#bulkLedgerName").value.trim(), rate=bulkNumber("#bulkGstRate");
  const ledgers={salesLedger0:"Sales 0%",salesLedger5:"Sales 5%",salesLedger12:"Sales 12%",salesLedger18:"Sales 18%",salesLedger28:"Sales 28%",igstLedger:"Output IGST",cgstLedger:"Output CGST",sgstLedger:"Output SGST",roundLedger:"Round Off"};
  ledgers[`salesLedger${rate}`]=ledgerName||`Sales ${rate}%`;
  if(type==="Credit Note"){
    ledgers.igstLedger=`Output IGST @ ${rate}%`;
    ledgers.cgstLedger=`Output CGST @ ${rate/2}%`;
    ledgers.sgstLedger=`Output SGST @ ${rate/2}%`;
  }
  const button=$("#bulkSendBtn");button.disabled=true;startTallySendProgress(`Sending ${rows.length} ${type} voucher(s)`,sendBulkVouchers);
  try {
    if (type === "Sales Invoice") {
      // Sales: verify vs Tally, then send ONLY missing vouchers one-by-one.
      const verifyInfo = await fetchMissingSalesVouchers(rows.map(row => ({ ...row, selected: true, ready_for_sales_tally: true })));
      const missing = verifyInfo.missing;
      if (!missing.length) {
        finishTallySendProgress(true, "", {
          counts: {
            created_confirmed: 0,
            already_exists_count: verifyInfo.already_exists_count || rows.length,
            missing_sent_count: 0,
            still_missing_count: 0,
            tally_sales_count_before: verifyInfo.tally_sales_count,
            tally_sales_count_after: verifyInfo.tally_sales_count,
            missing_only_mode: true,
            selected_total: 0,
          },
        });
        return;
      }
      if (!confirm(`About to send ${missing.length} missing vouchers\n\nAlready in Tally: ${verifyInfo.already_exists_count}\nTally Sales: ${verifyInfo.tally_sales_count}`)) {
        finishTallySendProgress(false, "Send cancelled.");
        return;
      }
      const sendInfo = await sendMissingSalesBulkFast(missing, ledgers);
      finishMissingSalesProgress(verifyInfo, sendInfo);
    } else {
      const endpoint = "/api/gst/notes/tally/send";
      const extra = { voucherType: type, tallyVoucherType: $("#bulkTallyVoucherType").value.trim(), ledgers };
      const result = await sendTallyInBatches(endpoint, rows, extra, type);
      const done = (result.created || 0) + (result.altered || 0);
      finishTallySendProgress(true, `${done} ${type} voucher(s) successfully accepted by Tally.`);
    }
  }
  catch(failure){finishTallySendProgress(false,failure.message);}
  finally{button.disabled=false;}
}
$("#openBulkSalesModule").onclick=()=>{const today=new Date().toISOString().slice(0,10);if(!$("#bulkStartDate").value)$("#bulkStartDate").value=today;if(!$("#bulkEndDate").value)$("#bulkEndDate").value=today;$("#bulkSalesDialog").showModal();};
function updateBulkEntryMode(){
  const type=$("#bulkVoucherType").value, credit=type==="Credit Note", input=$("#bulkItemName");
  $("#bulkTallyVoucherType").value = credit ? "Credit Note B2B" : (type==="Debit Note" ? "Debit Note" : "Sales");
  $("#bulkItemNameLabel").textContent = credit ? "Income Ledger" : "Item Name";
  input.setAttribute("list", credit ? "tallyLedgerList" : "tallyItemList");
  input.placeholder = credit ? "Select synced Tally income ledger" : "Tally stock item";
  if(credit && input.value==="Items") input.value="";
}
$("#bulkVoucherType").onchange=updateBulkEntryMode;
$("#bulkGstRate").onchange=()=>{$("#bulkLedgerName").value=`Sales ${$("#bulkGstRate").value}%`;};
$("#bulkGenerateBtn").onclick=generateBulkPreview;
$("#bulkSelectAllBtn").onclick=()=>{const all=bulkVoucherRows.every(row=>row.selected);bulkVoucherRows.forEach(row=>row.selected=!all);renderBulkPreview();};
$("#bulkSendBtn").onclick=sendBulkVouchers;
