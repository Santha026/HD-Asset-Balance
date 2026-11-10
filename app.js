"use strict";

const state = { raw: [], value: [], master: [], results: [], quality: {}, toastTimer: null };
const aliases = {
  raw: {
    assetId: ["assetid", "asseid", "assettagnumber", "tagnumber", "assetcode"],
    hdCode: ["hdcode", "hdid"],
    description: ["assetdescription", "assetdescipion", "description", "assetdesc"],
    capacity: ["capacity", "capaciy"],
    deliveryDate: ["deliverydate", "deliedate", "dateofdelivery", "deldate"]
  },
  value: { assetId: ["assetid", "asseid", "assettagnumber", "tagnumber", "assetcode"], assetValue: ["assetvalue", "bookvalue", "value"] },
  master: {
    hdCode: ["hdcode", "hdid"],
    hdName: ["hdname", "name"],
    vendorCode: ["vendorcode"],
    location: ["location"],
    sdmName: ["sdmname"],
    rsmName: ["rsmname"],
    zsmName: ["zsmname"],
    region: ["region", "state"],
    securityDeposit: ["securitydeposit", "securitydeposite", "securitydepositamount", "securitydepositeamount"],
    totalDeposit: ["totaldeposit", "totaldeposite", "deposittotal"]
  }
};
const requiredFields = {
  raw: ["assetId", "hdCode"],
  value: ["assetId", "assetValue"],
  master: ["hdCode", "securityDeposit", "totalDeposit"]
};

const $ = (selector) => document.querySelector(selector);
const normalized = (value) => String(value ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
const text = (value) => value == null ? "" : String(value).trim();
const number = (value) => {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const parsed = Number(String(value ?? "").replace(/[₹,\s]/g, "").replace(/^\((.*)\)$/, "-$1"));
  return Number.isFinite(parsed) ? parsed : 0;
};
const formatMoney = (value) => new Intl.NumberFormat("en-IN", {
  style: "currency", currency: "INR", maximumFractionDigits: 2
}).format(value || 0);
const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, (character) => ({
  "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
})[character]);

function showToast(message) {
  const toast = $("#toast");
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(state.toastTimer);
  state.toastTimer = setTimeout(() => toast.classList.remove("show"), 3600);
}

function setAlert(message = "") {
  const alert = $("#upload-alert");
  alert.textContent = message;
  alert.hidden = !message;
}

function findHeader(headers, fieldAliases) {
  return headers.find((header) => fieldAliases.includes(normalized(header)));
}

function parseRows(workbook, kind) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("The workbook does not contain a worksheet.");
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false, dateNF: "yyyy-mm-dd" });
  if (rows.length === 0) throw new Error("The first worksheet is empty.");
  const headers = Object.keys(rows[0]);
  const fields = aliases[kind];
  const columns = {};
  for (const [field, candidates] of Object.entries(fields)) {
    columns[field] = findHeader(headers, candidates);
    if (!columns[field] && requiredFields[kind].includes(field)) {
      throw new Error(`Missing required ${kind} column: ${candidates[0]}.`);
    }
  }
  return rows.map((row) => {
    const result = {};
    for (const field of Object.keys(fields)) result[field] = row[columns[field]];
    return result;
  }).filter((row) => kind === "master" ? text(row.hdCode) : text(row.assetId));
}

async function readFile(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: "array", cellDates: true });
}

function getAssetKey(value) {
  return normalized(value);
}

function calculate() {
  if (!state.raw.length || !state.value.length || !state.master.length) {
    state.results = [];
    render();
    return;
  }
  const valueById = new Map();
  for (const item of state.value) {
    const key = getAssetKey(item.assetId);
    if (!key) continue;
    valueById.set(key, (valueById.get(key) || 0) + number(item.assetValue));
  }
  const assetsByHd = new Map();
  let unmatched = 0;
  for (const item of state.raw) {
    const code = normalized(item.hdCode);
    if (!code) continue;
    const key = getAssetKey(item.assetId);
    const assetValue = valueById.get(key);
    if (assetValue === undefined) unmatched++;
    const list = assetsByHd.get(code) || [];
    list.push({ ...item, assetValue: assetValue ?? 0 });
    assetsByHd.set(code, list);
  }
  const seenHdCodes = new Set();
  const duplicateHds = [];
  state.results = state.master.map((hd) => {
    const codeKey = normalized(hd.hdCode);
    if (seenHdCodes.has(codeKey)) duplicateHds.push(text(hd.hdCode));
    seenHdCodes.add(codeKey);
    const assets = assetsByHd.get(codeKey) || [];
    const assetValue = assets.reduce((sum, asset) => sum + asset.assetValue, 0);
    const securityDeposit = number(hd.securityDeposit);
    const totalDeposit = number(hd.totalDeposit);
    return {
      ...hd,
      assets,
      assetCount: assets.length,
      assetValue,
      securityDeposit,
      totalDeposit,
      balance: totalDeposit - (assetValue + securityDeposit)
    };
  });
  state.quality = { unmatched, duplicateHds: [...new Set(duplicateHds)] };
  render();
}

function render() {
  const query = normalized($("#search-input").value);
  const filtered = state.results.filter((hd) => !query ||
    normalized(hd.hdCode).includes(query) || normalized(hd.hdName).includes(query));
  const hasData = state.results.length > 0;
  $("#empty-state").hidden = filtered.length > 0;
  $("#metric-hds").textContent = hasData ? state.results.length.toLocaleString("en-IN") : "—";
  $("#metric-assets").textContent = hasData ? state.results.reduce((sum, hd) => sum + hd.assetCount, 0).toLocaleString("en-IN") : "—";
  $("#metric-value").textContent = hasData ? formatMoney(state.results.reduce((sum, hd) => sum + hd.assetValue, 0)) : "—";
  const totalBalance = state.results.reduce((sum, hd) => sum + hd.balance, 0);
  $("#metric-balance").textContent = hasData ? formatMoney(totalBalance) : "—";
  $("#metric-balance").classList.toggle("negative", totalBalance < 0);
  $("#report-caption").textContent = hasData ? `${filtered.length.toLocaleString("en-IN")} HD${filtered.length === 1 ? "" : "s"} · Balance = Total deposit − (Asset value + Security deposit)` : "Upload all three files to build your report.";
  $("#export-button").disabled = !hasData;
  $("#rows-count").textContent = hasData ? `Showing ${filtered.length.toLocaleString("en-IN")} of ${state.results.length.toLocaleString("en-IN")} HDs` : "Waiting for files";
  const warnings = [];
  if (state.quality.unmatched) warnings.push(`${state.quality.unmatched} asset${state.quality.unmatched === 1 ? "" : "s"} without a matching value`);
  if (state.quality.duplicateHds?.length) warnings.push(`${state.quality.duplicateHds.length} duplicate HD code${state.quality.duplicateHds.length === 1 ? "" : "s"}`);
  $("#data-quality").textContent = warnings.join(" · ");
  $("#report-body").innerHTML = filtered.map((hd) => `
    <tr>
      <td class="hd-cell"><strong>${escapeHtml(hd.hdCode)}</strong><span>${escapeHtml(hd.hdName || "—")}</span></td>
      <td>${escapeHtml(hd.location || hd.region || "—")}</td>
      <td class="number-cell">${hd.assetCount.toLocaleString("en-IN")}</td>
      <td class="number-cell">${formatMoney(hd.assetValue)}</td>
      <td class="number-cell">${formatMoney(hd.securityDeposit)}</td>
      <td class="number-cell">${formatMoney(hd.totalDeposit)}</td>
      <td class="number-cell balance ${hd.balance < 0 ? "negative" : "positive"}">${formatMoney(hd.balance)}</td>
      <td><button class="detail-button" type="button" data-hd="${escapeHtml(hd.hdCode)}">Details</button></td>
    </tr>`).join("");
  $("#empty-state").querySelector("strong").textContent = hasData ? "No HDs match your search" : "Your HD report is ready to build";
  $("#empty-state").querySelector("span:last-child").textContent = hasData ? "Try a different HD code or name." : "Upload your Asset Raw, Asset Value and HD Master files above to see balances here.";
}

function openDetails(code) {
  const hd = state.results.find((item) => item.hdCode === code);
  if (!hd) return;
  $("#modal-title").textContent = `${hd.hdCode} · ${hd.hdName || "HD assets"}`;
  $("#modal-summary").innerHTML = [
    `Location: ${escapeHtml(hd.location || hd.region || "—")}`,
    `Assets: ${hd.assetCount.toLocaleString("en-IN")}`,
    `Asset value: ${formatMoney(hd.assetValue)}`,
    `Security deposit: ${formatMoney(hd.securityDeposit)}`,
    `Total deposit: ${formatMoney(hd.totalDeposit)}`,
    `Balance: ${formatMoney(hd.balance)}`
  ].map((item) => `<span>${item}</span>`).join("");
  $("#asset-body").innerHTML = hd.assets.length ? hd.assets.map((asset, index) => `
    <tr><td>${index + 1}</td><td>${escapeHtml(asset.assetId)}</td><td>${escapeHtml(asset.description || "—")}</td>
    <td>${escapeHtml(asset.capacity || "—")}</td><td>${escapeHtml(asset.deliveryDate || "—")}</td>
    <td class="number-cell">${formatMoney(asset.assetValue)}</td></tr>`).join("") :
    '<tr><td colspan="6">No assets are assigned to this HD.</td></tr>';
  $("#detail-modal").hidden = false;
  $("#close-modal").focus();
}

function csvCell(value) {
  const safe = String(value ?? "");
  return `"${safe.replace(/"/g, '""')}"`;
}

function exportReport() {
  const columns = ["HD Code", "HD Name", "Vendor Code", "Location", "SDM Name", "RSM Name", "ZSM Name", "Region", "Asset Count", "Total Asset Value", "Security Deposit", "Total Deposit", "Balance"];
  const rows = state.results.map((hd) => [
    hd.hdCode, hd.hdName, hd.vendorCode, hd.location, hd.sdmName, hd.rsmName, hd.zsmName, hd.region,
    hd.assetCount, hd.assetValue, hd.securityDeposit, hd.totalDeposit, hd.balance
  ]);
  const csv = [columns, ...rows].map((row) => row.map(csvCell).join(",")).join("\r\n");
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "hd-asset-balance-report.csv";
  link.click();
  URL.revokeObjectURL(link.href);
}

for (const kind of ["raw", "value", "master"]) {
  $(`#${kind}-file`).addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!window.XLSX) {
      setAlert("Excel reader could not load. Check your internet connection and reload the page, or upload a CSV file.");
      return;
    }
    setAlert("");
    try {
      const rows = parseRows(await readFile(file), kind);
      if (!rows.length) throw new Error("No usable data rows were found.");
      state[kind] = rows;
      $(`#${kind}-status`).textContent = `${file.name} · ${rows.length.toLocaleString("en-IN")} rows`;
      $(`[data-kind="${kind}"]`).classList.add("has-file");
      calculate();
      showToast(`${file.name} loaded successfully.`);
    } catch (error) {
      setAlert(`${file.name}: ${error.message}`);
      event.target.value = "";
    }
  });
}

$("#search-input").addEventListener("input", () => render());
$("#export-button").addEventListener("click", exportReport);
$("#report-body").addEventListener("click", (event) => {
  const button = event.target.closest("[data-hd]");
  if (button) openDetails(button.dataset.hd);
});
$("#close-modal").addEventListener("click", () => { $("#detail-modal").hidden = true; });
$("#detail-modal").addEventListener("click", (event) => {
  if (event.target === $("#detail-modal")) $("#detail-modal").hidden = true;
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") $("#detail-modal").hidden = true;
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
    event.preventDefault();
    $("#search-input").focus();
  }
});

render();
