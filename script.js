let mismatchResults = [];
let toOrderResults = [];
let filteredToOrder = [];
let showOnlyToOrder = false;

function processFile() {
  const fileInput = document.getElementById("file-input");
  const file = fileInput.files[0];

  if (!file) {
    alert("Please select an Excel file.");
    return;
  }

  document.getElementById("status").innerText = "⏳ Processing...";
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    try {
      const simaSheet = workbook.Sheets["Quantity on Order - SIMA System"];
      const allocationSheet = workbook.Sheets["Allocation File"];
      const ordersSheet = workbook.Sheets["Orders-SIMA System"];

      const simaData = XLSX.utils.sheet_to_json(simaSheet);
      const allocationData = XLSX.utils.sheet_to_json(allocationSheet);
      const ordersData = XLSX.utils.sheet_to_json(ordersSheet);

      analyzeData(simaData, allocationData, ordersData);
    } catch (err) {
      document.getElementById("status").innerText = "❌ Sheet Error: " + err;
    }
  };

  reader.readAsArrayBuffer(file);
}

function analyzeData(simaData, allocationData, ordersData) {
  const simaMap = {};
  simaData.forEach(row => {
    const code = String(row["Item Code"] || "").trim();
    if (!code || code.toLowerCase() === "undefined") return;
    const style = row["Material Code"] || "";
    const qty = Number(row["Qty On Order"] || 0);
    simaMap[code] = { qty, style };
  });

  const allocationMap = {};
  allocationData.forEach(row => {
    const code = String(row["Material Code"] || "").trim();
    if (!code || code.toLowerCase() === "undefined") return;
    const qty = Number(row["Pending order qty"] || 0);
    if (!allocationMap[code]) allocationMap[code] = 0;
    allocationMap[code] += qty;
  });

  const balanceMap = {};
  const styleCodeMap = {};
  ordersData.forEach(row => {
    const code = String(row["ITEMCODE"] || "").trim();
    if (!code || code.toLowerCase() === "undefined") return;
    const balance = Number(row["BALANCE"] || 0);
    const style = row["STYLECODE"] || "";
    if (!balanceMap[code]) balanceMap[code] = 0;
    balanceMap[code] += balance;
    if (!styleCodeMap[code]) styleCodeMap[code] = style;
  });

  mismatchResults = [];
  toOrderResults = [];
  filteredToOrder = [];

  const allCodes = new Set([...Object.keys(simaMap), ...Object.keys(allocationMap)]);
  allCodes.forEach(code => {
    const simaQty = simaMap[code]?.qty || 0;
    const allocQty = allocationMap[code] || 0;
    const style = simaMap[code]?.style || styleCodeMap[code] || "";

    let status = "OK";
    if (!simaMap[code]) status = "Only in Allocation File";
    else if (!allocationMap[code]) status = "Only in SIMA System";
    else if (simaQty !== allocQty) status = "Qty Mismatch";

    if (status !== "OK") {
      mismatchResults.push({
        "Item Code": code,
        "Style Code": style,
        "SIMA Qty": simaQty,
        "Allocation Qty": allocQty,
        "Status": status
      });
    }
  });

  Object.keys(balanceMap).forEach(code => {
    const balance = balanceMap[code];
    const allocQty = allocationMap[code] || 0;
    const simaQty = simaMap[code]?.qty || 0;
    const style = simaMap[code]?.style || styleCodeMap[code] || "";
    const qtyToOrder = balance - allocQty - simaQty;

    if (!code || code === "undefined") return;

    const result = {
      "Item Code": code,
      "Style Code": style,
      "Balance from Orders": balance,
      "Qty Allocated": allocQty,
      "Qty on Order (SIMA)": simaQty,
      "Qty to Order": qtyToOrder > 0 ? qtyToOrder : 0
    };

    toOrderResults.push(result);
  });

  filteredToOrder = toOrderResults.filter(row => row["Qty to Order"] > 0);

  displayTable("mismatch-table", mismatchResults, true);
  displayToOrderTable();
  document.getElementById("status").innerText = "✅ Analysis complete.";
  document.getElementById("download-btn").style.display = "inline-block";
  document.getElementById("filter-btn").style.display = "inline-block";
}

function displayToOrderTable() {
  const data = showOnlyToOrder ? filteredToOrder : toOrderResults;
  displayTable("order-table", data, false);
}

function displayTable(containerId, data, highlightMismatch) {
  const container = document.getElementById(containerId);
  container.innerHTML = "";

  if (!data.length) {
    container.innerText = "No records to display.";
    return;
  }

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");

  const headers = Object.keys(data[0]);
  const headerRow = document.createElement("tr");
  headers.forEach(h => {
    const th = document.createElement("th");
    th.innerText = h;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const numericTotals = {};

  data.forEach(row => {
    if (!row["Item Code"] || row["Item Code"] === "undefined") return;
    const tr = document.createElement("tr");

    if (highlightMismatch && row["Status"] !== "OK") {
      tr.classList.add("highlight-missing");
    } else if (!highlightMismatch && row["Qty to Order"] > 0) {
      tr.classList.add("highlight-to-order");
    }

    headers.forEach(key => {
      const td = document.createElement("td");
      td.innerText = row[key];
      tr.appendChild(td);

      if (!isNaN(row[key]) && typeof row[key] === "number") {
        if (!numericTotals[key]) numericTotals[key] = 0;
        numericTotals[key] += Number(row[key]);
      }
    });

    tbody.appendChild(tr);
  });

  // Totals row
  const totalsRow = document.createElement("tr");
  headers.forEach(key => {
    const td = document.createElement("td");
    if (numericTotals[key]) {
      td.innerText = `Total: ${numericTotals[key]}`;
      td.style.fontWeight = "bold";
    }
    totalsRow.appendChild(td);
  });
  tbody.appendChild(totalsRow);

  table.appendChild(thead);
  table.appendChild(tbody);
  container.appendChild(table);
}

function toggleFilter() {
  showOnlyToOrder = !showOnlyToOrder;
  document.getElementById("filter-btn").innerText = showOnlyToOrder ? "Show All" : "Show Only To Order";
  displayToOrderTable();
}

function downloadCSV() {
  const wb = XLSX.utils.book_new();
  const mismatchSheet = XLSX.utils.json_to_sheet(mismatchResults);
  const toOrderSheet = XLSX.utils.json_to_sheet(toOrderResults);
  XLSX.utils.book_append_sheet(wb, mismatchSheet, "Mismatch Report");
  XLSX.utils.book_append_sheet(wb, toOrderSheet, "To Order Report");
  XLSX.writeFile(wb, "stock_analysis_report.xlsx");
}

function openTab(tabId) {
  document.querySelectorAll(".tab-button").forEach(btn => btn.classList.remove("active"));
  document.querySelectorAll(".tab-content").forEach(tab => tab.classList.remove("active-tab"));
  document.querySelector(`.tab-button[onclick*="${tabId}"]`).classList.add("active");
  document.getElementById(tabId).classList.add("active-tab");
}

