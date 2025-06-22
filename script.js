let mismatchResults = [];
let toOrderResults = [];

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
    const item = String(row["Item Code"]).trim();
    simaMap[item] = {
      style: row["Style Code"] || "",
      qty: Number(row["Qty On Order"] || 0)
    };
  });

  const allocationMap = {};
  allocationData.forEach(row => {
    const item = String(row["Material code"]).trim();
    const qty = Number(row["Pending order qty"] || 0);
    if (!allocationMap[item]) allocationMap[item] = 0;
    allocationMap[item] += qty;
  });

  const ordersMap = {};
  ordersData.forEach(row => {
    const item = String(row["Item Code"]).trim();
    if (!ordersMap[item]) {
      ordersMap[item] = { ordered: 0, reserved: 0, confirmed: 0 };
    }
    ordersMap[item].ordered += Number(row["Total Qty Ordered"] || 0);
    ordersMap[item].reserved += Number(row["Reserved"] || 0);
    ordersMap[item].confirmed += Number(row["Confirmed"] || 0);
  });

  mismatchResults = [];
  toOrderResults = [];

  const allItemCodes = new Set([
    ...Object.keys(simaMap),
    ...Object.keys(allocationMap)
  ]);

  allItemCodes.forEach(code => {
    const simaQty = simaMap[code]?.qty || 0;
    const allocationQty = allocationMap[code] || 0;
    const style = simaMap[code]?.style || "";

    let status = "OK";
    if (!simaMap[code]) {
      status = "Only in Allocation File";
    } else if (!allocationMap[code]) {
      status = "Only in SIMA System";
    } else if (simaQty !== allocationQty) {
      status = "Qty Mismatch";
    }

    if (status !== "OK") {
      mismatchResults.push({
        "Item Code": code,
        "Style Code": style,
        "SIMA Qty": simaQty,
        "Allocation Qty": allocationQty,
        "Status": status
      });
    }
  });

  Object.keys(ordersMap).forEach(code => {
    const order = ordersMap[code];
    const balance = order.ordered - order.reserved - order.confirmed;
    const allocationQty = allocationMap[code] || 0;

    if (balance > allocationQty) {
      toOrderResults.push({
        "Item Code": code,
        "Total Ordered": order.ordered,
        "Reserved": order.reserved,
        "Confirmed": order.confirmed,
        "Balance": balance,
        "Allocated Qty": allocationQty,
        "Qty to Order": balance - allocationQty
      });
    }
  });

  displayTable("mismatch-table", mismatchResults, true);
  displayTable("order-table", toOrderResults, false);
  document.getElementById("status").innerText = "✅ Done.";
  document.getElementById("download-btn").style.display = "inline-block";
}

function displayTable(containerId, data, highlightMismatch) {
  const container = document.getElementById(containerId);
  container.innerHTML = "";

  if (!data.length) {
    container.innerText = "No issues found.";
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

  data.forEach(row => {
    const tr = document.createElement("tr");

    if (highlightMismatch && row["Status"] && row["Status"] !== "OK") {
      tr.classList.add("highlight-missing");
    } else if (!highlightMismatch) {
      tr.classList.add("highlight-to-order");
    }

    headers.forEach(key => {
      const td = document.createElement("td");
      td.innerText = row[key];
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  container.appendChild(table);
}

function downloadCSV() {
  const allData = [...mismatchResults, ...toOrderResults];
  if (!allData.length) return;

  const headers = Object.keys(allData[0]);
  const csv = [
    headers.join(","),
    ...allData.map(row => headers.map(h => `"${row[h]}"`).join(","))
  ].join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "stock_analysis_report.csv";
  link.click();
}

function openTab(tabId) {
  document.querySelectorAll(".tab-button").forEach(btn => btn.classList.remove("active"));
  document.querySelectorAll(".tab-content").forEach(tab => tab.classList.remove("active-tab"));

  document.querySelector(`.tab-button[onclick*="${tabId}"]`).classList.add("active");
  document.getElementById(tabId).classList.add("active-tab");
}
