let finalResults = [];
let sampleOnlyResults = [];

function processFile() {
  const fileInput = document.getElementById("file-input");
  const file = fileInput.files[0];

  if (!file) {
    alert("Please select an Excel file first.");
    return;
  }

  document.getElementById("status").innerText = "⏳ Processing file...";
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
      document.getElementById("status").innerText = "❌ Error reading sheets: " + err;
    }
  };

  reader.readAsArrayBuffer(file);
}

function analyzeData(simaData, allocationData, ordersData) {
  const simaMap = {};
  simaData.forEach(row => {
    const code = String(row["Item Code"]).trim();
    simaMap[code] = Number(row["Qty On Order"] || 0);
  });

  const allocationMap = {};
  const sampleItems = new Set();

  allocationData.forEach(row => {
    const code = String(row["Material code"]).trim();
    const qty = Number(row["Pending order qty"] || 0);
    const po = String(row["PO Reference"] || "").toLowerCase().trim();
    if (!po) return;

    if (!allocationMap[code]) allocationMap[code] = { qty: 0, po_refs: [] };
    allocationMap[code].qty += qty;
    if (!allocationMap[code].po_refs.includes(po)) {
      allocationMap[code].po_refs.push(po);
    }

    if (po.includes("sample")) {
      sampleItems.add(code);
    }
  });

  const orderMap = {};
  ordersData.forEach(row => {
    const code = String(row["Item Code"]).trim();
    if (!orderMap[code]) {
      orderMap[code] = { total: 0, reserved: 0, confirmed: 0, balance: 0 };
    }
    orderMap[code].total += Number(row["Total Qty Ordered"] || 0);
    orderMap[code].reserved += Number(row["Reserved"] || 0);
    orderMap[code].confirmed += Number(row["Confirmed"] || 0);
    orderMap[code].balance += Number(row["Balance"] || 0);
  });

  finalResults = [];
  sampleOnlyResults = [];

  Object.keys(orderMap).forEach(code => {
    const order = orderMap[code];
    const allocation = allocationMap[code] || { qty: 0, po_refs: [] };
    const onOrder = simaMap[code] || 0;
    const isSample = sampleItems.has(code);

    const qtyToOrder = Math.max(0, order.balance - allocation.qty - onOrder);

    const result = {
      "Item Code": code,
      "Total Ordered": order.total,
      "Reserved": order.reserved,
      "Confirmed": order.confirmed,
      "Balance": order.balance,
      "Allocation Qty": allocation.qty,
      "PO Refs": allocation.po_refs.join(", "),
      "On Order Qty": onOrder,
      "Qty to Order": qtyToOrder,
      "Sample PO?": isSample ? "Yes" : "No",
      "Comment": ""
    };

    if (isSample) {
      result["Comment"] = "Sample PO – Check before ordering";
      sampleOnlyResults.push(result);
    } else if (qtyToOrder > 0) {
      result["Comment"] = "Needs to be ordered";
      finalResults.push(result);
    } else {
      result["Comment"] = "OK or Already on Order";
      finalResults.push(result);
    }
  });

  displayResults(finalResults, sampleOnlyResults);
}

function displayResults(normal, samples) {
  const container = document.getElementById("results-container");
  container.innerHTML = "";

  const resultTitle = document.createElement("h3");
  resultTitle.innerText = "📋 Main Order Analysis";
  container.appendChild(resultTitle);
  container.appendChild(createTable(normal));

  if (samples.length > 0) {
    const sampleTitle = document.createElement("h3");
    sampleTitle.innerText = "⚠️ Sample PO Items (Separate)";
    container.appendChild(sampleTitle);
    container.appendChild(createTable(samples, true));
  }

  document.getElementById("status").innerText = `✅ Analysis complete. ${normal.length + samples.length} items processed.`;
  document.getElementById("download-btn").style.display = "inline-block";
}

function createTable(data, isSample = false) {
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

    // row coloring
    if (row["Sample PO?"] === "Yes") {
      tr.style.backgroundColor = "#fff8dc"; // light yellow
    }
    if (row["Comment"] === "Needs to be ordered") {
      tr.style.backgroundColor = "#ffe6e6"; // light red
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
  return table;
}

function downloadCSV() {
  const allData = finalResults.concat(sampleOnlyResults);
  const headers = Object.keys(allData[0]);
  const csv = [
    headers.join(","),
    ...allData.map(row => headers.map(h => `"${row[h]}"`).join(","))
  ].join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "ordering_report.csv";
  link.click();
}
