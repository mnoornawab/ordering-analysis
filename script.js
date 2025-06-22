let finalResults = [];

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
    const qty = Number(row["Qty On Order"] || 0);
    simaMap[code] = qty;
  });

  const allocationMap = {};
  const samplePOs = {};

  allocationData.forEach(row => {
    const code = String(row["Material code"]).trim();
    const qty = Number(row["Pending order qty"] || 0);
    const po = String(row["PO Reference"] || "").toLowerCase().trim();

    if (!allocationMap[code]) allocationMap[code] = { total: 0, real: 0, sample: [] };
    allocationMap[code].total += qty;

    if (po.includes("sample")) {
      allocationMap[code].sample.push({ po: po, qty: qty });
      if (!samplePOs[code]) samplePOs[code] = [];
      samplePOs[code].push(`${po} (${qty})`);
    } else {
      allocationMap[code].real += qty;
    }
  });

  const ordersMap = {};
  ordersData.forEach(row => {
    const code = String(row["Item Code"]).trim();
    if (!ordersMap[code]) {
      ordersMap[code] = { total: 0, reserved: 0, confirmed: 0 };
    }
    ordersMap[code].total += Number(row["Total Qty Ordered"] || 0);
    ordersMap[code].reserved += Number(row["Reserved"] || 0);
    ordersMap[code].confirmed += Number(row["Confirmed"] || 0);
  });

  finalResults = [];

  Object.keys(simaMap).forEach(code => {
    const onOrderQty = simaMap[code] || 0;
    const order = ordersMap[code] || { total: 0, reserved: 0, confirmed: 0 };
    const allocation = allocationMap[code] || { total: 0, real: 0, sample: [] };

    const balance = order.total - order.reserved - order.confirmed;
    const qtyToOrder = Math.max(0, balance - onOrderQty - allocation.real);
    const mismatch = onOrderQty !== allocation.total;

    finalResults.push({
      "Item Code": code,
      "Total Ordered": order.total,
      "Reserved": order.reserved,
      "Confirmed": order.confirmed,
      "Balance": balance,
      "On Order (SIMA)": onOrderQty,
      "Allocated Qty (Real)": allocation.real,
      "Qty to Order": qtyToOrder,
      "Allocation Mismatch?": mismatch ? "⚠️ Yes" : "No",
      "Sample PO Summary": (samplePOs[code] || []).join(", ")
    });
  });

  displayResults(finalResults);
}

function displayResults(data) {
  const container = document.getElementById("results-container");
  container.innerHTML = "";

  if (!data.length) {
    container.innerText = "No matching records.";
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

    // Highlight rules
    if (row["Qty to Order"] > 0) {
      tr.style.backgroundColor = "#ffe6e6"; // red
    } else if (row["Allocation Mismatch?"] === "⚠️ Yes") {
      tr.style.backgroundColor = "#fff3cd"; // yellow
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

  document.getElementById("status").innerText = `✅ Analysis complete. ${data.length} items processed.`;
  document.getElementById("download-btn").style.display = "inline-block";
}

function downloadCSV() {
  const headers = Object.keys(finalResults[0]);
  const csv = [
    headers.join(","),
    ...finalResults.map(row => headers.map(h => `"${row[h]}"`).join(","))
  ].join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "ordering_analysis_report.csv";
  link.click();
}
