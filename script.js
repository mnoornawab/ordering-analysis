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
    simaMap[code] = Number(row["Qty On Order"] || 0);
  });

  const allocationMap = {};
  allocationData.forEach(row => {
    const code = String(row["Material code"]).trim();
    const qty = Number(row["Pending order qty"] || 0);
    if (!allocationMap[code]) allocationMap[code] = 0;
    allocationMap[code] += qty;
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

  Object.keys(orderMap).forEach(code => {
    const order = orderMap[code];
    const allocationQty = allocationMap[code] || 0;
    const onOrder = simaMap[code] || 0;

    const qtyToOrder = Math.max(0, order.balance - allocationQty - onOrder);
    const comment = qtyToOrder > 0 ? "Needs to be ordered" : "OK or Already on Order";

    finalResults.push({
      "Item Code": code,
      "Total Ordered": order.total,
      "Reserved": order.reserved,
      "Confirmed": order.confirmed,
      "Balance": order.balance,
      "Allocation Qty": allocationQty,
      "On Order Qty": onOrder,
      "Qty to Order": qtyToOrder,
      "Comment": comment
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
  container.appendChild(table);

  document.getElementById("status").innerText = `✅ Found ${data.length} items.`;
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
  link.download = "ordering_report.csv";
  link.click();
}
