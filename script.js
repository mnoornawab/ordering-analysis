function processFile() {
  const fileInput = document.getElementById("file-input");
  const file = fileInput.files[0];
  if (!file) {
    alert("Please select an Excel file.");
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    const simaSheet = XLSX.utils.sheet_to_json(workbook.Sheets["Quantity on Order - SIMA System"]);
    const allocationSheet = XLSX.utils.sheet_to_json(workbook.Sheets["Allocation File"]);
    const ordersSheet = XLSX.utils.sheet_to_json(workbook.Sheets["Orders-SIMA System"]);

    // Step 1: Load SIMA as base
    const simaMap = {};
    simaSheet.forEach(row => {
      const code = String(row["Item Code"]).trim();
      if (!code || code.toLowerCase() === "undefined") return;
      simaMap[code] = {
        "Item Code": code,
        "Style Code": row["Style"] || "",
        "Qty On Order": Number(row["Qty On Order"]) || 0
      };
    });

    // Step 2: Merge allocation quantities
    const allocationTotals = {};
    allocationSheet.forEach(row => {
      const code = String(row["Material Code"]).trim();
      const qty = Number(row["Pending Order Qty"]) || 0;
      if (!allocationTotals[code]) allocationTotals[code] = 0;
      allocationTotals[code] += qty;
    });

    // Step 3: Merge orders balances
    const ordersTotals = {};
    ordersSheet.forEach(row => {
      const code = String(row["ITEMCODE"]).trim();
      const balance = Number(row["BALANCE"]) || 0;
      if (!ordersTotals[code]) ordersTotals[code] = 0;
      ordersTotals[code] += balance;
    });

    // Step 4: Merge all into result array
    const results = Object.keys(simaMap).map(code => {
      return {
        "Item Code": code,
        "Style Code": simaMap[code]["Style Code"],
        "Qty On Order": simaMap[code]["Qty On Order"],
        "On Allocation File": allocationTotals[code] || 0,
        "Balance from Orders": ordersTotals[code] || 0
      };
    });

    displayTable(results);
  };

  reader.readAsArrayBuffer(file);
}

function displayTable(data) {
  const container = document.getElementById("report");
  container.innerHTML = "";

  if (!data.length) {
    container.innerText = "No records to show.";
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
    headers.forEach(h => {
      const td = document.createElement("td");
      td.innerText = row[h];
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  container.appendChild(table);
}
