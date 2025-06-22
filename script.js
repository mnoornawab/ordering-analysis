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

    const allocationTotals = {};
    allocationSheet.forEach(row => {
      const code = String(row["Material Code"]).trim();
      const qty = Number(row["Pending Order Qty"]) || 0;
      if (!allocationTotals[code]) allocationTotals[code] = 0;
      allocationTotals[code] += qty;
    });

    const ordersTotals = {};
    ordersSheet.forEach(row => {
      const code = String(row["ITEMCODE"]).trim();
      const balance = Number(row["BALANCE"]) || 0;
      if (!ordersTotals[code]) ordersTotals[code] = 0;
      ordersTotals[code] += balance;
    });

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

  const checkbox = document.createElement("label");
  checkbox.innerHTML = `
    <input type="checkbox" id="hide-zero" onchange="filterTable()"> Hide rows with all 0 values
  `;
  container.appendChild(checkbox);

  const table = document.createElement("table");
  table.id = "results-table";

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

  let totalOrder = 0, totalAlloc = 0, totalBalance = 0;

  data.forEach(row => {
    const tr = document.createElement("tr");
    tr.classList.add("data-row");

    let rowSum = 0;
    headers.forEach(h => {
      const td = document.createElement("td");
      td.innerText = row[h];
      tr.appendChild(td);

      if (h === "Qty On Order") totalOrder += row[h];
      if (h === "On Allocation File") totalAlloc += row[h];
      if (h === "Balance from Orders") totalBalance += row[h];
      if (["Qty On Order", "On Allocation File", "Balance from Orders"].includes(h)) {
        rowSum += row[h];
      }
    });

    if (rowSum === 0) tr.classList.add("all-zero");
    tbody.appendChild(tr);
  });

  // Totals row
  const totalRow = document.createElement("tr");
  headers.forEach(h => {
    const td = document.createElement("td");
    if (h === "Style Code") td.innerText = "TOTAL";
    else if (h === "Qty On Order") td.innerText = totalOrder;
    else if (h === "On Allocation File") td.innerText = totalAlloc;
    else if (h === "Balance from Orders") td.innerText = totalBalance;
    else td.innerText = "";
    totalRow.appendChild(td);
  });
  tbody.appendChild(totalRow);

  table.appendChild(thead);
  table.appendChild(tbody);
  container.appendChild(table);
}

function filterTable() {
  const hide = document.getElementById("hide-zero").checked;
  document.querySelectorAll(".data-row").forEach(row => {
    row.style.display = (hide && row.classList.contains("all-zero")) ? "none" : "";
  });
}
