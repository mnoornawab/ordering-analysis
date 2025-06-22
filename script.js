// ✨ Updated script.js to include Mismatch Report tab

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

    const results = Object.keys(simaMap).map(code => {
      return {
        "Item Code": code,
        "Style Code": simaMap[code]["Style Code"],
        "Qty On Order": simaMap[code]["Qty On Order"],
        "On Allocation File": allocationTotals[code] || 0
      };
    });

    // Build Mismatch Report from above
    const mismatches = results.filter(row => row["Qty On Order"] !== row["On Allocation File"])
      .map(row => ({
        ...row,
        "Mismatch Type": "SIMA vs Allocation"
      }));

    displayMainReport(results);
    displayMismatchReport(mismatches);
  };

  reader.readAsArrayBuffer(file);
}

function displayMainReport(data) {
  // ... (existing logic from earlier displayTable function)
  // Create a table for "main-report" div
  // Use logic already provided in earlier steps
}

function displayMismatchReport(mismatches) {
  const container = document.getElementById("mismatch-report");
  container.innerHTML = "";

  if (!mismatches.length) {
    container.innerHTML = "<p>No mismatches found between SIMA and Allocation.</p>";
    return;
  }

  const table = document.createElement("table");
  const headers = Object.keys(mismatches[0]);

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  headers.forEach(h => {
    const th = document.createElement("th");
    th.innerText = h;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const tbody = document.createElement("tbody");
  mismatches.forEach(row => {
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

  const downloadBtn = document.createElement("button");
  downloadBtn.innerText = "Download Mismatch Report";
  downloadBtn.onclick = function () {
    const ws = XLSX.utils.json_to_sheet(mismatches);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Mismatch Report");
    XLSX.writeFile(wb, "Mismatch_Report.xlsx");
  };
  container.appendChild(downloadBtn);
}
