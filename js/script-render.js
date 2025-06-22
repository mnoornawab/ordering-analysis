function getColumnValue(row, possibleHeaders) {
    for (const header of possibleHeaders) {
        if (header in row) return row[header];
        for (const key in row) {
            if (key.trim().toLowerCase() === header.trim().toLowerCase()) {
                return row[key];
            }
        }
    }
    return undefined;
}

function downloadTemplate() {
    const templateSheets = {
        "Quantity on Order - SIMA System": [
            ['Item Code', 'Style', 'Qty on Order'],
            ['A1', 'S100', 10],
            ['A2', 'S200', 5]
        ],
        "Allocation File": [
            ['Item Code', 'Pending Order Qty'],
            ['A1', 8],
            ['A1', 2],
            ['A2', 4],
            ['A2', 1]
        ],
        "Orders-SIMA System": [
            ['Item Code', 'Style', 'BALANCE'],
            ['A1', 'S100', 12],
            ['A2', 'S200', 3],
            ['A2', 'S200', 2]
        ]
    };
    exportToExcel(templateSheets, 'Stock_Order_Template.xlsx');
}

let analyzedData = [];
let allocationMap = {};
let ordersMap = {};
let styleMap = {};
let qtyData = [];
let allocData = [];
let ordersData = [];

function analyzeFile() {
    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];
    if (!file) {
        alert("Please upload a file first.");
        return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const qtySheet = workbook.Sheets["Quantity on Order - SIMA System"];
        const allocSheet = workbook.Sheets["Allocation File"];
        const ordersSheet = workbook.Sheets["Orders-SIMA System"];
        if (!qtySheet || !allocSheet || !ordersSheet) {
            alert("Excel file is missing required sheets. Please use the template.");
            return;
        }

        qtyData = XLSX.utils.sheet_to_json(qtySheet);
        allocData = XLSX.utils.sheet_to_json(allocSheet);
        ordersData = XLSX.utils.sheet_to_json(ordersSheet);

        // Allocation Map: sum by Item Code
        allocationMap = {};
        allocData.forEach(row => {
            const item = getColumnValue(row, ["Item Code"]);
            const qty = Number(getColumnValue(row, ["Pending Order Qty", "Pending order qty"]) || 0);
            if (!item) return;
            allocationMap[item] = (allocationMap[item] || 0) + qty;
        });

        // Orders Map: sum by Item Code
        ordersMap = {};
        ordersData.forEach(row => {
            const item = getColumnValue(row, ["Item Code"]);
            const qty = Number(getColumnValue(row, ["BALANCE", "Balance"]) || 0);
            if (!item) return;
            ordersMap[item] = (ordersMap[item] || 0) + qty;
        });

        // Style Map: prefer Sheet 1, fallback to Sheet 3
        styleMap = {};
        qtyData.forEach(row => {
            const code = getColumnValue(row, ["Item Code"]);
            const style = getColumnValue(row, ["Style", "Style Code"]);
            if (code && style) styleMap[code] = style;
        });
        ordersData.forEach(row => {
            const code = getColumnValue(row, ["Item Code"]);
            const style = getColumnValue(row, ["Style", "Style Code"]);
            if (code && style && !styleMap[code]) styleMap[code] = style;
        });

        // Main Data: only Item Codes from Sheet 1
        analyzedData = qtyData.map(row => {
            const item = getColumnValue(row, ["Item Code"]);
            const style = styleMap[item] || "";
            let qtyOnOrder = Number(getColumnValue(row, ["Qty on Order", "Qty On Order", "Qty on order"]) || 0);
            return {
                "Item Code": item,
                "Style Code": style,
                "Qty on Order": qtyOnOrder,
                "On Allocation File": allocationMap[item] || 0,
                "Balance from Orders": ordersMap[item] || 0
            };
        });

        renderMainReport(analyzedData);
        renderMismatchReport();
        renderToOrderReport(analyzedData);

        switchTab('main');
    };
    reader.readAsArrayBuffer(file);
}

function calculateTotals(data, keys) {
    const totals = {};
    keys.forEach(k => totals[k] = 0);
    data.forEach(row => {
        keys.forEach(k => {
            let val = row[k];
            if (typeof val === 'string') val = Number(val.replace(/,/g, ''));
            totals[k] += Number(val || 0);
        });
    });
    return totals;
}

function renderMainReport(data) {
    const container = document.getElementById("results");
    if (!data || data.length === 0) {
        container.innerHTML = "<p>No data found.</p>";
        document.getElementById("totals").innerHTML = "";
        return;
    }

    // Filters for > 0
    const filterQty = document.getElementById("filterQtyCheckbox")?.checked;
    const filterAlloc = document.getElementById("filterAllocCheckbox")?.checked;
    const filterOrders = document.getElementById("filterOrdersCheckbox")?.checked;
    let filtered = data;
    if (filterQty) filtered = filtered.filter(row => row["Qty on Order"] > 0);
    if (filterAlloc) filtered = filtered.filter(row => row["On Allocation File"] > 0);
    if (filterOrders) filtered = filtered.filter(row => row["Balance from Orders"] > 0);

    const headers = ["Item Code", "Style Code", "Qty on Order", "On Allocation File", "Balance from Orders"];
    let html = "<table><thead><tr>";
    headers.forEach(h => html += `<th>${h}</th>`);
    html += "</tr></thead><tbody>";
    filtered.forEach(row => {
        html += "<tr>";
        headers.forEach(h => html += `<td>${row[h]}</td>`);
        html += "</tr>";
    });
    // Totals row
    const totals = calculateTotals(filtered, ["Qty on Order", "On Allocation File", "Balance from Orders"]);
    html += "<tr style='font-weight: bold; background: #eaf2fb;'>";
    headers.forEach(h => {
        if (h === "Item Code") html += "<td>TOTAL</td>";
        else if (["Qty on Order", "On Allocation File", "Balance from Orders"].includes(h))
            html += `<td>${totals[h]}</td>`;
        else html += "<td></td>";
    });
    html += "</tr>";

    html += "</tbody></table>";
    container.innerHTML = html;

    // Show stats for number of positive-value rows
    const totalDiv = document.getElementById("totals");
    totalDiv.innerHTML =
      `Rows with Qty on Order &gt; 0: ${data.filter(r => r["Qty on Order"] > 0).length} | ` +
      `On Allocation File &gt; 0: ${data.filter(r => r["On Allocation File"] > 0).length} | ` +
      `Balance from Orders &gt; 0: ${data.filter(r => r["Balance from Orders"] > 0).length}<br>` +
      "Totals — " +
      ["Qty on Order", "On Allocation File", "Balance from Orders"]
        .map(k => k + ": " + totals[k]).join(" | ");
}

function renderMismatchReport() {
    const container = document.getElementById("mismatchResults");
    const totalsDiv = document.getElementById("mismatchTotals");
    if (!qtyData || !allocData) {
        container.innerHTML = "<p>No data found.</p>";
        totalsDiv.innerHTML = "";
        return;
    }

    // Pivot Allocation File: sum by Item Code
    const allocSum = {};
    allocData.forEach(row => {
        const code = getColumnValue(row, ["Item Code"]);
        const qty = Number(getColumnValue(row, ["Pending Order Qty", "Pending order qty"]) || 0);
        allocSum[code] = (allocSum[code] || 0) + qty;
    });

    // For each Item Code from Sheet 1, compare and add appropriate note
    const mismatchRows = qtyData.map(row => {
        const code = getColumnValue(row, ["Item Code"]);
        const style = getColumnValue(row, ["Style", "Style Code"]) || "";
        const qtyOnOrder = Number(getColumnValue(row, ["Qty on Order", "Qty On Order", "Qty on order"]) || 0);
        const allocQty = allocSum[code] || 0;
        let note = "";
        if (allocQty > qtyOnOrder) note = "Remove from SIMA System";
        else if (allocQty < qtyOnOrder) note = "Add to SIMA System";
        else note = "No action needed";
        return {
            "Item Code": code,
            "Style Code": style,
            "Qty on Order": qtyOnOrder,
            "Total Allocation": allocQty,
            "Note": note
        };
    });

    // Show all rows for reconciliation
    const rowsToShow = mismatchRows;

    if (rowsToShow.length === 0) {
        container.innerHTML = "<p>No mismatches found.</p>";
        totalsDiv.innerHTML = "";
        return;
    }

    const headers = ["Item Code", "Style Code", "Qty on Order", "Total Allocation", "Note"];
    let html = "<table><thead><tr>";
    headers.forEach(h => html += `<th>${h}</th>`);
    html += "</tr></thead><tbody>";
    rowsToShow.forEach(row => {
        html += "<tr>";
        headers.forEach(h => html += `<td>${row[h]}</td>`);
        html += "</tr>";
    });

    // Totals row
    const totalQtyOnOrder = rowsToShow.reduce((sum, r) => sum + (Number(r["Qty on Order"]) || 0), 0);
    const totalAlloc = rowsToShow.reduce((sum, r) => sum + (Number(r["Total Allocation"]) || 0), 0);
    html += "<tr style='font-weight: bold; background: #eaf2fb;'>";
    headers.forEach(h => {
        if (h === "Item Code") html += "<td>TOTAL</td>";
        else if (h === "Qty on Order") html += `<td>${totalQtyOnOrder}</td>`;
        else if (h === "Total Allocation") html += `<td>${totalAlloc}</td>`;
        else html += "<td></td>";
    });
    html += "</tr>";

    html += "</tbody></table>";
    container.innerHTML = html;
    totalsDiv.innerHTML = `Totals — Qty on Order: ${totalQtyOnOrder} | Total Allocation: ${totalAlloc}`;
}

function renderToOrderReport(data) {
    const container = document.getElementById("toOrderResults");
    const totalsDiv = document.getElementById("toOrderTotals");
    const filtered = data.filter(row => row["Balance from Orders"] > row["On Allocation File"]);
    if (filtered.length === 0) {
        container.innerHTML = "<p>No items to order.</p>";
        totalsDiv.innerHTML = "";
        return;
    }
    const headers = ["Item Code", "Style Code", "On Allocation File", "Balance from Orders"];
    let html = "<table><thead><tr>";
    headers.forEach(h => html += `<th>${h}</th>`);
    html += "</tr></thead><tbody>";
    filtered.forEach(row => {
        html += "<tr>";
        headers.forEach(h => html += `<td>${row[h]}</td>`);
        html += "</tr>";
    });
    // Totals row
    const totals = calculateTotals(filtered, ["On Allocation File", "Balance from Orders"]);
    html += "<tr style='font-weight: bold; background: #eaf2fb;'>";
    headers.forEach(h => {
        if (h === "Item Code") html += "<td>TOTAL</td>";
        else if (["On Allocation File", "Balance from Orders"].includes(h))
            html += `<td>${totals[h]}</td>`;
        else html += "<td></td>";
    });
    html += "</tr>";

    html += "</tbody></table>";
    container.innerHTML = html;
    totalsDiv.innerHTML = "Totals — " +
      ["On Allocation File", "Balance from Orders"]
        .map(k => k + ": " + totals[k]).join(" | ");
}

function exportReport(type) {
    let headers, rows, filename;
    if (type === 'main') {
        headers = ["Item Code", "Style Code", "Qty on Order", "On Allocation File", "Balance from Orders"];
        rows = analyzedData.map(r => headers.reduce((obj, h) => (obj[h] = r[h], obj), {}));
        filename = "Stock_Order_Main_Report.xlsx";
    } else if (type === 'mismatch') {
        headers = ["Item Code", "Style Code", "Qty on Order", "Total Allocation", "Note"];
        // Use the same logic as renderMismatchReport
        const allocSum = {};
        allocData.forEach(row => {
            const code = getColumnValue(row, ["Item Code"]);
            const qty = Number(getColumnValue(row, ["Pending Order Qty", "Pending order qty"]) || 0);
            allocSum[code] = (allocSum[code] || 0) + qty;
        });
        rows = qtyData.map(row => {
            const code = getColumnValue(row, ["Item Code"]);
            const style = getColumnValue(row, ["Style", "Style Code"]) || "";
            const qtyOnOrder = Number(getColumnValue(row, ["Qty on Order", "Qty On Order", "Qty on order"]) || 0);
            const allocQty = allocSum[code] || 0;
            let note = "";
            if (allocQty > qtyOnOrder) note = "Remove from SIMA System";
            else if (allocQty < qtyOnOrder) note = "Add to SIMA System";
            else note = "No action needed";
            return {
                "Item Code": code,
                "Style Code": style,
                "Qty on Order": qtyOnOrder,
                "Total Allocation": allocQty,
                "Note": note
            };
        });
        filename = "Stock_Order_Mismatch_Report.xlsx";
    } else if (type === 'toorder') {
        headers = ["Item Code", "Style Code", "On Allocation File", "Balance from Orders"];
        rows = analyzedData
          .filter(r => r["Balance from Orders"] > r["On Allocation File"])
          .map(r => headers.reduce((obj, h) => (obj[h] = r[h], obj), {}));
        filename = "Stock_Order_ToOrder_Report.xlsx";
    } else {
        alert("Unknown report type.");
        return;
    }
    if (!rows || rows.length === 0) {
        alert("No data to export for this report.");
        return;
    }
    // Add totals row
    let totals = calculateTotals(rows, headers.filter(h =>
        typeof rows[0][h] === 'number' && h !== 'Item Code' && h !== 'Style Code' && h !== 'Note'
    ));
    let totalsRow = {};
    headers.forEach(h => {
        if (h === "Item Code") totalsRow[h] = "TOTAL";
        else if (totals[h] !== undefined) totalsRow[h] = totals[h];
        else totalsRow[h] = "";
    });
    rows.push(totalsRow);
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, type.charAt(0).toUpperCase() + type.slice(1) + " Report");
    XLSX.writeFile(wb, filename);
}

function exportAllReports() {
    // Create a workbook with all three reports
    const wb = XLSX.utils.book_new();
    // Main
    let mainHeaders = ["Item Code", "Style Code", "Qty on Order", "On Allocation File", "Balance from Orders"];
    let mainRows = analyzedData.map(r => mainHeaders.reduce((obj, h) => (obj[h] = r[h], obj), {}));
    let mainTotals = calculateTotals(mainRows, ["Qty on Order", "On Allocation File", "Balance from Orders"]);
    let mainTotalsRow = {};
    mainHeaders.forEach(h => {
        if (h === "Item Code") mainTotalsRow[h] = "TOTAL";
        else if (mainTotals[h] !== undefined) mainTotalsRow[h] = mainTotals[h];
        else mainTotalsRow[h] = "";
    });
    mainRows.push(mainTotalsRow);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mainRows), "Main Report");
    // Mismatch
    let mismatchHeaders = ["Item Code", "Style Code", "Qty on Order", "Total Allocation", "Note"];
    const allocSum = {};
    allocData.forEach(row => {
        const code = getColumnValue(row, ["Item Code"]);
        const qty = Number(getColumnValue(row, ["Pending Order Qty", "Pending order qty"]) || 0);
        allocSum[code] = (allocSum[code] || 0) + qty;
    });
    let mismatchRows = qtyData.map(row => {
        const code = getColumnValue(row, ["Item Code"]);
        const style = getColumnValue(row, ["Style", "Style Code"]) || "";
        const qtyOnOrder = Number(getColumnValue(row, ["Qty on Order", "Qty On Order", "Qty on order"]) || 0);
        const allocQty = allocSum[code] || 0;
        let note = "";
        if (allocQty > qtyOnOrder) note = "Remove from SIMA System";
        else if (allocQty < qtyOnOrder) note = "Add to SIMA System";
        else note = "No action needed";
        return {
            "Item Code": code,
            "Style Code": style,
            "Qty on Order": qtyOnOrder,
            "Total Allocation": allocQty,
            "Note": note
        };
    });
    let mismatchTotals = calculateTotals(mismatchRows, ["Qty on Order", "Total Allocation"]);
    let mismatchTotalsRow = {};
    mismatchHeaders.forEach(h => {
        if (h === "Item Code") mismatchTotalsRow[h] = "TOTAL";
        else if (mismatchTotals[h] !== undefined) mismatchTotalsRow[h] = mismatchTotals[h];
        else mismatchTotalsRow[h] = "";
    });
    mismatchRows.push(mismatchTotalsRow);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mismatchRows), "Mismatch Report");
    // To Order
    let toOrderHeaders = ["Item Code", "Style Code", "On Allocation File", "Balance from Orders"];
    let toOrderRows = analyzedData
      .filter(r => r["Balance from Orders"] > r["On Allocation File"])
      .map(r => toOrderHeaders.reduce((obj, h) => (obj[h] = r[h], obj), {}));
    let toOrderTotals = calculateTotals(toOrderRows, ["On Allocation File", "Balance from Orders"]);
    let toOrderTotalsRow = {};
    toOrderHeaders.forEach(h => {
        if (h === "Item Code") toOrderTotalsRow[h] = "TOTAL";
        else if (toOrderTotals[h] !== undefined) toOrderTotalsRow[h] = toOrderTotals[h];
        else toOrderTotalsRow[h] = "";
    });
    toOrderRows.push(toOrderTotalsRow);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toOrderRows), "To Order Report");

    XLSX.writeFile(wb, "Stock_Order_All_Reports.xlsx");
}

function switchTab(id) {
    document.querySelectorAll('.tab-button').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active-tab'));
    document.getElementById(id).classList.add('active-tab');
    const tabIndex = {main:0, mismatch:1, toorder:2}[id];
    if (typeof tabIndex !== "undefined") {
        document.querySelectorAll('.tab-button')[tabIndex].classList.add('active');
    }
}
