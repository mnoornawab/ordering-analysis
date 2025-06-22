function downloadTemplate() {
    const templateSheets = {
        "Quantity on Order - SIMA System": [
            ['Item Code', 'Style', 'Qty On Order'],
            ['ABC123', '1234A', 100]
        ],
        "Allocation File": [
            ['Item Code', 'Pending Order Qty'],
            ['ABC123', 25]
        ],
        "Orders-SIMA System": [
            ['Item Code', 'Style', 'BALANCE'],
            ['ABC123', '1234A', 30]
        ]
    };
    exportToExcel(templateSheets, 'Stock_Order_Template.xlsx');
}

let analyzedData = [];

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

        // Check for required sheets
        const simaSheet = workbook.Sheets["Quantity on Order - SIMA System"];
        const allocationSheet = workbook.Sheets["Allocation File"];
        const ordersSheet = workbook.Sheets["Orders-SIMA System"];
        if (!simaSheet || !allocationSheet || !ordersSheet) {
            alert("Excel file is missing required sheets. Please use the template.");
            return;
        }

        const simaData = XLSX.utils.sheet_to_json(simaSheet);
        const allocationData = XLSX.utils.sheet_to_json(allocationSheet);
        const ordersData = XLSX.utils.sheet_to_json(ordersSheet);

        // Build lookup maps for calculation
        const allocationMap = {};
        allocationData.forEach(row => {
            const itemCode = row["Item Code"];
            const qty = Number(row["Pending Order Qty"] || 0);
            allocationMap[itemCode] = (allocationMap[itemCode] || 0) + qty;
        });

        const ordersMap = {};
        ordersData.forEach(row => {
            const itemCode = row["Item Code"];
            const balance = Number(row["BALANCE"] || 0);
            ordersMap[itemCode] = (ordersMap[itemCode] || 0) + balance;
        });

        analyzedData = simaData.map(row => {
            const itemCode = row["Item Code"];
            const style = row["Style"];
            const qtyOnOrder = Number(row["Qty On Order"] || 0);
            const allocatedQty = allocationMap[itemCode] || 0;
            const balanceOrder = ordersMap[itemCode] || 0;
            const qtyToOrder = Math.max(allocatedQty - balanceOrder, 0);

            return {
                "Item Code": itemCode,
                "Style": style,
                "Qty On Order": qtyOnOrder,
                "On Allocation File": allocatedQty,
                "Balance from Orders": balanceOrder,
                "Qty to Order": qtyToOrder
            };
        });

        renderResults(analyzedData);
        renderMismatch(analyzedData, allocationMap, ordersMap);
        renderToOrder(analyzedData);
    };
    reader.readAsArrayBuffer(file);
}

function renderResults(data) {
    const container = document.getElementById("results");
    if (!data || data.length === 0) {
        container.innerHTML = "<p>No data found.</p>";
        document.getElementById("totals").innerHTML = "";
        return;
    }

    const filtered = document.getElementById("filterCheckbox").checked
        ? data.filter(row => row["Qty to Order"] > 0)
        : data;

    const table = document.createElement("table");
    table.border = "1";
    table.style.borderCollapse = "collapse";
    table.style.width = "100%";

    const thead = document.createElement("thead");
    const headers = Object.keys(filtered[0]);
    const headRow = document.createElement("tr");
    headers.forEach(header => {
        const th = document.createElement("th");
        th.innerText = header;
        th.style.padding = "5px";
        headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    filtered.forEach(row => {
        const tr = document.createElement("tr");
        headers.forEach(header => {
            const td = document.createElement("td");
            td.innerText = row[header];
            td.style.padding = "5px";
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    container.innerHTML = "";
    container.appendChild(table);

    // Totals
    const totalKeys = ["Qty On Order", "On Allocation File", "Balance from Orders", "Qty to Order"];
    const totals = calculateTotals(filtered, totalKeys);
    const totalDiv = document.getElementById("totals");
    totalDiv.innerHTML = "Totals — " + totalKeys.map(k => k + ": " + totals[k]).join(" | ");
}

function renderMismatch(data, allocationMap, ordersMap) {
    // Optional: Show mismatches between allocation and orders
    const mismatches = data.filter(row => {
        return row["On Allocation File"] !== row["Balance from Orders"];
    });
    const container = document.getElementById("mismatchResults");
    if (mismatches.length === 0) {
        container.innerHTML = "<p>No mismatches found.</p>";
        return;
    }
    let html = "<table><thead><tr>";
    Object.keys(mismatches[0]).forEach(h => html += `<th>${h}</th>`);
    html += "</tr></thead><tbody>";
    mismatches.forEach(row => {
        html += "<tr>";
        Object.values(row).forEach(val => html += `<td>${val}</td>`);
        html += "</tr>";
    });
    html += "</tbody></table>";
    container.innerHTML = html;
}

function renderToOrder(data) {
    // Show rows where Qty to Order > 0
    const filtered = data.filter(row => row["Qty to Order"] > 0);
    const container = document.getElementById("toOrderResults");
    if (filtered.length === 0) {
        container.innerHTML = "<p>No items to order.</p>";
        return;
    }
    let html = "<table><thead><tr>";
    Object.keys(filtered[0]).forEach(h => html += `<th>${h}</th>`);
    html += "</tr></thead><tbody>";
    filtered.forEach(row => {
        html += "<tr>";
        Object.values(row).forEach(val => html += `<td>${val}</td>`);
        html += "</tr>";
    });
    html += "</tbody></table>";
    container.innerHTML = html;
}

function exportResults() {
    if (!analyzedData || analyzedData.length === 0) {
        alert("Please analyze a file first.");
        return;
    }
    const filtered = document.getElementById("filterCheckbox").checked
        ? analyzedData.filter(row => row["Qty to Order"] > 0)
        : analyzedData;
    const ws = XLSX.utils.json_to_sheet(filtered);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, "Stock_Order_Report.xlsx");
}

function calculateTotals(data, keys) {
    const totals = {};
    keys.forEach(k => totals[k] = 0);
    data.forEach(row => {
        keys.forEach(k => {
            totals[k] += Number(row[k] || 0);
        });
    });
    return totals;
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
