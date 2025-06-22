function downloadTemplate() {
    const templateSheets = {
        "Quantity on Order - SIMA System": [
            ['Item Code', 'Style', 'Qty on Order'],
            ['ABC123', '1234A', 100],
            ['DEF456', '2345B', 200]
        ],
        "Allocation File": [
            ['Item Code', 'Pending Order Qty'],
            ['ABC123', 100],
            ['DEF456', 200]
        ],
        "Orders-SIMA System": [
            ['Item Code', 'Style', 'BALANCE'],
            ['ABC123', '1234A', 120],
            ['DEF456', '2345B', 190],
            ['DEF456', '2345B', 20]
        ]
    };
    exportToExcel(templateSheets, 'Stock_Order_Template.xlsx');
}

let analyzedData = [];
let allocationMap = {};
let ordersMap = {};
let styleMap = {};

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

        // ---- Load all sheets ----
        const qtySheet = workbook.Sheets["Quantity on Order - SIMA System"];
        const allocSheet = workbook.Sheets["Allocation File"];
        const ordersSheet = workbook.Sheets["Orders-SIMA System"];
        if (!qtySheet || !allocSheet || !ordersSheet) {
            alert("Excel file is missing required sheets. Please use the template.");
            return;
        }

        const qtyData = XLSX.utils.sheet_to_json(qtySheet);
        const allocData = XLSX.utils.sheet_to_json(allocSheet);
        const ordersData = XLSX.utils.sheet_to_json(ordersSheet);

        // ---- Style map ----
        styleMap = {};
        qtyData.forEach(row => {
            if (row["Item Code"] && row["Style"]) styleMap[row["Item Code"]] = row["Style"];
            // Support "Style" and "Style Code"
            if (row["Item Code"] && row["Style Code"]) styleMap[row["Item Code"]] = row["Style Code"];
        });
        ordersData.forEach(row => {
            if (row["Item Code"] && row["Style"] && !styleMap[row["Item Code"]]) styleMap[row["Item Code"]] = row["Style"];
            if (row["Item Code"] && row["Style Code"] && !styleMap[row["Item Code"]]) styleMap[row["Item Code"]] = row["Style Code"];
        });

        // ---- Allocation Map (sum by Item Code) ----
        allocationMap = {};
        allocData.forEach(row => {
            const item = getColumnValue(row, ["Item Code"]);
            const qty = Number(getColumnValue(row, ["Pending Order Qty", "Pending order qty"]) || 0);
            if (!item) return;
            allocationMap[item] = (allocationMap[item] || 0) + qty;
        });

        // ---- Orders Map (sum by Item Code) ----
        ordersMap = {};
        ordersData.forEach(row => {
            const item = getColumnValue(row, ["Item Code"]);
            const qty = Number(getColumnValue(row, ["BALANCE", "Balance"]));
            if (!item) return;
            ordersMap[item] = (ordersMap[item] || 0) + (qty || 0);
        });

        // ---- All unique Item Codes (from Qty on Order and Allocation File) ----
        const allItemCodes = new Set();
        qtyData.forEach(row => { const code = getColumnValue(row, ["Item Code"]); if (code) allItemCodes.add(code); });
        allocData.forEach(row => { const code = getColumnValue(row, ["Item Code"]); if (code) allItemCodes.add(code); });

        // ---- Main Data ----
        analyzedData = Array.from(allItemCodes).map(item => {
            const qtyRow = qtyData.find(row => getColumnValue(row, ["Item Code"]) === item);
            const style = styleMap[item] || "";
            let qtyOnOrder = 0;
            if (qtyRow) {
                qtyOnOrder = Number(getColumnValue(qtyRow, ["Qty on Order", "Qty On Order"]) || 0);
            }
            return {
                "Item Code": item,
                "Style Code": style,
                "Qty on Order": qtyOnOrder,
                "On Allocation File": allocationMap[item] || 0,
                "Balance from Orders": ordersMap[item] || 0
            };
        });

        renderMainReport(analyzedData);
        renderMismatchReport(analyzedData);
        renderToOrderReport(analyzedData);

        switchTab('main');
    };
    reader.readAsArrayBuffer(file);
}

function renderMainReport(data) {
    const container = document.getElementById("results");
    if (!data || data.length === 0) {
        container.innerHTML = "<p>No data found.</p>";
        document.getElementById("totals").innerHTML = "";
        return;
    }

    // Filtered view (show only those needing to order more if checkbox is checked)
    const showToOrder = document.getElementById("filterCheckbox").checked;
    const filtered = showToOrder ? data.filter(row => row["Balance from Orders"] > row["On Allocation File"]) : data;

    const headers = ["Item Code", "Style Code", "Qty on Order", "On Allocation File", "Balance from Orders"];
    let html = "<table><thead><tr>";
    headers.forEach(h => html += `<th>${h}</th>`);
    html += "</tr></thead><tbody>";
    filtered.forEach(row => {
        html += "<tr>";
        headers.forEach(h => html += `<td>${row[h]}</td>`);
        html += "</tr>";
    });
    html += "</tbody></table>";
    container.innerHTML = html;

    // Totals
    const totalKeys = ["Qty on Order", "On Allocation File", "Balance from Orders"];
    const totals = calculateTotals(filtered, totalKeys);
    const totalDiv = document.getElementById("totals");
    totalDiv.innerHTML = "Totals — " + totalKeys.map(k => k + ": " + totals[k]).join(" | ");
}

function renderMismatchReport(data) {
    const container = document.getElementById("mismatchResults");
    // Only rows where Qty on Order ≠ On Allocation File
    const mismatches = data.filter(row => row["Qty on Order"] !== row["On Allocation File"]);
    if (mismatches.length === 0) {
        container.innerHTML = "<p>No mismatches found.</p>";
        return;
    }
    const headers = ["Item Code", "Style Code", "Qty on Order", "On Allocation File", "Mismatch Note"];
    let html = "<table><thead><tr>";
    headers.forEach(h => html += `<th>${h}</th>`);
    html += "</tr></thead><tbody>";
    mismatches.forEach(row => {
        let note = "";
        if (row["Qty on Order"] !== row["On Allocation File"]) {
            note = `Qty on Order (${row["Qty on Order"]}) ≠ On Allocation File (${row["On Allocation File"]})`;
        }
        html += "<tr>";
        headers.forEach(h => {
            if (h === "Mismatch Note") {
                html += `<td>${note}</td>`;
            } else {
                html += `<td>${row[h]}</td>`;
            }
        });
        html += "</tr>";
    });
    html += "</tbody></table>";
    container.innerHTML = html;
}

function renderToOrderReport(data) {
    const container = document.getElementById("toOrderResults");
    // Only rows where Balance from Orders > On Allocation File
    const filtered = data.filter(row => row["Balance from Orders"] > row["On Allocation File"]);
    if (filtered.length === 0) {
        container.innerHTML = "<p>No items to order.</p>";
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
    html += "</tbody></table>";
    container.innerHTML = html;
}

function exportResults() {
    if (!analyzedData || analyzedData.length === 0) {
        alert("Please analyze a file first.");
        return;
    }
    const ws = XLSX.utils.json_to_sheet(analyzedData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Main Report");
    XLSX.writeFile(wb, "Stock_Order_Main_Report.xlsx");
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
