let csvRows = [];

function readExcel(file, callback) {
    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        callback(json);
    };
    reader.readAsArrayBuffer(file);
}

document.getElementById('uploadForm').addEventListener('submit', function (e) {
    e.preventDefault();

    const allocFile = document.getElementById('allocationFile').files[0];
    const orderFile = document.getElementById('orderFile').files[0];
    const pendingFile = document.getElementById('pendingFile').files[0];

    if (!allocFile || !orderFile || !pendingFile) {
        alert("Please upload all three Excel files.");
        return;
    }

    readExcel(allocFile, allocData => {
        readExcel(orderFile, orderData => {
            readExcel(pendingFile, pendingData => {
                generateReport(allocData, orderData, pendingData);
            });
        });
    });
});

function sumByStyleCode(data, qtyField) {
    const summed = {};
    data.forEach(item => {
        const code = item.StyleCode?.trim();
        const qty = parseInt(item[qtyField]) || 0;
        if (!summed[code]) summed[code] = 0;
        summed[code] += qty;
    });
    return summed;
}

function generateReport(allocRaw, orderRaw, pendingRaw) {
    const allocSum = sumByStyleCode(allocRaw, "AllocatedQty");
    const pendingSum = sumByStyleCode(pendingRaw, "PendingQty");
    const orderMap = {}; // unique, so no need to sum
    orderRaw.forEach(item => {
        const code = item.StyleCode?.trim();
        orderMap[code] = parseInt(item.OrderedQty) || 0;
    });

    const styleCodes = new Set([
        ...Object.keys(pendingSum),
        ...Object.keys(allocSum),
        ...Object.keys(orderMap)
    ]);

    const toOrderRows = [["StyleCode", "NeedToOrder"]];
    const missingInSimaRows = [["StyleCode", "AllocatedQty", "OrderedInSima"]];

    const table = document.getElementById("reportTable");
    table.innerHTML = "<tr><th>StyleCode</th><th>NeedToOrder</th></tr>";

    csvRows = [["StyleCode", "NeedToOrder"]]; // for CSV export

    styleCodes.forEach(code => {
        const pending = pendingSum[code] || 0;
        const allocated = allocSum[code] || 0;
        const ordered = orderMap[code] || 0;

        const toOrder = pending - allocated - ordered;
        if (toOrder > 0) {
            toOrderRows.push([code, toOrder]);
            csvRows.push([code, toOrder]);
            table.innerHTML += `<tr><td>${code}</td><td>${toOrder}</td></tr>`;
        }

        if (allocated > 0 && ordered < allocated) {
            missingInSimaRows.push([code, allocated, ordered]);
        }
    });

    document.getElementById("reportSection").style.display = 'block';

    document.getElementById("downloadBtn").onclick = () => {
        const csvContent = csvRows.map(e => e.join(",")).join("\\n");
        const blob = new Blob([csvContent], { type: "text/csv" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "kering_to_order.csv";
        a.click();
    };

    document.getElementById("downloadExcelBtn").onclick = () => {
        const wb = XLSX.utils.book_new();
        const ws1 = XLSX.utils.aoa_to_sheet(toOrderRows);
        const ws2 = XLSX.utils.aoa_to_sheet(missingInSimaRows);
        XLSX.utils.book_append_sheet(wb, ws1, "To Order from Kering");
        XLSX.utils.book_append_sheet(wb, ws2, "Missing in Sima System");
        XLSX.writeFile(wb, "kering_order_report.xlsx");
    };
}
