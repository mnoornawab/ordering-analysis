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

    readExcel(orderFile, orderData => {
        readExcel(allocFile, allocData => {
            readExcel(pendingFile, pendingData => {
                generateReport(orderData, allocData, pendingData);
            });
        });
    });
});

function sumByStyleCode(data, qtyField) {
    const summed = {};
    data.forEach(item => {
        const code = item.StyleCode?.trim();
        const qty = parseInt(item[qtyField]) || 0;
        if (!code) return;
        if (!summed[code]) summed[code] = 0;
        summed[code] += qty;
    });
    return summed;
}

function generateReport(orderRaw, allocRaw, pendingRaw) {
    const orderMap = {};
    orderRaw.forEach(item => {
        const code = item.StyleCode?.trim();
        orderMap[code] = parseInt(item.OrderedQty) || 0;
    });

    const allocSum = sumByStyleCode(allocRaw, "AllocatedQty");
    const pendingSum = sumByStyleCode(pendingRaw, "PendingQty");

    const allCodes = new Set([
        ...Object.keys(orderMap),
        ...Object.keys(allocSum),
        ...Object.keys(pendingSum)
    ]);

    const finalData = [["StyleCode", "OrderedQty", "AllocatedQty", "PendingQty", "AddToSimaQty", "OrderFromKeringQty"]];
    const table = document.getElementById("reportTable");
    table.innerHTML = "<tr><th>StyleCode</th><th>OrderedQty</th><th>AllocatedQty</th><th>PendingQty</th><th>AddToSimaQty</th><th>OrderFromKeringQty</th></tr>";

    allCodes.forEach(code => {
        const ordered = orderMap[code] || 0;
        const allocated = allocSum[code] || 0;
        const pending = pendingSum[code] || 0;

        const addToSima = Math.max(allocated - ordered, 0);
        const orderFromKering = Math.max(pending - allocated - ordered, 0);

        if (addToSima > 0 || orderFromKering > 0) {
            finalData.push([code, ordered, allocated, pending, addToSima, orderFromKering]);
            table.innerHTML += `<tr>
                <td>${code}</td><td>${ordered}</td><td>${allocated}</td><td>${pending}</td>
                <td>${addToSima}</td><td>${orderFromKering}</td>
            </tr>`;
        }
    });

    document.getElementById("reportSection").style.display = 'block';

    // Excel download
    document.getElementById("downloadExcelBtn").onclick = () => {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(finalData);
        XLSX.utils.book_append_sheet(wb, ws, "Ordering Report");
        XLSX.writeFile(wb, "kering_order_report.xlsx");
    };

    // CSV download
    document.getElementById("downloadBtn").onclick = () => {
        const csvContent = finalData.map(e => e.join(",")).join("\\n");
        const blob = new Blob([csvContent], { type: "text/csv" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "kering_order_report.csv";
        a.click();
    };
}
