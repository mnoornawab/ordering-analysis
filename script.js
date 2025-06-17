function parseCSV(file, callback) {
    const reader = new FileReader();
    reader.onload = function (e) {
        const rows = e.target.result.split(/\r?\n/).filter(Boolean).map(row => row.split(','));
        const headers = rows[0];
        const data = rows.slice(1).map(row => {
            const obj = {};
            headers.forEach((h, i) => obj[h.trim()] = row[i] ? row[i].trim() : "0");
            return obj;
        });
        callback(data);
    };
    reader.readAsText(file);
}

document.getElementById('uploadForm').addEventListener('submit', function (e) {
    e.preventDefault();

    const allocFile = document.getElementById('allocationFile').files[0];
    const orderFile = document.getElementById('orderFile').files[0];
    const pendingFile = document.getElementById('pendingFile').files[0];

    if (!allocFile || !orderFile || !pendingFile) {
        alert("Please upload all three files.");
        return;
    }

    let allocData, orderData, pendingData;

    parseCSV(allocFile, function (data) {
        allocData = data;
        parseCSV(orderFile, function (data) {
            orderData = data;
            parseCSV(pendingFile, function (data) {
                pendingData = data;
                generateReport(allocData, orderData, pendingData);
            });
        });
    });
});

function generateReport(allocData, orderData, pendingData) {
    const reportMap = {};

    pendingData.forEach(item => {
        const code = item.StyleCode;
        const pending = parseInt(item.PendingQty || "0");
        reportMap[code] = { pending, allocated: 0, ordered: 0 };
    });

    allocData.forEach(item => {
        const code = item.StyleCode;
        if (reportMap[code]) {
            reportMap[code].allocated = parseInt(item.AllocatedQty || "0");
        }
    });

    orderData.forEach(item => {
        const code = item.StyleCode;
        if (reportMap[code]) {
            reportMap[code].ordered = parseInt(item.OrderedQty || "0");
        }
    });

    const reportTable = document.getElementById('reportTable');
    reportTable.innerHTML = "<tr><th>StyleCode</th><th>NeedToOrder</th></tr>";

    const csvRows = [["StyleCode", "NeedToOrder"]];

    Object.entries(reportMap).forEach(([code, values]) => {
        const need = values.pending - values.allocated - values.ordered;
        const qty = need > 0 ? need : 0;
        reportTable.innerHTML += `<tr><td>${code}</td><td>${qty}</td></tr>`;
        csvRows.push([code, qty]);
    });

    document.getElementById('reportSection').style.display = 'block';

    document.getElementById('downloadBtn').onclick = () => {
        const csvContent = csvRows.map(e => e.join(",")).join("\n");
        const blob = new Blob([csvContent], { type: "text/csv" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "order_report.csv";
        a.click();
    };
}
