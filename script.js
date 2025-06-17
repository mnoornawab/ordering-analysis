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

  const progressBar = document.getElementById("progressBar");
  const progressBarWrapper = document.getElementById("progressBarWrapper");

  progressBarWrapper.style.display = "block";
  progressBar.style.width = "0%";
  document.getElementById("reportSection").style.display = "none";

  let orderData, allocData, pendingData;

  setTimeout(() => {
    readExcel(orderFile, data1 => {
      progressBar.style.width = "30%";
      orderData = data1;

      setTimeout(() => {
        readExcel(allocFile, data2 => {
          progressBar.style.width = "60%";
          allocData = data2;

          setTimeout(() => {
            readExcel(pendingFile, data3 => {
              progressBar.style.width = "80%";
              pendingData = data3;

              setTimeout(() => {
                generateReport(orderData, allocData, pendingData, () => {
                  progressBar.style.width = "100%";
                  setTimeout(() => {
                    progressBarWrapper.style.display = "none";
                    progressBar.style.width = "0%";
                  }, 500);
                });
              }, 100);
            });
          }, 100);
        });
      }, 100);
    });
  }, 100);
});

function generateReport(orderRaw, allocRaw, pendingRaw, doneCallback) {
  const orderMap = new Map();
  orderRaw.forEach(item => {
    const code = item.StyleCode?.trim();
    const qty = parseInt(item.OrderedQty) || 0;
    if (code) orderMap.set(code, qty);
  });

  const allocMap = new Map();
  allocRaw.forEach(item => {
    const code = item.StyleCode?.trim();
    const qty = parseInt(item.AllocatedQty) || 0;
    if (!code) return;
    allocMap.set(code, (allocMap.get(code) || 0) + qty);
  });

  const pendingMap = new Map();
  pendingRaw.forEach(item => {
    const code = item.StyleCode?.trim();
    const qty = parseInt(item.PendingQty) || 0;
    if (!code) return;
    pendingMap.set(code, (pendingMap.get(code) || 0) + qty);
  });

  const allCodes = new Set([
    ...orderMap.keys(),
    ...allocMap.keys(),
    ...pendingMap.keys()
  ]);

  const finalData = [["StyleCode", "OrderedQty", "AllocatedQty", "PendingQty", "AddToSimaQty", "OrderFromKeringQty"]];
  let htmlRows = "";

  for (const code of allCodes) {
    const ordered = orderMap.get(code) || 0;
    const allocated = allocMap.get(code) || 0;
    const pending = pendingMap.get(code) || 0;

    const addToSima = Math.max(allocated - ordered, 0);
    const orderFromKering = Math.max(pending - allocated - ordered, 0);

    if (addToSima > 0 || orderFromKering > 0) {
      finalData.push([code, ordered, allocated, pending, addToSima, orderFromKering]);
      htmlRows += `<tr>
        <td>${code}</td><td>${ordered}</td><td>${allocated}</td><td>${pending}</td>
        <td>${addToSima}</td><td>${orderFromKering}</td>
      </tr>`;
    }
  }

  const table = document.getElementById("reportTable");
  table.innerHTML = `<tr>
    <th>StyleCode</th><th>OrderedQty</th><th>AllocatedQty</th><th>PendingQty</th>
    <th>AddToSimaQty</th><th>OrderFromKeringQty</th>
  </tr>` + htmlRows;

  document.getElementById("reportSection").style.display = 'block';

  document.getElementById("downloadExcelBtn").onclick = () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(finalData);
    XLSX.utils.book_append_sheet(wb, ws, "Ordering Report");
    XLSX.writeFile(wb, "kering_order_report.xlsx");
  };

  document.getElementById("downloadBtn").onclick = () => {
    const csvContent = finalData.map(e => e.join(",")).join("\n");
    const blob = new Blob([csvContent], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "kering_order_report.csv";
    a.click();
  };

  if (doneCallback) doneCallback();
}
