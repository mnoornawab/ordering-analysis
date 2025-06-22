
function processFile() {
  const fileInput = document.getElementById('file-input');
  const file = fileInput.files[0];
  const statusDiv = document.getElementById('status');
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const sheet1 = XLSX.utils.sheet_to_json(workbook.Sheets['Quantity on Order - SIMA System'], { defval: '' });
    const sheet2 = XLSX.utils.sheet_to_json(workbook.Sheets['Allocation File'], { defval: '' });
    const sheet3 = XLSX.utils.sheet_to_json(workbook.Sheets['Orders-SIMA System'], { defval: '' });

    const allocationMap = {};
    sheet2.forEach(row => {
      const item = String(row['Item Code']).trim();
      const qty = parseInt(row['Pending Order Qty']) || 0;
      if (!allocationMap[item]) allocationMap[item] = 0;
      allocationMap[item] += qty;
    });

    const orderBalanceMap = {};
    sheet3.forEach(row => {
      const item = String(row['Item Code']).trim();
      const balance = parseInt(row['BALANCE']) || 0;
      if (!orderBalanceMap[item]) orderBalanceMap[item] = 0;
      orderBalanceMap[item] += balance;
    });

    const results = [];
    sheet1.forEach(row => {
      const itemCode = String(row['Item Code']).trim();
      const styleCode = row['Style'];
      const qtyOnOrder = parseInt(row['Qty On Order']) || 0;
      const qtyAllocated = allocationMap[itemCode] || 0;
      const balanceOrders = orderBalanceMap[itemCode] || 0;

      results.push({
        itemCode,
        styleCode,
        qtyOnOrder,
        qtyAllocated,
        balanceOrders
      });
    });

    renderMainReport(results);
    renderMismatchReport(results);
    renderToOrderReport(results);
    statusDiv.innerHTML = '<p style="color:green;">✅ File processed successfully.</p>';
  };
  reader.readAsArrayBuffer(file);
}
