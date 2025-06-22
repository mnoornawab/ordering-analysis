function processFile() {
  const fileInput = document.getElementById('file-input');
  const file = fileInput.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const sheet1 = workbook.Sheets['Quantity on Order - SIMA System'];
    const sheet2 = workbook.Sheets['Allocation File'];
    const sheet3 = workbook.Sheets['Orders-SIMA System'];

    const data1 = XLSX.utils.sheet_to_json(sheet1, { defval: '' });
    const data2 = XLSX.utils.sheet_to_json(sheet2, { defval: '' });
    const data3 = XLSX.utils.sheet_to_json(sheet3, { defval: '' });

    const allocationSums = {};
    data2.forEach(row => {
      const itemCode = String(row['Item Code']).trim();
      const qty = parseInt(row['Pending Order Qty']) || 0;
      if (itemCode) {
        allocationSums[itemCode] = (allocationSums[itemCode] || 0) + qty;
      }
    });

    const ordersSums = {};
    data3.forEach(row => {
      const itemCode = String(row['Item Code']).trim();
      const balance = parseInt(row['BALANCE']) || 0;
      if (itemCode) {
        ordersSums[itemCode] = (ordersSums[itemCode] || 0) + balance;
      }
    });

    const finalData = data1.map(row => {
      const itemCode = String(row['Item Code']).trim();
      const styleCode = row['Style'] || '';
      const qtyOnOrder = parseInt(row['Qty On Order']) || 0;

      return {
        itemCode,
        styleCode,
        qtyOnOrder,
        qtyAllocated: allocationSums[itemCode] || 0,
        balanceOrders: ordersSums[itemCode] || 0
      };
    });

    // Send data to renderer
    renderMainReport(finalData);
    renderMismatchReport(finalData);
    renderToOrderReport(finalData);

    document.getElementById('download-btn').style.display = 'block';
    document.getElementById('status').innerHTML = '<p style="color:green;">✅ File uploaded and analyzed.</p>';
  };

  reader.readAsArrayBuffer(file);
}
