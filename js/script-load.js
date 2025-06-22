function processFile() {
  const fileInput = document.getElementById('file-input');
  const file = fileInput.files[0];
  const statusDiv = document.getElementById('status');
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const sheet1 = workbook.Sheets[workbook.SheetNames[0]]; // Qty On Order
    const sheet2 = workbook.Sheets[workbook.SheetNames[1]]; // Allocation
    const sheet3 = workbook.Sheets[workbook.SheetNames[2]]; // Orders

    const baseData = XLSX.utils.sheet_to_json(sheet1, { defval: '' });
    const allocData = XLSX.utils.sheet_to_json(sheet2, { defval: '' });
    const ordersData = XLSX.utils.sheet_to_json(sheet3, { defval: '' });

    // Step 1: Build allocation map
    const allocMap = {};
    allocData.forEach(row => {
      const itemCode = row['Item Code']?.toString().trim();
      const qty = parseInt(row['Pending Order Qty']) || 0;
      if (itemCode) {
        allocMap[itemCode] = (allocMap[itemCode] || 0) + qty;
      }
    });

    // Step 2: Build balance map
    const balanceMap = {};
    ordersData.forEach(row => {
      const itemCode = row['Item Code']?.toString().trim();
      const bal = parseInt(row['BALANCE']) || 0;
      if (itemCode) {
        balanceMap[itemCode] = (balanceMap[itemCode] || 0) + bal;
      }
    });

    // Step 3: Merge data into base
    const finalData = baseData.map(row => {
      const itemCode = row['Item Code']?.toString().trim();
      const styleCode = row['Style']?.toString().trim();
      const qtyOnOrder = parseInt(row['Qty On Order']) || 0;
      const qtyAllocated = allocMap[itemCode] || 0;
      const balanceOrders = balanceMap[itemCode] || 0;

      return {
        itemCode,
        styleCode,
        qtyOnOrder,
        qtyAllocated,
        balanceOrders
      };
    });

    // Step 4: Render results
    renderMainReport(finalData);
    renderMismatchReport(finalData);
    renderToOrderReport(finalData);

    statusDiv.innerHTML = '<p style="color:green;">✅ File processed successfully.</p>';
  };

  reader.readAsArrayBuffer(file);
}
