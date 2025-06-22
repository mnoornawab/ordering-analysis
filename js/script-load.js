function processFile() {
  const fileInput = document.getElementById('file-input');
  const file = fileInput.files[0];
  const statusDiv = document.getElementById('status');
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const sheetSIMA = workbook.Sheets['Quantity on Order - SIMA System'];
    const sheetAllocation = workbook.Sheets['Allocation File'];
    const sheetOrders = workbook.Sheets['Orders-SIMA System'];

    const simaData = XLSX.utils.sheet_to_json(sheetSIMA, { defval: '' });
    const allocData = XLSX.utils.sheet_to_json(sheetAllocation, { defval: '' });
    const ordersData = XLSX.utils.sheet_to_json(sheetOrders, { defval: '' });

    // Map base data from SIMA system
    const itemMap = {};
    simaData.forEach(row => {
      const itemCode = row['Item Code']?.toString().trim();
      const styleCode = row['Material Code']?.toString().trim();
      const qtyOnOrder = parseInt(row['Qty On Order']) || 0;
      if (itemCode && itemCode !== 'undefined') {
        itemMap[itemCode] = {
          itemCode,
          styleCode,
          qtyOnOrder,
          qtyAllocated: 0,
          balanceOrders: 0,
        };
      }
    });

    // Merge Allocation File data
    allocData.forEach(row => {
      const itemCode = row['Material code']?.toString().trim();
      const qty = parseInt(row['Pending order qty']) || 0;
      if (itemCode && itemMap[itemCode]) {
        itemMap[itemCode].qtyAllocated += qty;
      }
    });

    // Merge Orders-SIMA System data
    ordersData.forEach(row => {
      const itemCode = row['ITEMCODE']?.toString().trim();
      const balance = parseInt(row['BALANCE']) || 0;
      if (itemCode && itemMap[itemCode]) {
        itemMap[itemCode].balanceOrders += balance;
      }
    });

    const finalData = Object.values(itemMap);

    // Show final table
    document.getElementById('main-report').innerHTML = generateMainTable(finalData);
    document.getElementById('mismatch-report').innerHTML = generateMismatchTable(finalData);
    statusDiv.innerHTML = '<p style="color:green;">✅ File processed successfully.</p>';
  };
  reader.readAsArrayBuffer(file);
}
