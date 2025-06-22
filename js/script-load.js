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

    if (!sheetSIMA || !sheetAllocation || !sheetOrders) {
      statusDiv.innerHTML = `<p style="color:red;">❌ One or more sheets not found.</p>`;
      return;
    }

    const simaData = XLSX.utils.sheet_to_json(sheetSIMA, { defval: '' });
    const allocData = XLSX.utils.sheet_to_json(sheetAllocation, { defval: '' });
    const ordersData = XLSX.utils.sheet_to_json(sheetOrders, { defval: '' });

    const itemMap = {};

    // Load SIMA sheet
    simaData.forEach(row => {
      const itemCode = row['Item Code']?.toString().trim();
      const styleCode = row['Style Code']?.toString().trim();
      const qtyOnOrder = parseInt(row['Qty On Order']) || 0;

      if (itemCode && itemCode !== 'undefined') {
        itemMap[itemCode] = {
          itemCode,
          styleCode,
          qtyOnOrder,
          qtyAllocated: 0,
          balanceOrders: 0
        };
      }
    });

    // Load Allocation File by Style Code
    allocData.forEach(row => {
      const styleCodeAlloc = row['Material code']?.toString().trim();
      const qtyAlloc = parseInt(row['Pending order qty']) || 0;

      Object.values(itemMap).forEach(item => {
        if (item.styleCode === styleCodeAlloc) {
          item.qtyAllocated += qtyAlloc;
        }
      });
    });

    // Load Orders sheet by Item Code
    ordersData.forEach(row => {
      const itemCode = row['ITEMCODE']?.toString().trim();
      const balance = parseInt(row['BALANCE']) || 0;

      if (itemMap[itemCode]) {
        itemMap[itemCode].balanceOrders += balance;
      }
    });

    const finalData = Object.values(itemMap);

    // Save for rendering
    window.processedData = finalData;

    // Render
    document.getElementById('main-report').innerHTML = generateMainTable(finalData);
    document.getElementById('mismatch-report').innerHTML = generateMismatchTable(finalData);
    document.getElementById('to-order-report').innerHTML = generateToOrderTable(finalData);

    statusDiv.innerHTML = `<p style="color:green;">✅ File processed successfully.</p>`;
  };

  reader.readAsArrayBuffer(file);
}
