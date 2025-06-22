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

    // STEP 1: Load base SIMA file
    simaData.forEach(row => {
      const itemCode = row['Item Code']?.toString().trim();
      const styleCode = row['Style Code']?.toString().trim(); // 🟢 From Column G
      const qtyOnOrder = parseInt(row['Qty On Order']) || 0;

      if (itemCode) {
        itemMap[itemCode] = {
          itemCode,
          styleCode,
          qtyOnOrder,
          qtyAllocated: 0,
          balanceOrders: 0
        };
      }
    });

    // STEP 2: Match Allocation File by Style Code
    allocData.forEach(row => {
      const allocStyle = row['Material code']?.toString().trim();
      const allocQty = parseInt(row['Pending order qty']) || 0;

      for (const code in itemMap) {
        if (itemMap[code].styleCode === allocStyle) {
          itemMap[code].qtyAllocated += allocQty;
        }
      }
    });

    // STEP 3: Match Orders-SIMA File by Item Code
    ordersData.forEach(row => {
      const itemCode = row['ITEMCODE']?.toString().trim();
      const balance = parseInt(row['BALANCE']) || 0;

      if (itemMap[itemCode]) {
        itemMap[itemCode].balanceOrders += balance;
      }
    });

    const finalData = Object.values(itemMap);

    // Pass to render script
    window.processedData = finalData;
    document.getElementById('main-report').innerHTML = generateMainTable(finalData);
    document.getElementById('mismatch-report').innerHTML = generateMismatchTable(finalData);
    document.getElementById('to-order-report').innerHTML = generateToOrderTable(finalData);

    statusDiv.innerHTML = `<p style="color:green;">✅ File processed successfully.</p>`;
  };

  reader.readAsArrayBuffer(file);
}
