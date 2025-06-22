function normalizeKey(key) {
  return key?.toString().trim().toLowerCase();
}

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

    const itemMap = {};

    // Extract from SIMA sheet
    simaData.forEach(row => {
      const keys = Object.keys(row).reduce((acc, k) => {
        acc[normalizeKey(k)] = k;
        return acc;
      }, {});

      const itemCode = row[keys['item code']]?.toString().trim();
      const styleCode = row[keys['material code']]?.toString().trim();
      const qtyOnOrder = parseInt(row[keys['qty on order']]) || 0;

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

    // Extract from Allocation sheet
    allocData.forEach(row => {
      const keys = Object.keys(row).reduce((acc, k) => {
        acc[normalizeKey(k)] = k;
        return acc;
      }, {});

      const itemCode = row[keys['material code']]?.toString().trim();
      const qty = parseInt(row[keys['pending order qty']]) || 0;

      if (itemMap[itemCode]) {
        itemMap[itemCode].qtyAllocated += qty;
      }
    });

    // Extract from Orders sheet
    ordersData.forEach(row => {
      const keys = Object.keys(row).reduce((acc, k) => {
        acc[normalizeKey(k)] = k;
        return acc;
      }, {});

      const itemCode = row[keys['itemcode']]?.toString().trim();
      const qty = parseInt(row[keys['balance']]) || 0;

      if (itemMap[itemCode]) {
        itemMap[itemCode].balanceOrders += qty;
      }
    });

    const finalData = Object.values(itemMap);

    // Send to renderer
    document.getElementById('main-report').innerHTML = generateMainTable(finalData);
    document.getElementById('mismatch-report').innerHTML = generateMismatchTable(finalData);
    statusDiv.innerHTML = `<p style="color:green;">✅ File processed successfully.</p>`;
  };
  reader.readAsArrayBuffer(file);
}
