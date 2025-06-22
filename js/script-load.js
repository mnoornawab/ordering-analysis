function processFile() {
  const fileInput = document.getElementById('file-input');
  const file = fileInput.files[0];
  const statusDiv = document.getElementById('status');
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const sheet1 = workbook.Sheets['Quantity on Order - SIMA System'];
    const sheet2 = workbook.Sheets['Allocation File'];
    const sheet3 = workbook.Sheets['Orders-SIMA System'];

    const simaData = XLSX.utils.sheet_to_json(sheet1, { defval: '' });
    const allocData = XLSX.utils.sheet_to_json(sheet2, { defval: '' });
    const ordersData = XLSX.utils.sheet_to_json(sheet3, { defval: '' });

    const itemMap = {};

    // Step 1: Load SIMA Sheet (base)
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
          balanceOrders: 0
        };
      }
    });

    // Step 2: Merge Allocation File (by Material code = styleCode)
    allocData.forEach(row => {
      const styleCode = row['Material code']?.toString().trim();
      const qty = parseInt(row['Pending order qty']) || 0;

      for (const code in itemMap) {
        if (itemMap[code].styleCode === styleCode) {
          itemMap[code].qtyAllocated += qty;
        }
      }
    });

    // Step 3: Merge Balance from Orders File (match by ITEMCODE)
    ordersData.forEach(row => {
      const itemCode = row['ITEMCODE']?.toString().trim();
      const balance = parseInt(row['BALANCE']) || 0;

      if (itemMap[itemCode]) {
        itemMap[itemCode].balanceOrders += balance;
      }
    });

    const finalData = Object.values(itemMap);

    // Render
    document.getElementById('main-report').innerHTML = generateMainTable(finalData);
    document.getElementById('mismatch-report').innerHTML = generateMismatchTable(finalData);
    statusDiv.innerHTML = `<p style="color:green;">✅ File processed successfully.</p>`;
  };

  reader.readAsArrayBuffer(file);
}
