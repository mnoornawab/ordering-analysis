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

    if (!sheet1 || !sheet2 || !sheet3) {
      statusDiv.innerHTML = `<p style="color:red;">❌ One or more sheets not found.</p>`;
      return;
    }

    const simaData = XLSX.utils.sheet_to_json(sheet1, { defval: '' });
    const allocData = XLSX.utils.sheet_to_json(sheet2, { defval: '' });
    const ordersData = XLSX.utils.sheet_to_json(sheet3, { defval: '' });

    const itemMap = {};

    // STEP 1: Load SIMA base file
    simaData.forEach(row => {
      const itemCode = row['Item Code']?.toString().trim();
      const styleCode = row['Style Code']?.toString().trim(); // ✅ Must be "Style Code"
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

    // STEP 2: Match Allocation File by style code
    allocData.forEach(row => {
      const allocStyle = row['Material code']?.toString().trim();
      const allocQty = parseInt(row['Pending order qty']) || 0;

      Object.values(itemMap).forEach(entry => {
        if (entry.styleCode === allocStyle) {
          entry.qtyAllocated += allocQty;
        }
      });
    });

    // STEP 3: Match Orders by item code
    ordersData.forEach(row => {
      const itemCode = row['ITEMCODE']?.toString().trim();
      const balance = parseInt(row['BALANCE']) || 0;

      if (itemMap[itemCode]) {
        itemMap[itemCode].balanceOrders += balance;
      }
    });

    const finalData = Object.values(itemMap);
    document.getElementById('main-report').innerHTML = generateMainTable(finalData);
    statusDiv.innerHTML = `<p style="color:green;">✅ File loaded successfully</p>`;
  };

  reader.readAsArrayBuffer(file);
}
