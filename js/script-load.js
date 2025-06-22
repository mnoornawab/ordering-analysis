function processFile() {
  const fileInput = document.getElementById('file-input');
  const file = fileInput.files[0];
  const statusDiv = document.getElementById('status');
  if (!file) {
    statusDiv.innerHTML = '<p style="color:red;">❌ Please select a file.</p>';
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const sheet1 = workbook.Sheets[workbook.SheetNames[0]]; // Quantity on Order - SIMA System
      const sheet2 = workbook.Sheets[workbook.SheetNames[1]]; // Allocation File
      const sheet3 = workbook.Sheets[workbook.SheetNames[2]]; // Orders Sheet

      const simaData = XLSX.utils.sheet_to_json(sheet1, { defval: '' });
      const allocData = XLSX.utils.sheet_to_json(sheet2, { defval: '' });
      const ordersData = XLSX.utils.sheet_to_json(sheet3, { defval: '' });

      // Create lookup maps from Sheet 2 and Sheet 3
      const allocationMap = {};
      allocData.forEach(row => {
        const itemCode = row['Item Code']?.toString().trim();
        const qty = parseInt(row['Qty Allocated']) || 0;
        if (itemCode) {
          allocationMap[itemCode] = (allocationMap[itemCode] || 0) + qty;
        }
      });

      const balanceMap = {};
      ordersData.forEach(row => {
        const itemCode = row['Item Code']?.toString().trim();
        const balance = parseInt(row['Balance']) || 0;
        if (itemCode) {
          balanceMap[itemCode] = (balanceMap[itemCode] || 0) + balance;
        }
      });

      // Merge data into a unified report using Sheet 1 as base
      const finalData = simaData.map(row => {
        const itemCode = row['Item Code']?.toString().trim();
        return {
          itemCode,
          styleCode: row['Style Code']?.toString().trim() || '',
          qtyOnOrder: parseInt(row['Qty On Order']) || 0,
          qtyAllocated: allocationMap[itemCode] || 0,
          balanceOrders: balanceMap[itemCode] || 0
        };
      });

      // Store in window for reuse
      window._stockData = finalData;

      // Render all reports
      renderMainReport(finalData);
      renderMismatchReport(finalData);
      renderToOrderReport(finalData);
      statusDiv.innerHTML = '<p style="color:green;">✅ File loaded and processed.</p>';
    } catch (err) {
      console.error(err);
      statusDiv.innerHTML = '<p style="color:red;">❌ Error reading file.</p>';
    }
  };

  reader.readAsArrayBuffer(file);
}

function openTab(tabId) {
  document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active-tab'));

  document.querySelector(`button[onclick="openTab('${tabId}')"]`).classList.add('active');
  document.getElementById(tabId).classList.add('active-tab');
}
