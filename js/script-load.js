import { renderMainReportTable, renderMismatchTable } from './script-render.js';

window.processFile = function () {
  const fileInput = document.getElementById('file-input');
  const file = fileInput.files[0];
  const statusDiv = document.getElementById('status');

  if (!file) {
    alert("Please upload a file.");
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const sheetSIMA = workbook.Sheets['Quantity on Order - SIMA System'];
      const sheetAllocation = workbook.Sheets['Allocation File'];
      const sheetOrders = workbook.Sheets['Orders-SIMA System'];

      if (!sheetSIMA || !sheetAllocation || !sheetOrders) {
        statusDiv.innerHTML = "<span style='color:red'>❌ One or more sheets not found!</span>";
        return;
      }

      const simaData = XLSX.utils.sheet_to_json(sheetSIMA, { defval: '' });
      const allocData = XLSX.utils.sheet_to_json(sheetAllocation, { defval: '' });
      const ordersData = XLSX.utils.sheet_to_json(sheetOrders, { defval: '' });

      const itemMap = {};

      // STEP 1: Load SIMA base
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

      // STEP 2: Add Allocation Quantities
      allocData.forEach(row => {
        const itemCode = row['Material code']?.toString().trim();
        const qty = parseInt(row['Pending order qty']) || 0;
        if (itemMap[itemCode]) {
          itemMap[itemCode].qtyAllocated += qty;
        }
      });

      // STEP 3: Add Order Balances
      ordersData.forEach(row => {
        const itemCode = row['ITEMCODE']?.toString().trim();
        const balance = parseInt(row['BALANCE']) || 0;
        if (itemMap[itemCode]) {
          itemMap[itemCode].balanceOrders += balance;
        }
      });

      const finalData = Object.values(itemMap);

      renderMainReportTable(finalData);
      renderMismatchTable(finalData);

      statusDiv.innerHTML = "<span style='color:green'>✅ File processed successfully.</span>";
    } catch (err) {
      statusDiv.innerHTML = `<span style='color:red'>❌ Error: ${err.message}</span>`;
    }
  };

  reader.readAsArrayBuffer(file);
};

// Optional utility: switch tab visibility
window.openTab = function (tabId) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active-tab'));
  document.getElementById(tabId).classList.add('active-tab');

  document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
  const button = Array.from(document.querySelectorAll('.tab-button'))
    .find(btn => btn.textContent.includes(tabId === 'mismatch' ? 'Mismatch' : 'Full'));
  if (button) button.classList.add('active');
};

// Optional utility: toggle 0-only rows
window.toggleZeroRows = function (checkbox) {
  const rows = document.querySelectorAll('.zero-row');
  rows.forEach(row => {
    row.style.display = checkbox.checked ? 'none' : '';
  });
};
