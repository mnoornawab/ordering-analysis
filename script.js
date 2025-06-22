// script.js

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

    // Create a base item map from SIMA sheet
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

    // Merge allocation file data
    allocData.forEach(row => {
      const itemCode = row['Material code']?.toString().trim();
      const qty = parseInt(row['Pending order qty']) || 0;
      if (itemMap[itemCode]) {
        itemMap[itemCode].qtyAllocated += qty;
      }
    });

    // Merge orders balance data
    ordersData.forEach(row => {
      const itemCode = row['ITEMCODE']?.toString().trim();
      const balance = parseInt(row['BALANCE']) || 0;
      if (itemMap[itemCode]) {
        itemMap[itemCode].balanceOrders += balance;
      }
    });

    const finalData = Object.values(itemMap);

    // Render both reports
    document.getElementById('main-report').innerHTML = generateMainTable(finalData);
    document.getElementById('mismatch-report').innerHTML = generateMismatchTable(finalData);
    statusDiv.innerHTML = '<p style="color:green;">✅ File processed successfully.</p>';
  };
  reader.readAsArrayBuffer(file);
}

function generateMainTable(data) {
  let html = `<label><input type="checkbox" id="hide-zero" onchange="toggleZeroRows(this)"> Hide rows with all 0 values</label>`;
  html += `<table><thead><tr>
    <th>Item Code</th><th>Style Code</th><th>Qty On Order</th><th>On Allocation File</th><th>Balance from Orders</th>
  </tr></thead><tbody>`;

  let totals = { order: 0, allocation: 0, balance: 0 };

  data.forEach(row => {
    const hideClass = (row.qtyOnOrder === 0 && row.qtyAllocated === 0 && row.balanceOrders === 0) ? 'zero-row' : '';
    html += `<tr class="${hideClass}">
      <td>${row.itemCode}</td>
      <td>${row.styleCode}</td>
      <td>${row.qtyOnOrder}</td>
      <td>${row.qtyAllocated}</td>
      <td>${row.balanceOrders}</td>
    </tr>`;

    totals.order += row.qtyOnOrder;
    totals.allocation += row.qtyAllocated;
    totals.balance += row.balanceOrders;
  });

  html += `<tr><th>Total</th><td></td><th>${totals.order}</th><th>${totals.allocation}</th><th>${totals.balance}</th></tr>`;
  html += `</tbody></table>`;
  return html;
}

function generateMismatchTable(data) {
  let html = `<table><thead><tr>
    <th>Item Code</th><th>Style Code</th><th>Qty On Order</th><th>On Allocation File</th>
  </tr></thead><tbody>`;

  let totalOrder = 0, totalAlloc = 0;

  data.forEach(row => {
    if (row.qtyOnOrder !== row.qtyAllocated) {
      html += `<tr style="background:#ffecec;">
        <td>${row.itemCode}</td>
        <td>${row.styleCode}</td>
        <td>${row.qtyOnOrder}</td>
        <td>${row.qtyAllocated}</td>
      </tr>`;
      totalOrder += row.qtyOnOrder;
      totalAlloc += row.qtyAllocated;
    }
  });

  html += `<tr><th>Total</th><td></td><th>${totalOrder}</th><th>${totalAlloc}</th></tr>`;
  html += `</tbody></table>`;
  return html;
}

function toggleZeroRows(checkbox) {
  const rows = document.querySelectorAll('.zero-row');
  rows.forEach(row => {
    row.style.display = checkbox.checked ? 'none' : '';
  });
}
