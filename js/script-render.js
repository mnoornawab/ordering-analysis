function renderMainReport(data) {
  let html = `<label><input type="checkbox" id="hide-zero" onchange="toggleZeroRows(this)"> Hide rows with all 0 values</label>`;
  html += `<table><thead><tr>
    <th>Item Code</th>
    <th>Style Code</th>
    <th>Qty On Order</th>
    <th>On Allocation File</th>
    <th>Balance from Orders</th>
  </tr></thead><tbody>`;

  let totalOrder = 0, totalAlloc = 0, totalBalance = 0;

  data.forEach(row => {
    const hideClass = (row.qtyOnOrder === 0 && row.qtyAllocated === 0 && row.balanceOrders === 0) ? 'zero-row' : '';
    html += `<tr class="${hideClass}">
      <td>${row.itemCode}</td>
      <td>${row.styleCode}</td>
      <td>${row.qtyOnOrder}</td>
      <td>${row.qtyAllocated}</td>
      <td>${row.balanceOrders}</td>
    </tr>`;
    totalOrder += row.qtyOnOrder;
    totalAlloc += row.qtyAllocated;
    totalBalance += row.balanceOrders;
  });

  html += `<tr class="total-row">
    <th>Total</th><td></td>
    <th>${totalOrder}</th>
    <th>${totalAlloc}</th>
    <th>${totalBalance}</th>
  </tr>`;
  html += `</tbody></table>`;

  document.getElementById('main-report').innerHTML = html;
}

function toggleZeroRows(checkbox) {
  const rows = document.querySelectorAll('.zero-row');
  rows.forEach(row => {
    row.style.display = checkbox.checked ? 'none' : '';
  });
}

function renderMismatchReport(data) {
  const filtered = data.filter(row => row.qtyOnOrder !== row.qtyAllocated);

  let html = `<table><thead><tr>
    <th>Item Code</th>
    <th>Style Code</th>
    <th>Qty On Order</th>
    <th>On Allocation File</th>
  </tr></thead><tbody>`;

  let totalOrder = 0, totalAlloc = 0;

  filtered.forEach(row => {
    html += `<tr style="background:#fff1f1;">
      <td>${row.itemCode}</td>
      <td>${row.styleCode}</td>
      <td>${row.qtyOnOrder}</td>
      <td>${row.qtyAllocated}</td>
    </tr>`;
    totalOrder += row.qtyOnOrder;
    totalAlloc += row.qtyAllocated;
  });

  html += `<tr class="total-row">
    <th>Total</th><td></td>
    <th>${totalOrder}</th><th>${totalAlloc}</th>
  </tr>`;
  html += `</tbody></table>`;

  document.getElementById('mismatch-report').innerHTML = html;
}

function renderToOrderReport(data) {
  const filtered = data.filter(row => row.balanceOrders > row.qtyAllocated);

  let html = `<table><thead><tr>
    <th>Item Code</th>
    <th>Style Code</th>
    <th>Balance from Orders</th>
    <th>On Allocation File</th>
    <th>Qty To Order</th>
  </tr></thead><tbody>`;

  let totalBalance = 0, totalAlloc = 0, totalToOrder = 0;

  filtered.forEach(row => {
    const toOrder = row.balanceOrders - row.qtyAllocated;
    html += `<tr>
      <td>${row.itemCode}</td>
      <td>${row.styleCode}</td>
      <td>${row.balanceOrders}</td>
      <td>${row.qtyAllocated}</td>
      <td>${toOrder}</td>
    </tr>`;
    totalBalance += row.balanceOrders;
    totalAlloc += row.qtyAllocated;
    totalToOrder += toOrder;
  });

  html += `<tr class="total-row">
    <th>Total</th><td></td>
    <th>${totalBalance}</th><th>${totalAlloc}</th><th>${totalToOrder}</th>
  </tr>`;
  html += `</tbody></table>`;

  document.getElementById('to-order-report').innerHTML = html;
}
