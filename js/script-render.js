
function renderMainReport(data) {
  const container = document.getElementById('main-report');
  let html = '<table><thead><tr><th>Item Code</th><th>Style Code</th><th>Qty On Order</th><th>On Allocation File</th><th>Balance from Orders</th></tr></thead><tbody>';

  let totalOrder = 0, totalAlloc = 0, totalBalance = 0;

  data.forEach(row => {
    const zeroClass = (row.qtyOnOrder === 0 && row.qtyAllocated === 0 && row.balanceOrders === 0) ? ' class="zero-row"' : '';
    html += `<tr${zeroClass}><td>${row.itemCode}</td><td>${row.styleCode}</td><td>${row.qtyOnOrder}</td><td>${row.qtyAllocated}</td><td>${row.balanceOrders}</td></tr>`;
    totalOrder += row.qtyOnOrder;
    totalAlloc += row.qtyAllocated;
    totalBalance += row.balanceOrders;
  });

  html += `<tr><th>Total</th><td></td><th>${totalOrder}</th><th>${totalAlloc}</th><th>${totalBalance}</th></tr>`;
  html += '</tbody></table>';

  container.innerHTML = html;
  toggleZeroRows(document.getElementById('hide-zero'));
}

function renderMismatchReport(data) {
  const mismatch = data.filter(row => row.qtyOnOrder !== row.qtyAllocated);
  let html = '<table><thead><tr><th>Item Code</th><th>Style Code</th><th>Qty On Order</th><th>On Allocation File</th></tr></thead><tbody>';

  let totalOrder = 0, totalAlloc = 0;
  mismatch.forEach(row => {
    html += `<tr><td>${row.itemCode}</td><td>${row.styleCode}</td><td>${row.qtyOnOrder}</td><td>${row.qtyAllocated}</td></tr>`;
    totalOrder += row.qtyOnOrder;
    totalAlloc += row.qtyAllocated;
  });

  html += `<tr><th>Total</th><td></td><th>${totalOrder}</th><th>${totalAlloc}</th></tr>`;
  html += '</tbody></table>';

  document.getElementById('mismatch-report').innerHTML = html;
}

function renderToOrderReport(data) {
  const filtered = data.filter(row => row.qtyAllocated < row.balanceOrders);
  let html = '<table><thead><tr><th>Item Code</th><th>Style Code</th><th>Balance from Orders</th><th>On Allocation File</th><th>Qty To Order</th></tr></thead><tbody>';

  let totalBal = 0, totalAlloc = 0, totalToOrder = 0;

  filtered.forEach(row => {
    const toOrder = row.balanceOrders - row.qtyAllocated;
    html += `<tr><td>${row.itemCode}</td><td>${row.styleCode}</td><td>${row.balanceOrders}</td><td>${row.qtyAllocated}</td><td>${toOrder}</td></tr>`;
    totalBal += row.balanceOrders;
    totalAlloc += row.qtyAllocated;
    totalToOrder += toOrder;
  });

  html += `<tr><th>Total</th><td></td><th>${totalBal}</th><th>${totalAlloc}</th><th>${totalToOrder}</th></tr>`;
  html += '</tbody></table>';

  document.getElementById('to-order-report').innerHTML = html;
}

function toggleZeroRows(checkbox) {
  const rows = document.querySelectorAll('.zero-row');
  rows.forEach(row => {
    row.style.display = checkbox.checked ? 'none' : '';
  });
}

function openTab(tabName) {
  const contents = document.querySelectorAll('.tab-content');
  contents.forEach(c => c.classList.remove('active-tab'));

  const buttons = document.querySelectorAll('.tab-button');
  buttons.forEach(b => b.classList.remove('active'));

  document.getElementById(tabName).classList.add('active-tab');
  event.target.classList.add('active');
}
