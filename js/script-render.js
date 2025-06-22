function renderMainReport(data) {
  let html = `<label><input type="checkbox" id="hide-zero" onchange="toggleZeroRows(this)"> Hide rows with all 0 values</label>`;
  html += `<table><thead><tr>
    <th>Item Code</th><th>Style Code</th><th>Qty On Order</th><th>On Allocation File</th><th>Balance from Orders</th>
  </tr></thead><tbody>`;

  let totals = { order: 0, allocation: 0, balance: 0 };

  data.forEach(row => {
    const isZeroRow = row.qtyOnOrder === 0 && row.qtyAllocated === 0 && row.balanceOrders === 0;
    html += `<tr class="${isZeroRow ? 'zero-row' : ''}">
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

  html += `<tr class="total-row">
    <th>Total</th><td></td><th>${totals.order}</th><th>${totals.allocation}</th><th>${totals.balance}</th>
  </tr>`;
  html += `</tbody></table>`;

  document.getElementById('main-report').innerHTML = html;
}

function renderMismatchReport(data) {
  let html = `<table><thead><tr>
    <th>Item Code</th><th>Style Code</th><th>Qty On Order</th><th>On Allocation File</th>
  </tr></thead><tbody>`;

  let totalOrder = 0;
  let totalAlloc = 0;

  data.forEach(row => {
    if (row.qtyOnOrder !== row.qtyAllocated) {
      html += `<tr class="highlight-mismatch">
        <td>${row.itemCode}</td>
        <td>${row.styleCode}</td>
        <td>${row.qtyOnOrder}</td>
        <td>${row.qtyAllocated}</td>
      </tr>`;
      totalOrder += row.qtyOnOrder;
      totalAlloc += row.qtyAllocated;
    }
  });

  html += `<tr class="total-row">
    <th>Total</th><td></td><th>${totalOrder}</th><th>${totalAlloc}</th>
  </tr>`;
  html += `</tbody></table>`;

  document.getElementById('mismatch-report').innerHTML = html;
}

function renderToOrderReport(data) {
  let html = `<table><thead><tr>
    <th>Item Code</th><th>Style Code</th><th>Balance from Orders</th><th>On Allocation File</th><th>Qty to Order</th>
  </tr></thead><tbody>`;

  let totalToOrder = 0;

  data.forEach(row => {
    const qtyToOrder = Math.max(0, row.balanceOrders - row.qtyAllocated);
    if (qtyToOrder > 0) {
      html += `<tr class="highlight-order">
        <td>${row.itemCode}</td>
        <td>${row.styleCode}</td>
        <td>${row.balanceOrders}</td>
        <td>${row.qtyAllocated}</td>
        <td>${qtyToOrder}</td>
      </tr>`;
      totalToOrder += qtyToOrder;
    }
  });

  html += `<tr class="total-row">
    <th>Total</th><td></td><td></td><td></td><th>${totalToOrder}</th>
  </tr>`;
  html += `</tbody></table>`;

  document.getElementById('to-order-report').innerHTML = html;
}

function toggleZeroRows(checkbox) {
  const rows = document.querySelectorAll('.zero-row');
  rows.forEach(row => {
    row.style.display = checkbox.checked ? 'none' : '';
  });
}
