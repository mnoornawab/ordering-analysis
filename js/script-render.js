function renderMainReport(data) {
  const container = document.getElementById('main-report');

  let html = `<label><input type="checkbox" id="hide-zero" onchange="toggleZeroRows(this, ${JSON.stringify(data).replace(/"/g, '&quot;')})"> Hide rows with all 0 values</label>`;
  html += `<table><thead><tr>
    <th>Item Code</th><th>Style Code</th><th>Qty On Order</th><th>On Allocation File</th><th>Balance from Orders</th>
  </tr></thead><tbody>`;

  data.forEach(row => {
    const isZeroRow = row.qtyOnOrder === 0 && row.qtyAllocated === 0 && row.balanceOrders === 0;
    html += `<tr class="${isZeroRow ? 'zero-row' : ''}">
      <td>${row.itemCode}</td>
      <td>${row.styleCode}</td>
      <td>${row.qtyOnOrder}</td>
      <td>${row.qtyAllocated}</td>
      <td>${row.balanceOrders}</td>
    </tr>`;
  });

  html += `</tbody><tfoot id="main-totals"></tfoot></table>`;
  container.innerHTML = html;

  updateMainTotals(data);
}

function updateMainTotals(data, hideZeros = false) {
  let totals = { order: 0, allocation: 0, balance: 0 };

  data.forEach(row => {
    const isZero = row.qtyOnOrder === 0 && row.qtyAllocated === 0 && row.balanceOrders === 0;
    if (!hideZeros || !isZero) {
      totals.order += row.qtyOnOrder;
      totals.allocation += row.qtyAllocated;
      totals.balance += row.balanceOrders;
    }
  });

  const totalHTML = `<tr class="total-row">
    <th>Total</th><td></td><th>${totals.order}</th><th>${totals.allocation}</th><th>${totals.balance}</th>
  </tr>`;

  document.getElementById('main-totals').innerHTML = totalHTML;
}

function toggleZeroRows(checkbox, jsonData) {
  const rows = document.querySelectorAll('.zero-row');
  rows.forEach(row => row.style.display = checkbox.checked ? 'none' : '');

  const parsed = JSON.parse(jsonData);
  updateMainTotals(parsed, checkbox.checked);
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
  </tr></tbody></table>`;

  document.getElementById('mismatch-report').innerHTML = html;
}

function renderToOrderReport(data) {
  let html = `<table><thead><tr>
    <th>Item Code</th><th>Style Code</th><th>Balance from Orders</th><th>On Allocation File</th><th>Qty to Order</th>
  </tr></thead><tbody>`;

  let totalToOrder = 0;

  data.forEach(row => {
    const toOrder = Math.max(0, row.balanceOrders - row.qtyAllocated);
    if (toOrder > 0) {
      html += `<tr class="highlight-order">
        <td>${row.itemCode}</td>
        <td>${row.styleCode}</td>
        <td>${row.balanceOrders}</td>
        <td>${row.qtyAllocated}</td>
        <td>${toOrder}</td>
      </tr>`;
      totalToOrder += toOrder;
    }
  });

  html += `<tr class="total-row">
    <th>Total</th><td></td><td></td><td></td><th>${totalToOrder}</th>
  </tr></tbody></table>`;

  document.getElementById('to-order-report').innerHTML = html;
}
