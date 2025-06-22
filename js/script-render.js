function generateMainTable(data) {
  let html = `<label><input type="checkbox" id="hide-zero" onchange="toggleZeroRows(this)"> Hide rows with all 0 values</label>`;
  html += `<table><thead><tr>
    <th>Item Code</th><th>Style Code</th><th>Qty On Order</th><th>On Allocation File</th><th>Balance from Orders</th>
  </tr></thead><tbody>`;

  let totals = { order: 0, allocation: 0, balance: 0 };

  data.forEach(row => {
    const { itemCode, styleCode, qtyOnOrder, qtyAllocated, balanceOrders } = row;
    const isZeroRow = (qtyOnOrder === 0 && qtyAllocated === 0 && balanceOrders === 0);
    const hideClass = isZeroRow ? 'zero-row' : '';

    html += `<tr class="${hideClass}">
      <td>${itemCode}</td>
      <td>${styleCode}</td>
      <td>${qtyOnOrder}</td>
      <td>${qtyAllocated}</td>
      <td>${balanceOrders}</td>
    </tr>`;

    totals.order += qtyOnOrder;
    totals.allocation += qtyAllocated;
    totals.balance += balanceOrders;
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
      html += `<tr style="background:#fff0f0;">
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

function generateToOrderTable(data) {
  let html = `<table><thead><tr>
    <th>Item Code</th><th>Style Code</th><th>Balance from Orders</th><th>Qty On Order</th><th>On Allocation File</th><th>Qty to Order</th>
  </tr></thead><tbody>`;

  let totalToOrder = 0;

  data.forEach(row => {
    const totalAvailable = row.qtyOnOrder + row.qtyAllocated;
    const required = row.balanceOrders;
    const toOrder = required - totalAvailable;

    if (toOrder > 0) {
      html += `<tr style="background:#fffbea;">
        <td>${row.itemCode}</td>
        <td>${row.styleCode}</td>
        <td>${required}</td>
        <td>${row.qtyOnOrder}</td>
        <td>${row.qtyAllocated}</td>
        <td>${toOrder}</td>
      </tr>`;
      totalToOrder += toOrder;
    }
  });

  html += `<tr><th>Total</th><td></td><td></td><td></td><td></td><th>${totalToOrder}</th></tr>`;
  html += `</tbody></table>`;
  return html;
}

function toggleZeroRows(checkbox) {
  const rows = document.querySelectorAll('.zero-row');
  rows.forEach(row => {
    row.style.display = checkbox.checked ? 'none' : '';
  });
}
