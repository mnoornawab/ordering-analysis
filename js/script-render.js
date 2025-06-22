function generateMainTable(data) {
  let html = `<label><input type="checkbox" id="hide-zero" onchange="toggleZeroRows(this)"> Hide rows with all 0 values</label>`;
  html += `<table><thead><tr>
    <th>Item Code</th><th>Style Code</th><th>Qty On Order</th><th>On Allocation File</th><th>Balance from Orders</th>
  </tr></thead><tbody>`;

  let totals = { order: 0, allocation: 0, balance: 0 };

  data.forEach(row => {
    const { itemCode, styleCode, qtyOnOrder, qtyAllocated, balanceOrders } = row;
    const isZero = qtyOnOrder === 0 && qtyAllocated === 0 && balanceOrders === 0;
    const rowClass = isZero ? 'zero-row' : '';

    html += `<tr class="${rowClass}">
      <td>${itemCode}</td>
      <td>${styleCode || '-'}</td>
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

  let mismatchOrder = 0, mismatchAlloc = 0;

  data.forEach(row => {
    if (row.qtyOnOrder !== row.qtyAllocated) {
      html += `<tr style="background-color:#fff3f3;">
        <td>${row.itemCode}</td>
        <td>${row.styleCode || '-'}</td>
        <td>${row.qtyOnOrder}</td>
        <td>${row.qtyAllocated}</td>
      </tr>`;
      mismatchOrder += row.qtyOnOrder;
      mismatchAlloc += row.qtyAllocated;
    }
  });

  html += `<tr><th>Total</th><td></td><th>${mismatchOrder}</th><th>${mismatchAlloc}</th></tr>`;
  html += `</tbody></table>`;
  return html;
}

function generateToOrderTable(data) {
  let html = `<table><thead><tr>
    <th>Item Code</th><th>Style Code</th><th>Balance from Orders</th><th>On Allocation File</th><th>Qty to Order</th>
  </tr></thead><tbody>`;

  let totalToOrder = 0;

  data.forEach(row => {
    const needed = row.balanceOrders - row.qtyAllocated;
    if (needed > 0) {
      html += `<tr style="background-color:#fffbe6;">
        <td>${row.itemCode}</td>
        <td>${row.styleCode || '-'}</td>
        <td>${row.balanceOrders}</td>
        <td>${row.qtyAllocated}</td>
        <td>${needed}</td>
      </tr>`;
      totalToOrder += needed;
    }
  });

  html += `<tr><th>Total</th><td></td><td></td><td></td><th>${totalToOrder}</th></tr>`;
  html += `</tbody></table>`;
  return html;
}

function toggleZeroRows(checkbox) {
  const rows = document.querySelectorAll('.zero-row');
  rows.forEach(row => {
    row.style.display = checkbox.checked ? 'none' : '';
  });
}
