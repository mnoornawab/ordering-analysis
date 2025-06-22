function generateMainTable(data) {
  let html = `<label><input type="checkbox" id="hide-zero" onchange="toggleZeroRows(this)"> Hide rows with all 0 values</label>`;
  html += `<table><thead><tr>
    <th>Item Code</th>
    <th>Style Code</th>
    <th>Qty On Order</th>
    <th>On Allocation File</th>
    <th>Balance from Orders</th>
  </tr></thead><tbody>`;

  let totalOrder = 0;
  let totalAlloc = 0;
  let totalBalance = 0;

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

    totalOrder += qtyOnOrder;
    totalAlloc += qtyAllocated;
    totalBalance += balanceOrders;
  });

  html += `<tr><th>Total</th><td></td><th>${totalOrder}</th><th>${totalAlloc}</th><th>${totalBalance}</th></tr>`;
  html += `</tbody></table>`;
  return html;
}

function toggleZeroRows(checkbox) {
  const rows = document.querySelectorAll('.zero-row');
  rows.forEach(row => {
    row.style.display = checkbox.checked ? 'none' : '';
  });
}
