
function generateMainTable(data) {
  let html = '<label><input type="checkbox" id="hide-zero" onchange="toggleZeroRows(this)"> Hide rows with all 0 values</label>';
  html += '<table><thead><tr><th>Item Code</th><th>Style</th><th>Qty On Order</th><th>On Allocation File</th><th>Balance from Orders</th></tr></thead><tbody>';

  let totalOrder = 0, totalAlloc = 0, totalBalance = 0;

  data.forEach(row => {
    const hideClass = (row['Qty On Order'] === 0 && row['Pending Order Qty'] === 0 && row['BALANCE'] === 0) ? 'zero-row' : '';
    html += `<tr class="${hideClass}"><td>${row['Item Code']}</td><td>${row['Style']}</td><td>${row['Qty On Order']}</td><td>${row['Pending Order Qty']}</td><td>${row['BALANCE']}</td></tr>`;
    totalOrder += row['Qty On Order'];
    totalAlloc += row['Pending Order Qty'];
    totalBalance += row['BALANCE'];
  });

  html += `<tr><th>Total</th><td></td><th>${totalOrder}</th><th>${totalAlloc}</th><th>${totalBalance}</th></tr>`;
  html += '</tbody></table>';
  return html;
}

function generateToOrderTable(data) {
  let html = '<table><thead><tr><th>Item Code</th><th>Style</th><th>Balance from Orders</th><th>On Allocation File</th><th>To Order</th></tr></thead><tbody>';
  let totalBalance = 0, totalAlloc = 0, totalToOrder = 0;

  data.forEach(row => {
    const toOrder = row['BALANCE'] - row['Pending Order Qty'];
    if (toOrder > 0) {
      html += `<tr><td>${row['Item Code']}</td><td>${row['Style']}</td><td>${row['BALANCE']}</td><td>${row['Pending Order Qty']}</td><td>${toOrder}</td></tr>`;
      totalBalance += row['BALANCE'];
      totalAlloc += row['Pending Order Qty'];
      totalToOrder += toOrder;
    }
  });

  html += `<tr><th>Total</th><td></td><th>${totalBalance}</th><th>${totalAlloc}</th><th>${totalToOrder}</th></tr>`;
  html += '</tbody></table>';
  return html;
}
