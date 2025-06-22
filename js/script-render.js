export function renderMainReportTable(data) {
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

  const container = document.getElementById('main-report');
  if (container) container.innerHTML = html;
}

export function renderMismatchTable(data) {
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

  const container = document.getElementById('mismatch-report');
  if (container) container.innerHTML = html;
}

