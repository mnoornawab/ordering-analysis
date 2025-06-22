function openTab(tabId) {
  document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active-tab'));

  document.querySelector(`button[onclick="openTab('${tabId}')"]`).classList.add('active');
  document.getElementById(tabId).classList.add('active-tab');
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
      html += `<tr class="highlight">
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
