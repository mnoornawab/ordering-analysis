function renderMainReport(data) {
  const container = document.getElementById("main-report-table");

  let html = `
    <label><input type="checkbox" id="hide-zero" onchange="toggleZeroRows(this)"> Hide rows with all 0 values</label>
    <table>
      <thead>
        <tr>
          <th>Item Code</th>
          <th>Style Code</th>
          <th>Qty On Order</th>
          <th>On Allocation File</th>
          <th>Balance from Orders</th>
        </tr>
      </thead>
      <tbody>
  `;

  let totals = { qtyOnOrder: 0, qtyAllocated: 0, balanceOrders: 0 };

  data.forEach(row => {
    const isZero = row.qtyOnOrder === 0 && row.qtyAllocated === 0 && row.balanceOrders === 0;
    const rowClass = isZero ? "zero-row" : "";

    html += `
      <tr class="${rowClass}">
        <td>${row.itemCode}</td>
        <td>${row.styleCode}</td>
        <td>${row.qtyOnOrder}</td>
        <td>${row.qtyAllocated}</td>
        <td>${row.balanceOrders}</td>
      </tr>
    `;

    totals.qtyOnOrder += row.qtyOnOrder;
    totals.qtyAllocated += row.qtyAllocated;
    totals.balanceOrders += row.balanceOrders;
  });

  html += `
      <tr>
        <th>Total</th><td></td>
        <th>${totals.qtyOnOrder}</th>
        <th>${totals.qtyAllocated}</th>
        <th>${totals.balanceOrders}</th>
      </tr>
    </tbody></table>
  `;

  container.innerHTML = html;
}

function toggleZeroRows(checkbox) {
  const rows = document.querySelectorAll(".zero-row");
  rows.forEach(row => {
    row.style.display = checkbox.checked ? "none" : "";
  });
}
