
function downloadCSV() {
  const activeTab = document.querySelector('.tab-content.active-tab');
  const table = activeTab.querySelector('table');
  if (!table) return;

  let csv = '';
  const rows = table.querySelectorAll('tr');
  rows.forEach(row => {
    const cols = row.querySelectorAll('th, td');
    const rowData = Array.from(cols).map(col => `"${col.innerText}"`).join(',');
    csv += rowData + '\n';
  });

  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);

  const a = document.createElement('a');
  a.href = url;
  a.download = 'report.csv';
  a.click();

  URL.revokeObjectURL(url);
}
