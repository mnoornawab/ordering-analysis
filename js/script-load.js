function processFile() {
  const fileInput = document.getElementById('file-input');
  const file = fileInput.files[0];
  const statusDiv = document.getElementById('status');
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    // Sheets
    const simaSheet = workbook.Sheets['Quantity on Order - SIMA System'];
    const allocSheet = workbook.Sheets['Allocation File'];
    const orderSheet = workbook.Sheets['Orders-SIMA System'];

    const simaData = XLSX.utils.sheet_to_json(simaSheet, { defval: '' });
    const allocData = XLSX.utils.sheet_to_json(allocSheet, { defval: '' });
    const orderData = XLSX.utils.sheet_to_json(orderSheet, { defval: '' });

    const simaPreview = simaData.slice(0, 10).map(row =>
      `${row['Item Code']} | ${row['Material Code']} | ${row['Qty On Order']}`).join('<br>');

    const allocPreview = allocData.slice(0, 10).map(row =>
      `${row['Material code']} | ${row['Pending order qty']}`).join('<br>');

    const orderPreview = orderData.slice(0, 10).map(row =>
      `${row['ITEMCODE']} | ${row['STYLECODE']} | ${row['BALANCE']}`).join('<br>');

    document.getElementById('main-report').innerHTML = `
      <h3>📄 Quantity on Order - SIMA System</h3>
      <div class="debug-box">${simaPreview}</div>
      <h3>📄 Allocation File</h3>
      <div class="debug-box">${allocPreview}</div>
      <h3>📄 Orders-SIMA System</h3>
      <div class="debug-box">${orderPreview}</div>
    `;

    statusDiv.innerHTML = `<p style="color:green;">✅ File loaded and preview shown.</p>`;
  };
  reader.readAsArrayBuffer(file);
}
