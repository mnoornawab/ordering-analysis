// script-load.js

async function handleFileUpload(event) {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = async function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const simaSheet = workbook.Sheets['Quantity on Order - SIMA System'];
    const allocationSheet = workbook.Sheets['Allocation File'];
    const ordersSheet = workbook.Sheets['Orders-SIMA System'];

    const simaData = XLSX.utils.sheet_to_json(simaSheet, { defval: '' });
    const allocationData = XLSX.utils.sheet_to_json(allocationSheet, { defval: '' });
    const ordersData = XLSX.utils.sheet_to_json(ordersSheet, { defval: '' });

    const base = simaData.map(row => {
      const itemCode = row['Item Code'] || row['ITEMCODE'] || row['ITEM CODE'] || row['itemcode'];
      const styleCode = row['Material Code'] || row['STYLECODE'] || row['Style Code'];
      const qtyOnOrder = Number(row['Qty On Order'] || row['QTY ON ORDER'] || row['Qty on Order'] || 0);
      return { itemCode, styleCode, qtyOnOrder };
    }).filter(item => item.itemCode);

    const allocations = {};
    allocationData.forEach(row => {
      const code = row['Material code'];
      const qty = Number(row['Pending order qty'] || 0);
      if (!allocations[code]) allocations[code] = 0;
      allocations[code] += qty;
    });

    const orderBalances = {};
    ordersData.forEach(row => {
      const code = row['ITEMCODE'];
      const balance = Number(row['BALANCE'] || 0);
      if (!orderBalances[code]) orderBalances[code] = 0;
      orderBalances[code] += balance;
    });

    const merged = base.map(row => {
      return {
        itemCode: row.itemCode,
        styleCode: row.styleCode,
        qtyOnOrder: row.qtyOnOrder,
        onAllocation: allocations[row.itemCode] || 0,
        balanceFromOrders: orderBalances[row.itemCode] || 0,
      };
    });

    window.stockData = merged;
    renderFullReport(merged);
    renderMismatchReport(merged);
  };

  reader.readAsArrayBuffer(file);
}

