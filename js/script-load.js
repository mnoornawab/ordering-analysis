function processFile() {
  const fileInput = document.getElementById('file-input');
  const file = fileInput.files[0];
  const statusDiv = document.getElementById('status');
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const sheet1 = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "" });
    const sheet2 = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[1]], { defval: "" });
    const sheet3 = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[2]], { defval: "" });

    // Clean and normalize all Item Codes as strings
    const normalize = (val) => val?.toString().trim();

    // Group Allocation File (Sheet 2)
    const allocMap = {};
    sheet2.forEach(row => {
      const itemCode = normalize(row["Item Code"]);
      const qty = parseInt(row["Qty Allocated"]) || 0;
      if (itemCode) {
        allocMap[itemCode] = (allocMap[itemCode] || 0) + qty;
      }
    });

    // Group Orders File (Sheet 3)
    const ordersMap = {};
    sheet3.forEach(row => {
      const itemCode = normalize(row["Item Code"]);
      const balance = parseInt(row["Balance from Orders"]) || 0;
      if (itemCode) {
        ordersMap[itemCode] = (ordersMap[itemCode] || 0) + balance;
      }
    });

    // Build final dataset from Sheet 1 (Base)
    const finalData = sheet1.map(row => {
      const itemCode = normalize(row["Item Code"]);
      const styleCode = normalize(row["Style Code"]);
      const qtyOnOrder = parseInt(row["Qty On Order"]) || 0;
      const qtyAllocated = allocMap[itemCode] || 0;
      const balanceOrders = ordersMap[itemCode] || 0;

      return {
        itemCode,
        styleCode,
        qtyOnOrder,
        qtyAllocated,
        balanceOrders,
      };
    });

    // Store to global state
    window.finalData = finalData;
    statusDiv.innerHTML = "<p style='color:green;'>✅ File uploaded successfully</p>";
    renderMainReport(finalData);
  };

  reader.readAsArrayBuffer(file);
}
