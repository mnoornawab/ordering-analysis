
document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("file-input").addEventListener("change", handleFile, false);
});

let sheetData = [];

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    const sheet1 = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "" });
    const sheet2 = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[1]], { defval: "" });
    const sheet3 = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[2]], { defval: "" });

    const itemMap = {};

    // Sheet 1: Base
    sheet1.forEach(row => {
      const itemCode = row['Item Code']?.toString().trim();
      const style = row['Style']?.toString().trim();
      const qty = parseInt(row['Qty On Order']) || 0;
      if (!itemCode) return;
      if (!itemMap[itemCode]) {
        itemMap[itemCode] = {
          'Item Code': itemCode,
          'Style': style,
          'Qty On Order': 0,
          'Pending Order Qty': 0,
          'BALANCE': 0
        };
      }
      itemMap[itemCode]['Qty On Order'] += qty;
    });

    // Sheet 2: Allocation File
    sheet2.forEach(row => {
      const itemCode = row['Item Code']?.toString().trim();
      const qty = parseInt(row['Pending Order Qty']) || 0;
      if (itemMap[itemCode]) {
        itemMap[itemCode]['Pending Order Qty'] += qty;
      }
    });

    // Sheet 3: Orders
    sheet3.forEach(row => {
      const itemCode = row['Item Code']?.toString().trim();
      const qty = parseInt(row['Balance from Orders']) || 0;
      if (itemMap[itemCode]) {
        itemMap[itemCode]['BALANCE'] += qty;
      }
    });

    sheetData = Object.values(itemMap);

    document.getElementById("main-report").innerHTML = generateMainTable(sheetData);
    document.getElementById("to-order-report").innerHTML = generateToOrderTable(sheetData);
    document.getElementById("status").innerHTML = "<p style='color:green;'>✅ File processed successfully.</p>";
  };
  reader.readAsArrayBuffer(file);
}
