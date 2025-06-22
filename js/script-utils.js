function exportToExcel(sheetsObj, fileName) {
    const wb = XLSX.utils.book_new();
    for (const sheetName in sheetsObj) {
        const ws = XLSX.utils.aoa_to_sheet(sheetsObj[sheetName]);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    }
    XLSX.writeFile(wb, fileName);
}
