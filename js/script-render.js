function renderMismatchReport() {
    const container = document.getElementById("mismatchResults");
    const totalsDiv = document.getElementById("mismatchTotals");
    if (!qtyData || !allocData) {
        container.innerHTML = "<p>No data found.</p>";
        totalsDiv.innerHTML = "";
        return;
    }

    // Build a lookup for Qty on Order and Style by Item Code from Sheet 1
    const qtyOnOrderMap = {};
    const styleMap = {};
    qtyData.forEach(row => {
        const code = getColumnValue(row, ["Item Code"]);
        qtyOnOrderMap[code] = Number(getColumnValue(row, ["Qty on Order", "Qty On Order", "Qty on order"]) || 0);
        styleMap[code] = getColumnValue(row, ["Style", "Style Code"]) || "";
    });

    // For every allocation row in Sheet 2, compare to Sheet 1 Qty on Order
    const mismatchRows = allocData.map(row => {
        const code = getColumnValue(row, ["Item Code"]);
        const allocQty = Number(getColumnValue(row, ["Pending Order Qty", "Pending order qty"]) || 0);
        const qtyOnOrder = qtyOnOrderMap[code] || 0;
        const style = styleMap[code] || "";
        return {
            "Item Code": code,
            "Style Code": style,
            "Qty on Order": qtyOnOrder,
            "On Allocation File (row)": allocQty,
            "Mismatch Note": qtyOnOrder !== allocQty
                ? `Qty on Order (${qtyOnOrder}) ≠ Allocation Row (${allocQty})`
                : ""
        };
    });

    if (mismatchRows.length === 0) {
        container.innerHTML = "<p>No mismatches found.</p>";
        totalsDiv.innerHTML = "";
        return;
    }

    const headers = ["Item Code", "Style Code", "Qty on Order", "On Allocation File (row)", "Mismatch Note"];
    let html = "<table><thead><tr>";
    headers.forEach(h => html += `<th>${h}</th>`);
    html += "</tr></thead><tbody>";
    mismatchRows.forEach(row => {
        html += "<tr>";
        headers.forEach(h => html += `<td>${row[h]}</td>`);
        html += "</tr>";
    });

    // Totals row
    const totalQtyOnOrder = mismatchRows.reduce((sum, r) => sum + (Number(r["Qty on Order"]) || 0), 0);
    const totalAlloc = mismatchRows.reduce((sum, r) => sum + (Number(r["On Allocation File (row)"]) || 0), 0);
    html += "<tr style='font-weight: bold; background: #eaf2fb;'>";
    headers.forEach(h => {
        if (h === "Item Code") html += "<td>TOTAL</td>";
        else if (h === "Qty on Order") html += `<td>${totalQtyOnOrder}</td>`;
        else if (h === "On Allocation File (row)") html += `<td>${totalAlloc}</td>`;
        else html += "<td></td>";
    });
    html += "</tr>";

    html += "</tbody></table>";
    container.innerHTML = html;
    totalsDiv.innerHTML = `Totals — Qty on Order: ${totalQtyOnOrder} | On Allocation File (row): ${totalAlloc}`;
}
