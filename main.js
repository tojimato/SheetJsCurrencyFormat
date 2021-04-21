/*
usage
----------------
<a class="btn btn-warning text-white cursor-pointer mr-3" onclick="ExportToExcelSalesReport('tbl-sales-list','tbl-sales-list')">
    <i class="fa fa-print"></i> Excel Olarak İndir
 </a>


*/

function formatColumnCurrency(worksheet, col) {           
      const range = XLSX.utils.decode_range(worksheet['!ref'])
      console.log(worksheet);
      // note: range.s.r + 1 skips the header row
      for (let row = range.s.r + 1; row <= range.e.r; ++row) {                
          const ref = XLSX.utils.encode_cell({ r: row, c: col })
          console.log(worksheet[ref]);
          if (worksheet[ref]) {
              worksheet[ref].v = worksheet[ref].v.replace(",", "");
              worksheet[ref].t = "n";
              worksheet[ref].z = "₺0.00";                                      
          }
      }
  }

function ExportToExcelSalesReport(tableId, fileName, fn, dl) {
    var elt = document.getElementById(tableId);
    var wb = XLSX.utils.table_to_book(elt, { sheet: "Sales", raw: true });
    formatColumnCurrency(wb.Sheets["Sales"], 8);
    formatColumnCurrency(wb.Sheets["Sales"], 11);

    return dl ?
        XLSX.write(wb, { bookType: "xlsx", bookSST: true, type: 'base64' }) :
        XLSX.writeFile(wb, fn || ('SaleReport.' + ("xlsx" || 'xlsx')));
}
