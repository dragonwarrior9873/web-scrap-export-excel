import XLSX from 'xlsx';
console.log(new Date(1712673665110));
const date = new Date(1712673665110);
const hours = date.getHours().toString().padStart(2, '0');
const minutes = date.getMinutes().toString().padStart(2, '0');
const seconds = date.getSeconds().toString().padStart(2, '0');
const day = date.getDate().toString().padStart(2, '0');
const month = (date.getMonth() + 1).toString().padStart(2, '0');
const year = date.getFullYear();

const formattedDate = `${hours}:${minutes}:${seconds} ${day}/${month}/${year}`;
console.log(formattedDate);

const workbook = XLSX.readFile("result-3.xlsx");
const workbookout = XLSX.utils.book_new();
const workbookSheet = workbook.SheetNames; 
let workbookResponse = {};
for( let i = 0 ; i < 1; i++ ){
    workbookResponse = XLSX.utils.sheet_to_json(workbook.Sheets[workbookSheet[i]]);
    workbookResponse.sort((a, b) => b.Marfa_total - a.Marfa_total);
    let workSheet = XLSX.utils.json_to_sheet(workbookResponse);
    XLSX.utils.book_append_sheet(workbookout, workSheet, workbookSheet[i]);    
}
console.log(workbookResponse);
XLSX.writeFile(workbookout, "output.xlsx")