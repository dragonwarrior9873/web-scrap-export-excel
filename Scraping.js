import axios from 'axios';
import Excel from 'exceljs';
import path from 'path';
import XLSX from 'xlsx';
import tqdm from 'tqdm';

const workbook = new Excel.Workbook();
const exportPath = path.resolve('result.xlsx');
let startDate = new Date('2024-03-11');
let endDate = new Date('2024-04-12');
let currentDate = new Date(startDate.getTime());
let worksheet = [];
let month_iterator = 0;
while (currentDate <= endDate) {
    const dateString = currentDate.toISOString().slice(0, 10);
    worksheet[month_iterator] = workbook.addWorksheet(dateString);
    worksheet[month_iterator].columns = [
        { key: 'codAviz', header: 'CodAviz' },
        { key: 'emitent_denumire', header: 'Emitent_denumire' },
        { key: 'emitent_cui', header: 'Emitent_cui' },
        { key: 'provenienta', header: 'Provenienta' },
        { key: 'marfa_grupeSpecii', header: 'Marfa_grupeSpecii' },
        { key: 'marfa_specii', header: 'Marfa_specii' },
        { key: 'marfa_sortimente', header: 'Marfa_sortimente' },
        { key: 'marfa_total', header: 'Marfa_total' },
        { key: 'nrIdentificare', header: 'NrIdentificare' },
        { key: 'transportator_denumire', header: 'Transportator_denumire' },
        { key: 'transportator_cui', header: 'Transportator_cui' },
        { key: 'valabilitate_emitere', header: 'Valabilitate_emitere' },
        { key: 'valabilitate_finalizare', header: 'Valabilitate_finalizare' },
        { key: 'volum_volumSpecie', header: 'Volum_volumSpecie' },
    ];
    month_iterator++;
    currentDate.setDate(currentDate.getDate() + 1);
}

const response = await axios.get('https://inspectorulpadurii.ro/api/aviz/locations');
console.log(response.data.codAviz.length);
for (let i = 0; i < response.data.codAviz.length; i++) {
    console.log(`Iterate ${i}th element.`);
    let temp = response.data.codAviz[i];
    if (temp.slice(0, 2).toLowerCase() === "dc") {
        try {
            let content = await axios.get(`https://inspectorulpadurii.ro/api/aviz/${temp}`);
            let data = content.data;
            let inserted = {  
                'codAviz': data.codAviz,
                'emitent_denumire': data.emitent.denumire,
                'emitent_cui': data.emitent.cui,
                'provenienta': data.provenienta,
                'marfa_grupeSpecii': data.marfa.grupeSpecii,
                'marfa_specii': data.marfa.specii,
                'marfa_sortimente': data.marfa.sortimente,
                'marfa_total': data.marfa.total,
                'nrIdentificare': data.nrIdentificare,
                'transportator_denumire': data.transportator.denumire,
                'transportator_cui': data.transportator.cui,
                'valabilitate_emitere': formateDate(data.valabilitate.emitere),
                'valabilitate_finalizare': formateDate(data.valabilitate.finalizare),
                'volum_volumSpecie': data.volum.volumSpecie,
            };
            const dateObj = new Date(data.valabilitate.emitere).toISOString().slice(0, 10);
            month_iterator = 0;
            currentDate = new Date(startDate.getTime());
            while (currentDate <= endDate) {
                const dateString = currentDate.toISOString().slice(0, 10);
                if (dateString == dateObj) {
                    console.log(`Inserted to the Excel ${i}th row.`);
                    console.log(`Inserted Row has ${dateString} date`);
                    worksheet[month_iterator].addRow(inserted);
                    await workbook.xlsx.writeFile(exportPath);
                }
                currentDate.setDate(currentDate.getDate() + 1);
                month_iterator++;
            }
        }
        catch (err) {
            console.error(err);
        }
    }
}
convertTableSorted();


function convertTableSorted() {
    try {
        const workbookResult = XLSX.readFile("result.xlsx");
        const workbookout = XLSX.utils.book_new();
        const workbookSheet = workbookResult.SheetNames; 
        
        for( let i = 0 ; i < workbookSheet.length; i++ ){
            let workbookResponse = {};
            workbookResponse = XLSX.utils.sheet_to_json(workbookResult.Sheets[workbookSheet[i]]);
            workbookResponse.sort((a, b) => b.Marfa_total - a.Marfa_total);
            let workSheet = XLSX.utils.json_to_sheet(workbookResponse);
            XLSX.utils.book_append_sheet(workbookout, workSheet, workbookSheet[i]);    
        }
        
        XLSX.writeFile(workbookout, "output.xlsx")
    }
    catch (err) {
        console.error("No result.xlsx");
    }
}

function formateDate(date) {
    date = new Date(date);
    const hours = date.getHours().toString().padStart(2, '0');
    const minutes = date.getMinutes().toString().padStart(2, '0');
    const seconds = date.getSeconds().toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear();
    const formattedDate = `${hours}:${minutes}:${seconds} ${day}/${month}/${year}`;
    return formattedDate;
}