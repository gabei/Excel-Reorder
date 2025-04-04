/* 

- Read in excel file from input folder
- Remove extra columns
- Reorder columns
- Save manipulated sheet to output folder


- Columns to keep (in this order):
Vendor 	PO 	Received	Fund Code	Pub Date	Top Note	Title	Sub-Title	Author	Qty	List Price	Unit Price	Line Price													

*/



const ExcelJS = require('exceljs');

const requiredColumns = ["Vendor", "PO Received", "Fund", "Code",	"Pub Date", "Top Note", "Title",	"Sub-Title", "Author", "Qty",	"List Price",	"Unit Price",	"Line Price"];

const resultingColumns = [];

async function readInFile(){
  const workbook = new ExcelJS.Workbook();
  const worksheet = await workbook.csv.readFile('./input/input.csv');

  worksheet.columns.forEach(column => {
    const thisColumnTitle = column.values[1];

    if( requiredColumns.includes(thisColumnTitle) ) {
      resultingColumns.push(column);
    }

  });
}

readInFile();