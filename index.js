/* 

- Read in excel file from input folder
- Remove extra columns
- Reorder columns
- Save manipulated sheet to output folder


- Columns to keep (in this order):
Vendor 	PO 	Received	Fund Code	Pub Date	Top Note	Title	Sub-Title	Author	Qty	List Price	Unit Price	Line Price													

*/



const ExcelJS = require('exceljs');

const requiredColumns = ["PO Received",	"Pub-Date", "Top Note", "Title", "Sub-Title", "Author-Artist", "Qty-Ordered", "List-Price", "Unit Price", "Line Price"];



async function readInFile(){
  const workbook = new ExcelJS.Workbook();
  const worksheet = await workbook.csv.readFile('./input/input.csv');

  const resultingColumns = [];

  worksheet.columns.forEach(column => {
    const thisColumnTitle = column.values[1];
    
    if( requiredColumns.includes(thisColumnTitle) ) {
      resultingColumns.push(column);
    }

  });

  return resultingColumns;
}

// build new worksheet
const workbook = new ExcelJS.Workbook();
const worksheet =  workbook.addWorksheet('sheet', {
  pageSetup:{paperSize: 9, orientation:'landscape'}
});

worksheet.columns = [
  {header: 'PO Received', key: 'poReceived', width: 10},
  {header: 'Pub-Date', key: 'pubDate', width: 10},
  {header: 'Top Note', key: 'topNote', width: 10},
  {header: 'Title', key: 'title', width: 10},
  {header: 'Subtitle', key: 'subtitle', width: 10},
  {header: 'Author-Artist', key: 'author', width: 10},
  {header: 'Qty-Ordered', key: 'qty', width: 10},
  {header: 'List-Price', key: 'listPrice', width: 10},
  {header: 'Unit-Price', key: 'unitPrice', width: 10},
  {header: 'Line-Price', key: 'linePrice', width: 10},
];

async function populateCells() {
  // use key to match input data with new columns
  // iterate through cell of column
  // input data from input CSV into corresponding cells
}

function readRelevantData(data){
  data.forEach((column) => {
    console.log(column.values[1]);
  });
}

async function init(){
  readInFile().then((data) => {
    readRelevantData(data);
  });
}

init();