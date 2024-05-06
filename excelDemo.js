const ExcelJs = require('exceljs');

//Way 1 via Promise
//Workbook
/*const workBook = new ExcelJs.Workbook();
//Read File
workBook.xlsx.readFile("/Users/shankykalra/Downloads/download.xlsx").then(function () {
    //Sheet
    const worksheet = workBook.getWorksheet('Sheet1');
    //Traversing Row and Cell
    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, collNumber) => {
            console.log(cell.value);
        });
    });
})*/


//Way 2 Via async and await
async function excelTravserse() {
console.log("Via Function");
    const workBook = new ExcelJs.Workbook();
    //Read File
    await workBook.xlsx.readFile("/Users/shankykalra/Downloads/download.xlsx")
    //Sheet
    const worksheet = workBook.getWorksheet('Sheet1');
    //Traversing Row and Cell
    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, collNumber) => {
            console.log(cell.value);
        });
    });
}
//excelTravserse();


//Getting Specific Cell Number and Value
 
let output ={row:1, col:1};
let appleRow =0;

async function excelGetCoordinates() {
        const workBook = new ExcelJs.Workbook();
        //Read File
        await workBook.xlsx.readFile("/Users/shankykalra/Downloads/download.xlsx")
        //Sheet
        const worksheet = workBook.getWorksheet('Sheet1');
        //Traversing Row and Cell
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, collNumber) => {
                if(cell.value === 'Banana'){
                    output.row = rowNumber;
                    output.col = collNumber;
                }
                
            });
        });

        //Writing Specific Cell
        const cellToWrite = worksheet.getCell(output.row,output.col);
        cellToWrite.value = 'New Avocado';
        await workBook.xlsx.writeFile("/Users/shankykalra/Downloads/download.xlsx");
    }

    excelGetCoordinates();