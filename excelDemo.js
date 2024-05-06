const ExcelJs = require('exceljs');

//Way 1 via Promise
//Workbook
const workBook = new ExcelJs.Workbook();
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
})


//Way 2 Via async and await

async function excelTravserse() {
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

