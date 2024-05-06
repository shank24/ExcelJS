const ExcelJs = require('exceljs');

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
