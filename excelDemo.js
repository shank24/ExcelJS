const ExcelJs = require('exceljs');

const workBook = new ExcelJs.workBook();
const workSheet = workBook.getWorksheet('Sheet1');