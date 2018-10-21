const Excel = require('exceljs'); 
const nanp = require('./nanp-script');

let input_filename = "./input.xlsx";



function readWorkSheet(worksheet){

 // let number_regions = {};
 let numbers_to_compare = [];

    worksheet.eachRow(function (row,rowNumber) {
      if (rowNumber !== 1) {
        var inp = ""+(row.getCell(1).value);
        numbers_to_compare.push(inp.slice(-10, inp.length));

    
      } else {
      }
    });

  return numbers_to_compare;

}

exports.readFile = function(){
   
////////////////////////
let workbook_nanp = new Excel.Workbook();
let workbook_input = new Excel.Workbook();

const promise1 = workbook_input.xlsx.readFile(input_filename)
  .then(function () {
    let worksheet_nos = workbook_input.getWorksheet();
    
    number_list = readWorkSheet(worksheet_nos);

    nanp.readFile(number_list);
    
  });


Promise.all([promise1]).then(function(result){ 
  
});

}

// readFile();

