const Excel = require('exceljs');
let input_filename = "./input.xlsx";

let nanp_filename = "./nanp.xlsx";
let number_regions_9xx = {};
let number_regions_8xx = {};
let number_regions_7xx = {};
let number_regions_6xx = {};
var regions = []
function readWorkSheet(worksheet){

 let number_regions = {};

    worksheet.eachRow(function (row,rowNumber) {
      if (rowNumber !== 1) {
         let num = ""+row.getCell(1).value;
        let reg = row.getCell(3).value;
        number_regions[num] = reg;  
      } else {
      }
    });
    
  return number_regions;

}

exports.readFile = function(numbers_to_compare){
 
let workbook_nanp = new Excel.Workbook();
let workbook_numbers = new Excel.Workbook();

const promise1 = workbook_nanp.xlsx.readFile(nanp_filename)
  .then(function () {
    
    let worksheet_9xx = workbook_nanp.getWorksheet("9xxx");
    let worksheet_8xx = workbook_nanp.getWorksheet("8xxx");
    let worksheet_7xx = workbook_nanp.getWorksheet("7xxx");
    let worksheet_6xx = workbook_nanp.getWorksheet("6xxx");


    number_regions_9xx = readWorkSheet(worksheet_9xx);
    number_regions_8xx = readWorkSheet(worksheet_8xx);
    number_regions_7xx = readWorkSheet(worksheet_7xx);
    number_regions_6xx = readWorkSheet(worksheet_6xx);

    // calling compareNumber function to populate regions corresponding to numbers in numbers_to_compare
    regions = compareNumber(numbers_to_compare,number_regions_9xx,number_regions_8xx,number_regions_7xx,number_regions_6xx);
 
   //write in file
   let workbook_input = new Excel.Workbook();
const promise2 = workbook_input.xlsx.readFile(input_filename)
  .then(function () {
    let worksheet = workbook_input.getWorksheet();
    var cell = worksheet.getCell('D1');
    var cell1 = worksheet.getCell('A1');
    cell.value = 'region'
    cell1.value = 'phoneNo'



    //file write with regions
    var i = -1
     worksheet.eachRow(function (row,rowNumber) {
      if (rowNumber !== 1) {
      
        i=i+1
        let num = row.getCell(4);
        let num1 =row.getCell(1);
        num1.value = numbers_to_compare[i]
        num.value = regions[i]
 
        
      } else {
      }
    });

     //each row ends here
     return workbook_input.xlsx.writeFile('output.xlsx');
  });
  //promise2 readfile ends here

  });

   //return promise1;
  Promise.all([promise1]).then(function(result){ 
     
});
//readMaria();
}

compareNumber = function(numbers_to_compare,number_regions_9xx,number_regions_8xx,number_regions_7xx,number_regions_6xx){
  let regions = [];
  for (var i = 0 ; i<numbers_to_compare.length;i++){

     let number = ""+numbers_to_compare[i];
     let region = "empty";

     if(number.startsWith('9')){
         region = compareWithRegex(number,number_regions_9xx);
     } else if(number.startsWith('8')){
         region = compareWithRegex(number,number_regions_8xx);
 
     } else if(number.startsWith('7')){
         region = compareWithRegex(number,number_regions_7xx);
     } else if(number.startsWith('6')){
         region = compareWithRegex(number,number_regions_6xx);
     }
     regions.push(region);
  }
  return regions;
}

function compareWithRegex(number, numbers){
   return numbers[number.slice(0,4)];
}
