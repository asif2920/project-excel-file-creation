var excel = require('excel4node');

// Create a new instance of a Workbook class
var workbook = new excel.Workbook();

// Add Worksheets to the workbook
var worksheet = workbook.addWorksheet('Sheet 1');
var worksheet2 = workbook.addWorksheet('Sheet 2');


// Create a reusable style
var style = workbook.createStyle({
  font: {
    color: '#FF0800',
    size: 12
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -'
});

let data=[{'s.no':'1','Name':'xxx','Age':'22'},
{'s.no':'2','Name':'yyy','Age':'12'},
    {'s.no':'3','Name':'zzz','Age':'32'}]
    let tempArr = data[0]
    let i=1;
        //read key
        for (var key in tempArr) {
          console.log(key);
          worksheet.cell(1,i).string(key).style(style);
          i++;
          //console.log(tempArr[key]);
      }
      let rowNumber=2;
      data.forEach(async (value,idx)=>{
        
        let tempDataArr;
        tempDataArr = value
        writeToExcel(tempDataArr,rowNumber);
        rowNumber++;
      })
      

      function writeToExcel(tempDataArr,rowNumber){
        console.log("Row number: ",rowNumber,tempDataArr)
      let columnNumber=1;
      for (var key in tempDataArr) {
        worksheet.cell(rowNumber,columnNumber).string(tempDataArr[key]).style(style);
        columnNumber++;
    }
     }
workbook.write('Excel.xlsx');