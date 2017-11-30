var Excel = require('exceljs');
var Papa = require('./papaparse.min.js');
var fs = require('fs');
try{
    var workbook = new Excel.Workbook();
    workbook.csv.readFile('./normal.csv')
        .then(function(workbookdata) {
            console.log(workbookdata);
            // var sheet = workbook.getWorksheet("Sheet1")
            // var dobIndex= sheet.getRow(1).values.indexOf("YOB")
            // var sum = 0;
            // for(var i = 2;i<=sheet.rowCount;i++ ){
            //     sum +=sheet.getRow(i).values[dobIndex];
            // }
            // console.log(sheet,sum);
        });
    Papa.SCRIPT_PATH	= "./papaparse.min.js";
}catch(ex){
    console.error(ex);
}

// Papa.parse("http://127.0.0.1:8074/normal.csv", {
//     download: true,
//     worker:false,
//     error: function(err, file, inputElem, reason)
// 	{
// 		// executed if an error occurs while loading the file,
// 		// or if before callback aborted for some reason
// 	},
// 	complete: function(fata)
// 	{
//         console.log(arguments);
// 		// executed after all files are complete
//     },
//     step: function(results, parser) {
//         //console.log("Row data:", results.data);
//         //console.log("Row errors:", results.errors);
//     }
// 	// rest of config ...
// })
// var workbook = new Excel.Workbook();
// var fileStreem = fs.createReadStream('./big_co.csv')
// workbook.csv.read(fileStreem)
//     .then(function(worksheet) {
//         console.log(worksheet);
//         // use workbook or worksheet
//     });

// var workbook = new Excel.Workbook();
// fileStreem.on("end",function(){
//     console.log(workbook.getWorksheet("INELIGIBLE_STATE"))
// })
// workbook.csv.read(fileStreem).then((workbook)=> {
//     console.log(workbook)
// })
    
//fileStreem.pipe(workbook.xlsx.createInputStream());

// var options = {
//     filename: './streamed-workbook.xlsx',
//     useStyles: true,
//     useSharedStrings: true
// };
// var workbook = new Excel.stream.xlsx.WorkbookWriter(options);