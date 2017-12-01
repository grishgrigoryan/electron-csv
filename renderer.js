
var Excel = require('exceljs');
var Papa = require('./papaparse.min.js');
var FS = require('fs');
var temp = require('temp');
var PATH = require('path')
window.PATH =PATH;
document.getElementById("input").addEventListener('change', function (e) {
    var file;
    if (e.target.files) {
        file = e.target.files[0];
        console.log(file);
        readFile(file.path);
    }
}, false);
var readFile = function(path){
        var fileStreem =FS.createReadStream(path)
        var workbook = new Excel.Workbook();
        var method = null;
        if(PATH.extname(path)==".csv") {
            method = "csv"
        }
        if(PATH.extname(path)==".xlsx") {
            method = "xlsx"
        }
        if(!method){
            throw new Error("File not suport");
        }
        // var workbook = new Excel.stream.xlsx.WorkbookReader();
        // var sharedString = [];
        // workbook.on('error', function (error) {
        //         console.log('An error occurred while writing reading excel', error);        
        // });

        // workbook.on('entry', function (entry) {
        //     //console.log("entry", entry);
        // });

        // workbook.on('shared-string', function (data) {
        //     sharedString[data.index] = data.text;
        //     // console.log("index:", index, "text:", text);
        // });

        // workbook.on('worksheet', function (worksheet) {
        //         console.log("worksheet", worksheet.name);           
        //         //if(worksheet.name=="Sheet1"){
        //             worksheet.on('row', function (row) {
        //                 row = row.values.map((r)=>{
        //                     if( r && typeof r === 'object' ){
        //                         console.log("Should take value from sharedString",sharedString)
        //                         return sharedString[r.sharedString];
        //                     }else{
        //                         return r;
        //                     }
        //                 })
        //                 console.log("Sheet1 row.values", row);
        //                 //console.log("Sheet1 row.model", row.model);
        //             });
        //         //}
                

        //         worksheet.on('close', function () {
        //             console.log("worksheet close");         
        //         });

        //         // worksheet.on('finished', function () {
        //         //     console.log("worksheet finished");          
        //         // });
        // });

        // workbook.on('finished', function () {
        //         console.log("finished",workbook.sharedString);
        // });

        // workbook.on('close', function () {
        //         console.log("close",workbook.sharedString);
        // });
        // workbook.read(fileStreem,
        //      {  entries: "emit",
        //         sharedStrings: "emit",
        //         styles: "emit",
        //         hyperlinks: "emit",
        //         worksheets: "emit"
        //     }
        // );
        // console.log(workbook);
        workbook[method].read(fileStreem)
            .then(function(workbookdata) {
                var sheet = workbook.getWorksheet("Sheet1")
                var dobIndex= sheet.getRow(1).values.indexOf("YOB")
                var sum = 0;
                for(var i = 2;i<=sheet.rowCount;i++ ){
                    sum +=sheet.getRow(i).values[dobIndex];
                }
                console.log(sheet,sum);
                createFile([{key:"Id"},{key:"SUM"}],[[1,sum]])
            });  
}
var createFile = function(columns,rows){
    let absPath =PATH.join(__dirname,'calculated.xlsx');
    //var writeStream = FS.createWriteStream(absPath);
    var workbookWriter = new Excel.stream.xlsx.WorkbookWriter({
        filename : absPath
    });
    console.log(workbookWriter);
    workbookWriter.creator = 'Application';
    workbookWriter.lastModifiedBy = 'Application';
    workbookWriter.created =  new Date();
    workbookWriter.modified = new Date();
    //var sheet = workbook.addWorksheet('my_sheet');
    var sheet = workbookWriter.addWorksheet('sheet');
    
    sheet.columns = columns.map((column)=>{
        return { header: column.header, key: column.key, width: column.width ||10 };
    })

    sheet.columns = columns.map((column)=>{
        return { header: column.key, key: column.key, width: column.width ||10 };
    });
    rows.map((row)=>{
        sheet.addRow(row).commit();    
    })
    //sheet.addRows(rows).commit();    
    sheet.commit();
    workbookWriter.commit()
        .then(function() {
            let fileDownloadLink = document.getElementById("file-download");
            fileDownloadLink.innerHTML = "Click to download";
            fileDownloadLink.setAttribute('href',absPath);
    });
    // workbook.xlsx.writeFile(absPath)
    //     .then(function() {
    //         let fileDownloadLink = document.getElementById("file-download");
    //         fileDownloadLink.innerHTML = "Click to download";
    //         fileDownloadLink.setAttribute('href',absPath);
    // })
    
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