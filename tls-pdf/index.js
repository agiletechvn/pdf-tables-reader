const Excel = require('exceljs')

var workbook = new Excel.Workbook();
workbook.xlsx.readFile('./sample.xlsx').then(function() {
    // use workbook
    var worksheet = workbook.getWorksheet('BC nang suat các ST')
    worksheet.addRow([3, 'Sam', new Date()])
    workbook.xlsx.writeFile('./test.xlsx').then(function() {
        // done
    });
});

