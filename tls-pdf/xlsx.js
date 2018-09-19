var data = [
  {
    "CAMPAIGN_ID": 17,
    "STT": 1,
    "SHOP_ID": 96159,
    "DONVI": "Nhân viên",
    "SOLUONG": "14",
    "TITLE": "SL NV cần triển khai chăm sóc",
    "ID": 34254,
    "STT_SUB": 1,
    "RPT_TYPE": "Tonghop"
  },
  {
    "CAMPAIGN_ID": 17,
    "STT": 2,
    "SHOP_ID": 96159,
    "DONVI": "Nhân viên",
    "SOLUONG": "1",
    "TITLE": "SL NV báo cáo kết quả chăm sóc trên PM",
    "ID": 34255,
    "STT_SUB": 2,
    "RPT_TYPE": "Tonghop"
  },
  {
    "CAMPAIGN_ID": 17,
    "STT": 3,
    "SHOP_ID": 96159,
    "DONVI": "Khách",
    "SOLUONG": "6371",
    "TITLE": "SL KH cần chăm sóc",
    "ID": 34256,
    "STT_SUB": 2,
    "RPT_TYPE": "Tonghop"
  },
  {
    "CAMPAIGN_ID": 17,
    "STT": 4,
    "SHOP_ID": 96159,
    "DONVI": "Khách",
    "SOLUONG": "4",
    "TITLE": "SL KH siêu thị đã chăm sóc",
    "ID": 34257,
    "STT_SUB": 2,
    "RPT_TYPE": "Tonghop"
  },
  {
    "CAMPAIGN_ID": 17,
    "STT": 5,
    "SHOP_ID": 96159,
    "DONVI": "Khách",
    "SOLUONG": "0",
    "TITLE": "Tỉ lệ KH đã chăm sóc/tổng KH",
    "ID": 34258,
    "STT_SUB": 2,
    "RPT_TYPE": "Tonghop"
  }
]

function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

function writeWorkbook(workbook, name){
  /* bookType can be 'xlsx' or 'xlsm' or 'xlsb' or 'ods' */
  var wopts = { 
    bookType:'xlsx', 
    bookSST:true, 

    type:'binary',

  };

  var wbout = XLSX.write(workbook, wopts);

  /* the saveAs call downloads a file on the local machine */
  saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), name);

}



function readExcel(){
  var url = "./sample.xlsx";
  var oReq = new XMLHttpRequest();
  oReq.open("GET", url, true);
  oReq.responseType = "arraybuffer";

  oReq.onload = function(e) {
    var arraybuffer = oReq.response;

    /* convert data to binary string */
    var data = new Uint8Array(arraybuffer);
    var arr = new Array();
    for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
    var bstr = arr.join("");

    /* Call XLSX */
    var workbook = XLSX.read(bstr, {
      type:"binary",
      cellStyles: true,
      // bookDeps: true,

    });

    /* DO SOMETHING WITH workbook HERE */
    readWorkbook(workbook)
  }

  oReq.send();
}

function setCellValue(cell, value){
  if(cell){
    cell.w = undefined
    cell.v = value.toString()
  }
}

function addData(sheet, data, range){
  var sheetRange = XLS.utils.decode_range(sheet['!ref']) // {c:A,r:2}, {c:AE, r: 336}  A2:AE336    
  // move all value from start range to bottom by distance
  var distance = range.e.r - range.s.r;
  sheetRange.e.r += distance
  sheet['!ref'] = XLS.utils.encode_range(sheetRange)

  var merges = sheet['!merges']
  
  merges.forEach(function(merge){    
    if(merge.e.r > range.s.r){      
      merge.e.r += distance
      merge.s.r += distance
    }
  });
  

  for(var R = sheetRange.e.r; R>= range.s.r + distance; --R) {
    for(var C = sheetRange.s.c; C <= sheetRange.e.c; ++C) {
      var cellAddress = XLS.utils.encode_cell({c:C, r:R - distance});
      var shiftCellAddress = XLS.utils.encode_cell({c:C, r: R});
      if(sheet[cellAddress]){
        sheet[shiftCellAddress] = Object.assign({}, sheet[cellAddress])
        // delete sheet[cellAddress]
        // console.log(cellAddress)
      }
      
    }
  }
  // set current row to data at range
  // console.log(range)
  var i=0,j=0;
  for(var R = range.s.r; R <= range.e.r; ++R) {
    for(var C = sheetRange.s.c; C <= sheetRange.e.c; ++C) {
      var cellAddress = XLS.utils.encode_cell({c:C, r:R});
      // setCellValue(sheet[cellAddress], '')
      sheet[cellAddress] = {t: 's', v: ''}
    }
    var row = data[i]
    j = 0;
    for(var C = range.s.c; C <= range.e.c; ++C) {
      var cellAddress = XLS.utils.encode_cell({c:C, r:R});
      // setCellValue(sheet[cellAddress], row[j])   
      sheet[cellAddress] = {t: 's', v: row[j].toString()}      
      j++;      
          
    }
    i++;    
  }
}

function readWorkbook(workbook){
  var firstSheetName = workbook.SheetNames[0]
  var firstSheet = workbook.Sheets[firstSheetName] // BC nang suat các ST
  window.sheet = firstSheet
  window.workbook = workbook
  // start from column D
  var c = 'D'
  var r = 1
  while(true && r < 1000){
    var cell = firstSheet[c+r]
    if(cell && cell.v.trim() === '[DATA]')
      break;
    r++
  }

  var rangeData = data.map(function(item){
    return[item.STT, item.TITLE, item.DONVI, item.SOLUONG]    
  })

  var start = XLS.utils.decode_cell('D' + r)
  var end = {r: start.r + rangeData.length-1, c: start.c + rangeData[0].length-1}        

  addData(firstSheet, rangeData, {s: start, e: end})
  // firstSheet['!merges'][1].s.r+=4
  // firstSheet['!merges'][1].e.r+=4
  writeWorkbook(workbook, "export.xlsx")
}


readExcel()

$(function(){

  $('#test').click(function(){
    readExcel()
  })
})