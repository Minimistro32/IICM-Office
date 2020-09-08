function _exportBlob(blob, fileName){
  blob = blob.setName(fileName)
  var today = new Date();
  var date = today.getFullYear() + '.' + (today.getMonth() + 1) + '.' + today.getDate();
  var subDir = "IICM Directory - " + date;
  

  try{ //and search for the parent folder "IICM Directory History"
    var folder1, folders1 = DriveApp.getFoldersByName("IICM Directory History");
    if (folders1.hasNext()) {
      //create and input file
      folder1 = folders1.next();
      folder1.createFile(blob);
    } else {
      // If no directory is found create it and the file
      DriveApp.createFolder("IICM Directory History").createFile(blob);
    }
  } catch (err1) {
    throw err1;
  }
}



function _getAsBlob(url, sheet, range) {
  var rangeParam = ''
  var sheetParam = ''

  if (sheet) {
    sheetParam = '&gid=' + sheet.getSheetId()
  }
  var exportUrl = url.replace(/\/edit.*$/, '')
      + '/export?exportFormat=pdf&format=pdf'
      + '&size=LETTER'
      + '&portrait=true'
      + '&fitw=true'       
      + '&top_margin=0.3'              
      + '&bottom_margin=0.65'          
      + '&left_margin=0.2'             
      + '&right_margin=0.2'           
      + '&sheetnames=false&printtitle=false'
      + '&pagenum=RIGHT'
      + '&gridlines=false'
      + '&fzr=FALSE'      
      + '&printnotes=false'
      + sheetParam
      + rangeParam
      
  //Logger.log('exportUrl=' + exportUrl)
  var response = UrlFetchApp.fetch(exportUrl, {
    headers: { 
      Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
    },
    //muteHttpExceptions: true
  })
  
  //Logger.log(response);
  
  return response.getBlob()
}

function printDirectory() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = spreadsheet.getSheetByName('Sheet1');
  
  var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet);
  var today = new Date();
  var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
  _exportBlob(blob, "IICM Directory " + date + ".pdf");
}