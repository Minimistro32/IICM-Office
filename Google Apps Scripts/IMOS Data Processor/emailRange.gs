//MADE BY CHEMI ADEL
//EMAIL : pace3man@gmail.com
//VISIT US : www.sheetmania.com



function EmailRange() {
  RANGE="A1:L";
  SHEET_NAME="Contacts";
  
  //Types available : pdf,csv or xlsx
  EXPORT_TYPE="csv";
  
  //Assign The Spreadsheet,Sheet,Range to variables
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName(SHEET_NAME);
  var range=sheet.getRange(RANGE);
  
  //Range values to export
  var values=range.getValues();
  
  //Create temporary sheet
  var sheetName=Utilities.formatDate(new Date(), "GMT", "MM-dd-YYYY hh:mm:ss");
  var tempSheet=ss.insertSheet(sheetName);
  
  //Copy range onto that sheet
  tempSheet.getRange(1, 1, values.length, values[0].length).setValues(values);  
  
  //Save active sheets (Unhidden)
  var unhidden=[];
  for(var i in ss.getSheets()){
    if(ss.getSheets()[i].getName()==sheetName) continue;
    if(ss.getSheets()[i].isSheetHidden()) continue;
    unhidden.push(ss.getSheets()[i].getName());
    ss.getSheets()[i].hideSheet();
  }
  
  //Authentification 
  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  
  var url="https://docs.google.com/spreadsheets/d/"+ss.getId()+"/export?format="+EXPORT_TYPE;
  
  //Fetch URL of active spreadsheet
  var fetch=UrlFetchApp.fetch(url,params);
  
  //Get content as blob
  var blob=fetch.getBlob(); 
  
  var mimetype;
  if(EXPORT_TYPE=="pdf"){
    mimetype="application/pdf";      
  }else if(EXPORT_TYPE=="csv"){
    mimetype="text/csv";    
  }else if(EXPORT_TYPE=="xlsx"){
    mimetype="application/xlsx";   
  }else{
    return;
  }
  
  //Send Email
  GmailApp.sendEmail('iowa.iowacity@missionary.org',
                     'Google Contact Import', 
                     '1. Download the attached CSV file.\n2. Go to www.contacts.google.com\n3. Delete the old contacts.\n4. Click Import and select the attached CSV File.' ,
                     {
                       attachments: [{
                         fileName: "Attach title" + "."+EXPORT_TYPE,
                         content: blob.getBytes(),
                         mimeType: mimetype
                       }]
                     });
  
  //Reshow the sheets
  for(var i in unhidden){
    ss.getSheetByName(unhidden[i]).showSheet();
  }
  
  //Delete the temporary sheet
  ss.deleteSheet(tempSheet);
}