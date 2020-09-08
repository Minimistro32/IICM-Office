function copyFullSheet(sourceFile, sourceSheetName, targetSheetName) {

  // Tries to open the file with SpreadsheetApp 
  try {
    var source = SpreadsheetApp.openById(sourceFile);
  } catch(err) {
    throw err;
  };

  var destination = SpreadsheetApp.getActiveSpreadsheet();
  
  // Source Sheet
  var sourceSheet = source.getSheetByName(sourceSheetName);

  // Collects all the data and sets it onto the "targetSheet"
  var newTargetSheet = sourceSheet.copyTo(destination);
  var targetSheet = destination.getSheetByName(targetSheetName);
  if (targetSheet != null) {
    destination.deleteSheet(targetSheet);
  }
  newTargetSheet.setName(targetSheetName);
  
  // returns the targetSheet for further editing
  return newTargetSheet;
};

//copy src sheet to tgt spreadSheet
//del old sheet
//rename new sheet -> old sheet name