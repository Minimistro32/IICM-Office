//KI stands for Key Indicator

//data to get unique from; rowNum and columnNum indexed at 1
function getUniqueDateFromColumn(data, rowNum, columnNum){
  var col = columnNum - 1 ; // choose the column you want to use as data source (0 indexed, it works at array level)
  //var data = sheet.getDataRange().getValues();// get all data

  var newdata = new Array();
  for(var row = rowNum - 1; row < data.length; row++){
    var duplicate = false;
    for(j in newdata){
      //Logger.log(data[row][col]);
      //Logger.log(j);
      if(Date.parse(data[row][col]) == Date.parse(newdata[j])){
        duplicate = true;
      }
    }
    if(!duplicate){
      newdata.push(data[row][col]);
    }
  }
  return newdata;
}

function getUniqueFromTable(data, rowNum, columnNum){
  var col = columnNum - 1 ; // choose the column you want to use as data source (0 indexed, it works at array level)

  var newdata = new Array();
  for(var row = rowNum - 1; row < data.length; row++){
    var duplicate = false;
    for(j in newdata){
      if(data[row][col] == newdata[j][col] || data[row][col] === ""){
        duplicate = true;
      }
    }
    if(!duplicate){
      newdata.push([...data[row]]);
    }
  }

  /*newdata.sort(function(x,y){
    var xp = Number(x[0]);// ensure you get numbers
    var yp = Number(y[0]);
    return xp == yp ? 0 : xp < yp ? -1 : 1;// sort on numeric ascending
  });

  sh.getRange(1,5,newdata.length,newdata[0].length).setValues(newdata);// paste new values sorted in column of your choice (here column 5, indexed from 1, we are on a sheet))
*/Logger.log(newdata);
  return newdata;
}

function findCellByValue(sheet, cellValue){ //returns (row, column) in a tuple
  let dataRange = sheet.getDataRange();
  let rangeValues = dataRange.getValues();
  
  for (i = 0; i < dataRange.getLastColumn() - 1; i++){
    for (j = 0; j < dataRange.getLastRow() - 1; j++){
      if (rangeValues[j][i] == cellValue){
        return [j, i];
      }
    }
  }
  throw `Cell with value: ${cellValue} not found`;
}

function handleKIData(kiSheet) {
  kiSheet = kiSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Key Indicators");
  
  kiSheet.getRange(1,1,kiSheet.getMaxRows(),kiSheet.getMaxColumns()).breakApart();
  kiSheet.deleteRow(1);
  let dataRange = kiSheet.getDataRange();

  //sort with oldest date on top
  let datePosition = findCellByValue(kiSheet, "Date");
  let sortRange = kiSheet.getRange((datePosition[0] + 2), 1, (dataRange.getLastRow() - (datePosition[0] + 1)), dataRange.getLastColumn());
  sortRange.sort([{column: datePosition[1]+1, ascending: false}, 2]);


  //deleterows if no data or more than a transfer old
  let rangeValues1 = dataRange.getValues();

  for (j = rangeValues1.length - 1; j >= 0 ; j--){
    emptyCounter = 0;
    for (i = 0; i < 6; i++){
      if (rangeValues1[j][i] === ""){
        emptyCounter++;
      }
      if (emptyCounter > 2){
        kiSheet.deleteRow(j + 1);
        break;
      }
    }
  }
  
  //remove first row from rangeValues
  let rangeValues = kiSheet.getDataRange().getValues();
  rangeValues.splice(0,1);
  
    //sort and get table of unique areas
  let areaPosition = findCellByValue(kiSheet, "Area");
  let uniqueAreas = getUniqueFromTable(rangeValues, 1, areaPosition[1] + 1);
  
  dataRange = kiSheet.getDataRange();//Range(1,1,kiSheet.getMaxRows()-1,kiSheet.getMaxColumns()-1);
  rangeValues1 = dataRange.getValues();
  
  /*DELETE DATES THAT ARE OVER A TRANSFER OLD
  var transferAgoDate = new Date();
  transferAgoDate.setDate(transferAgoDate.getDate() - 42);
  
  let rowcount = dataRange.getLastRow();
  let rowsToDelete = []
  
  rangeValues1.forEach((row, index) => {
      if (row[datePosition[1]] < transferAgoDate){
         rowsToDelete.push(index + 1);
      }
    });  

  if (rowsToDelete.length > 0){
    kiSheet.deleteRows(rowsToDelete[0],rowsToDelete.length);
  }

  dataRange = kiSheet.getDataRange();//Range(1,1,kiSheet.getMaxRows()-1,kiSheet.getMaxColumns()-1);
  rangeValues1 = dataRange.getValues();

  */

  var maxRows = kiSheet.getMaxRows(); 
  var lastRow = kiSheet.getLastRow();
  kiSheet.deleteRows(lastRow+1, maxRows-lastRow);
  
  //FILL IN ZERO'S
  let kiPosition = findCellByValue(kiSheet, "New People");
  
  //Potentially Faster way to fill zero's. Doesn't work yet:
  //dataRange = kiSheet.getDataRange();
  
  /*var range = kiSheet.getRange(kiPosition[0] + 2,kiPosition[1] + 1,kiSheet.getMaxRows() - (kiPosition[0] + 2),4)
      .offset(0, 0, kiSheet.getDataRange()
          .getNumRows());
  range.setValues(range.getValues()
      .map(function (row) {
          return row.map(function (cell) {
              return cell === '' ? 0 : cell;
          });
      }));*/

  dataRange = kiSheet.getRange(kiPosition[0] + 2,kiPosition[1] + 1,dataRange.getLastRow() - (kiPosition[0] + 2),4);
  rangeValues = dataRange.getValues();
  
  for (i = 0; i < dataRange.getLastColumn() - 1; i++){
    for (j = 0; j < dataRange.getLastRow() -1; j++){
      if (rangeValues[j][i] === ""){
        kiSheet.getRange(j + dataRange.getRow(),i + dataRange.getColumn()).setValue(0);
      }
    }
  }//*/
  
  //ADD IN BAGELS
  rangeValues = kiSheet.getDataRange().getValues()
  rangeValues.splice(0,1);
  let uniqueDates = getUniqueDateFromColumn(rangeValues, 1, datePosition[1] + 1);
  Logger.log(uniqueDates);
  //set isReportedForThisUniqueDate to empty array
  for(x = 0; x < uniqueAreas.length; x++){
    uniqueAreas[x][datePosition[1]] = [];
  }
  
  //identify areas that have been reported
  let dateIndex = 0;
  for (j = 0; j < rangeValues.length; j++){
    if (dateIndex > uniqueDates.length){
      throw "chief that aint it";
    }
    for (i = 0; i < uniqueAreas.length; i++){
      if (rangeValues[j][areaPosition[1]] == uniqueAreas[i][areaPosition[1]] && !(dateIndex in uniqueAreas[i][datePosition[1]])){
        uniqueAreas[i][datePosition[1]].push(dateIndex);
      }
    }
    if(j + 1 < rangeValues.length && Date.parse(rangeValues[j][datePosition[1]]) != Date.parse(rangeValues[j + 1][datePosition[1]])){
      dateIndex += 1;
    }
  }
  
  let dateIndexArray = [];
  for (i = 0; i < uniqueDates.length; i++){
    dateIndexArray.push(i);
  }
  
  for (j = 0; j < uniqueAreas.length; j++){
    let invertedIndiciesArray = [];
    let uniqueAreasDateIndex = 0;
    for (i = 0; i < dateIndexArray.length; i++){
      if (uniqueAreas[j][datePosition[1]][uniqueAreasDateIndex] == dateIndexArray[i]){
        if (uniqueAreasDateIndex + 1 < uniqueAreas[j][datePosition[1]].length) {
          uniqueAreasDateIndex += 1;
        }
      } else {
        invertedIndiciesArray.push(i);
      }
    }
    uniqueAreas[j][datePosition[1]] = invertedIndiciesArray;
  }
  
  //insert bagels
  let bagels = [];
  for(x = 0; x < uniqueAreas.length; x++){
    //insert new row for every dateIndex not present in uniqueAreas[x][datePosition[1]] array
    for(i = 0; i < uniqueAreas[x][datePosition[1]].length; i++){
      rowCopy = [...uniqueAreas[x]];
      rowCopy[datePosition[1]] = uniqueDates[uniqueAreas[x][datePosition[1]][i]];
      rowCopy[kiPosition[1]] = 0;
      rowCopy[kiPosition[1] + 1] = 0;
      rowCopy[kiPosition[1] + 2] = 0;
      rowCopy[kiPosition[1] + 3] = 0;
      bagels.push(rowCopy);
    } 
  }

  Logger.log(bagels);
  kiSheet.insertRowsAfter(1,bagels.length).getRange(2,1,bagels.length,bagels[0].length).setValues(bagels);
}