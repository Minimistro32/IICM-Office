/**
 * Pulls data from the 'Contacts' sheet using ssID.
 * Formats that information into an array that spills
 * into a 6x2 grid.
 *
 * @return An array that spills
 * into a 6x2 grid
 * @customfunction
 */
const cardinals = {
  " North ":" N ",
  " South ":" S ",
  " East ":" E ",
  " West ":" W "
}

function spillArrayByArea(areaName, dataValues) {
  //areaName = "Macomb";
  
  var areaRowNum = -1;
  for (i = 1; i < dataValues.length; i++){
    if (dataValues[i][0] == areaName){
      areaRowNum = i;
      break;
    }
  }
  
  //title, first, last, pos
  var missionary1 = dataValues[i][14].split(";");
  var missionary2 = dataValues[i][15].split(";");
  var missionary3 = dataValues[i][16].split(";");

  //[20-09-01 10:52:45:162 CDT] Sister;Ashlyn;Albert;JC
  //[20-09-01 10:52:45:164 CDT] Sister;Audry;Ricks;TR
  //[20-09-01 10:52:45:166 CDT] 
  var regex = new RegExp(' (South|North|East|West) ', 'gm')
  var streetAddress = dataValues[i][6].replace(regex, cardinals[dataValues[i][6].match(regex)]);
  
  //var activeCell = SpreadsheetApp.getActiveSheet().getActiveCell();
  //var col = activeCell.getColumn();
  //var row = activeCell.getRow();
  
  //SpreadsheetApp.getActiveSheet().getRange(col, row + 1, 3).setBackground("Pink");
  
  if(missionary1[0] == "Sister"){
    if(missionary3[3]){
      return [[areaName],
              [" " + missionary1[3] + " ", missionary1[2]],
              [" " + missionary2[3] + " ", missionary2[2]],
              [" " + missionary3[3] + " ", missionary3[2]],
              [streetAddress],
              [dataValues[i][7] + ", " + dataValues[i][8] + " " + dataValues[i][9]],
              [dataValues[i][3]]]
    } else if (missionary2[3]) {
      return [[areaName],
              [" " + missionary1[3] + " ", missionary1[2]],
              [" " + missionary2[3] + " ", missionary2[2]],
              [missionary3[3], missionary3[2]],
              [streetAddress],
              [dataValues[i][7] + ", " + dataValues[i][8] + " " + dataValues[i][9]],
              [dataValues[i][3]]]
    } else {
      return [[areaName],
              [" " + missionary1[3] + " ", missionary1[2]],
              [missionary2[3], missionary2[2]],
              [missionary3[3], missionary3[2]],
              [streetAddress],
              [dataValues[i][7] + ", " + dataValues[i][8] + " " + dataValues[i][9]],
              [dataValues[i][3]]]
    }
  } else {
    return [[areaName],
            [missionary1[3], missionary1[2]],
            [missionary2[3], missionary2[2]],
            [missionary3[3], missionary3[2]],
            [streetAddress],
            [dataValues[i][7] + ", " + dataValues[i][8] + " " + dataValues[i][9]],
            [dataValues[i][3]]]
  }
}

function updateAreas(){
  var SpreadSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1kxPf4UHMcRerizZzfL8mTPSKpdGTJRWtz7_Z2hDQmcQ/edit#gid=615767414");
  var Sheet = SpreadSheet.getSheetByName("Contacts");
  var dataRange = Sheet.getDataRange();
  var rangeValues = dataRange.getValues();
  rangeValues.slice(0,1);
  SpreadsheetApp.getActiveSheet().getRange("A2:A").breakApart();
  SpreadsheetApp.getActiveSheet().unhideRow(SpreadsheetApp.getActiveSheet().getRange("A1:A"));
  
  rangeValues.sort((a, b) => {
    var zoneColumn = 12;
    var areaColumn = 0;
    if (a[zoneColumn] === b[zoneColumn]) {
      if (a[areaColumn] === b[areaColumn]) {
        return 0;
      } else {
        return (a[areaColumn] < b[areaColumn]) ? -1 : 1;
      }
    } else {
        return (a[zoneColumn] < b[zoneColumn]) ? -1 : 1;
    }
  });
  
  var areaNames = [];
  for (i = 0; i < rangeValues.length; i++){
    areaNames.push([rangeValues[i][12],rangeValues[i][0]]);
  }

  var updateGrid = [];
  var areaIndex = 0;
  var correctPositionToStartAdding = false;
  var startZoneRow = 2;
  var endZoneRow = 1;
  //loop through 20 rows of areas
  for(j = 0; j < 21; j++){
    //on last iteration simply merge the zone cell
    if (j != 20){
      //if there is no more areas push 7 empty rows
      if (areaIndex + 1 >= areaNames.length){
        for(numEmptyRowCounter = 0; numEmptyRowCounter < 7; numEmptyRowCounter++){
          updateGrid.push(["","","","","","","","","","","","","",""]);
        }
      } else {
        //Push new areas onto first row
        updateGrid.push([]);
        for(i = 0; i < 7; i++){
          var endOfZone = false;
          //determine if we've reached the end of a zone
          if(areaIndex != 0 && areaNames[areaIndex - 1][0] != areaNames[areaIndex][0]){
            endOfZone = true;
          }
          //if it's "not the end of a zone" or "this is actually the right place to resume adding areas", go ahead and add areas 
          if (!endOfZone || correctPositionToStartAdding) {
            updateGrid[j*7].push("=spillArrayByArea(\"" + areaNames[areaIndex][1] + "\",IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1kxPf4UHMcRerizZzfL8mTPSKpdGTJRWtz7_Z2hDQmcQ/edit#gid=615767414\",\"'Contacts'!A1:Q\"))","");
            if (endOfZone){
              //setValue to zone name in column 'A'. Merge the cells.
              SpreadsheetApp.getActiveSheet().getRange(startZoneRow, 1, (endZoneRow - startZoneRow) + 1).mergeVertically().setValue(areaNames[areaIndex-1][0]);
              startZoneRow = j*7+2
            }
            areaIndex += 1;
            correctPositionToStartAdding = false;
          } else {
            updateGrid[j*7].push("","");
          }
        }
        for(numEmptyRowCounter = 0; numEmptyRowCounter < 6; numEmptyRowCounter++){
          updateGrid.push(["","","","","","","","","","","","","",""]);
        }
        correctPositionToStartAdding = true;
        endZoneRow = ((j+1)*7)+1;
      }
    } else {
      SpreadsheetApp.getActiveSheet().getRange(startZoneRow, 1, (endZoneRow - startZoneRow) + 1).mergeVertically().setValue(areaNames[areaIndex-1][0]);
      Logger.log(endZoneRow, SpreadsheetApp.getActiveSheet().getLastRow());
      SpreadsheetApp.getActiveSheet().hideRows(endZoneRow + 1, SpreadsheetApp.getActiveSheet().getMaxRows() - endZoneRow);
    }
  }

  SpreadsheetApp.getActiveSheet().getRange(2, 2, 140, 14).setFormulas(updateGrid);

}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('IICM Directory')
      .addItem('Update Areas', 'updateAreas')
      .addItem('Print Directory', 'printDirectory')
      /*.addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('Second item', 'menuItem2'))*/
      .addToUi();
}
