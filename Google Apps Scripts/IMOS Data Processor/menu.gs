var ui = SpreadsheetApp.getUi();

function onOpen(e) {
  // Define menu configuration
  const customMenu = {
    'ui': ui,
    'title': 'IICM Data Processor',
    'menuItems': [
      ['Upload Key Indicators', 'importKIDialog'],
      {
        'title' : 'Contacts',
        'menuItems' : [
          ['Upload from File', 'importContactsDialog'],
          {
            'title' : 'Push to Google Contacts',
            'menuItems' : [
              ['Automatic (BROKEN)', 'pushToContactsApp'],
              ['Manually via Email', 'EmailRange'],
            ]
              }
            ],
          }
            
      /*,
      ['Arrivals', 'importArrivalsDialog'],
      ['Departures', 'importDeparturesDialog'],
      ['Handle KI', 'handleKIData']*/
    ]
  };
  
  // Create the menu with the defined configuration
  createCustomMenu(customMenu);
};

function importKIDialog() {
  // Creates an IFrame from the form.html file
  const html = HtmlService.createHtmlOutputFromFile('KI Form.html')
  .setWidth(450)
  .setHeight(300);
  
  // Shows the created IFrame as a popup dialog
  let dialog = ui.showModalDialog(html, 'Import Data');
};

function importContactsDialog() {
  // Creates an IFrame from the form.html file
  const html = HtmlService.createHtmlOutputFromFile('contactsForm.html')
  .setWidth(450)
  .setHeight(300);
  
  // Shows the created IFrame as a popup dialog
  let dialog = ui.showModalDialog(html, 'Import Data');
};

function importArrivalsDialog() {
  // Creates an IFrame from the form.html file
  const html = HtmlService.createHtmlOutputFromFile('unimplemented.html')
  .setWidth(450)
  .setHeight(300);
  
  // Shows the created IFrame as a popup dialog
  let dialog = ui.showModalDialog(html, 'Import Data');
};

function importDeparturesDialog() {
  // Creates an IFrame from the form.html file
  const html = HtmlService.createHtmlOutputFromFile('unimplemented.html')
  .setWidth(450)
  .setHeight(300);
  
  // Shows the created IFrame as a popup dialog
  let dialog = ui.showModalDialog(html, 'Import Data');
};

function importData(formData) {
  
  // Deconstruct the formData
  const {
    fileData,
    fileName
  } = formData;
  let {
    sourceSheetName,
    targetSheetName
  } = formData;
  
  // Upload the specified file and convert it to google sheets
  let sourceFile = uploadFileToDrive(fileData, fileName, 'Previous Imports', (sourceSheetName || fileName));
  
  // If no sourceSheetName specified, use the name of the first sheet of the sourceFile
  sourceSheetName = sourceSheetName || SpreadsheetApp.openById(sourceFile.id).getSheets()[0].getSheetName()
  
  // Copies the data from the uploaded Sheet to the specified Sheet on the current Spreadsheet
  let modifiedSheet = copyFullSheet(sourceFile.id, sourceSheetName, targetSheetName);
  
  //data Handlers
  switch (targetSheetName){
    case "Key Indicators":
      handleKIData(modifiedSheet);
      break;
    case "Contacts":
      handleContactData(modifiedSheet);
      break;
    case "Arrivals":
      break;
    case "Departures":
      break;
  }
  
  return;
};
