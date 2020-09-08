// Copies relevent information from the dataSet into the contactObject
function copyRawContactData(contactObject, dataSet, cols) {
  // Filter the dataSet based on the properties of the contactObject
  const filteredData = dataSet.filter((dataRow) => dataRow[cols['Area']] == contactObject['Given Name']);

  // Itterate through the people in the filteredData and add them as the missionaries in the contactObject
  filteredData.forEach((person, index) => {
    if (person[cols["Last Name"]] == '') return;
    contactObject["Missionary Count"] += 1;
    contactObject[`Missionary ${index + 1}`] = `${person[cols['Type']]};${person[cols['First Name']]};${person[cols['Last Name']]};${person[cols['Position']]}`;
  });

  // Obtain all other direct copy properties from the first dataRow
  contactObject['Phone 1 - Value'] = filteredData[0][cols['Phone1']];
  contactObject['Phone 2 - Value'] = filteredData[0][cols['Phone2']];
  contactObject['E-mail 1 - Value'] = filteredData[0][cols['Area Email']];
  contactObject['Address 1 - Street'] = filteredData[0][cols['Street']];
  contactObject['Address 1 - City'] = filteredData[0][cols['City']];
  contactObject['Address 1 - Region'] = filteredData[0][cols['State/Province']];
  contactObject['Address 1 - Postal Code'] = filteredData[0][cols['Postal Code']];
  contactObject['Zone'] = filteredData[0][cols['Zone']].replace(' Zone', '');

};


// Creates Notes and adds neccessary groups based on missionary and zone data
function createContactNotes(contactObject, zonePrefixes) {
  // Key the zonePrefixes to assign the contact Name Prefix
  contactObject['Name Prefix'] = zonePrefixes[contactObject['Zone']];

  // Add contacts to the correct Zone contact group
  contactObject['Group Membership'] += ` ::: ${contactObject['Zone']}`

  // Loop through the missionaries and append their Position and Last Name to the Notes
  for (let i = 0; i < contactObject['Missionary Count']; i++) {
    let infoArray = contactObject[`Missionary ${i + 1}`].split(';');
    contactObject['Notes'] += `${infoArray[3]} ${infoArray[2]}; `;
  }

  // If in the MLC, add them to the group
  if (['ZL', 'STL', 'AP'].some((substr) => {
      return contactObject['Notes'].includes(substr)
    })) {
    contactObject['Group Membership'] += ' ::: #MLC';
  }
};


// Takes an input dataArray[][] and formats the information into a contacArray[][]
function parseContactData(data, dataHeaders, zoneAbbreviations) {
  
  // Find the row that contains the data headers
  const headerRowIndex = data.findIndex((row) => {
    return row[0] == 'Missionary Name';
  })
  const headerRow = data[headerRowIndex];
  // Remove the headerRow and any preceeding rows from the data
  data.splice(0, (headerRowIndex + 1));

  // Extract the column indexes of all the required fields for easy refrence
  const cols = headerRow.reduce((cols, header, index) => {
    if (dataHeaders.includes(header)) {
      cols[header] = index
    };
    return cols;
  }, {});

  let contactObjects = [];
  // Extract all unique areaNames from data and then create a contactObject for each one
  data.map((dataRow) => dataRow[cols['Area']])
    .filter((value, index, dataColumn) => (dataColumn.indexOf(value) === index && value != ''))
    .forEach((areaName) => {
      // Initiallize all contactObjects with fields that correspond to the needed column headers
      contactObjects.push({
        'Given Name': areaName,
        'Notes': '',
        'Name Prefix': '',
        'Phone 1 - Value': '',
        'Phone 2 - Value': '',
        'E-mail 1 - Value': '',
        'Address 1 - Street': '',
        'Address 1 - City': '',
        'Address 1 - Region': '',
        'Address 1 - Postal Code': '',
        'Address 1 - Country': 'United States',
        'Group Membership': '* myContacts ::: @IICM',
        'Zone': '',
        'Missionary Count': 0,
        'Missionary 1': '',
        'Missionary 2': '',
        'Missionary 3': ''
      });
    });

  // Create the contactArray and push the header row
  let contactArray = [];
  contactArray.push(Object.keys(contactObjects[0]));
  // Take all areas' contactObjects and fill them with formatted data. Push the values to the contactArray
  contactObjects.forEach(contact => {
    copyRawContactData(contact, data, cols);
    createContactNotes(contact, zoneAbbreviations);
    contactArray.push(Object.values(contact));
  });

  return contactArray;
};


/** 
 * The Contact Data Handling Method
 * Built to work with a Sheet that has had the Organization Roster from IMOS copied to it
 * @param {GoogleAppsScript.Spreadsheet.Sheet} contactSheet
 */
function handleContactData(contactSheet) {

  /*
  These strings need to exactly match the header strings in the Organizational Roster.
  Order does not matter, but if the string changes due to an IMOS update, you will need to
  refactor all instances to the new string value.
  Almost all instances of these headers being used are in the copyRawContactData() method.
  The only exception to this is the 'Area' header, which also occurs in the parseContactData() method.
  */
  const dataHeaders = [
    'Last Name',
    'First Name',
    'Type',
    'Position',
    'Phone1',
    'Phone2',
    'Zone',
    'Area',
    'Area Email',
    'Street',
    'City',
    'State/Province',
    'Postal Code'
  ];

  /*
  ZoneAbbreviations can be altered or edited as needed.
  The only guideline is that the Key String for a Zone must match the string in the 'Zone' column of the 
  Organizational Roster (Minus the " Zone" at the end).
  */
  const zoneAbbreviations = {
    'Ames': 'AM',
    'Cedar Rapids': 'CR',
    'Davenport': 'DP',
    'Des Moines': 'DM',
    'Iowa City': 'IC',
    'Mt. Pisgah': 'MP',
    'Nauvoo': 'NA',
    'Peoria': 'PE'
  };


  let rosterData = contactSheet.getDataRange().getValues();
  contactSheet.clear();
  const contactData = parseContactData(rosterData, dataHeaders, zoneAbbreviations);
  let destinationRange = contactSheet.getRange(1, 1, contactData.length, contactData[0].length);
  destinationRange.setValues(contactData);

}
