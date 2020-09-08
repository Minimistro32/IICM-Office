function nukeContacts(){
  var groups = ContactsApp.getContactGroups();
  
  for (i = 0; i < groups.length; i++){
    try {
      groups[i].deleteGroup();
    } catch (err) {
      i += 1;
    }
  }
}

function generateContacts() {
  var sheetLastColumn = 12;
  var sheet = SpreadsheetApp.getActiveSheet();
  var contactsToPush = [];
  var colHeaders = sheet.getRange(1, 1, 1, sheetLastColumn).getValues();
  var iColStr = "1";
  var iColInt = 1;
  var colDict = {};
  
  colHeaders.forEach(function (header) {
    header.forEach(function (subHeader) {
      colDict[iColStr] = subHeader;
      iColInt++;
      iColStr = iColInt.toString();
    });
  });
  
  for (i = 2; i <= sheet.getLastRow(); i++) {
    iColStr = "1";
    iColInt = 1;
    var contactDict = {};
    sheet.getRange(i, 1, 1, sheetLastColumn).getValues().forEach(function (rawContactRow) {
      rawContactRow.forEach(function (rawContactData) {
        contactDict[colDict[iColStr]] = rawContactData;
        iColInt++;
        iColStr = iColInt.toString();
      });
    });
    contactsToPush.push(contactDict);
  }
  return contactsToPush;
}

function pushToContactsApp() {
  var contactsToPush = generateContacts();
  
  var newlyCreatedOrPreviouslyDeletedGroups = [];
  var systemMyContactsGroup = ContactsApp.getContactGroup("System Group: My Contacts");
  newlyCreatedOrPreviouslyDeletedGroups.push(systemMyContactsGroup);
  
  contactsToPush.forEach(function (contactData) {
    createContactFromData(contactData, newlyCreatedOrPreviouslyDeletedGroups, [systemMyContactsGroup]);
  });
}
 
function createContactFromData(contactData, freshGroups = [], customGroups = [], errIteration = 0) {
  var desiredGroups = []
  
  try{
    var newContact = ContactsApp.createContact(contactData["Name Prefix"], contactData["Given Name"], contactData["E-mail 1 - Value"]).setNotes(contactData["Notes"]);
    newContact.addAddress(ContactsApp.Field.HOME_ADDRESS, contactData["Address 1 - Street"] + ", " + contactData["Address 1 - City"] + ", " + contactData["Address 1 - Region"] + " " + contactData["Address 1 - Postal Code"] + " " + contactData["Address 1 - Country"]);
    newContact.addPhone(ContactsApp.Field.MOBILE_PHONE, contactData["Phone 1 - Value"]);
    if (contactData["Phone 2 - Value"] != "" && contactData["Phone 2 - Value"] !== null && contactData["Phone 2 - Value"] !== undefined) {
      newContact.addPhone(ContactsApp.Field.WORK_PHONE, contactData["Phone 2 - Value"]);
    }
    if (contactData["Phone 3 - Value"] != "" && contactData["Phone 3 - Value"] !== null && contactData["Phone 3 - Value"] !== undefined) {
      try {
        newContact.addPhone(ContactsApp.Field.HOME_PHONE, contactData["Phone 3 - Value"]);
      } catch(err) {
        throw err + "|" + contactData["Phone 3 - Value"];
      }
    }
    
    desiredGroups.push(contactData["Group Membership"].split(" ::: ").slice(1, contactData["Group Membership"].split(":::").length));
    desiredGroups.push(customGroups);
    
    desiredGroups.forEach(function (groupArray) {
      groupArray.forEach(function (group) {
        if (ContactsApp.getContactGroup(`${group}`)){
          //Logger.log(contactData["Given Name"] + " " + group + " it exists");
          if (!freshGroups.includes(group)) {
            //Logger.log(contactData["Given Name"] + " " + group + " it has never been deleted before");
            ContactsApp.getContactGroup(group).getContacts().forEach(function (contact) {
              contact.deleteContact()
            });
            ContactsApp.getContactGroup(group).deleteGroup();
            freshGroups.push(group);
            newContact.addToGroup(ContactsApp.createContactGroup(group));
          } else {
            //Logger.log(contactData["Given Name"] + " " + group + " it has been deleted before");
            newContact.addToGroup(ContactsApp.getContactGroup(group));
          }
        } else {
          //Logger.log(contactData["Given Name"] + " " + group + " it doesn't exist");
          newContact.addToGroup(ContactsApp.createContactGroup(group));
          freshGroups.push(group);
        }
      });
    });
    
  } catch(err) {
    Logger.log(errIteration);
    if (errIteration > 10){
      throw err;
    } else {
      if (typeof newContact !== 'undefined'){
        newContact.deleteContact();
      }
      Utilities.sleep(3000);
      createContactFromData(contactData, errIteration + 1);
    }
  }
}