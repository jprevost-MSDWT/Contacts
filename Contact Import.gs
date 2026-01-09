/*
Project Name: Contact Import
Project Version: 6.01
Filename: Contact import.gs
File Version: 4.01
Chat link: [Insert Link]
*/

/**
 * The main script file for handling contact imports.
 * Note: CONFIG and onOpen are located in Initialization.gs
 */

/**
 * Fetches all Google Contacts and imports them into a structured,
 * import-friendly format. Names, orgs, phones, and emails are split
 * into separate columns.
 */
function importContactsToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.Import_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.Import_SHEET_NAME);
  }
  sheet.clear();

  // Create a lookup map for contact group (label) names.
  let groupMap = {};
  try {
    const groups = People.ContactGroups.list().contactGroups;
    if (groups && groups.length > 0) {
      groups.forEach(group => {
        groupMap[group.resourceName] = group.formattedName;
      });
    }
  } catch (e) {
    console.log("Could not fetch contact groups (labels): " + e.message);
  }

  // Get ALL contacts using the People API, handling pagination.
  let allConnections = [];
  let pageToken;
  try {
    do {
      const response = People.People.Connections.list('people/me', {
        personFields: 'names,emailAddresses,phoneNumbers,organizations,memberships',
        pageSize: 100,
        pageToken: pageToken
      });
      if (response.connections && response.connections.length > 0) {
        allConnections = allConnections.concat(response.connections);
      }
      pageToken = response.nextPageToken;
    } while (pageToken);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Failed to get contacts. Please ensure the "People API" is enabled in Services. Error: ' + e.message);
    return;
  }

  if (allConnections.length === 0) {
      SpreadsheetApp.getUi().alert('No contacts found to import.');
      return;
  }

  // --- Dynamic Header and Row Creation ---

  // Determine the maximum number of phone numbers and emails for any contact.
  let maxPhones = 0;
  let maxEmails = 0;
  allConnections.forEach(person => {
    if (person.phoneNumbers && person.phoneNumbers.length > maxPhones) {
      maxPhones = person.phoneNumbers.length;
    }
    if (person.emailAddresses && person.emailAddresses.length > maxEmails) {
      maxEmails = person.emailAddresses.length;
    }
  });

  // Build the dynamic header row.
  const headers = [
    "Resource ID", // Unique ID for Sync Back
    "First Name",
    "Last Name",
    "Labels",
    "Organization 1 - Company", 
    "Organization 1 - Title", 
    "Organization 1 - Department"
  ];
  for (let i = 1; i <= maxEmails; i++) {
    headers.push(`Email ${i} - Type`, `Email ${i} - Value`);
  }
  for (let i = 1; i <= maxPhones; i++) {
    headers.push(`Phone ${i} - Type`, `Phone ${i} - Value`);
  }

  // Process each contact into the new structured format.
  const contactData = allConnections.map(person => {
    const row = new Array(headers.length).fill(''); // Pre-fill row with empty strings

    // --- Static Fields ---
    row[0] = person.resourceName || ''; // Resource ID

    if (person.names && person.names.length > 0) {
      row[1] = person.names[0].givenName || '';  // First Name
      row[2] = person.names[0].familyName || ''; // Last Name
    }
    
    if (person.memberships && person.memberships.length > 0) {
      const labelNames = person.memberships
        .map(m => m.contactGroupMembership && groupMap[m.contactGroupMembership.contactGroupResourceName])
        .filter(name => name);
      row[3] = labelNames.join(', '); // Labels
    }
    
    if (person.organizations && person.organizations.length > 0) {
      const org = person.organizations[0];
      row[4] = org.name || '';         // Company
      row[5] = org.title || '';        // Title
      row[6] = org.department || '';   // Department
    }

    // --- Dynamic Fields ---
    // Start index is 7 (Resource ID + First + Last + Labels + 3 Org fields)
    const emailStartIndex = 7; 
    if (person.emailAddresses) {
      person.emailAddresses.forEach((email, i) => {
        const colIndex = emailStartIndex + (i * 2);
        row[colIndex] = email.formattedType || 'Other';
        row[colIndex + 1] = email.value || '';
      });
    }

    const phoneStartIndex = emailStartIndex + (maxEmails * 2);
    if (person.phoneNumbers) {
      person.phoneNumbers.forEach((phone, i) => {
        const colIndex = phoneStartIndex + (i * 2);
        row[colIndex] = phone.formattedType || 'Other';
        row[colIndex + 1] = phone.value || '';
      });
    }
    
    return row;
  }).filter(row => row[1] || row[2]); // Filter out contacts with no first or last name (Indices shifted by 1).

  // --- Write Data to Sheet ---
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (contactData.length > 0) {
    sheet.getRange(2, 1, contactData.length, headers.length).setValues(contactData);
  }

  // Create the list of labels and apply formatting.
  createLabelsdata(headers, contactData);
  applySheetFormatting(sheet);

  SpreadsheetApp.getUi().alert('Contact import complete! ' + contactData.length + ' contacts were imported.');
}


/**
 * Finds a column with a specific header, or creates it in the first empty column.
 * Populates this column with a unique list of contact labels and creates a named range.
 * @param {Array<string>} headers The header row of the data.
 * @param {Array<Array<string>>} contactData The 2D array of contact information.
 */
function createLabelsdata(headers, contactData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET_NAME);

  if (!dataSheet) {
    dataSheet = ss.insertSheet(CONFIG.DATA_SHEET_NAME);
  }
  
  let targetColumn;
  if (dataSheet.getLastColumn() > 0) {
    const dataHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
    targetColumn = dataHeaders.indexOf(CONFIG.Imported_LABELS_HEADER) + 1;
  } else {
    targetColumn = 0;
  }

  if (targetColumn === 0) {
    targetColumn = dataSheet.getLastColumn() + 1;
    dataSheet.getRange(1, targetColumn).setValue(CONFIG.Imported_LABELS_HEADER);
  }

  if (dataSheet.getLastRow() > 1) {
    dataSheet.getRange(2, targetColumn, dataSheet.getLastRow() - 1, 1).clearContent();
  }
  
  const labelsColumnIndex = headers.indexOf("Labels");
  if (labelsColumnIndex === -1) {
      return; // "Labels" column not found
  }

  const allLabelStrings = contactData.map(row => row[labelsColumnIndex]).filter(labelCell => labelCell);
  if (allLabelStrings.length === 0) {
    return; // No labels to process
  }

  const uniqueLabels = [...new Set(allLabelStrings.flatMap(s => s.split(', ')))].sort();
  
  if (uniqueLabels.length > 0) {
    const formattedLabels = uniqueLabels.map(label => [label]);
    const range = dataSheet.getRange(2, targetColumn, formattedLabels.length, 1);
    range.setValues(formattedLabels);
    
    const existingNamedRange = ss.getRangeByName(CONFIG.Imported_NAMED_RANGE);
    if (existingNamedRange) {
      ss.removeNamedRange(CONFIG.Imported_NAMED_RANGE);
    }
    ss.setNamedRange(CONFIG.Imported_NAMED_RANGE, range);
  }
}

/**
 * Applies formatting to the contact sheet.
 * Sets the header to bold and creates a filter for the data.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to apply formatting to.
 */
function applySheetFormatting(sheet) {
  if (!sheet || sheet.getLastRow() === 0) return;
  const dataRange = sheet.getDataRange();
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  dataRange.createFilter();
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setFontWeight("bold");
}
