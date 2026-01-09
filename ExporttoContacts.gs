/*
Project Name: Contact Import
Project Version: 7.01
Filename: ExporttoContacts.gs
File Version: 3.01
Chat link: [Insert Link]
*/

/**
 * Syncs the structured data from the Setup Sheet back to Google Contacts.
 * - Uses BATCH OPERATIONS to prevent Quota limits and Timeouts.
 * - Updates existing contacts in chunks of 50.
 * - Manages Labels (Creates new ones if missing).
 * - Deletes contacts if the "Delete from contacts" label is applied.
 * - NOTE: Contact Creation is currently DISABLED.
 */
function syncSheetsToContacts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.Setup_SHEET_NAME);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${CONFIG.Setup_SHEET_NAME}" not found.`);
    return;
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0];
  
  // 1. Map Headers to Column Indices
  const colMap = mapHeaders(headers);
  const missingCols = [];
  if (colMap.id === -1) missingCols.push("Resource ID");
  if (colMap.labels === -1) missingCols.push("Labels");
  if (missingCols.length > 0) {
    SpreadsheetApp.getUi().alert(`Error: Missing columns: ${missingCols.join(", ")}`);
    return;
  }

  // 2. Fetch Existing Contact Groups (Labels) to memory
  let groupMap = fetchAllContactGroups();

  // 3. Stats & Queues
  let stats = { updated: 0, created: 0, deleted: 0, labelsCreated: 0, errors: [], skipped: 0 };
  let updateQueue = []; // Array of row objects to process in batch
  const BATCH_SIZE = 50; // API limit is often 50 or 100

  // 4. Iterate Rows (Skip Header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Aggressive sanitization of Resource ID
    let sheetResourceName = row[colMap.id] ? String(row[colMap.id]).trim().replace(/[\s\u200B-\u200D\uFEFF]/g, '') : "";
    const labelString = row[colMap.labels] ? String(row[colMap.labels]) : "";
    
    // --- A. HANDLE DELETE (Single Operation) ---
    // Deletes are rarer, so we process them individually to keep logic simple
    if (labelString.includes("Delete from contacts")) {
      if (sheetResourceName) {
        try {
          People.People.deleteContact(sheetResourceName);
          stats.deleted++;
          sheet.getRange(i + 1, colMap.id + 1).clearContent(); 
        } catch (e) {
          if (!e.message.includes("404")) {
             stats.errors.push(`Row ${i+1} (Delete): ${e.message}`);
          }
        }
      }
      continue; 
    }

    // --- B. PREPARE LABELS ---
    // We do this per row to create missing labels on the fly if needed
    const labels = labelString.split(',').map(s => s.trim()).filter(s => s);
    const groupResourceNames = [];
    
    labels.forEach(label => {
      if (label === "Delete from contacts") return;
      if (!groupMap[label]) {
        const newGroup = createContactGroup(label);
        if (newGroup) {
          groupMap[label] = newGroup;
          stats.labelsCreated++;
        }
      }
      if (groupMap[label]) groupResourceNames.push(groupMap[label]);
    });

    // --- C. QUEUE UPDATE ---
    if (sheetResourceName) {
      if (sheetResourceName.startsWith("people/")) {
        // Add to batch queue
        updateQueue.push({
          row: row,
          rowIndex: i + 1, // 1-based index for logging
          resourceName: sheetResourceName,
          groupResourceNames: groupResourceNames
        });
      } else {
        stats.errors.push(`Row ${i+1}: Invalid ID format. Skipping.`);
      }
    } else {
      // Creation DISABLED
      continue;
    }

    // --- D. PROCESS BATCH IF FULL ---
    if (updateQueue.length >= BATCH_SIZE) {
      processUpdateBatch(updateQueue, colMap, stats);
      updateQueue = []; // Reset queue
      Utilities.sleep(1000); // Brief pause between batches
    }
  }

  // --- E. PROCESS REMAINING ITEMS ---
  if (updateQueue.length > 0) {
    processUpdateBatch(updateQueue, colMap, stats);
  }

  // --- REPORT ---
  let msg = `Sync Complete.
Updated: ${stats.updated}
Deleted: ${stats.deleted}
Skipped (Read-Only/Domain): ${stats.skipped}
New Labels: ${stats.labelsCreated}
(Creation Disabled)`;

  if (stats.errors.length > 0) {
    msg += `\n\nErrors (${stats.errors.length}):\n` + stats.errors.slice(0, 5).join("\n");
  }
  
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Processes a batch of up to 50 updates.
 * 1. Fetches ETags for all IDs in the batch (1 API Call).
 * 2. Matches Sheet Data to API Data.
 * 3. Sends Updates (1 API Call).
 */
function processUpdateBatch(queue, colMap, stats) {
  if (queue.length === 0) return;

  const resourceNames = queue.map(item => item.resourceName);
  const resourceToSheetMap = {}; // Map ID -> Queue Item
  queue.forEach(item => { resourceToSheetMap[item.resourceName] = item; });

  try {
    // 1. BATCH GET: Fetch metadata for all IDs to check existence and get ETags
    // getBatchGet returns { responses: [ { httpStatusCode, person, requestedResourceName } ] }
    const getResponse = People.People.getBatchGet({
      resourceNames: resourceNames,
      personFields: 'metadata'
    });

    const batchUpdatePayload = {};
    const validIds = [];

    // Process GET results
    if (getResponse.responses) {
      getResponse.responses.forEach(response => {
        const id = response.requestedResourceName;
        const sheetItem = resourceToSheetMap[id];

        if (response.httpStatusCode === 200 && response.person) {
          const person = response.person;
          
          // Check Read-Only
          const sources = person.metadata.sources || [];
          const isContact = sources.some(s => s.type === 'CONTACT');
          
          if (!isContact) {
            stats.skipped++;
            return;
          }

          // Construct Payload
          // Note: sheetItem.row is the raw array from the sheet
          const payload = constructContactPayload(sheetItem.row, colMap, sheetItem.groupResourceNames, person.etag);
          
          batchUpdatePayload[id] = payload;
          validIds.push(id);

        } else if (response.httpStatusCode === 404) {
          stats.errors.push(`Row ${sheetItem.rowIndex}: Contact not found (404).`);
        } else {
           stats.errors.push(`Row ${sheetItem.rowIndex}: API Error ${response.httpStatusCode}`);
        }
      });
    }

    // 2. BATCH UPDATE: Send all valid payloads
    if (validIds.length > 0) {
      // Fields to update
      const fields = ['memberships', 'names', 'organizations', 'emailAddresses', 'phoneNumbers'];

      const batchRequest = {
        contacts: batchUpdatePayload,
        updateMask: fields.join(',')
      };

      try {
        const updateResponse = People.People.batchUpdateContacts(batchRequest);
        // updateResponse.updateResult is a map of ID -> PersonResponse
        if (updateResponse.updateResult) {
           // Count successful updates
           const successes = Object.keys(updateResponse.updateResult).length;
           stats.updated += successes;
        }
      } catch (batchErr) {
        stats.errors.push(`Batch Update Failed: ${batchErr.message.substring(0, 100)}`);
      }
    }

  } catch (e) {
    stats.errors.push(`Batch Processing Error: ${e.message}`);
  }
}


// --- HELPER FUNCTIONS ---

function constructContactPayload(row, colMap, groupResourceNames, etag) {
  const getVal = (idx) => (idx !== -1 && row[idx]) ? String(row[idx]) : "";

  const payload = {};
  if (etag) payload.etag = etag;

  // 1. Names
  const given = getVal(colMap.firstName);
  const family = getVal(colMap.lastName);
  if (given || family) {
    payload.names = [{ givenName: given, familyName: family }];
  }

  // 2. Organizations
  const orgName = getVal(colMap.orgName);
  const orgTitle = getVal(colMap.orgTitle);
  const orgDept = getVal(colMap.orgDept);
  if (orgName || orgTitle || orgDept) {
    payload.organizations = [{ name: orgName, title: orgTitle, department: orgDept }];
  }

  // 3. Memberships
  payload.memberships = groupResourceNames.map(grp => ({
    contactGroupMembership: { contactGroupResourceName: grp }
  }));

  // 4. Emails
  const emails = [];
  colMap.emails.forEach(pair => {
    const val = getVal(pair.valIdx);
    const type = getVal(pair.typeIdx);
    if (val) emails.push({ value: val, type: type || 'other' });
  });
  payload.emailAddresses = emails; 

  // 5. Phones
  const phones = [];
  colMap.phones.forEach(pair => {
    const val = getVal(pair.valIdx);
    const type = getVal(pair.typeIdx);
    if (val) phones.push({ value: val, type: type || 'other' });
  });
  payload.phoneNumbers = phones;

  return payload;
}

function mapHeaders(headers) {
  const map = {
    id: headers.indexOf("Resource ID"),
    firstName: headers.indexOf("First Name"),
    lastName: headers.indexOf("Last Name"),
    labels: headers.indexOf("Labels"),
    orgName: headers.indexOf("Organization 1 - Company"),
    orgTitle: headers.indexOf("Organization 1 - Title"),
    orgDept: headers.indexOf("Organization 1 - Department"),
    emails: [],
    phones: []
  };

  headers.forEach((h, i) => {
    if (/^Email \d+ - Value$/.test(h)) {
      const num = h.match(/\d+/)[0];
      const typeIdx = headers.indexOf(`Email ${num} - Type`);
      if (typeIdx > -1) map.emails.push({ valIdx: i, typeIdx: typeIdx });
    }
    if (/^Phone \d+ - Value$/.test(h)) {
      const num = h.match(/\d+/)[0];
      const typeIdx = headers.indexOf(`Phone ${num} - Type`);
      if (typeIdx > -1) map.phones.push({ valIdx: i, typeIdx: typeIdx });
    }
  });
  
  return map;
}

function fetchAllContactGroups() {
  const map = {};
  let pageToken;
  try {
    do {
      const response = People.ContactGroups.list({
        groupFields: 'name',
        pageSize: 1000,
        pageToken: pageToken
      });
      if (response.contactGroups) {
        response.contactGroups.forEach(g => {
          map[g.formattedName] = g.resourceName;
        });
      }
      pageToken = response.nextPageToken;
    } while (pageToken);
  } catch (e) {
    console.log("Error fetching groups: " + e.message);
  }
  return map;
}

function createContactGroup(name) {
  try {
    const response = People.ContactGroups.create({
      contactGroup: { name: name }
    });
    return response.resourceName;
  } catch (e) {
    console.log(`Error creating label '${name}': ` + e.message);
    return null;
  }
}
