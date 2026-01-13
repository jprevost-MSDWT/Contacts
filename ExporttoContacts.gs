/*
Project Name: Contact Import
Project Version: 9.01
Filename: ExporttoContacts.gs
File Version: 5.02
Chat link: [Insert Link]
*/

/**
 * Syncs the structured data from the Setup Sheet back to Google Contacts.
 * - Uses BATCH OPERATIONS.
 * - Reports specific REASONS for updates to help debug phantom changes.
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
  
  const colMap = mapHeaders(headers);
  const missingCols = [];
  if (colMap.id === -1) missingCols.push("Resource ID");
  if (colMap.labels === -1) missingCols.push("Labels");
  if (missingCols.length > 0) {
    SpreadsheetApp.getUi().alert(`Error: Missing columns: ${missingCols.join(", ")}`);
    return;
  }

  let groupMap = fetchAllContactGroups();

  // Stats include 'changes' array to log specific reasons
  let stats = { updated: 0, created: 0, deleted: 0, labelsCreated: 0, errors: [], skipped: 0, unchanged: 0, changes: [] };
  let updateQueue = []; 
  const BATCH_SIZE = 50; 

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let sheetResourceName = row[colMap.id] ? String(row[colMap.id]).trim().replace(/[\s\u200B-\u200D\uFEFF]/g, '') : "";
    const labelString = row[colMap.labels] ? String(row[colMap.labels]) : "";
    
    // --- DELETE ---
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

    // --- PREPARE LABELS ---
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

    // --- QUEUE ---
    if (sheetResourceName) {
      if (sheetResourceName.startsWith("people/")) {
        updateQueue.push({
          row: row,
          rowIndex: i + 1, 
          resourceName: sheetResourceName,
          groupResourceNames: groupResourceNames
        });
      } else {
        stats.errors.push(`Row ${i+1}: Invalid ID. Skipping.`);
      }
    } else {
      continue; // Creation Disabled
    }

    if (updateQueue.length >= BATCH_SIZE) {
      processUpdateBatch(updateQueue, colMap, stats, groupMap);
      updateQueue = []; 
      Utilities.sleep(1000); 
    }
  }

  if (updateQueue.length > 0) {
    processUpdateBatch(updateQueue, colMap, stats, groupMap);
  }

  // --- REPORT ---
  let msg = `Sync Complete.
Updated: ${stats.updated}
Unchanged: ${stats.unchanged}
Deleted: ${stats.deleted}
New Labels: ${stats.labelsCreated}`;

  if (stats.changes.length > 0) {
    msg += `\n\n--- Update Reasons (First 10) ---\n` + stats.changes.slice(0, 10).join("\n");
  }

  if (stats.errors.length > 0) {
    msg += `\n\n--- Errors (First 5) ---\n` + stats.errors.slice(0, 5).join("\n");
  }
  
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Processes a batch of contacts to update.
 * It fetches the full contact data from the People API, compares it with the sheet data,
 * and sends a batch update request only for contacts with detected changes.
 * @param {Array<Object>} queue The queue of items to process.
 * @param {Object} colMap A map of header names to column indices.
 * @param {Object} stats The statistics object to update.
 * @param {Object} groupMap A map of contact group names to their resource names.
 */
function processUpdateBatch(queue, colMap, stats, groupMap) {
  if (queue.length === 0) return;

  const resourceNames = queue.map(item => item.resourceName);
  const resourceToSheetMap = {}; 
  queue.forEach(item => { resourceToSheetMap[item.resourceName] = item; });
  const knownGroupIds = new Set(Object.values(groupMap));

  try {
    const getResponse = People.People.getBatchGet({
      resourceNames: resourceNames,
      personFields: 'metadata,names,organizations,emailAddresses,phoneNumbers,memberships'
    });

    const batchUpdatePayload = {};
    const validIds = [];

    if (getResponse.responses) {
      getResponse.responses.forEach(response => {
        const requestedId = response.requestedResourceName;
        const sheetItem = resourceToSheetMap[requestedId];

        if (response.httpStatusCode === 200 && response.person) {
          const person = response.person;
          const canonicalId = person.resourceName;
          
          const sources = person.metadata.sources || [];
          const isContact = sources.some(s => s.type === 'CONTACT');
          if (!isContact) {
            stats.skipped++;
            return;
          }

          const payload = constructContactPayload(sheetItem.row, colMap, sheetItem.groupResourceNames, person.etag);
          
          // CHECK DIFFERENCES
          const diffReason = getContentDiffReason(payload, person, knownGroupIds, sheetItem.rowIndex);
          
          if (diffReason) {
            batchUpdatePayload[canonicalId] = payload;
            validIds.push(canonicalId);
            stats.changes.push(`Row ${sheetItem.rowIndex}: ${diffReason}`);
            Logger.log(`[Diff Row ${sheetItem.rowIndex}] ${diffReason}`); 
          } else {
            stats.unchanged++;
          }

        } else if (response.httpStatusCode === 404) {
          stats.errors.push(`Row ${sheetItem.rowIndex}: Contact ID 404. skipped.`);
        } else {
           stats.errors.push(`Row ${sheetItem.rowIndex}: GET Error ${response.httpStatusCode}`);
        }
      });
    }

    if (validIds.length > 0) {
      const fields = ['memberships', 'names', 'organizations', 'emailAddresses', 'phoneNumbers'];
      const batchRequest = {
        contacts: batchUpdatePayload,
        updateMask: fields.join(',')
      };

      try {
        const updateResponse = People.People.batchUpdateContacts(batchRequest);
        if (updateResponse && updateResponse.updateResult) {
           const successes = Object.keys(updateResponse.updateResult).length;
           stats.updated += successes;
        } else {
           // Fallback with Warning
           Logger.log(`Batch update response missing 'updateResult' for ${validIds.length} contacts. Optimistically assuming success.`);
           stats.updated += validIds.length;
        }
      } catch (batchErr) {
        stats.errors.push(`Batch Update Failed: ${batchErr.message.substring(0, 100)}`);
      }
    }

  } catch (e) {
    stats.errors.push(`Batch Processing Error: ${e.message}`);
  }
}

/**
 * Returns a string reason if content is different, or null if identical.
 * Accumulates multiple differences.
 */
function getContentDiffReason(payload, contact, knownGroupIds, rowNum) {
  const norm = (str) => (str || "").trim();
  // Filter helper: Only keep fields that are editable CONTACT type (ignore Domain/System fields)
  const isEditable = (item) => {
    if (!item.metadata || !item.metadata.source) return true; // Assume true if no metadata (rare)
    return item.metadata.source.type === 'CONTACT';
  };

  const diffs = [];

  // 1. Names
  const pName = payload.names ? payload.names[0] : {};
  const cName = (contact.names && contact.names.length > 0) ? contact.names[0] : {};
  if (norm(pName.givenName) !== norm(cName.givenName)) {
    diffs.push(`Name (Given): "${pName.givenName}" != "${cName.givenName}"`);
  }
  if (norm(pName.familyName) !== norm(cName.familyName)) {
    diffs.push(`Name (Family): "${pName.familyName}" != "${cName.familyName}"`);
  }

  // 2. Organizations
  const pOrg = payload.organizations ? payload.organizations[0] : {};
  const cOrg = (contact.organizations && contact.organizations.length > 0) ? contact.organizations[0] : {};
  if (norm(pOrg.name) !== norm(cOrg.name)) {
    diffs.push(`Org Name: "${pOrg.name}" != "${cOrg.name}"`);
  }
  if (norm(pOrg.title) !== norm(cOrg.title)) {
    diffs.push(`Org Title: "${pOrg.title}" != "${cOrg.title}"`);
  }
  if (norm(pOrg.department) !== norm(cOrg.department)) {
    diffs.push(`Org Dept: "${pOrg.department}" != "${cOrg.department}"`);
  }

  // 3. Emails (Compare VALUES ONLY)
  // FILTER: Only compare against Google emails that are source: CONTACT
  const getEmailSig = (e) => norm(e.value).toLowerCase();
  
  const pEmails = new Set((payload.emailAddresses || []).map(getEmailSig));
  
  // Filter Google emails to only those we can edit (User Contacts)
  const editableGoogleEmails = (contact.emailAddresses || []).filter(isEditable);
  const cEmails = new Set(editableGoogleEmails.map(getEmailSig));
  
  // Check for additions
  const emailsToAdd = [...pEmails].filter(e => !cEmails.has(e));
  if (emailsToAdd.length > 0) {
    diffs.push(`Email(s) to add: ${emailsToAdd.join(', ')}`);
  }

  // Check for removals
  const emailsToRemove = [...cEmails].filter(e => !pEmails.has(e));
  if (emailsToRemove.length > 0) {
    diffs.push(`Email(s) to remove: ${emailsToRemove.join(', ')}`);
  }

  // 4. Phones (Compare Digits ONLY)
  const cleanPhone = (p) => {
    let s = norm(p.value).replace(/\D/g, "");
    if (s.length === 11 && s.startsWith("1")) s = s.substring(1); 
    return s;
  };
  
  const pPhones = new Set((payload.phoneNumbers || []).map(cleanPhone));
  
  // Filter Google phones to only those we can edit
  const editableGooglePhones = (contact.phoneNumbers || []).filter(isEditable);
  const cPhones = new Set(editableGooglePhones.map(cleanPhone));
  
  if (pPhones.size !== cPhones.size) {
    diffs.push(`Phone Count: ${pPhones.size} vs ${cPhones.size}`);
  } else {
    for (let p of pPhones) {
      if (!cPhones.has(p)) {
        diffs.push(`Phone mismatch: "${p}"`);
      }
    }
  }

  // 5. Memberships
  // Filter API memberships to only include groups we know about (managed in the sheet)
  const getGroups = (arr) => (arr || [])
    .map(m => m.contactGroupMembership ? m.contactGroupMembership.contactGroupResourceName : null)
    .filter(id => id && knownGroupIds.has(id)); 
  
  const pGroups = new Set(
     (payload.memberships || []).map(m => m.contactGroupMembership.contactGroupResourceName)
  );
  const cGroups = new Set(getGroups(contact.memberships));
  
  if (pGroups.size !== cGroups.size) {
    diffs.push(`Label Count: ${pGroups.size} vs ${cGroups.size}`);
  } else {
    for (let g of pGroups) {
      if (!cGroups.has(g)) {
        diffs.push(`Label mismatch: Sheet has ${g} which Google lacks`);
      }
    }
  }

  if (diffs.length > 0) return diffs.join('; ');
  return null; // No difference found
}

function constructContactPayload(row, colMap, groupResourceNames, etag) {
  const getVal = (idx) => (idx !== -1 && row[idx]) ? String(row[idx]).trim() : "";

  const payload = {};
  if (etag) payload.etag = etag;

  // Names
  const given = getVal(colMap.firstName);
  const family = getVal(colMap.lastName);
  if (given || family) {
    payload.names = [{ givenName: given, familyName: family }];
  }

  // Organizations
  const orgName = getVal(colMap.orgName);
  const orgTitle = getVal(colMap.orgTitle);
  const orgDept = getVal(colMap.orgDept);
  if (orgName || orgTitle || orgDept) {
    payload.organizations = [{ name: orgName, title: orgTitle, department: orgDept }];
  }

  // Memberships
  payload.memberships = groupResourceNames.map(grp => ({
    contactGroupMembership: { contactGroupResourceName: grp }
  }));

  // Emails
  const emails = [];
  colMap.emails.forEach(pair => {
    const val = getVal(pair.valIdx);
    const type = getVal(pair.typeIdx);
    if (val) emails.push({ value: val, type: type || 'other' });
  });
  payload.emailAddresses = emails; 

  // Phones
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
