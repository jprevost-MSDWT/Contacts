/*
Project Name: Contact Import
Project Version: 6.02
Filename: ImportSetup.gs
File Version: 3.03
Chat link: [Insert Link]
*/

/**
 * ImportSetup.gs
 * This file handles the setup stage, moving data from the raw import
 * to the staging/setup sheet for processing.
 */

/**
 * Copies the raw imported contact data to the setup sheet.
 * Initializes the destination sheet if it doesn't exist.
 * Uses PASTE_VALUES to ensure manual formatting/validation on the destination
 * is not overwritten by the raw source data.
 */
function copyImportToSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.Import_SHEET_NAME);
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`Source sheet "${CONFIG.Import_SHEET_NAME}" not found.`);
    return;
  }

  let destSheet = ss.getSheetByName(CONFIG.Setup_SHEET_NAME);
  if (!destSheet) {
    destSheet = ss.insertSheet(CONFIG.Setup_SHEET_NAME);
  } else {
    // Only clear contents (values), keeping validation and formatting intact.
    destSheet.clearContents();
  }

  // Copy values from source to destination
  const sourceRange = sourceSheet.getDataRange();
  if (sourceRange.isBlank()) {
    SpreadsheetApp.getUi().alert("Source sheet is empty.");
    return;
  }
  
  // Use PASTE_VALUES so we don't overwrite the destination's validation rules
  sourceRange.copyTo(destSheet.getRange(1, 1), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}

/**
 * Applies formatting and data validation to the Setup sheet.
 * This should be run once, or when validation needs to be reset.
 * Separated from copy function to preserve manual "Chip" style settings during data updates.
 * Clears manual coloring before applying validation.
 */
function SetupSheetFormatting_DataValidation_Only() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.Setup_SHEET_NAME);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Target sheet "${CONFIG.Setup_SHEET_NAME}" not found. Run ImportCopy first.`);
    return;
  }

  // Clear any manual cell or text coloring to remove legacy highlights
  const fullRange = sheet.getDataRange();
  fullRange.setBackground(null);
  fullRange.setFontColor(null);

  const dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET_NAME);
  
  if (!dataSheet) {
    console.log(`Data sheet "${CONFIG.DATA_SHEET_NAME}" not found. Skipping validation.`);
    return;
  }

  // 1. Find the "All_Labels" column in Data sheet (Source of Truth)
  const dataHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
  const sourceColumnIndex = dataHeaders.indexOf(CONFIG.ALL_LABELS_HEADER);

  if (sourceColumnIndex === -1) {
    console.log(`Header "${CONFIG.ALL_LABELS_HEADER}" not found in Data sheet. Skipping validation.`);
    return;
  }

  // Define the range of valid labels (excluding header)
  const validationRange = dataSheet.getRange(2, sourceColumnIndex + 1, dataSheet.getMaxRows() - 1, 1);

  // 2. Find the "Labels" column in Setup sheet (Target)
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const targetColumnIndex = headers.indexOf("Labels");

  if (targetColumnIndex === -1) {
    console.log(`Header "Labels" not found in Setup sheet. Skipping validation.`);
    return;
  }

  // 3. Build and Apply the Validation Rule
  // Note: Apps Script cannot currently set "Chip" style programmatically.
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(validationRange, true) // true = show dropdown
    .setAllowInvalid(true) // Show warning only, allowing multiple/comma-separated items
    .setHelpText("For multi-select, manually change this column's Data Validation style to 'Chip' in the Data menu.")
    .build();

  const targetRange = sheet.getRange(2, targetColumnIndex + 1, sheet.getLastRow() - 1, 1);
  targetRange.setDataValidation(rule);
  
  SpreadsheetApp.getUi().alert("Formatting applied. You may now manually set Data Validation to 'Chip' style for multi-select support.");
}

/**
 * Applies general visual formatting to the Setup sheet.
 * Adds a filter to the data range.
 * Standardizes phone numbers to (xxx) xxx-xxxx format.
 * Sorts "Labels" cell values based on priority list.
 * Sorts ROWS based on priority list.
 * Highlights invalid phone numbers with red TEXT.
 * Formats @warren.k12.in.us emails: Sets Type to "Work" and highlights text gray.
 * Highlights Row ID (Col A) based on row status (Error, Contractor, Clean, etc).
 * Trims sheet exactly to the last row of data.
 * Copies Validation from Row 2 of Labels column down to the rest.
 */
function SetupSheetFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.Setup_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Target sheet "${CONFIG.Setup_SHEET_NAME}" not found.`);
    return;
  }

  // Clear any manual cell or text coloring to remove legacy highlights
  const fullRange = sheet.getDataRange();
  fullRange.setBackground(null);
  fullRange.setFontColor(null);

  const dataRange = sheet.getDataRange();
  
  // Apply Filter
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  dataRange.createFilter();

  // --- Formatting ---
  // Note: We use 'let' for values because we will modify it (sort rows)
  let values = dataRange.getValues();
  const headers = values[0];
  
  // 1. Identify Columns
  const phonePairs = [];
  const emailPairs = [];
  let labelsColumnIndex = -1;

  headers.forEach((header, index) => {
    if (/^Phone \d+ - Value$/.test(header)) {
      // Find matching Type column
      const match = header.match(/^Phone (\d+) - Value$/);
      const phoneNum = match[1];
      const typeHeader = `Phone ${phoneNum} - Type`;
      const typeIndex = headers.indexOf(typeHeader);
      
      if (typeIndex !== -1) {
        phonePairs.push({ valueIndex: index, typeIndex: typeIndex });
      }
    } else if (/^Email (\d+) - Value$/.test(header)) {
      const match = header.match(/^Email (\d+) - Value$/);
      const emailNum = match[1];
      const typeHeader = `Email ${emailNum} - Type`;
      const typeIndex = headers.indexOf(typeHeader);
      
      if (typeIndex !== -1) {
        emailPairs.push({ valueIndex: index, typeIndex: typeIndex });
      }
    } else if (header === "Labels") {
      labelsColumnIndex = index;
    }
  });

  if (phonePairs.length === 0 && emailPairs.length === 0) {
    console.log("No phone or email columns found to format.");
    return; 
  }

  // Label Sorting Priority
  const labelPriority = CONFIG.Labels_Values_Sort_Order.split(',').map(s => s.trim());
  const rowSortTop = CONFIG.Row_Labels_Sort_Order_Top.split(',').map(s => s.trim());
  const rowSortBottom = CONFIG.Row_Labels_Sort_Order_Bottom.split(',').map(s => s.trim());

  // Iterate rows (skip header) for Data Cleanup (In-Cell)
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    
    // --- Phone Logic (Normalize Values) ---
    phonePairs.forEach(pair => {
      let val = row[pair.valueIndex];
      if (val) {
        let strVal = String(val);
        let clean = strVal.replace(/\D/g, '');
        if (clean.length === 11 && clean.startsWith('1')) clean = clean.substring(1);
        if (clean.length === 10) {
          const formatted = `(${clean.substring(0, 3)}) ${clean.substring(3, 6)}-${clean.substring(6, 10)}`;
          if (strVal !== formatted) {
            values[r][pair.valueIndex] = formatted;
          }
        }
      }
    });

    // --- Email Logic (Update Types) ---
    emailPairs.forEach(pair => {
      const emailVal = row[pair.valueIndex];
      if (emailVal && String(emailVal).toLowerCase().includes(CONFIG.WARREN_DOMAIN)) {
        if (values[r][pair.typeIndex] !== 'Work') {
           values[r][pair.typeIndex] = 'Work';
        }
      }
    });

    // --- Label Sorting Logic (Sort string inside cell) ---
    if (labelsColumnIndex !== -1) {
      let labelVal = row[labelsColumnIndex];
      if (labelVal && typeof labelVal === 'string') {
        // Split, trim, and filter empty strings
        let labels = labelVal.split(',').map(s => s.trim()).filter(s => s);
        
        labels.sort((a, b) => {
          const idxA = labelPriority.indexOf(a);
          const idxB = labelPriority.indexOf(b);

          if (idxA !== -1 && idxB !== -1) return idxA - idxB; 
          if (idxA !== -1) return -1; 
          if (idxB !== -1) return 1;  
          return a.localeCompare(b); 
        });

        const newLabelString = labels.join(', ');
        if (newLabelString !== labelVal) {
          values[r][labelsColumnIndex] = newLabelString;
        }
      }
    }
  }

  // --- Row Sorting Logic (Sort rows by priority) ---
  if (labelsColumnIndex !== -1) {
    // Separate header
    const headerRow = values.shift(); 

    values.sort((a, b) => {
      const getPriorityScore = (row) => {
        const labelStr = (row[labelsColumnIndex] || "").toString();
        
        // Priority Top (Score 1..99)
        for (let i = 0; i < rowSortTop.length; i++) {
           if (labelStr.includes(rowSortTop[i])) return i + 1;
        }
        
        // Priority Bottom (Score 200..299)
        for (let i = 0; i < rowSortBottom.length; i++) {
           if (labelStr.includes(rowSortBottom[i])) return 200 + i;
        }
        
        // Default (Middle - Score 100)
        return 100; 
      };
      
      const scoreA = getPriorityScore(a);
      const scoreB = getPriorityScore(b);
      
      return scoreA - scoreB;
    });

    // Add header back
    values.unshift(headerRow);
  }

  // Write sorted and cleaned values back to sheet
  dataRange.setValues(values);

  // --- Apply Conditional Formatting ---
  sheet.clearConditionalFormatRules();
  const rules = [];
  const rangeA = sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1);

  // 1. DATA COLUMNS: Phone Error (Red BG)
  phonePairs.forEach(pair => {
    const ranges = [
      sheet.getRange(2, pair.valueIndex + 1, sheet.getMaxRows() - 1, 1),
      sheet.getRange(2, pair.typeIndex + 1, sheet.getMaxRows() - 1, 1)
    ];
    
    const valColLetter = sheet.getRange(1, pair.valueIndex + 1).getA1Notation().replace(/\d+/g, '');
    const cellRef = `$${valColLetter}2`;
    
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${cellRef})), NOT(REGEXMATCH(TO_TEXT(${cellRef}), "^\\(\\d{3}\\) \\d{3}-\\d{4}$")))`)
      .setBackground(CONFIG.ERROR_COLOR_Red)
      .setRanges(ranges)
      .build();
    rules.push(rule);
  });

  // 2. DATA COLUMNS: Email Error (Red Text)
  const emailRegex = "^[\\w.%+-]+@[\\w.-]+\\.[a-zA-Z]{2,}$";
  emailPairs.forEach(pair => {
    const ranges = [
      sheet.getRange(2, pair.valueIndex + 1, sheet.getMaxRows() - 1, 1),
      sheet.getRange(2, pair.typeIndex + 1, sheet.getMaxRows() - 1, 1)
    ];
    
    const valColLetter = sheet.getRange(1, pair.valueIndex + 1).getA1Notation().replace(/\d+/g, '');
    const cellRef = `$${valColLetter}2`;

    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${cellRef})), NOT(REGEXMATCH(TO_TEXT(${cellRef}), "${emailRegex}")))`)
      .setFontColor(CONFIG.ERROR_COLOR_Red)
      .setRanges(ranges)
      .build();
    rules.push(rule);
  });

  // 3. DATA COLUMNS: Warren Email Gray Text (Value & Type)
  emailPairs.forEach(pair => {
    const ranges = [
      sheet.getRange(2, pair.valueIndex + 1, sheet.getMaxRows() - 1, 1),
      sheet.getRange(2, pair.typeIndex + 1, sheet.getMaxRows() - 1, 1)
    ];
    
    const valColLetter = sheet.getRange(1, pair.valueIndex + 1).getA1Notation().replace(/\d+/g, '');
    const cellRef = `$${valColLetter}2`;

    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=ISNUMBER(SEARCH("${CONFIG.WARREN_DOMAIN}", ${cellRef}))`)
      .setFontColor(CONFIG.WARREN_COLOR)
      .setRanges(ranges)
      .build();
    rules.push(rule);
  });

  // 4. DATA COLUMNS: Non-Warren Email Purple Text (Value & Type)
  if (emailPairs.length > 0) {
    emailPairs.forEach(pair => {
      const ranges = [
        sheet.getRange(2, pair.valueIndex + 1, sheet.getMaxRows() - 1, 1),
        sheet.getRange(2, pair.typeIndex + 1, sheet.getMaxRows() - 1, 1)
      ];
      
      const valColLetter = sheet.getRange(1, pair.valueIndex + 1).getA1Notation().replace(/\d+/g, '');
      const cellRef = `$${valColLetter}2`;

      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${cellRef})), ISERROR(SEARCH("${CONFIG.WARREN_DOMAIN}", ${cellRef})))`)
        .setFontColor(CONFIG.NonWarren_Email_Color)
        .setRanges(ranges)
        .build();
      rules.push(rule);
    });
  }

  // --- COLUMN A RULES (Resource ID Dashboard) ---

  // A. Non-Warren Email (Text #9900ff)
  if (emailPairs.length > 0) {
    const nonWarrenConditions = emailPairs.map(pair => {
      const colLetter = sheet.getRange(1, pair.valueIndex + 1).getA1Notation().replace(/\d+/g, '');
      return `AND($${colLetter}2<>"", ISERROR(SEARCH("${CONFIG.WARREN_DOMAIN}", $${colLetter}2)))`;
    });
    
    if (nonWarrenConditions.length > 0) {
      const nonWarrenFormula = `=OR(${nonWarrenConditions.join(',')})`;
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(nonWarrenFormula)
        .setFontColor(CONFIG.NonWarren_Email_Color)
        .setRanges([rangeA])
        .build());
    }
  }

  // B. Clean Status (Text #C0C0C0)
  if (emailPairs.length > 0) {
    const hasWarrenConditions = emailPairs.map(pair => {
      const colLetter = sheet.getRange(1, pair.valueIndex + 1).getA1Notation().replace(/\d+/g, '');
      return `ISNUMBER(SEARCH("${CONFIG.WARREN_DOMAIN}", $${colLetter}2))`;
    });
    
    const noNonWarrenConditions = emailPairs.map(pair => {
      const colLetter = sheet.getRange(1, pair.valueIndex + 1).getA1Notation().replace(/\d+/g, '');
      return `OR($${colLetter}2="", ISNUMBER(SEARCH("${CONFIG.WARREN_DOMAIN}", $${colLetter}2)))`;
    });

    let noPhonesFormula = "TRUE";
    if (phonePairs.length > 0) {
      const phoneConditions = phonePairs.map(pair => {
        const colLetter = sheet.getRange(1, pair.valueIndex + 1).getA1Notation().replace(/\d+/g, '');
        return `$${colLetter}2=""`;
      });
      noPhonesFormula = `AND(${phoneConditions.join(',')})`;
    }

    const cleanFormula = `=AND(OR(${hasWarrenConditions.join(',')}), AND(${noNonWarrenConditions.join(',')}), ${noPhonesFormula})`;
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(cleanFormula)
      .setFontColor(CONFIG.WARREN_COLOR) 
      .setRanges([rangeA])
      .build());
  }

  // C. Contractor (BG #9900ff)
  if (labelsColumnIndex !== -1) {
    const colLetter = sheet.getRange(1, labelsColumnIndex + 1).getA1Notation().replace(/\d+/g, '');
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=ISNUMBER(SEARCH("Contractor", $${colLetter}2))`)
      .setBackground(CONFIG.Contractor_Color)
      .setRanges([rangeA])
      .build());
  }

  // D. Error Warning (BG #ffcc99) - Highest Priority Background
  let errorConditions = [];
  
  if (phonePairs.length > 0) {
    phonePairs.forEach(pair => {
      const colLetter = sheet.getRange(1, pair.valueIndex + 1).getA1Notation().replace(/\d+/g, '');
      errorConditions.push(`AND($${colLetter}2<>"", NOT(REGEXMATCH(TO_TEXT($${colLetter}2), "^\\(\\d{3}\\) \\d{3}-\\d{4}$")))`);
    });
  }
  
  if (emailPairs.length > 0) {
    emailPairs.forEach(pair => {
      const colLetter = sheet.getRange(1, pair.valueIndex + 1).getA1Notation().replace(/\d+/g, '');
      errorConditions.push(`AND($${colLetter}2<>"", NOT(REGEXMATCH(TO_TEXT($${colLetter}2), "${emailRegex}")))`);
    });
  }

  if (errorConditions.length > 0) {
    const errorFormula = `=OR(${errorConditions.join(',')})`;
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(errorFormula)
      .setBackground(CONFIG.COLUMN_A_WARNING_COLOR)
      .setRanges([rangeA])
      .build());
  }

  // --- LOW PRIORITY ROW-WIDE RULES ---
  if (labelsColumnIndex !== -1) {
    const colLetter = sheet.getRange(1, labelsColumnIndex + 1).getA1Notation().replace(/\d+/g, '');
    const allDataRange = sheet.getDataRange().offset(1, 0, sheet.getLastRow() - 1); 

    // 1. Archive Row (Cyan)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=ISNUMBER(SEARCH("Archive", $${colLetter}2))`)
      .setBackground(CONFIG.ARCHIVE_ROW_COLOR)
      .setRanges([allDataRange])
      .build());

    // 2. Delete Row (Light Red)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=ISNUMBER(SEARCH("Delete from contacts", $${colLetter}2))`)
      .setBackground(CONFIG.DELETE_ROW_COLOR)
      .setRanges([allDataRange])
      .build());
  }

  sheet.setConditionalFormatRules(rules);
  SpreadsheetApp.flush();

  // --- TRIM SHEET TO EXACT DATA ---
  const lastDataRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  
  if (maxRows > lastDataRow) {
    sheet.deleteRows(lastDataRow + 1, maxRows - lastDataRow);
  }

  // --- COPY DOWN VALIDATION ---
  // Only copy if we have data below Row 2
  if (labelsColumnIndex !== -1 && lastDataRow > 2) {
    const sourceCell = sheet.getRange(2, labelsColumnIndex + 1);
    
    // Fill from Row 3 to end
    const rowsToFill = lastDataRow - 2; 
    
    const targetRange = sheet.getRange(3, labelsColumnIndex + 1, rowsToFill, 1);
    sourceCell.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }
}
