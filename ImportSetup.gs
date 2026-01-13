/*
Project Name: Contact Import
Project Version: 9.10
Filename: ImportSetup.gs
File Version: 4.03
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
 * Highlights DUPLICATE phone/emails with orange TEXT (via Conditional Formatting).
 * Formats @warren.k12.in.us emails: Sets Type to "Work" and highlights text gray.
 * Highlights Row ID (Col A) based on row status (Error, Duplicate, Contractor, Clean, etc).
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
  let values = dataRange.getValues();
  const headers = values[0];
  const backgrounds = dataRange.getBackgrounds();
  const fontColors = dataRange.getFontColors();
  
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

  let hasChanges = false;
  let hasFormatChanges = false;

  // Iterate rows (skip header) for Data Cleanup (In-Cell) 
  // Note: We removed manual duplicate highlighting from here to use CF instead
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    
    // --- Phone Logic (Normalize) ---
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
            hasChanges = true;
          }
        }
      }
    });

    // --- Email Logic (Update Types) ---
    emailPairs.forEach(pair => {
      const emailVal = row[pair.valueIndex];
      if (emailVal) {
        if (String(emailVal).toLowerCase().includes(CONFIG.WARREN_DOMAIN)) {
          if (values[r][pair.typeIndex] !== 'Work') {
             values[r][pair.typeIndex] = 'Work';
             hasChanges = true;
          }
        }
      }
    });

    // --- Label Sorting Logic (Sort string inside cell) ---
    if (labelsColumnIndex !== -1) {
      let labelVal = row[labelsColumnIndex];
      if (labelVal && typeof labelVal === 'string') {
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
          hasChanges = true;
        }
      }
    }
  }

  // --- Row Sorting Logic (Sort rows by priority) ---
  if (labelsColumnIndex !== -1) {
    const headerRow = values.shift(); 
    const headerBg = backgrounds.shift();
    const headerFc = fontColors.shift();

    const combined = values.map((row, i) => ({
       row: row,
       bg: backgrounds[i],
       fc: fontColors[i]
    }));

    combined.sort((a, b) => {
      const getPriorityScore = (item) => {
        const labelStr = (item.row[labelsColumnIndex] || "").toString();
        for (let i = 0; i < rowSortTop.length; i++) {
           if (labelStr.includes(rowSortTop[i])) return i + 1;
        }
        for (let i = 0; i < rowSortBottom.length; i++) {
           if (labelStr.includes(rowSortBottom[i])) return 200 + i;
        }
        return 100; 
      };
      
      return getPriorityScore(a) - getPriorityScore(b);
    });

    values = combined.map(c => c.row);
    backgrounds = combined.map(c => c.bg);
    fontColors = combined.map(c => c.fc);

    values.unshift(headerRow);
    backgrounds.unshift(headerBg);
    fontColors.unshift(headerFc);

    hasChanges = true; // Always write if we sorted
    hasFormatChanges = true;
  }

  if (hasChanges) {
    dataRange.setValues(values);
  }
  if (hasFormatChanges) {
    dataRange.setBackgrounds(backgrounds);
    dataRange.setFontColors(fontColors);
  }

  // --- Apply Conditional Formatting ---
  sheet.clearConditionalFormatRules();
  const rules = [];
  const rangeA = sheet.getRange(2, 1, sheet.getMaxRows() - 1, 1);
  const emailRegex = "^[\\w.%+-]+@[\\w.-]+\\.[a-zA-Z]{2,}$";
  
  // Track all error/duplicate conditions for Column A formula
  let colA_ErrorConditions = [];

  // 1. DATA COLUMNS: Phone Rules
  // Loop through phones to create per-column rules
  // For duplicates, we check current col vs previous cols
  phonePairs.forEach((pair, index) => {
    const ranges = [
      sheet.getRange(2, pair.valueIndex + 1, sheet.getMaxRows() - 1, 1),
      sheet.getRange(2, pair.typeIndex + 1, sheet.getMaxRows() - 1, 1)
    ];
    const valColLetter = sheet.getRange(1, pair.valueIndex + 1).getA1Notation().replace(/\d+/g, '');
    const cellRef = `$${valColLetter}2`;

    // A. Phone Error (Red BG)
    const errFormula = `AND(NOT(ISBLANK(${cellRef})), NOT(REGEXMATCH(TO_TEXT(${cellRef}), "^\\(\\d{3}\\) \\d{3}-\\d{4}$")))`;
    colA_ErrorConditions.push(errFormula);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=${errFormula}`)
      .setBackground(CONFIG.ERROR_COLOR_Red)
      .setRanges(ranges)
      .build());
      
    // B. Phone Duplicate (Orange Text)
    // Compare current column index vs all PREVIOUS phone pair value indices
    if (index > 0) {
      let dupCheckParts = [];
      for (let k = 0; k < index; k++) {
        const prevColLetter = sheet.getRange(1, phonePairs[k].valueIndex + 1).getA1Notation().replace(/\d+/g, '');
        dupCheckParts.push(`${cellRef}=$${prevColLetter}2`);
      }
      
      const dupFormula = `AND(NOT(ISBLANK(${cellRef})), OR(${dupCheckParts.join(',')}))`;
      colA_ErrorConditions.push(dupFormula); // Duplicates trigger Col A warning too
      
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=${dupFormula}`)
        .setFontColor(CONFIG.Dup_Color)
        .setRanges(ranges)
        .build());
    }
  });

  // 2. DATA COLUMNS: Email Rules
  emailPairs.forEach((pair, index) => {
    const ranges = [
      sheet.getRange(2, pair.valueIndex + 1, sheet.getMaxRows() - 1, 1),
      sheet.getRange(2, pair.typeIndex + 1, sheet.getMaxRows() - 1, 1)
    ];
    const valColLetter = sheet.getRange(1, pair.valueIndex + 1).getA1Notation().replace(/\d+/g, '');
    const cellRef = `$${valColLetter}2`;

    // A. Email Error (Red Text)
    const errFormula = `AND(NOT(ISBLANK(${cellRef})), NOT(REGEXMATCH(TO_TEXT(${cellRef}), "${emailRegex}")))`;
    colA_ErrorConditions.push(errFormula);

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=${errFormula}`)
      .setFontColor(CONFIG.ERROR_COLOR_Red)
      .setRanges(ranges)
      .build());

    // B. Warren Email (Gray Text)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=ISNUMBER(SEARCH("${CONFIG.WARREN_DOMAIN}", ${cellRef}))`)
      .setFontColor(CONFIG.WARREN_COLOR)
      .setRanges(ranges)
      .build());

    // C. Non-Warren Email (Purple Text)
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(NOT(ISBLANK(${cellRef})), ISERROR(SEARCH("${CONFIG.WARREN_DOMAIN}", ${cellRef})))`)
      .setFontColor(CONFIG.NonWarren_Email_Color)
      .setRanges(ranges)
      .build());
      
    // D. Email Duplicate (Orange Text)
    if (index > 0) {
      let dupCheckParts = [];
      for (let k = 0; k < index; k++) {
        const prevColLetter = sheet.getRange(1, emailPairs[k].valueIndex + 1).getA1Notation().replace(/\d+/g, '');
        dupCheckParts.push(`${cellRef}=$${prevColLetter}2`);
      }
      
      const dupFormula = `AND(NOT(ISBLANK(${cellRef})), OR(${dupCheckParts.join(',')}))`;
      colA_ErrorConditions.push(dupFormula);
      
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=${dupFormula}`)
        .setFontColor(CONFIG.Dup_Color)
        .setRanges(ranges)
        .build());
    }
  });

  // --- COLUMN A RULES ---

  // A. Non-Warren Email (Text Purple)
  if (emailPairs.length > 0) {
    const nonWarrenConditions = emailPairs.map(pair => {
      const colLetter = sheet.getRange(1, pair.valueIndex + 1).getA1Notation().replace(/\d+/g, '');
      return `AND($${colLetter}2<>"", ISERROR(SEARCH("${CONFIG.WARREN_DOMAIN}", $${colLetter}2)))`;
    });
    
    if (nonWarrenConditions.length > 0) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=OR(${nonWarrenConditions.join(',')})`)
        .setFontColor(CONFIG.NonWarren_Email_Color)
        .setRanges([rangeA])
        .build());
    }
  }

  // B. Clean Status (Text Gray)
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

  // C. Contractor (BG Purple)
  if (labelsColumnIndex !== -1) {
    const colLetter = sheet.getRange(1, labelsColumnIndex + 1).getA1Notation().replace(/\d+/g, '');
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=ISNUMBER(SEARCH("Contractor", $${colLetter}2))`)
      .setBackground(CONFIG.Contractor_Color)
      .setRanges([rangeA])
      .build());
  }

  // D. Error Warning (BG #ffcc99) - INCLUDES DUPLICATES
  if (colA_ErrorConditions.length > 0) {
    const errorFormula = `=OR(${colA_ErrorConditions.join(',')})`;
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

  // --- TRIM SHEET ---
  const lastDataRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  if (maxRows > lastDataRow) {
    sheet.deleteRows(lastDataRow + 1, maxRows - lastDataRow);
  }

  // --- COPY DOWN VALIDATION ---
  if (labelsColumnIndex !== -1 && lastDataRow > 2) {
    const sourceCell = sheet.getRange(2, labelsColumnIndex + 1);
    const rowsToFill = lastDataRow - 2; 
    const targetRange = sheet.getRange(3, labelsColumnIndex + 1, rowsToFill, 1);
    sourceCell.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }
}
