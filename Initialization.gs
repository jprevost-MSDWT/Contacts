/*
Project Name: Contact Import
Project Version: 9.01
Filename: Initialization.gs
File Version: 4.01
Chat link: [Insert Link]
*/

/**
 * CONFIGURATION
 * Centralized settings for sheet names and data structure keys.
 */
const CONFIG = {
  Import_SHEET_NAME: "Contacts-Import",
  Setup_SHEET_NAME: "SetupSheet",
  DATA_SHEET_NAME: "Data",
  Imported_LABELS_HEADER: "ImportedLabels",
  Imported_NAMED_RANGE: "ImportedLabels",
  ALL_LABELS_HEADER: "All_Labels",
  // Formatting Configuration
  ERROR_COLOR_Red: '#ff6666',
  COLUMN_A_WARNING_COLOR: '#ffcc99',
  WARREN_COLOR: '#C0C0C0',
  NonWarren_Email_Color: '#9900ff',
  Contractor_Color: '#9900ff',
  DELETE_ROW_COLOR: '#f4c7c3', // Light Red for "Delete from contacts"
  ARCHIVE_ROW_COLOR: '#d0e0e3', // Light Cyan for "Archive"
  WARREN_DOMAIN: '@warren.k12.in.us',
  // Sorting Configuration (Comma Separated Values)
  Labels_Values_Sort_Order: "Work Phone List, Contractors, Personal, Porter, Archive, Delete from contacts",
  Row_Labels_Sort_Order_Top: "Work Phone List, Contractors, Personal,Porter",
  Row_Labels_Sort_Order_Bottom: "Archive, Delete from contacts"
};

/**
 * The onOpen function runs when the spreadsheet is opened.
 * It adds a custom menu to the spreadsheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Contact Tools')
      .addItem('Import All Contacts', 'importContactsToSheet')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Setup')
          .addItem('ImportCopy', 'copyImportToSetup')
          .addItem('Setup Validation', 'SetupSheetFormatting_DataValidation_Only')
          .addItem('Setup Formatting', 'SetupSheetFormatting'))
      .addItem('Re-upload', 'syncSheetsToContacts')
      .addToUi();
}
