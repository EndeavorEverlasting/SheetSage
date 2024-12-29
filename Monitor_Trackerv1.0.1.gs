/**
 * Main function to coordinate protection and updating of Automated Notes columns.
 */
function myFunction() {
  const traceId = generateTraceId(); // Unique ID for trace logs
  startSpan(traceId, "Main Execution");
  try {
    protectAutomatedNotesColumn(traceId); // Protects the Automated Notes column
    updateAutomatedNotes(traceId);       // Updates notes dynamically
  } catch (error) {
    logError(traceId, "Main Execution", error); // Log errors for debugging
  }
  endSpan(traceId, "Main Execution");
}

/**
 * Utility to generate a unique trace ID for logging.
 * @returns {string} A random string to uniquely identify a trace.
 */
function generateTraceId() {
  return Math.random().toString(36).substr(2, 9);
}

/**
 * Logs the start of a span for tracing execution.
 * @param {string} traceId - Unique trace identifier.
 * @param {string} spanName - Name of the execution span.
 */
function startSpan(traceId, spanName) {
  Logger.log(`[${traceId}] Starting span: ${spanName}`);
}

/**
 * Logs the end of a span for tracing execution.
 * @param {string} traceId - Unique trace identifier.
 * @param {string} spanName - Name of the execution span.
 */
function endSpan(traceId, spanName) {
  Logger.log(`[${traceId}] Ending span: ${spanName}`);
}

/**
 * Logs an error during a specific span.
 * @param {string} traceId - Unique trace identifier.
 * @param {string} spanName - Name of the execution span where the error occurred.
 * @param {Error} error - The error to log.
 */
function logError(traceId, spanName, error) {
  Logger.log(`[${traceId}] Error in span "${spanName}": ${error.message}`);
}

/**
 * Protects the "Automated Notes" column in all sheets except the penultimate tab.
 * @param {string} traceId - Unique trace identifier for logging.
 */
function protectAutomatedNotesColumn(traceId) {
  startSpan(traceId, "Protect Automated Notes Column");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  sheets.forEach((sheet, index) => {
    const sheetName = sheet.getName();

    // Skip protection for certain sheets
    if (isProtectedSheet(sheetName, sheets.length, index)) {
      Logger.log(`[${traceId}] Skipping protection for sheet "${sheetName}".`);
      return;
    }

    // Retrieve header row dynamically
    const headerRow = getHeaderRow(sheet);
    if (!headerRow.includes("Automated Notes")) {
      Logger.log(`[${traceId}] Header "Automated Notes" not found in sheet "${sheetName}".`);
      return;
    }

    // Get column index of "Automated Notes"
    const lastRow = sheet.getLastRow();
    const columnIndex = headerRow.indexOf("Automated Notes") + 1;
    const columnRange = sheet.getRange(1, columnIndex, lastRow);

    // Apply protection to the range
    let protection = getExistingProtection(sheet, columnRange);
    if (!protection) {
      protection = columnRange.protect();
      protection.setDescription(`Protecting "Automated Notes" column in "${sheetName}".`);
    }
    protection.setWarningOnly(true); // Allow edits but show warnings
    Logger.log(`[${traceId}] Protection applied to column "Automated Notes" in sheet "${sheetName}".`);
  });

  endSpan(traceId, "Protect Automated Notes Column");
}

/**
 * Updates the "Automated Notes" column in all applicable sheets.
 * @param {string} traceId - Unique trace identifier for logging.
 */
function updateAutomatedNotes(traceId) {
  startSpan(traceId, "Update Automated Notes");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Retrieve header row dynamically
    const headerRow = getHeaderRow(sheet);
    const formulaCell = sheet.getRange("G2");

    // Skip updates if formula is present in "Automated Notes" column
    if (formulaCell.getFormula()) {
      Logger.log(`[${traceId}] Formula detected in column G for sheet "${sheetName}". Skipping update.`);
      return;
    }

    // Validate presence of required headers
    const locationIndex = headerRow.indexOf("Location");
    const serialIndex = headerRow.indexOf("Serial");
    const assetIndex = headerRow.indexOf("Asset Tag");
    const hostIndex = headerRow.indexOf("Hostname");

    if (
      locationIndex === -1 ||
      serialIndex === -1 ||
      assetIndex === -1 ||
      hostIndex === -1
    ) {
      Logger.log(`[${traceId}] Missing critical headers in sheet "${sheetName}".`);
      return;
    }

    // Process rows and populate notes
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, headerRow.length);
    const data = dataRange.getValues();

    const notesColumn = sheet.getRange(2, headerRow.indexOf("Automated Notes") + 1, sheet.getLastRow() - 1);
    const notesValues = notesColumn.getValues();

    data.forEach((row, i) => {
      const location = row[locationIndex];
      const serial = row[serialIndex];
      const asset = row[assetIndex];
      const host = row[hostIndex];

      if (!location && !serial && !asset && !host) {
        notesValues[i][0] = ""; // Clear notes for empty rows
      } else if (location && serial && asset && host) {
        notesValues[i][0] = "Complete"; // All tuple components present
      } else if (serial && asset && host) {
        notesValues[i][0] = "Sufficient: Missing Loc"; // Missing location
      } else {
        notesValues[i][0] = "Missing: " +
          (location ? "" : "Loc") +
          (serial ? "" : ", Serial") +
          (asset ? "" : ", Asset") +
          (host ? "" : ", Host");
      }
    });

    notesColumn.setValues(notesValues);
    Logger.log(`[${traceId}] Updated Automated Notes in sheet "${sheetName}".`);
  });

  endSpan(traceId, "Update Automated Notes");
}

/**
 * Retrieves the header row dynamically.
 * @param {Sheet} sheet - The sheet to retrieve headers from.
 * @returns {string[]} Array of header names.
 */
function getHeaderRow(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  return range.getValues()[0];
}

/**
 * Checks if a sheet should be skipped from protection.
 * @param {string} sheetName - Name of the sheet.
 * @param {number} totalSheets - Total number of sheets in the spreadsheet.
 * @param {number} index - Index of the sheet in the list.
 * @returns {boolean} True if the sheet should be skipped, false otherwise.
 */
function isProtectedSheet(sheetName, totalSheets, index) {
  return sheetName === "Count" || index === totalSheets - 2;
}

/**
 * Finds existing protections for a range.
 * @param {Sheet} sheet - The sheet to check protections in.
 * @param {Range} range - The range to check protections for.
 * @returns {Protection|null} The existing protection if found, otherwise null.
 */
function getExistingProtection(sheet, range) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  return protections.find(p => p.getRange().getA1Notation() === range.getA1Notation());
}
