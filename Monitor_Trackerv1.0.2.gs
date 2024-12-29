function myFunction() {
  const traceId = generateTraceId();
  startSpan(traceId, "Main Execution");

  try {
    preserveHeaders(traceId);
    applyWarningOnlyProtection(traceId);
    applyOwnerOnlyProtection(traceId);
    applyFormulaToColumnG(traceId);
    checkExistingProtections(traceId);
  } catch (error) {
    logError(traceId, "Main Execution", error);
  }

  endSpan(traceId, "Main Execution");
}

// Generates a unique trace ID for tracking execution logs.
function generateTraceId() {
  return Math.random().toString(36).substr(2, 9);
}

// Logs the start of an execution span.
function startSpan(traceId, spanName) {
  Logger.log(`[${traceId}] Starting span: ${spanName}`);
}

// Logs the end of an execution span.
function endSpan(traceId, spanName) {
  Logger.log(`[${traceId}] Ending span: ${spanName}`);
}

// Logs errors during execution.
function logError(traceId, context, error) {
  Logger.log(`[${traceId}] Error in context "${context}": ${error.message}`);
}

// Ensures headers exist in all sheets except "Count".
function preserveHeaders(traceId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName === "Count") {
      Logger.log(`[${traceId}] Skipping header verification for sheet "${sheetName}".`);
      return;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.includes("")) {
      throw new Error(`[${traceId}] Headers are missing in sheet "${sheetName}". Please verify.`);
    }

    Logger.log(`[${traceId}] Headers verified for sheet "${sheetName}".`);
  });
}

// Applies warning-only protection to "Automated Notes" column.
function applyWarningOnlyProtection(traceId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName === "Count") {
      Logger.log(`[${traceId}] Skipping protection for sheet "${sheetName}".`);
      return;
    }

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = headerRow.indexOf("Automated Notes") + 1;
    if (colIndex === 0) {
      Logger.log(`[${traceId}] "Automated Notes" column not found in sheet "${sheetName}". Skipping.`);
      return;
    }

    const range = sheet.getRange(1, colIndex, sheet.getMaxRows());
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    const existingProtection = protections.find(protection => protection.getRange().getA1Notation() === range.getA1Notation());

    if (!existingProtection) {
      const protection = range.protect();
      protection.setWarningOnly(true);
      Logger.log(`[${traceId}] Warning-only protection added to sheet "${sheetName}", column "${colIndex}".`);
    } else {
      Logger.log(`[${traceId}] Warning-only protection already exists for sheet "${sheetName}", column "${colIndex}".`);
    }
  });
}

// Applies owner-only protection to "Automated Notes" column.
function applyOwnerOnlyProtection(traceId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const owner = spreadsheet.getOwner().getEmail();
  const sheets = spreadsheet.getSheets();

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName === "Count") {
      Logger.log(`[${traceId}] Skipping owner-only protection for sheet "${sheetName}".`);
      return;
    }

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = headerRow.indexOf("Automated Notes") + 1;
    if (colIndex === 0) {
      Logger.log(`[${traceId}] "Automated Notes" column not found in sheet "${sheetName}". Skipping.`);
      return;
    }

    const range = sheet.getRange(1, colIndex, sheet.getMaxRows());
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    const existingProtection = protections.find(protection => protection.getRange().getA1Notation() === range.getA1Notation());

    if (!existingProtection) {
      const protection = range.protect();
      protection.addEditor(owner);
      protection.removeEditors(protection.getEditors().filter(editor => editor !== owner));
      Logger.log(`[${traceId}] Owner-only protection added to sheet "${sheetName}", column "${colIndex}".`);
    } else {
      Logger.log(`[${traceId}] Owner-only protection already exists for sheet "${sheetName}", column "${colIndex}".`);
    }
  });
}

// Checks existing protections for logging.
function checkExistingProtections(traceId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  sheets.forEach(sheet => {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(protection => {
      Logger.log(`[${traceId}] Protection found in sheet "${sheet.getName()}": Range "${protection.getRange().getA1Notation()}", Description: "${protection.getDescription()}".`);
    });
  });
}

// Applies the formula to "Automated Notes" column G only if it is missing.
function applyFormulaToColumnG(traceId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName === "Count") {
      Logger.log(`[${traceId}] Skipping formula application for sheet "${sheetName}".`);
      return;
    }

    const formulaCell = sheet.getRange("G2");
    if (formulaCell.getFormula()) {
      Logger.log(`[${traceId}] Formula already exists in "G2" of sheet "${sheetName}". Skipping application.`);
      return;
    }

    const formula = `=ARRAYFORMULA(... /* Your formula here */)`;
    formulaCell.setFormula(formula);
    Logger.log(`[${traceId}] Formula applied to column G in sheet "${sheetName}".`);
  });
}
