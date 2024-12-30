/**
 * Monitor Tracker Script (Monitor_Trackerv1.0.4.gs)
 * Version: 1.0.4
 * Features:
 * - Dynamic insertion of the "Automated Notes" header if missing.
 * - Verification and application of the correct formula under "Automated Notes."
 * - Dual protection settings: owner-only and warning-only for the column.
 * - Comprehensive logging for debugging and auditing.
 */

// Entry point for the script
function myFunction() {
  const traceId = generateTraceId();
  startSpan(traceId, "Main Execution");

  try {
    verifyAndInsertHeader(traceId);
    verifyAndApplyFormula(traceId);
    applyProtections(traceId);
  } catch (error) {
    logError(traceId, "Main Execution", error);
  }

  endSpan(traceId, "Main Execution");
}

/**
 * Generates a unique trace ID for tracking execution logs.
 */
function generateTraceId() {
  return Math.random().toString(36).substr(2, 9);
}

/**
 * Logs the start of an execution span.
 */
function startSpan(traceId, spanName) {
  Logger.log(`[${traceId}] Starting span: ${spanName}`);
}

/**
 * Logs the end of an execution span.
 */
function endSpan(traceId, spanName) {
  Logger.log(`[${traceId}] Ending span: ${spanName}`);
}

/**
 * Logs errors during execution.
 */
function logError(traceId, context, error) {
  Logger.log(`[${traceId}] Error in context "${context}": ${error.message}`);
}

/**
 * Verifies if the "Automated Notes" header exists and inserts it if missing.
 */
function verifyAndInsertHeader(traceId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName === "Count") {
      Logger.log(`[${traceId}] Skipping header verification for sheet "${sheetName}".`);
      return;
    }

    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    const headers = headerRange.getValues()[0];
    const headerColumnIndex = headers.indexOf("Automated Notes");

    if (headerColumnIndex === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue("Automated Notes");
      Logger.log(`[${traceId}] Added "Automated Notes" header to sheet "${sheetName}".`);
    } else {
      Logger.log(`[${traceId}] "Automated Notes" header already exists in sheet "${sheetName}".`);
    }
  });
}

/**
 * Verifies if the formula is correctly applied under "Automated Notes" and applies it if missing.
 */
function verifyAndApplyFormula(traceId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const formula = `=ARRAYFORMULA(
    IF(
      ROW(A2:A) = 1,
      "Automated Notes",
      IF(
        (INDEX(A2:Z, 0, MATCH("Serial", A1:Z1, 0)) <> "") *
        (INDEX(A2:Z, 0, MATCH("Asset Tag", A1:Z1, 0)) <> "") *
        (INDEX(A2:Z, 0, MATCH("Hostname", A1:Z1, 0)) <> "") *
        (INDEX(A2:Z, 0, MATCH("Location", A1:Z1, 0)) = ""),
        "Sufficient: Missing Loc",
        IF(
          (INDEX(A2:Z, 0, MATCH("Location", A1:Z1, 0)) <> "") *
          (INDEX(A2:Z, 0, MATCH("Serial", A1:Z1, 0)) = "") *
          (INDEX(A2:Z, 0, MATCH("Asset Tag", A1:Z1, 0)) = "") *
          (INDEX(A2:Z, 0, MATCH("Hostname", A1:Z1, 0)) = ""),
          "Missing: Deployment",
          IF(
            (INDEX(A2:Z, 0, MATCH("Location", A1:Z1, 0)) = "") *
            (INDEX(A2:Z, 0, MATCH("Serial", A1:Z1, 0)) = "") *
            (INDEX(A2:Z, 0, MATCH("Asset Tag", A1:Z1, 0)) = "") *
            (INDEX(A2:Z, 0, MATCH("Hostname", A1:Z1, 0)) = ""),
            "",
            IF(
              (INDEX(A2:Z, 0, MATCH("Location", A1:Z1, 0)) <> "") *
              (INDEX(A2:Z, 0, MATCH("Serial", A1:Z1, 0)) <> "") *
              (INDEX(A2:Z, 0, MATCH("Asset Tag", A1:Z1, 0)) <> "") *
              (INDEX(A2:Z, 0, MATCH("Hostname", A1:Z1, 0)) <> ""),
              "Complete",
              "Missing: " &
              IF(INDEX(A2:Z, 0, MATCH("Location", A1:Z1, 0)) = "", "Loc", "") &
              IF(INDEX(A2:Z, 0, MATCH("Serial", A1:Z1, 0)) = "", ", Serial", "") &
              IF(INDEX(A2:Z, 0, MATCH("Asset Tag", A1:Z1, 0)) = "", ", Asset", "") &
              IF(INDEX(A2:Z, 0, MATCH("Hostname", A1:Z1, 0)) = "", ", Host", "")
            )
          )
        )
      )
    )
  )`;

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName === "Count") {
      Logger.log(`[${traceId}] Skipping formula verification for sheet "${sheetName}".`);
      return;
    }

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = headerRow.indexOf("Automated Notes") + 1;

    if (colIndex === 0) {
      Logger.log(`[${traceId}] "Automated Notes" header is missing in sheet "${sheetName}". Skipping formula application.`);
      return;
    }

    const formulaCell = sheet.getRange(2, colIndex);
    if (!formulaCell.getFormula().includes("ARRAYFORMULA")) {
      formulaCell.setFormula(formula);
      Logger.log(`[${traceId}] Formula applied in column "${colIndex}" of sheet "${sheetName}".`);
    } else {
      Logger.log(`[${traceId}] Formula already exists in column "${colIndex}" of sheet "${sheetName}".`);
    }
  });
}

/**
 * Applies protections: owner-only and warning-only, to the "Automated Notes" column.
 */
function applyProtections(traceId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const owner = spreadsheet.getOwner().getEmail();
  const sheets = spreadsheet.getSheets();

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName === "Count") {
      Logger.log(`[${traceId}] Skipping protections for sheet "${sheetName}".`);
      return;
    }

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colIndex = headerRow.indexOf("Automated Notes") + 1;

    if (colIndex === 0) {
      Logger.log(`[${traceId}] "Automated Notes" header missing in sheet "${sheetName}". Skipping protections.`);
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

      // Apply warning-only protection
      protection.setWarningOnly(true);
      Logger.log(`[${traceId}] Warning-only protection applied to sheet "${sheetName}", column "${colIndex}".`);
    } else {
      Logger.log(`[${traceId}] Protection already exists for sheet "${sheetName}", column "${colIndex}".`);
    }
  });
}
