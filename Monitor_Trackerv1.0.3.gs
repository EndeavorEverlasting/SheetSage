/**
 * Monitor_Trackerv1.0.4.gs
 * This script automates the following tasks for a Google Sheet:
 * 1. Ensures the "Automated Notes" header exists in all tabs except "Count".
 * 2. Verifies and inserts a formula in the "Automated Notes" column if missing.
 * 3. Applies both owner-only and warning protections to the "Automated Notes" column.
 * 4. Logs all steps, errors, and decisions for debugging and traceability.
 */

// Main function to execute all tasks
function myFunction() {
  const traceId = generateTraceId();
  startSpan(traceId, "Main Execution");

  try {
    preserveHeaders(traceId); // Add missing "Automated Notes" headers dynamically.
    ensureFormulaInColumnG(traceId); // Ensure the formula exists in Column G.
    applyProtectionToAutomatedNotes(traceId); // Apply owner-only and warning protections.
    checkExistingProtections(traceId); // Log existing protections for verification.
  } catch (error) {
    logError(traceId, "Main Execution", error); // Log errors for debugging.
  }

  endSpan(traceId, "Main Execution");
}

/**
 * Generates a unique trace ID for logging purposes.
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
 * Ensures headers exist in all sheets except "Count".
 * Dynamically adds the "Automated Notes" header if missing.
 */
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
    if (!headers.includes("Automated Notes")) {
      sheet.getRange(1, headers.length + 1).setValue("Automated Notes");
      Logger.log(`[${traceId}] Added missing "Automated Notes" header in sheet "${sheetName}".`);
    } else {
      Logger.log(`[${traceId}] "Automated Notes" header exists in sheet "${sheetName}".`);
    }
  });
}

/**
 * Ensures the formula is correctly inserted in the "Automated Notes" column (Column G).
 * Dynamically verifies if the formula exists; adds it if missing.
 */
function ensureFormulaInColumnG(traceId) {
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

    const formulaCell = sheet.getRange("G2");
    if (!formulaCell.getFormula() || formulaCell.getFormula() !== formula) {
      formulaCell.setFormula(formula);
      Logger.log(`[${traceId}] Formula added to "G2" in sheet "${sheetName}".`);
    } else {
      Logger.log(`[${traceId}] Formula already exists in "G2" of sheet "${sheetName}".`);
    }
  });
}

/**
 * Applies both owner-only and warning protections to the "Automated Notes" column.
 */
function applyProtectionToAutomatedNotes(traceId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ownerEmail = spreadsheet.getOwner().getEmail();
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
      protection.setDescription(`Protected "Automated Notes" column in sheet "${sheetName}".`);

      // Set owner access only
      protection.addEditor(ownerEmail);
      protection.removeEditors(protection.getEditors().filter(editor => editor !== ownerEmail));

      // Enable warning mode for the owner
      protection.setWarningOnly(true);

      Logger.log(`[${traceId}] Owner-only and warning protection applied to sheet "${sheetName}", column "${colIndex}".`);
    } else {
      Logger.log(`[${traceId}] Protection already exists for sheet "${sheetName}", column "${colIndex}". Skipping.`);
    }
  });
}

/**
 * Logs all existing protections in the spreadsheet for debugging.
 */
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
