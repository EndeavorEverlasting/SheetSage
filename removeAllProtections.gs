function removeAllProtections() {
  const traceId = generateTraceId();
  startSpan(traceId, "Remove All Protections");

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();

    Logger.log(`[${traceId}] Starting protection removal process for all sheets.`);

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      Logger.log(`[${traceId}] Processing sheet: "${sheetName}".`);

      // Remove range protections
      const rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      rangeProtections.forEach(protection => {
        try {
          Logger.log(`[${traceId}] Found range protection on "${sheetName}" for range: "${protection.getRange().getA1Notation()}".`);
          protection.remove();
          Logger.log(`[${traceId}] Removed range protection on "${sheetName}" for range: "${protection.getRange().getA1Notation()}".`);
        } catch (error) {
          Logger.log(`[${traceId}] Error removing range protection on "${sheetName}": ${error.message}`);
        }
      });

      // Remove sheet protections
      const sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      sheetProtections.forEach(protection => {
        try {
          Logger.log(`[${traceId}] Found sheet protection on "${sheetName}".`);
          protection.remove();
          Logger.log(`[${traceId}] Removed sheet protection on "${sheetName}".`);
        } catch (error) {
          Logger.log(`[${traceId}] Error removing sheet protection on "${sheetName}": ${error.message}`);
        }
      });

      Logger.log(`[${traceId}] Completed processing for sheet "${sheetName}".`);
    });

    Logger.log(`[${traceId}] Protection removal process completed for all sheets.`);
  } catch (error) {
    logError(traceId, "Remove All Protections", error);
  }

  endSpan(traceId, "Remove All Protections");
}

// Utility functions for logging
function generateTraceId() {
  return Math.random().toString(36).substr(2, 9);
}

function startSpan(traceId, spanName) {
  Logger.log(`[${traceId}] Starting span: ${spanName}`);
}

function endSpan(traceId, spanName) {
  Logger.log(`[${traceId}] Ending span: ${spanName}`);
}

function logError(traceId, context, error) {
  Logger.log(`[${traceId}] Error in context "${context}": ${error.message}`);
}
