function protectAllExceptCount() {
  const traceId = generateTraceId();
  startSpan(traceId, "Protect All Except 'Count'");
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const excludedTab = "Count"; // Tab to exclude from protection
    const sheets = spreadsheet.getSheets();

    sheets.forEach(sheet => {
      const tabName = sheet.getName();

      // Skip the excluded tab
      if (tabName === excludedTab) {
        Logger.log(`[${traceId}] Skipping protection for sheet '${tabName}'.`);
        return;
      }

      // Check if a protection already exists for the entire sheet
      const existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      if (existingProtections.length > 0) {
        Logger.log(`[${traceId}] Protection already exists for sheet '${tabName}'. Skipping.`);
        return;
      }

      // Apply sheet-level protection
      const protection = sheet.protect();
      protection.setDescription(`Protected sheet: ${tabName}`);
      Logger.log(`[${traceId}] Protection applied to sheet '${tabName}'.`);
    });
  } catch (error) {
    logError(traceId, "Protect All Except 'Count'", error);
  }
  
  endSpan(traceId, "Protect All Except 'Count'");
}

// Generates a unique trace ID for tracking execution logs
function generateTraceId() {
  return Math.random().toString(36).substr(2, 9);
}

// Logs the start of an execution span
function startSpan(traceId, spanName) {
  Logger.log(`[${traceId}] Starting span: ${spanName}`);
}

// Logs the end of an execution span
function endSpan(traceId, spanName) {
  Logger.log(`[${traceId}] Ending span: ${spanName}`);
}

// Logs errors during execution
function logError(traceId, context, error) {
  Logger.log(`[${traceId}] Error in context "${context}": ${error.message}`);
}
