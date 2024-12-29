function protectAllExceptCount() {
  const traceId = generateTraceId(); // Generate unique ID for traceability
  startSpan(traceId, "Main Execution");

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const excludedTab = "Count"; // Tab to exclude from protection
    const sheets = spreadsheet.getSheets();

    sheets.forEach(sheet => {
      const tabName = sheet.getName();

      // Skip the excluded tab
      if (tabName === excludedTab) {
        Logger.log(`[${traceId}] Skipping protection for sheet "${tabName}".`);
        return;
      }

      try {
        // Check if a protection already exists for the entire sheet
        const existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        if (existingProtections.length > 0) {
          Logger.log(`[${traceId}] Protection already exists for sheet "${tabName}". Skipping.`);
          return;
        }

        // Apply sheet-level protection
        const protection = sheet.protect();
        protection.setDescription(`Protected sheet: ${tabName}`);
        Logger.log(`[${traceId}] Protection applied to sheet "${tabName}".`);

      } catch (error) {
        // Log specific errors related to applying protection
        Logger.log(`[${traceId}] Failed to apply protection to sheet "${tabName}". Error: ${error.message}`);
      }
    });
  } catch (error) {
    // Log any unexpected errors during script execution
    Logger.log(`[${traceId}] Unexpected error in Main Execution: ${error.message}`);
  }

  endSpan(traceId, "Main Execution");
}

// Helper functions for enhanced logging
function generateTraceId() {
  return Math.random().toString(36).substr(2, 9);
}

function startSpan(traceId, spanName) {
  Logger.log(`[${traceId}] Starting span: ${spanName}`);
}

function endSpan(traceId, spanName) {
  Logger.log(`[${traceId}] Ending span: ${spanName}`);
}
