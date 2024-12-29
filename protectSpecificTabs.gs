function protectSpecificTabs() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const tabsToProtect = ["12/24", "12/25", "12/26", "12/27"]; // List of specific tabs to protect
  const ownerEmail = spreadsheet.getOwner().getEmail(); // Spreadsheet owner's email

  tabsToProtect.forEach(tabName => {
    const sheet = spreadsheet.getSheetByName(tabName);
    if (!sheet) {
      Logger.log(`Sheet "${tabName}" not found. Skipping.`);
      return;
    }

    // Check if a protection already exists for the entire sheet
    const existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (existingProtections.length > 0) {
      Logger.log(`Protection already exists for sheet "${tabName}". Skipping.`);
      return;
    }

    // Apply sheet-level protection
    const protection = sheet.protect();
    protection.setDescription(`Protected sheet: ${tabName}`);
    protection.addEditor(ownerEmail); // Allow the owner to edit
    protection.removeEditors(protection.getEditors().filter(editor => editor !== ownerEmail)); // Remove other editors

    // Log success
    Logger.log(`Protection applied to sheet "${tabName}".`);
  });
}
