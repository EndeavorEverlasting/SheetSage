function myFunction() {
  // Main function to call the protectAutomatedNotesColumn function
  protectAutomatedNotesColumn();
  updateAutomatedNotes();
}

function protectAutomatedNotesColumn() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // Get the active spreadsheet
  const sheets = spreadsheet.getSheets(); // Get all sheets in the spreadsheet
  const columnToProtect = "G"; // Specify the column to protect (adjust if needed)
  const titleToMatch = "Automated Notes"; // The title to search for in the sheet

  Logger.log("Starting protection process for all sheets..."); // Log the start of the process

  sheets.forEach(sheet => {
    const sheetName = sheet.getName(); // Get the name of the current sheet
    Logger.log(`Processing sheet: "${sheetName}"`); // Log the name of the sheet being processed

    const lastRow = sheet.getLastRow(); // Get the last row with content in the sheet
    if (lastRow === 0) {
      // Skip the sheet if it is empty
      Logger.log(`Sheet "${sheetName}" is empty. Skipping.`);
      return;
    }

    const lastColumn = sheet.getLastColumn(); // Get the last column with content in the sheet
    Logger.log(`Sheet "${sheetName}" has ${lastRow} rows and ${lastColumn} columns.`); // Log row and column count

    const rangeToSearch = sheet.getRange(1, 1, lastRow, lastColumn); // Get the range of all rows and columns
    const values = rangeToSearch.getValues(); // Get the data as a 2D array
    
    let rowIndex = -1; // Initialize the row index to -1 (indicating not found)
    for (let i = 0; i < values.length; i++) {
      if (values[i].includes(titleToMatch)) {
        rowIndex = i + 1; // Convert to 1-based index if the title is found
        Logger.log(`Found title "${titleToMatch}" in sheet "${sheetName}" at row ${rowIndex}.`);
        break; // Exit the loop once the title is found
      }
    }

    if (rowIndex !== -1) { // Proceed if the title was found
      const columnRange = sheet.getRange(`${columnToProtect}1:${columnToProtect}${lastRow}`); // Get the range of the column to protect
      let protection;

      // Check if protection already exists for the range
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE); // Get all range protections in the sheet
      const existingProtection = protections.find(p => p.getRange().getA1Notation() === columnRange.getA1Notation());
      
      if (existingProtection) {
        protection = existingProtection; // Use the existing protection if it exists
        Logger.log(`Existing protection found for column "${columnToProtect}" in sheet "${sheetName}".`);
      } else {
        protection = columnRange.protect(); // Create a new protection if none exists
        Logger.log(`Created new protection for column "${columnToProtect}" in sheet "${sheetName}".`);
      }
      
      // Update protection settings
      protection.setDescription(`Protected column ${columnToProtect} for Automated Notes`); // Set a description for the protection
      protection.removeEditors(protection.getEditors()); // Remove all editors, leaving only the script owner
      Logger.log(`Updated protection settings for column "${columnToProtect}" in sheet "${sheetName}".`);
    } else {
      // Log if the title was not found in the current sheet
      Logger.log(`Title "${titleToMatch}" not found in sheet "${sheetName}". Skipping column protection.`);
    }
  });

  Logger.log("Protection process completed for all sheets."); // Log the end of the process
}// End of protectAutomatedNotesColumn()

function updateAutomatedNotes() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const automatedNotesColumn = "G" 

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    Logger.log(`Updating Automated Notes for sheet: "${sheetName}"`);

    const lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      Logger.log(`Sheet "${sheetName}" is empty. Skipping.`);
      return;
    }

    const locationColumn = 1; //Column A
    const serialColumn = 2; // Column B
    const assetColumn = 3; // Column C
    const hostColumn = 4; // Column D

    const notesColumn = sheet.getRange(`${automatedNotesColumn}1:${automatedNotesColumn}${lastRow}`);
    const notesValues = notesColumn.getValues();

    const dataRange = sheet.getRange(1, 1, lastRow, 4); // First 4 columns (Location, Serial, Asset Tag, Hostname)
    const data = dataRange.getValues();

    for (let i = 0; i < data.length; i++) {
      const location = data[i][locationColumn - 1];
      const serial = data[i][serialColumn - 1];
      const asset = data[i][assetColumn - 1];
      const host = data[i][hostColumn - 1];

      // Determine the status based on available tuple components
      if (serial && asset && host && location) {
        notesValues[i][0] = "Complete"; // All four components are present
      } else if (serial && asset && host) {
        notesValues[i][0] = "Sufficient"; // Serial, Asset, and Host are present
      } else {
        notesValues[i][0] = "Missing: Loc, Host"; // Specify missing components
      }
    }// End of for loop

    // Update the notes column with new values
    notesColumn.setValues(notesValues);
    Logger.log(`Automated Notes updated for sheet "${sheetName}".`);      
  }); // End of arrow function

  Logger.log("Automated Notes update completed for all sheets.");
} // End of updateAutomatedNotes
