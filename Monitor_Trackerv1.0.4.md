Here’s the **Markdown (MD)** documentation for **Monitor_Trackerv1.0.4.gs**:

---

# Monitor Tracker Script (Monitor_Trackerv1.0.4.gs)

### **Version**
- **1.0.4**

---

## **Overview**
This script automates the management of monitor tracking spreadsheets, ensuring proper setup, consistent formulas, and robust protections for sensitive data columns. 

---

## **Features**
1. **Dynamic Header Insertion**
   - Automatically adds the "Automated Notes" header if it is missing.

2. **Formula Verification**
   - Ensures the correct formula is applied under the "Automated Notes" header if missing.

3. **Protection Settings**
   - Applies **owner-only protection** to the "Automated Notes" column to ensure data integrity.
   - Adds **warning-only protection** to prevent accidental edits.

4. **Comprehensive Logging**
   - Logs all actions and potential issues for traceability and debugging.

---

## **Usage**
1. **Run the Script**
   - Execute the function `myFunction` from the Apps Script Editor.
2. **View Logs**
   - Open the Apps Script Execution Log to debug or review operations.

---

## **Notes**
- The script is modular for easy updates and maintenance.
- Designed for deployment in Google Sheets as part of monitor tracking operations.
- Compatible with sheets that include a column for "Automated Notes."

---

## **How It Works**
1. **Dynamic Header Verification**
   - Checks if the "Automated Notes" header exists. If not, it adds the header to the next available column.

2. **Formula Application**
   - Applies the formula for "Automated Notes" dynamically if it isn't already present.

3. **Protections**
   - Ensures owner-only access for editing the "Automated Notes" column.
   - Adds a warning-only mode for the owner to avoid accidental overwrites.

4. **Logs Execution**
   - Provides detailed logs for every step, including errors or missing headers.

---

## **Functionality Flow**
1. **Header Verification**
   - Adds missing headers.
2. **Formula Application**
   - Inserts or verifies the `ARRAYFORMULA` for automated notes.
3. **Protection Management**
   - Applies owner-only and warning-only protections.
4. **Logs and Debugging**
   - Captures all execution details for easy troubleshooting.

---

## **Requirements**
- **Google Sheets**: The spreadsheet must include relevant headers like "Serial," "Asset Tag," "Hostname," and "Location."
- **Apps Script Editor**: Accessible via Google Workspace.

---

## **Changelog**
### **Version 1.0.4**
- Added:
  - Dynamic header insertion for "Automated Notes."
  - Dual protection (owner-only and warning-only) for sensitive columns.
  - Comprehensive logging for easier debugging and traceability.

- Enhanced:
  - Formula verification logic for dynamic adaptability.

---

## **Contact**
For questions or feedback, contact the developer:
- **Developer**: Richard Perez  
- **Email**: [sage.holly7825@eagereverest.com](mailto:sage.holly7825@eagereverest.com)

--- 

This documentation provides all the necessary information for deploying, maintaining, and troubleshooting the script.
