# Monitor_Trackerv1.0.3.gs

## Description
This script automates several tasks for managing Google Sheets used in tracking monitor deployments. It ensures data integrity, applies necessary protections, and provides robust logging for traceability. The main features include:

1. **Dynamic Header Insertion**: Ensures the "Automated Notes" header exists in all tabs except "Count".
2. **Formula Verification**: Verifies and inserts a live formula under the "Automated Notes" header if missing.
3. **Comprehensive Logging**: Logs all operations for debugging and traceability.
4. **Dual Protection**: Applies both owner-only protection and warning protection to the "Automated Notes" column.

## Functions
### `myFunction`
Executes all major tasks:
- Preserves headers.
- Ensures the formula is correctly inserted in the "Automated Notes" column.
- Applies protections to the "Automated Notes" column.
- Logs all existing protections for review.

### `generateTraceId`
Generates a unique trace ID for logging.

### `startSpan`
Logs the start of an execution span for better traceability.

### `endSpan`
Logs the end of an execution span.

### `logError`
Logs errors with the trace ID and context.

### `preserveHeaders`
- Checks for the "Automated Notes" header in all sheets except "Count".
- Dynamically adds the header if missing.

### `ensureFormulaInColumnG`
- Verifies that the live formula is correctly inserted in the "Automated Notes" column (Column G).
- Inserts the formula if missing or incorrect.

### `applyProtectionToAutomatedNotes`
- Applies **owner-only** protection to the "Automated Notes" column.
- Sets **warning-only** mode for the column to notify users of changes.

### `checkExistingProtections`
Logs all existing protections for debugging purposes.

## Formula Used
The script ensures the following formula is applied under "Automated Notes":

```plaintext
=ARRAYFORMULA(
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
)
