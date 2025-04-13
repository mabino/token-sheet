/**
 *
 * This script automatically generates a UUID in the first column (A) of any sheet
 * if that column's header is "UUID" (case-insensitive) and the row contains data
 * in other columns. The UUID is prefixed with "uuid-".
 * It uses both onEdit and an installable onChange trigger for robustness.
 * 
 * Install the onChange trigger by running the createOnChangeTrigger once.
 * 
 * Includes a locking mechanism to prevent concurrent modifications and rate limiting.
 */

// ========= CONSTANTS =========

/**
 * The header text expected in the first column to trigger UUID generation.
 * Comparison is case-insensitive.
 */
const UUID_HEADER_TEXT_GEN = "UUID"; 

/**
 * The prefix added to each generated UUID.
 */
const UUID_PREFIX_GEN = "uuid-"; 

/**
 * Minimum time in milliseconds between function executions for rate limiting.
 * Helps avoid hitting Google's quotas.
 */
const RATE_LIMIT_MS = 500;

/**
 * Script property key used for timestamp-based rate limiting.
 */
const LAST_EXECUTION_PROP = "lastExecutionTimestamp";

/**
 * Lock timeout in milliseconds. Maximum time to wait for acquiring a lock.
 */
const LOCK_TIMEOUT_MS = 5000;

// ========= TRIGGER HANDLERS =========

/**
 * Simple trigger that runs automatically when a user changes the value of any cell
 * in the spreadsheet.
 * @param {Object} e The event parameter for a simple onEdit trigger.
 */
function onEdit(e) {
  // Ensure the event object and range property exist.
  if (!e || !e.range) {
    Logger.log("onEdit event object or range property is missing.");
    return;
  }

  // Avoid acting on edits within the UUID column itself
  if (e.range.getColumn() === 1) {
    return;
  }

  const sheet = e.range.getSheet();
  
  // Implement rate limiting and locking
  if (respectRateLimitAndLock()) {
    try {
      // Call the main logic function for the sheet where the edit occurred.
      generateUuidsIfNeeded(sheet);
    } catch (error) {
      Logger.log(`Error in onEdit: ${error.message}`);
    } finally {
      releaseLock();
    }
  }
}

/**
 * An installable trigger handler function. Runs on broader changes.
 * Needs to be manually installed via the `createOnChangeTrigger` function.
 * @param {Object} e The event parameter for an installable onChange trigger.
 */
function onChangeHandler(e) {
  Logger.log(`onChangeHandler triggered by change type: ${e.changeType}`);
  
  // Implement rate limiting and locking
  if (respectRateLimitAndLock()) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const allSheets = ss.getSheets();
      allSheets.forEach(sheet => {
        generateUuidsIfNeeded(sheet);
      });
    } catch (error) {
      Logger.log(`Error in onChangeHandler: ${error.message}`);
    } finally {
      releaseLock();
    }
  }
}

// ========= MAIN UUID GENERATION LOGIC =========

/**
 * Checks a specific sheet for rows that require a UUID and generates them.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to process.
 */
function generateUuidsIfNeeded(sheet) {
  // Get the header value from cell A1
  const headerRange = sheet.getRange(1, 1);
  const headerValue = headerRange.getValue();

  // Check if the header matches (case-insensitive)
  if (typeof headerValue !== 'string' || headerValue.trim().toUpperCase() !== UUID_HEADER_TEXT_GEN.toUpperCase()) {
    return; // Not the sheet we're interested in
  }

  // Get the full data range of the sheet
  const dataRange = sheet.getDataRange();
  // Check if there's more than just the header row
  if (dataRange.getNumRows() <= 1) {
    return; // Nothing to process
  }

  const values = dataRange.getValues(); // Includes header
  let updatesMade = 0;

  // Iterate through rows, starting from the second row (index 1)
  for (let i = 1; i < values.length; i++) {
    const rowData = values[i];
    const uuidCell = rowData[0]; // Value in Column A (UUID column)
    const rowNumber = i + 1; // Sheet row number (1-based index)

    // Condition 1: UUID cell must be empty
    if (uuidCell === '' || uuidCell === null) {
      // Condition 2: Check if *any other* cell in the row has data
      let otherDataExists = false;
      for (let j = 1; j < rowData.length; j++) { // Start checking from column B (index 1)
        if (rowData[j] !== '' && rowData[j] !== null) {
          otherDataExists = true;
          break; // Found data, no need to check further in this row
        }
      }

      // If both conditions met, generate and set the UUID
      if (otherDataExists) {
        const newUuid = UUID_PREFIX_GEN + Utilities.getUuid();
        sheet.getRange(rowNumber, 1).setValue(newUuid); // Set value in Column A
        updatesMade++;
      }
    }
  }
  if (updatesMade > 0) {
    Logger.log(`Sheet "${sheet.getName()}": Generated ${updatesMade} UUID(s).`);
  }
}

// ========= LOCKING AND RATE LIMITING =========

/**
 * Checks rate limiting and acquires a lock if allowed.
 * @return {boolean} True if rate limit is respected and lock is acquired, false otherwise.
 */
function respectRateLimitAndLock() {
  // Check rate limit first
  const props = PropertiesService.getScriptProperties();
  const lastExecution = props.getProperty(LAST_EXECUTION_PROP);
  const now = Date.now();
  
  if (lastExecution && (now - parseInt(lastExecution)) < RATE_LIMIT_MS) {
    Logger.log(`Rate limit reached. Last execution: ${new Date(parseInt(lastExecution))}`);
    return false;
  }
  
  // Update timestamp regardless of lock acquisition to enforce rate limiting
  props.setProperty(LAST_EXECUTION_PROP, now.toString());
  
  // Try to acquire lock
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(LOCK_TIMEOUT_MS);
    return true;
  } catch (e) {
    Logger.log(`Could not acquire lock: ${e.message}`);
    return false;
  }
}

/**
 * Releases the script lock
 */
function releaseLock() {
  try {
    LockService.getScriptLock().releaseLock();
  } catch (e) {
    Logger.log(`Error releasing lock: ${e.message}`);
  }
}

/**
 * Creates an installable onChange trigger for the spreadsheet
 * if one doesn't already exist for the onChangeHandler function.
 * Run this function manually once from the script editor.
 */
function createOnChangeTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getUserTriggers(ss);
  let triggerExists = false;

  triggers.forEach(trigger => {
    if (trigger.getEventType() === ScriptApp.EventType.ON_CHANGE &&
        trigger.getHandlerFunction() === 'onChangeHandler') {
      triggerExists = true;
    }
  });

  if (!triggerExists) {
    try {
      ScriptApp.newTrigger('onChangeHandler')
        .forSpreadsheet(ss)
        .onChange()
        .create();
      Logger.log("Successfully created installable onChange trigger for 'onChangeHandler'.");
    } catch (err) {
       Logger.log(`Error creating onChange trigger: ${err}`);
    }
  } else {
     Logger.log("onChange trigger already exists.");
  }
}
