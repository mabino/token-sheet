/**
 * @OnlyCurrentDoc
 *
 * This script provides a web app endpoint (doPost) to update rows based on UUID.
 * It requires a valid SECRET_TOKEN stored in Script Properties and passed
 * in the JSON payload as 'authToken'.
 * 
 * Includes locking mechanism to prevent concurrent modifications and rate limiting.
 */

/**
 * The header text expected in the first column for UUID lookup.
 * Comparison is case-insensitive.
 */
const UUID_HEADER_TEXT_UPD = "UUID";

/**
 * The prefix expected on UUIDs being looked up.
 */
const UUID_PREFIX_UPD = "uuid-";

/**
 * The key name expected in the JSON payload for the secret token.
 */
const AUTH_TOKEN_KEY = "authToken";

/**
 * Property key for API request rate limiting.
 */
const API_LAST_REQUEST_PROP = "lastApiRequestTimestamp";

/**
 * Minimum time in milliseconds between API requests.
 */
const API_RATE_LIMIT_MS = 1000; // 1 second

/**
 * Handles HTTP POST requests to update sheet data based on UUID, requiring authentication.
 * Implements locking and rate limiting.
 *
 * @param {Object} e The event parameter for a POST request.
 * @return {ContentService.TextOutput} A JSON response indicating success or failure.
 */
function doPost(e) {
  let responsePayload;
  
  // Check rate limit first
  if (!checkApiRateLimit()) {
    responsePayload = {
      status: "error",
      message: "Rate limit exceeded. Please try again later."
    };
    return ContentService.createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Try to acquire lock
  const lock = LockService.getScriptLock();
  try {
    // Wait up to 10 seconds to acquire the lock
    lock.waitLock(10000);
  } catch (e) {
    responsePayload = {
      status: "error",
      message: "Server busy. Please try again later."
    };
    return ContentService.createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    // --- Main Processing Logic ---
    let sheetToUpdate = null;
    let rowIndexToUpdate = -1;
    let headers = [];
    let headerMap = {}; // { 'HEADER NAME UPPERCASE': columnIndex (1-based) }

    // --- Authentication Check ---
    const expectedToken = PropertiesService.getScriptProperties().getProperty('SECRET_TOKEN');
    if (!expectedToken) {
        Logger.log("FATAL: SECRET_TOKEN is not set in Script Properties.");
        throw new Error("Server configuration error: Authentication token not set.");
    }

    // 1. Parse the JSON payload
    if (!e || !e.postData || !e.postData.contents) {
        throw new Error("Invalid request: No POST data received.");
    }
    const jsonData = JSON.parse(e.postData.contents);

    // 2. Check for and validate the authentication token
    const receivedToken = jsonData[AUTH_TOKEN_KEY];
    if (!receivedToken) {
        throw new Error(`Authentication failed: Missing '${AUTH_TOKEN_KEY}' field in JSON payload.`);
    }
    if (receivedToken !== expectedToken) {
        Logger.log(`Authentication failed: Received token '${receivedToken}' does not match expected token.`);
        throw new Error("Authentication failed: Invalid token.");
    }
    Logger.log("Authentication successful.");

    // --- Proceed with Update Logic (if authenticated) ---

    // 3. Validate required UUID field
    const uuidToFind = jsonData[UUID_HEADER_TEXT_UPD];
    if (!uuidToFind || typeof uuidToFind !== 'string' || !uuidToFind.startsWith(UUID_PREFIX_UPD)) {
      throw new Error(`Invalid or missing '${UUID_HEADER_TEXT_UPD}' field in JSON payload. It should start with '${UUID_PREFIX_UPD}'.`);
    }

    // 4. Find the sheet and row matching the UUID
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();

    for (const sheet of allSheets) {
      const headerValue = sheet.getRange(1, 1).getValue();
      // Check if this sheet has the UUID header in A1
      if (typeof headerValue === 'string' && headerValue.trim().toUpperCase() === UUID_HEADER_TEXT_UPD.toUpperCase()) {
        const uuidColumnValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues(); // Get all UUIDs (Col A, starting row 2)

        for (let i = 0; i < uuidColumnValues.length; i++) {
          if (uuidColumnValues[i][0] === uuidToFind) {
            sheetToUpdate = sheet;
            rowIndexToUpdate = i + 2; // +1 for 0-based index, +1 because we started from row 2
            Logger.log(`Found UUID ${uuidToFind} in sheet "${sheet.getName()}" at row ${rowIndexToUpdate}`);
            break;
          }
        }
      }
      if (sheetToUpdate) break; // Found it
    }

    // 5. Handle UUID not found
    if (!sheetToUpdate) {
      throw new Error(`UUID '${uuidToFind}' not found in any sheet.`);
    }

    // 6. Get headers and create a map for the found sheet
    headers = sheetToUpdate.getRange(1, 1, 1, sheetToUpdate.getLastColumn()).getValues()[0];
    headers.forEach((header, index) => {
      if (typeof header === 'string' && header.trim() !== '') {
        headerMap[header.trim().toUpperCase()] = index + 1; // Store column index (1-based)
      }
    });

    // 7. Iterate through JSON data and update corresponding cells
    let updatesPerformed = 0;
    for (const key in jsonData) {
      // Skip the UUID and authToken fields
      const keyUpper = key.toUpperCase();
      if (keyUpper === UUID_HEADER_TEXT_UPD.toUpperCase() || keyUpper === AUTH_TOKEN_KEY.toUpperCase()) {
        continue;
      }

      const headerKeyUpper = key.trim().toUpperCase();
      const columnIndex = headerMap[headerKeyUpper];

      if (columnIndex) {
        // Found a matching header column
        // const valueToSet = jsonData[key];
        // Sanitize the value before storing.
        const valueToSet = sanitizeValue(key, jsonData[key]);
        sheetToUpdate.getRange(rowIndexToUpdate, columnIndex).setValue(valueToSet);
        updatesPerformed++;
        Logger.log(`Updated sheet "${sheetToUpdate.getName()}", row ${rowIndexToUpdate}, column ${columnIndex} (${key})`);
      } else {
        Logger.log(`Warning: Header '${key}' from POST data not found in sheet "${sheetToUpdate.getName()}". Skipping update for this field.`);
      }
    }

    // 8. Prepare success response
    responsePayload = {
      status: "success",
      message: `Row updated successfully for UUID ${uuidToFind}. ${updatesPerformed} field(s) processed.`,
      uuid: uuidToFind,
      sheet: sheetToUpdate.getName(),
      row: rowIndexToUpdate
    };

  } catch (error) {
    // 9. Handle errors (auth, parsing, not found, etc.)
    const status = (error.message.startsWith("Authentication failed") || error.message.startsWith("Server configuration error")) ? 401 : 400;
    Logger.log(`Error in doPost (Status ${status}): ${error.message} ${error.stack || ''}`);
    responsePayload = {
      status: "error",
      message: error.message
    };
  } finally {
    // Always release the lock when done
    lock.releaseLock();
  }

  // 10. Return JSON response
  return ContentService.createTextOutput(JSON.stringify(responsePayload))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Sanitizes a value before writing to the sheet.
 * - Rejects formulas (strings starting with '=')
 * - Add more rules as needed
 *
 * @param {string} key Field name (for error context)
 * @param {*} value Value to sanitize
 * @return {*} Sanitized value
 * @throws {Error} If value is invalid
 */
function sanitizeValue(key, value) {
  if (typeof value === 'string' && value.trim().startsWith('=')) {
    throw new Error(`Rejected field '${key}': Formulas are not allowed.`);
  }
  // Future rules can go here

  return value;
}

/**
 * Checks and updates rate limiting for API requests.
 * @return {boolean} True if the request is allowed, false if rate limited
 */
function checkApiRateLimit() {
  const props = PropertiesService.getScriptProperties();
  const lastRequest = props.getProperty(API_LAST_REQUEST_PROP);
  const now = Date.now();
  
  if (lastRequest && (now - parseInt(lastRequest)) < API_RATE_LIMIT_MS) {
    Logger.log(`API rate limit reached. Last request: ${new Date(parseInt(lastRequest))}`);
    return false;
  }
  
  // Update timestamp to enforce rate limiting
  props.setProperty(API_LAST_REQUEST_PROP, now.toString());
  return true;
}
