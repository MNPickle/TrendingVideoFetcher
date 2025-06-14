function fetchWithRetry(url, options = {}, maxRetries = 3, retryDelay = 1000) {
  if (!url || typeof url !== 'string') throw new Error("URL must be a valid string");
  if (typeof options !== 'object' || options === null) throw new Error("Options must be an object");
  if (typeof maxRetries !== 'number' || maxRetries < 0 || maxRetries > 10) throw new Error("maxRetries must be between 0-10");
  if (typeof retryDelay !== 'number' || retryDelay < 0 || retryDelay > 10000) throw new Error("retryDelay must be between 0-10000ms");
  if (!url.match(/^https?:\/\/.+/i)) throw new Error("Invalid URL format");

  let attempt = 0;
  let lastError;

  while (attempt < maxRetries) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      if (code >= 200 && code < 300) {
        return response;
      }
      throw new Error(`HTTP ${code}: ${response.getContentText().substring(0, 500)}`);
    } catch (error) {
      lastError = error;
      attempt++;
      if (attempt < maxRetries) {
        Utilities.sleep(Math.min(retryDelay * Math.pow(2, attempt - 1), 10000));
      }
    }
  }
  throw lastError;
}

/**
 * Robust configuration getter/setter with error handling
 * @param {string} key - Configuration key
 * @param {any} [value] - Optional value to set
 * @return {any} Current value when getting, or set value when setting
 */
function config(key, value) {
  if (!key || typeof key !== 'string') throw new Error("Key must be a non-empty string");

  const scriptProperties = PropertiesService.getScriptProperties();
  
  if (arguments.length === 1) {
    try {
      const storedValue = scriptProperties.getProperty(key);
      return storedValue ? JSON.parse(storedValue) : null;
    } catch (e) {
      console.error(`Failed to parse config value for key ${key}`, e);
      return null;
    }
  } else {
    try {
      scriptProperties.setProperty(key, JSON.stringify(value));
      return value;
    } catch (e) {
      console.error(`Failed to set config value for key ${key}`, e);
      throw e;
    }
  }
}

/**
 * Safe structured logging to sheet with error handling
 * @param {string} sheetName - Target sheet name
 * @param {Object} logData - Log data object
 */
function logToSheet(sheetName, logData) {
  if (!sheetName || typeof sheetName !== 'string') throw new Error("sheetName must be a valid string");
  if (!logData || typeof logData !== 'object') throw new Error("logData must be an object");

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, 4).setValues([["Timestamp", "Level", "Message", "Data"]]);
      sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    }
    
    const timestamp = new Date();
    const row = [
      timestamp,
      logData.level || "INFO",
      logData.message || "",
      JSON.stringify(logData.data || {})
    ];
    
    sheet.appendRow(row);
    
    // Auto-resize columns for better readability
    for (let i = 1; i <= 4; i++) {
      sheet.autoResizeColumn(i);
    }
  } catch (e) {
    console.error("Failed to log to sheet", e);
    throw e;
  }
}

/**
 * Date/time helper functions with input validation
 * @param {Date|string} [date] - Input date (default: now)
 * @return {Object} Helpers object
 */
function dateHelpers(date = new Date()) {
  try {
    if (!(date instanceof Date) && typeof date !== 'string') throw new Error("Invalid date input");
    const parsedDate = date instanceof Date ? new Date(date) : new Date(date);
    if (isNaN(parsedDate.getTime())) throw new Error("Invalid date value");

    return {
      formatDate: function() {
        return Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      },
      
      formatDateTime: function() {
        return Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      },
      
      addDays: function(days) {
        if (typeof days !== 'number') throw new Error("days must be a number");
        const newDate = new Date(parsedDate);
        newDate.setDate(parsedDate.getDate() + days);
        return newDate;
      },
      
      startOfDay: function() {
        const start = new Date(parsedDate);
        start.setHours(0, 0, 0, 0);
        return start;
      },
      
      endOfDay: function() {
        const end = new Date(parsedDate);
        end.setHours(23, 59, 59, 999);
        return end;
      }
    };
  } catch (e) {
    console.error("Date helper initialization failed", e);
    throw e;
  }
}

/**
 * Validates YouTube API response data structure
 * @param {Object} response - API response object
 * @return {boolean} True if valid response structure
 */
function validateYouTubeResponse(response) {
  if (!response || typeof response !== 'object') return false;
  if (response.error) {
    logToSheet("ErrorLogs", {
      level: "ERROR",
      message: "YouTube API Error",
      data: response.error
    });
    return false;
  }
  return true;
}