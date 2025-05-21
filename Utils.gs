// Utils.gs - V15
// This file contains general utility functions that are not directly tied to the CardService UI flow
// or specific Gmail actions, but are used across different parts of the add-on.

/**
 * Parses a JSON string, handling common formatting issues (like markdown code blocks).
 * @param {string} raw The raw string potentially containing JSON.
 * @returns {object|Array} The parsed JSON object or array.
 * @throws {Error} If the string cannot be parsed as JSON.
 */
function _parseJson(raw) {
  if (!raw || typeof raw !== 'string') {
    console.error("Error in _parseJson: Input is not a valid string or is empty. Input: " + raw);
    return {}; // Return an empty object for invalid input
  }

  const trimmedRaw = raw.trim(); // Trim the raw input first

  try {
    // Attempt 1: Clean markdown and parse
    // This regex handles "```json ... ```" or "``` ... ```"
    const jsonBlockMatch = trimmedRaw.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/);
    let jsonToParse;

    if (jsonBlockMatch && jsonBlockMatch[1]) {
      jsonToParse = jsonBlockMatch[1].trim(); // Content within the markdown block
    } else {
      // If no markdown block is found, assume the trimmedRaw might be the JSON itself
      jsonToParse = trimmedRaw;
    }
    return JSON.parse(jsonToParse);

  } catch (e) {
    console.warn("Warning in _parseJson (first attempt failed or no markdown block): " + e.toString() + ". Attempting direct parse or broad match. Original raw: " + raw.substring(0, 200) + "...");
    
    // Fallback Attempt: Try to find the first valid JSON object or array in the string
    // This is more forgiving if the JSON is embedded or not perfectly formatted
    const broadJsonMatch = trimmedRaw.match(/(\{[\s\S]*\})|(\[[\s\S]*\])/);
    if (broadJsonMatch && (broadJsonMatch[1] || broadJsonMatch[2])) {
      const potentialJson = (broadJsonMatch[1] || broadJsonMatch[2]).trim();
      try {
        return JSON.parse(potentialJson);
      } catch (e2) {
        console.error("Error in _parseJson (fallback attempt failed): " + e2.toString() + ". String attempted for parse: " + potentialJson.substring(0,200) + "...");
        throw new Error('Invalid JSON response from AI after all cleaning attempts. Original raw (start): ' + raw.substring(0, 200) + "...");
      }
    }
    // If still no valid JSON found
    console.error("Error in _parseJson: Could not find valid JSON in the input. Original raw (start): " + raw.substring(0, 200) + "...");
    throw new Error('Invalid JSON response from AI, and no JSON object/array found after all attempts. Original raw (start): ' + raw.substring(0, 200) + "...");
  }
}

/**
 * Formats a phone number string into (XXX) XXX-XXXX format.
 * @param {string} phone The phone number string.
 * @returns {string} The formatted phone number or the original if unformattable.
 */
function _formatPhone(phone) {
  if (!phone) return '';
  const digits = phone.replace(/\D/g, ''); // Remove all non-digits
  if (digits.length === 10) {
    return `(${digits.substring(0, 3)}) ${digits.substring(3, 6)}-${digits.substring(6, 10)}`;
  }
  return phone; // Return original if not a 10-digit number
}

/**
 * Extracts a name from an email string (e.g., "Display Name <email@example.com>" or "email@example.com").
 * Attempts to get "Display Name" first, then a cleaned version of the email local part.
 * @param {string} email The full email string.
 * @returns {string} The extracted name.
 */
function _extractNameFromEmail(email) {
  if (!email) return '';
  const match = email.match(/^(.*?)</); // "Display Name <email>"
  if (match && match[1]) return match[1].trim();
  const namePart = email.split('@')[0]; // "email" part
  // Clean up common email separators (., _, numbers at end) and capitalize
  return namePart.replace(/[._\d]+$/, '').replace(/[._]/g, ' ').split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ').trim();
}

/**
 * Extracts just the email address from a full email string (e.g., "Display Name <email@example.com>" or "email@example.com").
 * @param {string} senderEmailField The full email string.
 * @returns {string} The extracted email address.
 */
function _extractActualEmail(senderEmailField) {
  if (!senderEmailField) return '';
  const emailMatch = senderEmailField.match(/<([^>]+)>/);
  if (emailMatch && emailMatch[1]) {
    return emailMatch[1].trim();
  }
  return senderEmailField.trim(); // Assume it's already just the email
}

/**
 * Parses a date string (handling YYYY-MM-DD and MM/DD/YYYY primarily) into milliseconds since epoch.
 * The goal is to represent midnight on the given date in the script's configured timezone ("America/New_York").
 * Handles MM/DD for current/next year. Defaults to current date at midnight if parsing fails.
 * @param {string} dateString The date string to parse.
 * @returns {number} The date in milliseconds since epoch, representing midnight in the script's timezone.
 */
function _parseDateToMsEpoch(dateString) {
  if (!dateString || typeof dateString !== 'string' || dateString.trim() === "") {
    // Default to today at midnight in script's timezone
    const today = new Date();
    return new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  }
  const trimmedDateString = dateString.trim();
  let year, monthIndex, day;

  // Try YYYY-MM-DD
  let match = trimmedDateString.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) {
    year = parseInt(match[1]);
    monthIndex = parseInt(match[2]) - 1; // JS months are 0-indexed
    day = parseInt(match[3]);
    console.log(`_parseDateToMsEpoch: Parsed YYYY-MM-DD: ${year}-${monthIndex + 1}-${day} from "${trimmedDateString}"`);
  } else {
    // Try MM/DD/YYYY
    match = trimmedDateString.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (match) {
      year = parseInt(match[3]);
      monthIndex = parseInt(match[1]) - 1;
      day = parseInt(match[2]);
      console.log(`_parseDateToMsEpoch: Parsed MM/DD/YYYY: ${monthIndex + 1}/${day}/${year} from "${trimmedDateString}"`);
    } else {
      // Try MM/DD (assumes current or next year)
      match = trimmedDateString.match(/^(\d{1,2})\/(\d{1,2})$/);
      if (match) {
        const currentJsDate = new Date();
        year = currentJsDate.getFullYear();
        monthIndex = parseInt(match[1]) - 1;
        day = parseInt(match[2]);
        
        // Check if this date (in current year) has already passed
        const tempDateThisYear = new Date(year, monthIndex, day);
        const todayAtMidnight = new Date(currentJsDate.getFullYear(), currentJsDate.getMonth(), currentJsDate.getDate());
        if (tempDateThisYear.getTime() < todayAtMidnight.getTime()) {
          year++; // Assume next year
        }
        console.log(`_parseDateToMsEpoch: Parsed MM/DD: ${monthIndex + 1}/${day}, resolved to year ${year} from "${trimmedDateString}"`);
      } else {
        // Fallback for other textual formats (e.g., "May 20", "April 23, 2025")
        // This can be risky due to JavaScript's Date parsing quirks.
        let parsedAttempt = new Date(trimmedDateString);
        if (!isNaN(parsedAttempt.getTime())) {
          // Extract year, month, day from this potentially timezone-ambiguous parse
          // To ensure we are setting it to local midnight of that *intended calendar day*
          // It's safer to extract components if possible
          year = parsedAttempt.getFullYear(); 
          monthIndex = parsedAttempt.getMonth();
          day = parsedAttempt.getDate();
          console.log(`_parseDateToMsEpoch: Parsed textual date "${trimmedDateString}" to Y:${year}, M:${monthIndex}, D:${day}`);
        } else {
          console.warn(`_parseDateToMsEpoch: Could not parse date string: "${trimmedDateString}". Defaulting to today at midnight.`);
          const today = new Date();
          return new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
        }
      }
    }
  }

  // Construct the date as midnight in the script's local timezone
  // new Date(year, monthIndex, day) does exactly this.
  const finalDate = new Date(year, monthIndex, day);
  
  if (isNaN(finalDate.getTime())) {
      console.warn(`_parseDateToMsEpoch: Final constructed date is invalid for input "${trimmedDateString}". Defaulting.`);
      const today = new Date();
      return new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  }
  
  console.log(`_parseDateToMsEpoch: Input "${trimmedDateString}", successfully parsed to local Date object: ${finalDate}, returning ms: ${finalDate.getTime()}`);
  return finalDate.getTime();
}


/**
 * Combines a date (in milliseconds since epoch) and a time string (e.g., "7:00 PM") into a single timestamp.
 * @param {number} dateMs Date in milliseconds since epoch.
 * @param {string} timeStr Time string (e.g., "7:00 PM", "14:30").
 * @returns {number} Combined date and time in milliseconds since epoch.
 */
function _combineDateAndTime(dateMs, timeStr) {
  if (!dateMs || !timeStr) return dateMs || new Date().getTime();

  const dateObj = new Date(dateMs);
  let hours = 0;
  let minutes = 0;

  const timeParts = timeStr.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
  if (timeParts) {
    hours = parseInt(timeParts[1]);
    minutes = parseInt(timeParts[2]);
    const ampm = timeParts[3] ? timeParts[3].toUpperCase() : null;

    if (ampm === 'PM' && hours < 12) {
      hours += 12;
    } else if (ampm === 'AM' && hours === 12) { // Midnight case (12 AM)
      hours = 0;
    }
  } else { // Try parsing if it's just HH:MM (24-hour)
    const simpleTimeParts = timeStr.match(/(\d{1,2}):(\d{2})/);
    if (simpleTimeParts) {
      hours = parseInt(simpleTimeParts[1]);
      minutes = parseInt(simpleTimeParts[2]);
    } else {
      return dateMs; // Cannot parse time, return original dateMs
    }
  }
  dateObj.setHours(hours, minutes, 0, 0); // Set hours and minutes, reset seconds/ms
  return dateObj.getTime();
}

/**
 * Matches an email address to a predefined client based on rules.
 * Includes enhanced logging for troubleshooting.
 * @param {string} senderEmailField The full sender email string (e.g., "Display Name <email@example.com>").
 * @returns {string} The matched client name or 'Unknown'.
 */
function _matchClient(senderEmailField) {
  // Log the CLIENT_RULES_LOOKUP as it's seen by this function at runtime
  // Ensure Constants.gs is saved and its values are up-to-date if this log doesn't match your expectations.
  console.log("_matchClient: CLIENT_RULES_LOOKUP being used:", JSON.stringify(CLIENT_RULES_LOOKUP));
  console.log("_matchClient: Received senderEmailField:", senderEmailField);

  if (!senderEmailField) {
    console.log("_matchClient: senderEmailField is null or empty, returning 'Unknown'.");
    return 'Unknown';
  }

  const emailAddress = _extractActualEmail(senderEmailField); // Uses your existing helper
  console.log("_matchClient: Extracted emailAddress with _extractActualEmail:", emailAddress);

  if (!emailAddress) { // Add a check for empty extracted email
    console.log("_matchClient: _extractActualEmail returned empty or null, returning 'Unknown'.");
    return 'Unknown';
  }

  const emailLower = emailAddress.toLowerCase().trim();
  console.log("_matchClient: Normalized emailLower for matching:", emailLower);

  // CLIENT_RULES_LOOKUP should be pre-sorted by rule length (descending) in Constants.gs
  for (let i = 0; i < CLIENT_RULES_LOOKUP.length; i++) {
    const clientRuleEntry = CLIENT_RULES_LOOKUP[i];
    if (!clientRuleEntry || !clientRuleEntry.rule || !clientRuleEntry.clientName) {
      console.warn("_matchClient: Skipping invalid client rule entry at index " + i + ": " + JSON.stringify(clientRuleEntry));
      continue;
    }
    const rule = clientRuleEntry.rule.toLowerCase().trim();
    const clientName = clientRuleEntry.clientName.trim();

    console.log(`_matchClient: Checking rule: "${rule}" (for client: "${clientName}") against email: "${emailLower}"`);
    
    // The core matching logic: does the lowercase email *include* the lowercase rule string?
    if (rule && emailLower.includes(rule)) {
      console.log(`_matchClient: Match found! Client: "${clientName}" for email "${emailAddress}" with rule "${rule}"`);
      return clientName;
    }
  }

  console.log("_matchClient: No client match found for email: " + emailAddress);
  return 'Unknown';
}

/**
 * Calls the Gemini API with a given prompt.
 * @param {string} prompt The text prompt for the Gemini model.
 * @returns {string} The text response from Gemini.
 * @throws {Error} If the API key is missing, or the API call fails or returns an invalid response.
 */
function callGemini(prompt) {
  if (!API_KEY) {
    console.error("Error: GL_API_KEY is not set.");
    throw new Error("No API key set. Please set 'GL_API_KEY' in Script Properties.");
  }
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + encodeURIComponent(API_KEY);
  const payload = {
    contents: [{
      role: "user",
      parts: [{
        text: prompt
      }]
    }]
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // Capture errors in response body
  };

  console.log("Calling Gemini. Prompt length: " + prompt.length);
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode === 200) {
    const json = JSON.parse(responseBody);
    if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts.length > 0) {
      return json.candidates[0].content.parts[0].text;
    } else {
      console.error("Gemini response missing expected structure: " + responseBody);
      if (json.candidates && json.candidates[0] && json.candidates[0].finishReason) {
        console.error("Gemini finishReason: " + json.candidates[0].finishReason);
        if (json.candidates[0].safetyRatings) {
          console.error("SafetyRatings: " + JSON.stringify(json.candidates[0].safetyRatings));
        }
        if (json.candidates[0].finishReason === "SAFETY" || json.candidates[0].finishReason === "OTHER") {
          throw new Error("Gemini request blocked due to: " + json.candidates[0].finishReason + ". Check safety ratings in log.");
        }
      }
      throw new Error("Invalid Gemini response structure. See logs.");
    }
  } else {
    console.error("Gemini API Error - Code: " + responseCode + " Body: " + responseBody);
    throw new Error("Gemini API request failed. Code: " + responseCode + ". See logs for details.");
  }
}

/**
 * Retrieves form input values safely.
 * This helper is used to prevent "Missing initializer" errors by ensuring
 * that the access path to the value is always valid, returning default
 * values (empty string or null) if a part of the path is missing.
 * @param {object} inputs The formInputs object from the event.
 * @param {string} fieldName The name of the input field.
 * @param {boolean} isDate True if the input is a date picker, to return msSinceEpoch.
 * @returns {string|number|null} The input value or a default (empty string for text, null for date).
 */
function _getFormInputValue(inputs, fieldName, isDate = false) {
  if (!inputs || !inputs[fieldName]) {
    return isDate ? null : '';
  }
  if (isDate) {
    // For date pickers, the value is in dateInput.msSinceEpoch
    return inputs[fieldName].dateInput?.msSinceEpoch || null;
  }
  // For text inputs or dropdowns (stringInputs.value is an array)
  return inputs[fieldName].stringInputs?.value?.[0] || '';
}

/**
 * Retrieves master QuickBooks items from the designated Google Sheet.
 * @returns {Array<object>} An array of item objects.
 */
function getMasterQBItems() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID); const sheet = spreadsheet.getSheetByName(ITEM_LOOKUP_SHEET_NAME);
    if (!sheet) { console.warn("Master item sheet '" + ITEM_LOOKUP_SHEET_NAME + "' not found."); return []; }
    const data = sheet.getDataRange().getValues(); if (data.length < 2) { console.warn("No data in master item sheet."); return []; }
    const headers = data.shift().map(header => header.toString().trim());
    return data.map(row => {
      let item = {}; let hasSKU = false;
      headers.forEach((header, index) => {
        let value = row[index];
        if (header === 'SKU' || header === 'id') { value = (value !== undefined && value !== "") ? value.toString().trim() : null; if (header === 'SKU' && value) hasSKU = true; }
        else if (header === 'Price' || header === 'Add-on Cost 1' || header === 'Add-on Cost 2') { value = (value !== undefined && value !== "") ? parseFloat(value) : 0; }
        else if (header === 'Quantity') { value = (value !== undefined && value !== "") ? parseInt(value) : 1; }
        else if (typeof value === 'string') { value = value.trim(); }
        item[header] = value;
      });
      return hasSKU ? item : null; 
    }).filter(item => item !== null);
  } catch (e) { console.error("Error in getMasterQBItems: " + e.toString()); return []; }
}

/**
 * Uses Gemini to match email items to master QuickBooks items.
 * @param {Array<object>} emailItems Items extracted from the email.
 * @param {Array<object>} masterQBItems The master list of QuickBooks items.
 * @returns {Array<object>} Matched items with QB details.
 */
function getGeminiItemMatches(emailItems, masterQBItems) {
  if (!emailItems || emailItems.length === 0) return [];
  if (!masterQBItems || masterQBItems.length === 0) { 
    console.warn("Master QB Items list is empty in getGeminiItemMatches. Falling back to custom SKU.");
    return emailItems.map(item => ({
      original_email_description: item.description, extracted_main_quantity: item.quantity,
      matched_qb_item_id: FALLBACK_CUSTOM_ITEM_SKU, matched_qb_item_name: "Custom Item (No Master List)",
      match_confidence: "Low", parsed_flavors_or_details: item.description,
      identified_flavors: [] 
    }));
  }
  
  const prompt = _buildItemMatchingPrompt(emailItems, masterQBItems);
  console.log("Prompt for getGeminiItemMatches (first 500 chars):\n" + prompt.substring(0, 500)); 
  const geminiResponseText = callGemini(prompt);
  console.log("Raw response from getGeminiItemMatches: " + geminiResponseText);
  
  try {
    const parsedResponse = _parseJson(geminiResponseText);
    if (Array.isArray(parsedResponse)) {
      // Ensure identified_flavors is always an array
      return parsedResponse.map(item => ({ ...item, identified_flavors: item.identified_flavors || [] })) ;
    } else {
      console.error("Parsed Gemini response is not an array: " + JSON.stringify(parsedResponse));
      throw new Error("Parsed response is not an array.");
    }
  } catch (e) {
    console.error("Error parsing Gemini item matching response: " + e.toString() + " Raw response for parse error: " + geminiResponseText);
    const fallbackQBItem = masterQBItems.find(item => item.SKU === FALLBACK_CUSTOM_ITEM_SKU);
    const fallbackSkuForError = fallbackQBItem ? fallbackQBItem.SKU : FALLBACK_CUSTOM_ITEM_SKU;
    const fallbackNameForError = fallbackQBItem ? fallbackQBItem.Name : "Custom Unmatched Item";
    return emailItems.map(item => ({ 
      original_email_description: item.description, extracted_main_quantity: item.quantity,
      matched_qb_item_id: fallbackSkuForError, matched_qb_item_name: fallbackNameForError + " (AI Error)",
      match_confidence: "Low", parsed_flavors_or_details: "AI matching error occurred.", identified_flavors: []
    }));
  }
}

/**
 * Populates a new kitchen sheet with order details and confirmed items.
 * MODIFIED: Handles YYYY-MM-DD date format.
 * @param {string} orderNum The order number.
 * @returns {{id: string, url: string, name: string}} Object containing new sheet ID, URL, and Name.
 * @throws {Error} If order data or templates are not found, or population fails.
 */
function populateKitchenSheet(orderNum) {
  const userProps = PropertiesService.getUserProperties(); 
  const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { 
    console.error("KitchenSheet: Order data for " + orderNum + " not found."); 
    throw new Error("Order data not found for kitchen sheet: " + orderNum); 
  }
  const orderData = JSON.parse(orderDataString); 
  const confirmedItems = orderData['ConfirmedQBItems']; 
  const masterAllItems = getMasterQBItems();
  if (!confirmedItems || !Array.isArray(confirmedItems)) { 
    console.error("KitchenSheet: Confirmed items not found for order " + orderNum); 
    throw new Error("Confirmed items not found for kitchen sheet generation."); 
  }
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID); 
    const templateSheet = spreadsheet.getSheetByName(KITCHEN_SHEET_TEMPLATE_NAME);
    if (!templateSheet) { 
      console.error("Kitchen sheet template '" + KITCHEN_SHEET_TEMPLATE_NAME + "' not found."); 
      throw new Error("Kitchen sheet template not found."); 
    }
    const newSheetName = `Kitchen - ${orderNum} - ${orderData['Contact Person'] || orderData['Customer Name'] || 'Unknown'}`.substring(0,100); 
    const newSheet = templateSheet.copyTo(spreadsheet).setName(newSheetName);
    
    const customerNameForKitchen = orderData['Contact Person'] || orderData['Customer Name'] || '';
    const contactPhoneForKitchen = _formatPhone(orderData['Contact Phone'] || orderData['Customer Address Phone'] || '');
    newSheet.getRange(KITCHEN_CUSTOMER_PHONE_CELL).setValue(`${customerNameForKitchen} - Ph: ${contactPhoneForKitchen}`);
    
    // MODIFICATION: Handle YYYY-MM-DD date format from orderData
    let deliveryDateFormattedForSheet = orderData['Delivery Date'] || '';
    if (deliveryDateFormattedForSheet && deliveryDateFormattedForSheet.match(/^\d{4}-\d{2}-\d{2}$/)) {
        const parts = deliveryDateFormattedForSheet.split('-');
        const dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        deliveryDateFormattedForSheet = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");
    } // else, if it's already MM/DD/YYYY or other, use as is or add more parsing if needed.
      // For now, assuming it will be YYYY-MM-DD from storage.

    newSheet.getRange(KITCHEN_DELIVERY_DATE_CELL).setValue(deliveryDateFormattedForSheet);
    newSheet.getRange(KITCHEN_DELIVERY_TIME_CELL).setValue(orderData['Delivery Time'] || '');
    
    let currentRow = KITCHEN_ITEM_START_ROW;
    confirmedItems.forEach(item => {
      const masterItem = masterAllItems.find(mi => mi.SKU === item.sku); 
      const itemSize = masterItem ? (masterItem.Size || '') : '';
      newSheet.getRange(KITCHEN_QTY_COL + currentRow).setValue(item.quantity);
      newSheet.getRange(KITCHEN_SIZE_COL + currentRow).setValue(itemSize);
      newSheet.getRange(KITCHEN_ITEM_NAME_COL + currentRow).setValue(item.quickbooks_item_name);
      newSheet.getRange(KITCHEN_FILLING_COL + currentRow).setValue(item.kitchen_notes_and_flavors);
      newSheet.getRange(KITCHEN_NOTES_COL + currentRow).setValue(''); 
      currentRow++;
    });
    SpreadsheetApp.flush();
    return { id: newSheet.getSheetId(), url: newSheet.getParent().getUrl() + '#gid=' + newSheet.getSheetId(), name: newSheetName };
  } catch (e) { 
    console.error("Error in populateKitchenSheet for order " + orderNum + ": " + e.toString() + (e.stack ? ("\nStack: " + e.stack) : "")); 
    throw new Error("Failed to populate kitchen sheet: " + e.message); 
  }
}

/**
 * Populates a new invoice sheet with order details and confirmed items.
 * MODIFIED: Handles YYYY-MM-DD date format.
 * @param {string} orderNum The order number.
 * @returns {{id: string, url: string, name: string}} Object containing new sheet ID, URL, and Name.
 * @throws {Error} If order data or templates are not found, or population fails.
 */
function populateInvoiceSheet(orderNum) {
  const userProps = PropertiesService.getUserProperties(); 
  const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { 
    console.error("InvoiceSheet: Order data for " + orderNum + " not found."); 
    throw new Error("Order data not found for " + orderNum); 
  }
  const orderData = JSON.parse(orderDataString); 
  const confirmedItems = orderData['ConfirmedQBItems'];
  if (!confirmedItems || !Array.isArray(confirmedItems)) { 
    console.error("InvoiceSheet: Confirmed items not found for order " + orderNum); 
    throw new Error("Confirmed items not found for invoice generation."); 
  }
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID); 
    const templateSheet = spreadsheet.getSheetByName(INVOICE_TEMPLATE_SHEET_NAME);
    if (!templateSheet) { 
      console.error("Invoice template sheet '" + INVOICE_TEMPLATE_SHEET_NAME + "' not found."); 
      throw new Error("Invoice template sheet not found."); 
    }
    const newSheetName = `Invoice - ${orderNum} - ${orderData['Customer Name'] || 'Unknown'}`.substring(0,100); 
    const newSheet = templateSheet.copyTo(spreadsheet).setName(newSheetName);
    
    newSheet.getRange(ORDER_NUM_CELL).setValue(orderData.orderNum);
    newSheet.getRange(CUSTOMER_NAME_CELL).setValue(orderData['Customer Name'] || ''); 
    newSheet.getRange(ADDRESS_LINE_1_CELL).setValue(orderData['Customer Address Line 1'] || '');
    newSheet.getRange(ADDRESS_LINE_2_CELL).setValue(orderData['Customer Address Line 2'] || '');
    
    const cityStateZip = `${orderData['Customer Address City'] || ''}${orderData['Customer Address City'] && (orderData['Customer Address State'] || orderData['Customer Address ZIP']) ? ', ' : ''}${orderData['Customer Address State'] || ''} ${orderData['Customer Address ZIP'] || ''}`.trim();
    newSheet.getRange(CITY_STATE_ZIP_CELL).setValue(cityStateZip);
    
    // MODIFICATION: Handle YYYY-MM-DD date format from orderData
    let deliveryDateFormattedForSheet = orderData['Delivery Date'] || '';
    if (deliveryDateFormattedForSheet && deliveryDateFormattedForSheet.match(/^\d{4}-\d{2}-\d{2}$/)) {
        const parts = deliveryDateFormattedForSheet.split('-');
        const dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        deliveryDateFormattedForSheet = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");
    } // else, if it's already MM/DD/YYYY or other, use as is.

    newSheet.getRange(DELIVERY_DATE_CELL_INVOICE).setValue(deliveryDateFormattedForSheet);
    newSheet.getRange(DELIVERY_TIME_CELL_INVOICE).setValue(orderData['Delivery Time'] || '');
    
    let currentRow = ITEM_START_ROW_INVOICE;
    let grandTotal = 0;
    confirmedItems.forEach(item => {
      const descriptionCell = newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow);
      descriptionCell.setValue(item.original_email_description || item.quickbooks_item_name).setWrap(false); 
      newSheet.getRange(ITEM_QTY_COL_INVOICE + currentRow).setValue(item.quantity);
      newSheet.getRange(ITEM_UNIT_PRICE_COL_INVOICE + currentRow).setValue(item.unit_price).setNumberFormat("$#,##0.00"); 
      const lineTotal = (item.quantity || 0) * (item.unit_price || 0);
      newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setValue(lineTotal).setNumberFormat("$#,##0.00"); 
      grandTotal += lineTotal; currentRow++;
    });

    let tipAmount = orderData['TipAmount'] || 0;
    let otherChargesAmount = orderData['OtherChargesAmount'] || 0;
    let otherChargesDescription = orderData['OtherChargesDescription'] || "Other Charges";

    if (tipAmount > 0) {
        newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow).setValue("Tip").setWrap(false);
        newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setValue(tipAmount).setNumberFormat("$#,##0.00");
        grandTotal += tipAmount; currentRow++;
    }
    if (otherChargesAmount > 0) {
        newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow).setValue(otherChargesDescription).setWrap(false);
        newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setValue(otherChargesAmount).setNumberFormat("$#,##0.00");
        grandTotal += otherChargesAmount; currentRow++;
    }
    
    let deliveryFee = BASE_DELIVERY_FEE;
    if (orderData['master_delivery_time_ms']) {
        const deliveryHour = new Date(orderData['master_delivery_time_ms']).getHours();
        if (deliveryHour >= DELIVERY_FEE_CUTOFF_HOUR) {
            deliveryFee = AFTER_4PM_DELIVERY_FEE;
        }
    }
    newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow).setValue("Delivery Fee").setWrap(false);
    newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setValue(deliveryFee).setNumberFormat("$#,##0.00");
    grandTotal += deliveryFee; currentRow++;

    if (orderData['Include Utensils?'] === 'Yes') {
        const numUtensils = parseInt(orderData['If yes: how many?']) || 0;
        if (numUtensils > 0) {
            const utensilTotalCost = numUtensils * COST_PER_UTENSIL_SET;
            newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow).setValue(`Utensils (${numUtensils} sets)`).setWrap(false);
            newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setValue(utensilTotalCost).setNumberFormat("$#,##0.00");
            grandTotal += utensilTotalCost; currentRow++;
        }
    }

    const grandTotalDescCell = newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow);
    grandTotalDescCell.setValue("Grand Total:").setFontWeight("bold").setWrap(false);
    const grandTotalValueCell = newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow);
    grandTotalValueCell.setValue(grandTotal).setNumberFormat("$#,##0.00").setFontWeight("bold").setHorizontalAlignment("right").setWrap(false);
    newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow + ":" + ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setBorder(true, null, null, null, null, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    SpreadsheetApp.flush();
    return { id: newSheet.getSheetId(), url: newSheet.getParent().getUrl() + '#gid=' + newSheet.getSheetId(), name: newSheetName };
  } catch (e) { 
    console.error("Error in populateInvoiceSheet for order " + orderNum + ": " + e.toString() + (e.stack ? ("\nStack: " + e.stack) : "")); 
    throw new Error("Failed to populate invoice sheet: " + e.message); 
  }
}

/**
 * Creates a PDF blob of the specified sheet using UrlFetchApp.
 * @param {string} orderNum The order number (for logging).
 * @param {string} populatedSheetSpreadsheetId The ID of the spreadsheet.
 * @param {string} populatedSheetName The name of the populated invoice sheet.
 * @returns {GoogleAppsScript.Base.Blob|null} The PDF blob or null if an error occurs.
 * @throws {Error} If processing fails to obtain a sheet GID or PDF.
 */
function createPdfBlobOnly(orderNum, populatedSheetSpreadsheetId, populatedSheetName) {
  console.log(`createPdfBlobOnly: Initiating PDF blob creation for order ${orderNum}. SpreadsheetID: ${populatedSheetSpreadsheetId}, SheetName: "${populatedSheetName}"`);
  
  try {
    const spreadsheet = SpreadsheetApp.openById(populatedSheetSpreadsheetId); 
    if (!spreadsheet) {
      console.error(`createPdfBlobOnly: Critical - Failed to open spreadsheet with ID: ${populatedSheetSpreadsheetId}.`);
      throw new Error("Failed to open spreadsheet for PDF generation.");
    }
    console.log(`createPdfBlobOnly: Spreadsheet "${spreadsheet.getName()}" (ID: ${spreadsheet.getId()}) opened.`);

    let sheetForPdf = null;
    let sheetGid = null;
    let attempts = 0;
    const MAX_ATTEMPTS = 3;

    while (attempts < MAX_ATTEMPTS) {
        attempts++;
        console.log(`createPdfBlobOnly: Attempt ${attempts}/${MAX_ATTEMPTS} to getSheetByName("${populatedSheetName}").`);
        let currentSheetTry = spreadsheet.getSheetByName(populatedSheetName);
        
        if (currentSheetTry) {
            let currentGidAttempt = null;
            try {
                currentGidAttempt = currentSheetTry.getSheetId();
            } catch (gidErr) {
                console.warn(`createPdfBlobOnly: Attempt ${attempts} - Error getting GID for sheet "${currentSheetTry.getName()}": ${gidErr.message}`);
            }

            const hasGetAs = typeof currentSheetTry.getAs === 'function'; // Log for info
            console.log(`createPdfBlobOnly: Attempt ${attempts} - Sheet named "${currentSheetTry.getName()}" found. GID: ${currentGidAttempt}. Has getAs: ${hasGetAs}`);
            
            if (currentGidAttempt !== null && currentGidAttempt !== undefined) {
                sheetForPdf = currentSheetTry;
                sheetGid = currentGidAttempt;
                console.log(`createPdfBlobOnly: Successfully obtained sheet GID ${sheetGid} on attempt ${attempts}.`);
                break; 
            } else {
                console.warn(`createPdfBlobOnly: Attempt ${attempts} - Sheet GID is null/undefined. Activating and will retry.`);
                try {
                    spreadsheet.setActiveSheet(currentSheetTry);
                    SpreadsheetApp.flush();
                    console.log(`createPdfBlobOnly: Attempt ${attempts} - Activated sheet "${currentSheetTry.getName()}".`);
                } catch (activateErr) {
                    console.warn(`createPdfBlobOnly: Attempt ${attempts} - Could not activate sheet: ${activateErr.message}.`);
                }
            }
        } else {
             console.warn(`createPdfBlobOnly: Attempt ${attempts} - Sheet "${populatedSheetName}" not found.`);
        }

        if (attempts < MAX_ATTEMPTS && (sheetGid === null || sheetGid === undefined) ) {
            const delayMs = 1000 * attempts; 
            console.log(`createPdfBlobOnly: Waiting ${delayMs}ms before next attempt.`);
            Utilities.sleep(delayMs);
        }
    } 
    
    if (!sheetForPdf || sheetGid === null || sheetGid === undefined) { 
      const availableSheets = spreadsheet.getSheets().map(s => s.getName()).join(', ');
      console.error(`createPdfBlobOnly: Critical - After ${MAX_ATTEMPTS} attempts, could not get GID for "${populatedSheetName}". Available: [${availableSheets}].`);
      throw new Error(`Sheet "${populatedSheetName}" GID could not be obtained for PDF export.`);
    }
    
    // Generate PDF using UrlFetch
    const pdfExportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export` +
                         `?format=pdf&gid=${sheetGid}&scale=4&top_margin=0.50&bottom_margin=0.50` +
                         `&left_margin=0.50&right_margin=0.50&horizontal_alignment=CENTER&vertical_alignment=TOP` +
                         `&gridlines=false&printnotes=false&pageorder=1&sheetnames=false&printtitle=false` +
                         `&attachment=true&portrait=true&size=letter`;
    console.log(`createPdfBlobOnly: PDF Export URL (first 150): ${pdfExportUrl.substring(0,150)}...`);
    
    const response = UrlFetchApp.fetch(pdfExportUrl, {
        headers: { 'Authorization': `Bearer ${ScriptApp.getOAuthToken()}` },
        muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
        const pdfBlob = response.getBlob().setName(`${populatedSheetName}.pdf`);
        console.log(`createPdfBlobOnly: PDF blob created: ${pdfBlob.getName()}, Size: ${pdfBlob.getBytes().length} bytes.`);
        return pdfBlob;
    } else {
        const errorContent = response.getContentText();
        console.error(`createPdfBlobOnly: UrlFetch PDF export failed. Code: ${responseCode}. Response: ${errorContent.substring(0, 500)}`);
        throw new Error(`PDF export via UrlFetch failed with code ${responseCode}.`);
    }

  } catch (e) { 
    console.error(`Error in createPdfBlobOnly for order ${orderNum}, sheet "${populatedSheetName}": ${e.toString()}${(e.stack ? ("\nStack: " + e.stack) : "")}`); 
    // Return null or rethrow, depending on how handleGenerateInvoiceAndEmail should react
    return null; 
  }
}

/**
 * Helper to safely get all sheets from a spreadsheet object.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The spreadsheet object.
 * @return {Array<GoogleAppsScript.Spreadsheet.Sheet>} Array of sheet objects, or empty array if error.
 */
function allSheetsFromSpreadsheet(spreadsheet) {
    try {
        if (spreadsheet && typeof spreadsheet.getSheets === 'function') {
            return spreadsheet.getSheets();
        }
        console.warn("allSheetsFromSpreadsheet: Spreadsheet object was null or did not have getSheets method.");
        return [];
    } catch (e) {
        console.error("Error in allSheetsFromSpreadsheet: " + e.toString());
        return [];
    }
}

/**
 * Helper function to generate a PDF from a given sheet using UrlFetchApp, 
 * and prepare an email draft with the PDF attached.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetForPdf The Sheet object (used to get GID and for context like name).
 * @param {string} pdfSheetName The desired name for the PDF file (usually the sheet name from sheetForPdf.getName()).
 * @param {object} orderDataForEmail The order data object containing details for the email.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The parent spreadsheet object (used for its ID).
 * @returns {{pdfBlob: GoogleAppsScript.Base.Blob, draft: GoogleAppsScript.Gmail.GmailDraft, draftId: string}}
 * @throws {Error} If PDF generation or email drafting fails.
 */
function generatePdfFromSheet(sheetForPdf, pdfSheetName, orderDataForEmail, spreadsheet) {
    const orderNum = orderDataForEmail.orderNum; // Assuming orderNum is in orderDataForEmail
    const spreadsheetId = spreadsheet.getId();
    const sheetGid = sheetForPdf.getSheetId(); // GID of the sheet to export

    console.log(`generatePdfFromSheet: Attempting PDF generation for sheet: "${sheetForPdf.getName()}" (GID: ${sheetGid}) in Spreadsheet ID: ${spreadsheetId} using UrlFetchApp.`);

    let pdfBlob;
    try {
        const pdfExportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export` +
                             `?format=pdf` +
                             `&gid=${sheetGid}` +
                             `&scale=4` + 
                             `&top_margin=0.50` +  
                             `&bottom_margin=0.50` +
                             `&left_margin=0.50` +
                             `&right_margin=0.50` +
                             `&horizontal_alignment=CENTER` +
                             `&vertical_alignment=TOP` +
                             `&gridlines=false` + 
                             `&printnotes=false` + 
                             `&pageorder=1` + 
                             `&sheetnames=false` + 
                             `&printtitle=false` + 
                             `&attachment=true` +  
                             `&portrait=true` +    
                             `&size=letter`;       

        console.log(`generatePdfFromSheet: Constructed PDF Export URL (first 150 chars, excludes token): ${pdfExportUrl.substring(0,150)}...`);
        
        const response = UrlFetchApp.fetch(pdfExportUrl, {
            headers: {
                'Authorization': `Bearer ${ScriptApp.getOAuthToken()}` 
            },
            muteHttpExceptions: true 
        });

        const responseCode = response.getResponseCode();
        if (responseCode === 200) {
            pdfBlob = response.getBlob().setName(`${pdfSheetName}.pdf`); 
            console.log(`generatePdfFromSheet: PDF blob created via UrlFetch: ${pdfBlob.getName()}, Size: ${pdfBlob.getBytes().length} bytes.`);
        } else {
            const errorContent = response.getContentText();
            console.error(`generatePdfFromSheet: UrlFetch PDF export failed for sheet "${pdfSheetName}". Code: ${responseCode}. Response: ${errorContent.substring(0, 500)}`);
            throw new Error(`PDF export via UrlFetch failed with HTTP code ${responseCode}.`);
        }
    } catch (e) {
        console.error(`generatePdfFromSheet: Exception during UrlFetch PDF export for order ${orderNum}, sheet "${pdfSheetName}": ${e.toString()}${(e.stack ? ("\nStack: " + e.stack) : "")}`);
        throw new Error(`Failed to generate PDF via UrlFetch for sheet "${pdfSheetName}": ${e.message}`);
    }
    const threadId = orderDataForEmail.threadId;
    if (!threadId) { 
      console.error(`generatePdfFromSheet: Thread ID missing for order ${orderNum}`);
      throw new Error("Original email thread ID not found for reply."); 
    }
    const thread = GmailApp.getThreadById(threadId);
    if (!thread) { 
      console.error(`generatePdfFromSheet: Could not retrieve Gmail thread ID: ${threadId} for order ${orderNum}`);
      throw new Error("Could not retrieve original email thread."); 
    }
    const messages = thread.getMessages(); 
    const messageToReplyTo = messages[messages.length - 1]; 
    
    const recipient = orderDataForEmail['Contact Email'] || orderDataForEmail['Customer Address Email'] || orderDataForEmail['Internal Sender Email'] || messageToReplyTo.getFrom();
    const subject = `Catering Order Confirmed & Invoice - El Merkury - Order #${orderNum}`;
    
    // MODIFICATION: Date formatting for email, ensuring YYYY-MM-DD is parsed correctly
    let deliveryDateForEmail = "Not specified";
    if (orderDataForEmail['Delivery Date']) { // Expecting YYYY-MM-DD format from orderData
        try {
            const dateStr = orderDataForEmail['Delivery Date'];
            let tempDate;
            if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) { // YYYY-MM-DD
                const parts = dateStr.split('-');
                // new Date(year, monthIndex, day) for local midnight
                tempDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
            } else if (dateStr.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) { // MM/DD/YYYY fallback
                const parts = dateStr.split('/');
                tempDate = new Date(parseInt(parts[2]), parseInt(parts[0]) - 1, parseInt(parts[1]));
            } else { // Fallback for other unparsed or already formatted strings
                 tempDate = new Date(_parseDateToMsEpoch(dateStr)); // Use full parse if not YYYY-MM-DD
            }
            if (!isNaN(tempDate.getTime())) {
                 deliveryDateForEmail = Utilities.formatDate(tempDate, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy");
            } else {
                 deliveryDateForEmail = orderDataForEmail['Delivery Date']; // Use raw if parsing failed
            }
        } catch (dateParseErr) {
            console.warn(`generatePdfFromSheet: Error parsing deliveryDateForEmail for order ${orderNum}: "${orderDataForEmail['Delivery Date']}". Error: ${dateParseErr}. Using raw value.`);
            deliveryDateForEmail = orderDataForEmail['Delivery Date']; 
        }
    }
    // END MODIFICATION for date formatting

    const deliveryTimeForEmail = orderDataForEmail['Delivery Time'] ? _normalizeTimeFormat(orderDataForEmail['Delivery Time']) : "Not specified";

    const body = `Dear ${orderDataForEmail['Contact Person'] || orderDataForEmail['Customer Name'] || 'Valued Customer'},\n\nThank you for your El Merkury catering order!\n\nPlease find attached the invoice (#${orderNum}) for your order.\n\nDelivery is scheduled for ${deliveryDateForEmail} around ${deliveryTimeForEmail}.\n\nWe look forward to serving you!\n\nBest regards,\nSofia & The El Merkury Team`;
    
    console.log(`generatePdfFromSheet: Preparing draft reply to: ${recipient} for order ${orderNum}. Subject: ${subject}`);
    const draft = messageToReplyTo.createDraftReply(body, { 
        htmlBody: body.replace(/\n/g, '<br>'), 
        attachments: [pdfBlob], 
        to: recipient,
    });
    console.log(`generatePdfFromSheet: Draft email created for order ${orderNum}. ID: ${draft.getId()}`);
    return { pdfBlob: pdfBlob, draft: draft, draftId: draft.getId() };
}

/**
 * Clears user properties for the current order and closes the sidebar.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse} An action response to close the add-on.
 */
function handleClearAndClose(e) {
    const orderNum = e.parameters.orderNum;
    if (orderNum) { PropertiesService.getUserProperties().deleteProperty(orderNum); console.log("Cleared data for orderNum: " + orderNum); }
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().popToRoot()).setNotification(CardService.newNotification().setText("Order data cleared. Add-on is ready for the next email.")).build();
}

/**
 * Normalizes a time string to "h:mm a" format (e.g., "7:00 PM").
 * Attempts to parse various common time string inputs.
 * @param {string} timeString The time string to normalize.
 * @returns {string} The normalized time string in "h:mm a" format, or the original trimmed string if parsing/formatting fails.
 */
function _normalizeTimeFormat(timeString) {
  if (!timeString || typeof timeString !== 'string' || timeString.trim() === "") {
    return ""; // Return empty if input is empty or not a string
  }
  const trimmedTime = timeString.trim();

  // Attempt to parse the time string into a Date object.
  // Prepending a fixed date helps JavaScript's Date constructor parse time-only strings.
  // Common formats like "7:00 PM", "7pm", "19:00" are often handled.
  let dateObj = new Date(`01/01/2000 ${trimmedTime}`);

  // Check if the direct parsing worked
  if (isNaN(dateObj.getTime())) {
    // Direct parsing failed, try specific regexes for common patterns
    // to construct the date object more manually.
    let hours = -1;
    let minutes = 0;
    let ampmDesignator = null; // AM/PM

    // Try formats like "7:30pm", "07:30PM", "7:30", "19:30"
    const complexMatch = trimmedTime.match(/^(\d{1,2})(?:[:.](\d{2}))?\s*(AM|PM)?$/i);
    if (complexMatch) {
      hours = parseInt(complexMatch[1]);
      minutes = complexMatch[2] ? parseInt(complexMatch[2]) : 0; // Default to 00 if minutes are not present
      ampmDesignator = complexMatch[3] ? complexMatch[3].toUpperCase() : null;

      // Validate hours and minutes
      if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) {
        return trimmedTime; // Invalid time components
      }

      // If AM/PM is present, adjust hours for 24-hour format for Date constructor
      if (ampmDesignator) {
        if (ampmDesignator === 'PM' && hours > 0 && hours < 12) {
          hours += 12;
        } else if (ampmDesignator === 'AM' && hours === 12) { // Midnight case: 12 AM is 00 hours
          hours = 0;
        }
      }
      // If no AM/PM and hours are like 1-11, it's ambiguous without more context,
      // but Date constructor might assume AM or based on current time.
      // We assume valid 24-hour if AM/PM missing and hours >=0 <=23.
      dateObj = new Date(2000, 0, 1, hours, minutes);
    } else {
      // If no regex matched, return the original trimmed string
      return trimmedTime;
    }
  }

  // If dateObj is now valid (either from direct parse or regex-assisted parse)
  if (!isNaN(dateObj.getTime())) {
    try {
      // Format the valid Date object to the desired "h:mm a" string
      return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'h:mm a');
    } catch (e) {
      // Fallback if formatting fails for an unexpected reason
      console.warn("Could not format dateObj for time: " + trimmedTime + ", Error: " + e.toString());
      return trimmedTime; // Return original trimmed string on formatting error
    }
  }

  // If all parsing attempts failed
  return trimmedTime;
}

/**
 * Attempts to find the best match for an AI-identified flavor string within a list of standard flavor strings.
 * Prioritizes exact match (case-insensitive, trimmed), then checks if a standard flavor includes the AI flavor,
 * then if the AI flavor includes a standard flavor.
 * @param {string} aiFlavorString The flavor string identified by the AI from the email.
 * @param {Array<string>} itemStandardFlavorsArray An array of standard flavor strings for the item.
 * @returns {string|null} The matching standard flavor string from the array, or null if no good match is found.
 */
function _findBestStandardFlavorMatch(aiFlavorString, itemStandardFlavorsArray) {
  if (!aiFlavorString || !itemStandardFlavorsArray || itemStandardFlavorsArray.length === 0) {
    return null;
  }

  const normalizedAiFlavor = aiFlavorString.toLowerCase().trim();

  // Exact match (case-insensitive, trimmed)
  for (const stdFlavor of itemStandardFlavorsArray) {
    if (stdFlavor.toLowerCase().trim() === normalizedAiFlavor) {
      return stdFlavor; // Return the original casing of the standard flavor
    }
  }

  // Check if standard flavor text *contains* the AI flavor text (e.g., Std: "Jalapeno Black Bean (vegan)", AI: "Jalapeno Black Bean")
  for (const stdFlavor of itemStandardFlavorsArray) {
    if (stdFlavor.toLowerCase().trim().includes(normalizedAiFlavor)) {
      return stdFlavor;
    }
  }
  
  // Check if AI flavor text *contains* a standard flavor text (e.g., Std: "Chicken", AI: "Chicken and Cheese")
  // This is less precise, so it's a lower priority match.
  // For this to be useful, we'd want the shortest standard flavor that's a substring.
  let bestSubstringMatch = null;
  for (const stdFlavor of itemStandardFlavorsArray) {
    if (normalizedAiFlavor.includes(stdFlavor.toLowerCase().trim())) {
      if (!bestSubstringMatch || stdFlavor.length > bestSubstringMatch.length) { // Prefer longer (more specific) standard flavor match
        bestSubstringMatch = stdFlavor;
      }
    }
  }
  if (bestSubstringMatch) {
      return bestSubstringMatch;
  }

  // Optional: Add Levenshtein distance or other fuzzy matching here if needed
  // For now, the above checks cover many common cases.

  return null; // No good match found
}

/**
 * Generates an invoice PDF from an HTML template and order data.
 * @param {string} orderNum The order number for naming and logging.
 * @param {object} orderData The complete order data object from UserProperties.
 * @return {GoogleAppsScript.Base.Blob|null} The generated PDF blob, or null if an error occurs.
 */
function generateInvoicePdfFromHtml(orderNum, orderData) {
  console.log(`generateInvoicePdfFromHtml: Starting HTML-based invoice generation for order ${orderNum}`);
  try {
    const htmlTemplate = HtmlService.createTemplateFromFile('invoice.html'); // Name of your HTML file in the project

    // --- Prepare data object for the HTML Template ---
    const invoiceData = {};
    invoiceData.number = orderData.orderNum;
    invoiceData.dateGenerated = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy");
    
    invoiceData.customerName = orderData['Customer Name'] || 'N/A';
    let custAddressParts = [
        orderData['Customer Address Line 1'],
        orderData['Customer Address Line 2'],
        `${orderData['Customer Address City'] || ''}${orderData['Customer Address City'] && (orderData['Customer Address State'] || orderData['Customer Address ZIP']) ? ', ' : ''}${orderData['Customer Address State'] || ''} ${orderData['Customer Address ZIP'] || ''}`.trim()
    ];
    invoiceData.customerAddress = custAddressParts.filter(Boolean).join('\n'); // Filter out empty parts before joining

    // Format delivery date (assuming orderData['Delivery Date'] is 'YYYY-MM-DD')
    invoiceData.deliveryDateFormatted = "Not specified";
    if (orderData['Delivery Date']) {
        try {
            const dateStr = orderData['Delivery Date'];
            let tempDate;
            if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
                const parts = dateStr.split('-');
                tempDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
            } else { tempDate = new Date(_parseDateToMsEpoch(dateStr)); } // Fallback
            if (!isNaN(tempDate.getTime())) {
                 invoiceData.deliveryDateFormatted = Utilities.formatDate(tempDate, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy");
            } else { invoiceData.deliveryDateFormatted = dateStr; }
        } catch (e) { invoiceData.deliveryDateFormatted = orderData['Delivery Date'] || "Not specified"; }
    }
    invoiceData.deliveryTimeFormatted = orderData['Delivery Time'] ? _normalizeTimeFormat(orderData['Delivery Time']) : "Not specified";
    
    // For delivery location, using customer address for now. Adjust if separate fields are used.
    invoiceData.deliveryFullAddress = invoiceData.customerAddress; 

    invoiceData.items = (orderData['ConfirmedQBItems'] || []).map(item => {
      return {
        description: item.original_email_description || item.quickbooks_item_name,
        qty: item.quantity,
        unitPrice: parseFloat(item.unit_price || 0),
        total: (parseFloat(item.quantity) || 0) * (parseFloat(item.unit_price) || 0),
        notes: item.kitchen_notes_and_flavors // Pass kitchen notes to HTML template
      };
    });

    // Calculate Totals for the invoice object
    invoiceData.itemsSubtotal = 0;
    invoiceData.items.forEach(item => invoiceData.itemsSubtotal += item.total);

    invoiceData.tip = parseFloat(orderData['TipAmount'] || 0);
    invoiceData.otherChargesAmount = parseFloat(orderData['OtherChargesAmount'] || 0);
    invoiceData.otherChargesDescription = orderData['OtherChargesDescription'] || "Other Charges";
    
    invoiceData.deliveryFee = parseFloat(BASE_DELIVERY_FEE); // From Constants.gs
    if (orderData['master_delivery_time_ms']) {
        const deliveryHour = new Date(orderData['master_delivery_time_ms']).getHours();
        if (deliveryHour >= DELIVERY_FEE_CUTOFF_HOUR) { // From Constants.gs
            invoiceData.deliveryFee = parseFloat(AFTER_4PM_DELIVERY_FEE); // From Constants.gs
        }
    }
    
    invoiceData.utensilsCost = 0;
    invoiceData.utensilsCount = 0; // For display in template if needed
    if (orderData['Include Utensils?'] === 'Yes') {
        const numUtensils = parseInt(orderData['If yes: how many?']) || 0;
        if (numUtensils > 0) {
            invoiceData.utensilsCount = numUtensils;
            invoiceData.utensilsCost = numUtensils * parseFloat(COST_PER_UTENSIL_SET); // From Constants.gs
        }
    }

    invoiceData.grandTotal = invoiceData.itemsSubtotal + 
                         invoiceData.tip + 
                         invoiceData.otherChargesAmount + 
                         invoiceData.deliveryFee + 
                         invoiceData.utensilsCost;
    // --- End Prepare data for HTML Template ---

    htmlTemplate.invoice = invoiceData; // Pass the data object to the HTML template
    const htmlContent = htmlTemplate.evaluate().getContent();
    console.log(`generateInvoicePdfFromHtml: HTML content generated for order ${orderNum}. Length: ${htmlContent.length}`);

    // Convert HTML to PDF
    const pdfBlob = Utilities.newBlob(htmlContent, MimeType.HTML, `Invoice-${orderNum}.pdf`)
                           .getAs(MimeType.PDF);
    // Set a more descriptive name for the PDF file
    const pdfFileName = `Invoice-${orderNum}-${(orderData['Customer Name'] || 'Unknown').replace(/[^a-zA-Z0-9\s]/g, "_").replace(/\s+/g, "_")}.pdf`;
    pdfBlob.setName(pdfFileName); 
    
    console.log(`generateInvoicePdfFromHtml: PDF blob created from HTML: ${pdfBlob.getName()}, Size: ${pdfBlob.getBytes().length}`);
    return pdfBlob;

  } catch (e) {
    console.error(`Error in generateInvoicePdfFromHtml for order ${orderNum}: ${e.toString()}${(e.stack ? ("\nStack: " + e.stack) : "")}`);
    return null; 
  }
}

/**
 * Prepares the data and populates the HTML invoice template.
 * MODIFIED: Structures data for two-column layout, adds PO, refines notes.
 * @param {string} orderNum The order number.
 * @param {object} orderData The complete order data object.
 * @return {string} The populated HTML string for the invoice.
 * @throws {Error} If the HTML template cannot be created or data is missing.
 */
function getPopulatedInvoiceHtmlForWebApp(orderNum, orderData) {
  console.log(`getPopulatedInvoiceHtmlForWebApp: Preparing HTML for order ${orderNum}`);
  if (!orderData) {
    console.error(`getPopulatedInvoiceHtmlForWebApp: Order data missing for order ${orderNum}.`);
    throw new Error("Order data is missing for HTML invoice.");
  }

  let htmlTemplate;
  try {
    htmlTemplate = HtmlService.createTemplateFromFile('invoice.html');
  } catch (e) {
    console.error(`getPopulatedInvoiceHtmlForWebApp: Failed to create template from 'invoice.html'. Error: ${e.toString()}`);
    throw new Error("Invoice HTML template file not found or error.");
  }
  
  const invoiceData = {};
  invoiceData.number = orderData.orderNum || 'N/A';
  invoiceData.dateGenerated = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy");
  invoiceData.poNumber = orderData['PurchaseOrderNumber'] || ''; // Will be blank if not present in orderData

  invoiceData.customerName = orderData['Customer Name'] || 'N/A';
  invoiceData.customerClient = orderData['Client'] || ''; // Added Client for Billed To
  let custAddressParts = [
      orderData['Customer Address Line 1'],
      orderData['Customer Address Line 2'],
      `${orderData['Customer Address City'] || ''}${orderData['Customer Address City'] && (orderData['Customer Address State'] || orderData['Customer Address ZIP']) ? ', ' : ''}${orderData['Customer Address State'] || ''} ${orderData['Customer Address ZIP'] || ''}`.trim()
  ];
  invoiceData.customerAddress = custAddressParts.filter(Boolean).join('\n');
  invoiceData.customerPhone = _formatPhone(orderData['Customer Address Phone'] || '');
  invoiceData.customerEmail = orderData['Customer Address Email'] || '';

  // Delivery Information
  invoiceData.deliveryContactName = orderData['Contact Person'] || invoiceData.customerName;
  invoiceData.deliveryContactPhone = _formatPhone(orderData['Contact Phone'] || orderData['Customer Address Phone'] || ''); // Use specific delivery phone or fallback
  // For HTML invoice, delivery address is same as customer address unless specific delivery fields are populated and preferred.
  // If you have distinct orderData['Delivery Address Line 1'], etc. fields, use them here.
  invoiceData.deliveryFullAddress = invoiceData.customerAddress; 


  invoiceData.deliveryDateFormatted = "Not specified";
  if (orderData['Delivery Date']) { 
      try {
          const dateStr = orderData['Delivery Date'];
          let tempDate;
          if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
              const parts = dateStr.split('-');
              tempDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
          } else { tempDate = new Date(_parseDateToMsEpoch(dateStr)); }
          if (!isNaN(tempDate.getTime())) {
               invoiceData.deliveryDateFormatted = Utilities.formatDate(tempDate, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy");
          } else { invoiceData.deliveryDateFormatted = dateStr; }
      } catch (e) { invoiceData.deliveryDateFormatted = orderData['Delivery Date']; }
  }
  invoiceData.deliveryTimeFormatted = orderData['Delivery Time'] ? _normalizeTimeFormat(orderData['Delivery Time']) : "Not specified";
  
  invoiceData.items = (orderData['ConfirmedQBItems'] || []).map(item => {
    let customerNotesOnly = '';
    if (item.kitchen_notes_and_flavors) {
      const notesString = item.kitchen_notes_and_flavors;
      // Try to extract only the part after "Customer Notes:"
      const customerNotesMatch = notesString.match(/Customer Notes:\s*(.*)/i);
      if (customerNotesMatch && customerNotesMatch[1]) {
        customerNotesOnly = customerNotesMatch[1].trim();
        // If "Selected Flavors" was also present, ensure it's not part of customerNotesOnly here
        const selectedFlavorsPrefix = "Selected Flavors:";
        if (customerNotesOnly.toLowerCase().startsWith(selectedFlavorsPrefix.toLowerCase())) {
            // This case should ideally not happen if parsing was clean, but as a safeguard
            customerNotesOnly = ""; // Or re-evaluate logic if "Customer Notes:" can appear inside "Selected Flavors:"
        }
      } else if (!notesString.toLowerCase().includes("selected flavors:")) {
        // If "Selected Flavors:" is not present at all, assume the whole string is customer notes
        customerNotesOnly = notesString.trim();
      }
      // If only "Selected Flavors:" is present, customerNotesOnly will remain empty, which is correct.
    }
    return {
      description: item.original_email_description || item.quickbooks_item_name,
      qty: item.quantity,
      unitPrice: parseFloat(item.unit_price || 0),
      total: (parseFloat(item.quantity) || 0) * (parseFloat(item.unit_price) || 0),
      customerNotesOnly: customerNotesOnly 
    };
  });

  invoiceData.itemsSubtotal = 0;
  invoiceData.items.forEach(item => invoiceData.itemsSubtotal += item.total);
  invoiceData.tip = parseFloat(orderData['TipAmount'] || 0);
  invoiceData.otherChargesAmount = parseFloat(orderData['OtherChargesAmount'] || 0);
  invoiceData.otherChargesDescription = orderData['OtherChargesDescription'] || "Other Charges";
  
  const baseDeliveryFee = typeof BASE_DELIVERY_FEE !== 'undefined' ? BASE_DELIVERY_FEE : 0;
  const cutoffHour = typeof DELIVERY_FEE_CUTOFF_HOUR !== 'undefined' ? DELIVERY_FEE_CUTOFF_HOUR : 16;
  const after4PMFee = typeof AFTER_4PM_DELIVERY_FEE !== 'undefined' ? AFTER_4PM_DELIVERY_FEE : 0;
  const costPerUtensil = typeof COST_PER_UTENSIL_SET !== 'undefined' ? COST_PER_UTENSIL_SET : 0;

  invoiceData.deliveryFee = parseFloat(baseDeliveryFee);
  if (orderData['master_delivery_time_ms']) {
      const deliveryHour = new Date(orderData['master_delivery_time_ms']).getHours();
      if (deliveryHour >= cutoffHour) {
          invoiceData.deliveryFee = parseFloat(after4PMFee);
      }
  }
  
  invoiceData.utensilsCost = 0;
  invoiceData.utensilsCount = 0; 
  if (orderData['Include Utensils?'] === 'Yes') {
      const numUtensils = parseInt(orderData['If yes: how many?']) || 0;
      if (numUtensils > 0) {
          invoiceData.utensilsCount = numUtensils;
          invoiceData.utensilsCost = numUtensils * costPerUtensil;
      }
  }
  invoiceData.grandTotal = invoiceData.itemsSubtotal + invoiceData.tip + invoiceData.otherChargesAmount + invoiceData.deliveryFee + invoiceData.utensilsCost;

  htmlTemplate.invoice = invoiceData;
  const htmlContent = htmlTemplate.evaluate().getContent();
  console.log(`getPopulatedInvoiceHtmlForWebApp: HTML content generated for order ${orderNum}. Length: ${htmlContent.length}`);
  return htmlContent;
}

