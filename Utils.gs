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
 * Parses a date string into milliseconds since epoch. Handles MM/DD/YYYY, Kalimantan-MM-DD, and MM/DD.
 * Defaults to current date if parsing fails. When parsing MM/DD, it intelligently determines the year.
 * @param {string} dateString The date string to parse.
 * @returns {number} The date in milliseconds since epoch.
 */
function _parseDateToMsEpoch(dateString) {
  if (!dateString || typeof dateString !== 'string') { return new Date().getTime(); }
  let date; const currentYear = new Date().getFullYear();
  const now = new Date();

  // Handle common formats
  if (dateString.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) { // MM/DD/YYYY
    date = new Date(dateString);
  } else if (dateString.match(/^\d{4}-\d{2}-\d{2}$/)) { //YYYY-MM-DD
    date = new Date(dateString.replace(/-/g, '/'));
  } else if (dateString.match(/^\d{1,2}\/\d{1,2}$/)) { // MM/DD
    const parts = dateString.split('/');
    const month = parseInt(parts[0]) - 1; // Months are 0-indexed
    const day = parseInt(parts[1]);
    
    date = new Date(currentYear, month, day);

    // If the parsed date is in the past, assume it's for next year
    if (date.getTime() < now.getTime() && (month < now.getMonth() || (month === now.getMonth() && day < now.getDate()))) {
      date.setFullYear(currentYear + 1);
    }
  } else { // Try direct parsing for other formats (e.g., "May 20", "May 20 2025")
    date = new Date(dateString);
    // If direct parsing results in a year far in the past (e.g. 1970 for "May 20"),
    // intelligently set to current or next year.
    if (date && date.getFullYear() < 2000) {
      date.setFullYear(currentYear);
      // If setting to current year still makes it a past date (e.g., "May 20" parsed on May 21),
      // and it's not a leap year issue or something, roll over to next year.
      if (date.getTime() < now.getTime() && (date.getMonth() < now.getMonth() || (date.getMonth() === now.getMonth() && date.getDate() < now.getDate()))) {
        date.setFullYear(currentYear + 1);
      }
    }
  }

  if (isNaN(date.getTime())) { // If still invalid
    return new Date().getTime(); // Default to now
  }
  return date.getTime();
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
 * @param {string} senderEmailField The full sender email string (e.g., "Display Name <email@example.com>").
 * @returns {string} The matched client name or 'Unknown'.
 */
function _matchClient(senderEmailField) {
  if (!senderEmailField) return 'Unknown';
  let emailAddress = senderEmailField; // Assume it's already just the email
  const emailMatch = senderEmailField.match(/<([^>]+)>/);
  if (emailMatch && emailMatch[1]) {
    emailAddress = emailMatch[1];
  }
  
  const emailLower = emailAddress.toLowerCase().trim();
  console.log("Matching client for extracted email: " + emailLower);

  // CLIENT_RULES_LOOKUP is sorted by rule length, descending.
  for (let i = 0; i < CLIENT_RULES_LOOKUP.length; i++) {
    const clientRule = CLIENT_RULES_LOOKUP[i];
    const rule = clientRule.rule.toLowerCase().trim();
    const name = clientRule.clientName.trim();
    if (rule && emailLower.includes(rule)) {  
      console.log("Client matched: " + name + " for email " + emailAddress + " with rule " + rule);
      return name;  
    }
  }
  console.log("No client match for email: " + emailAddress);
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
 * @param {string} orderNum The order number.
 * @returns {{id: string, url: string, name: string}} Object containing new sheet ID, URL, and Name.
 * @throws {Error} If order data or templates are not found, or population fails.
 */
function populateKitchenSheet(orderNum) {
  const userProps = PropertiesService.getUserProperties(); const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { console.error("KitchenSheet: Order data for " + orderNum + " not found."); throw new Error("Order data not found for kitchen sheet: " + orderNum); }
  const orderData = JSON.parse(orderDataString); const confirmedItems = orderData['ConfirmedQBItems']; const masterAllItems = getMasterQBItems();
  if (!confirmedItems || !Array.isArray(confirmedItems)) { console.error("KitchenSheet: Confirmed items not found for order " + orderNum); throw new Error("Confirmed items not found for kitchen sheet generation."); }
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID); const templateSheet = spreadsheet.getSheetByName(KITCHEN_SHEET_TEMPLATE_NAME);
    if (!templateSheet) { console.error("Kitchen sheet template '" + KITCHEN_SHEET_TEMPLATE_NAME + "' not found."); throw new Error("Kitchen sheet template not found."); }
    const newSheetName = `Kitchen - ${orderNum} - ${orderData['Contact Person'] || orderData['Ordering Person Name'] || 'Unknown'}`.substring(0,100); 
    const newSheet = templateSheet.copyTo(spreadsheet).setName(newSheetName);
    
    const customerNameForKitchen = orderData['Contact Person'] || orderData['Customer Name'] || orderData['Ordering Person Name'] || ''; // Prioritize Contact Person
    const contactPhoneForKitchen = _formatPhone(orderData['Contact Phone'] || orderData['Customer Address Phone'] || ''); // Prioritize Contact Phone from form
    newSheet.getRange(KITCHEN_CUSTOMER_PHONE_CELL).setValue(`${customerNameForKitchen} - Ph: ${contactPhoneForKitchen}`);
    
    let deliveryDateFormatted = orderData['Delivery Date'];
    if (deliveryDateFormatted && typeof deliveryDateFormatted === 'string' && !deliveryDateFormatted.includes('/')) { 
        deliveryDateFormatted = Utilities.formatDate(new Date(parseInt(deliveryDateFormatted)), Session.getScriptTimeZone(), "MM/dd/yyyy");
    } else if (deliveryDateFormatted && typeof deliveryDateFormatted === 'string' && deliveryDateFormatted.match(/^\d{4}-\d{2}-\d{2}/)) {
        deliveryDateFormatted = Utilities.formatDate(new Date(deliveryDateFormatted.replace(/-/g, '/')), Session.getScriptTimeZone(), "MM/dd/yyyy");
    }
    newSheet.getRange(KITCHEN_DELIVERY_DATE_CELL).setValue(deliveryDateFormatted || '');
    newSheet.getRange(KITCHEN_DELIVERY_TIME_CELL).setValue(orderData['Delivery Time'] || '');
    
    let currentRow = KITCHEN_ITEM_START_ROW;
    confirmedItems.forEach(item => {
      const masterItem = masterAllItems.find(mi => mi.SKU === item.sku); const itemSize = masterItem ? (masterItem.Size || '') : '';
      newSheet.getRange(KITCHEN_QTY_COL + currentRow).setValue(item.quantity);
      newSheet.getRange(KITCHEN_SIZE_COL + currentRow).setValue(itemSize);
      newSheet.getRange(KITCHEN_ITEM_NAME_COL + currentRow).setValue(item.quickbooks_item_name);
      newSheet.getRange(KITCHEN_FILLING_COL + currentRow).setValue(item.kitchen_notes_and_flavors);
      newSheet.getRange(KITCHEN_NOTES_COL + currentRow).setValue(''); 
      currentRow++;
    });
    SpreadsheetApp.flush();
    return { id: newSheet.getSheetId(), url: newSheet.getParent().getUrl() + '#gid=' + newSheet.getSheetId(), name: newSheetName };
  } catch (e) { console.error("Error in populateKitchenSheet for order " + orderNum + ": " + e.toString() + (e.stack ? ("\nStack: " + e.stack) : "")); throw new Error("Failed to populate kitchen sheet: " + e.message); }
}

/**
 * Populates a new invoice sheet with order details and confirmed items.
 * @param {string} orderNum The order number.
 * @returns {{id: string, url: string, name: string}} Object containing new sheet ID, URL, and Name.
 * @throws {Error} If order data or templates are not found, or population fails.
 */
function populateInvoiceSheet(orderNum) {
  const userProps = PropertiesService.getUserProperties(); const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { console.error("InvoiceSheet: Order data for " + orderNum + " not found."); throw new Error("Order data not found for " + orderNum); }
  const orderData = JSON.parse(orderDataString); const confirmedItems = orderData['ConfirmedQBItems'];
  if (!confirmedItems || !Array.isArray(confirmedItems)) { console.error("InvoiceSheet: Confirmed items not found for order " + orderNum); throw new Error("Confirmed items not found for invoice generation."); }
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID); const templateSheet = spreadsheet.getSheetByName(INVOICE_TEMPLATE_SHEET_NAME);
    if (!templateSheet) { console.error("Invoice template sheet '" + INVOICE_TEMPLATE_SHEET_NAME + "' not found."); throw new Error("Invoice template sheet not found."); }
    const newSheetName = `Invoice - ${orderNum} - ${orderData['Customer Name'] || 'Unknown'}`.substring(0,100); 
    const newSheet = templateSheet.copyTo(spreadsheet).setName(newSheetName);
    
    newSheet.getRange(ORDER_NUM_CELL).setValue(orderData.orderNum);
    newSheet.getRange(CUSTOMER_NAME_CELL).setValue(orderData['Customer Name'] || ''); 
    newSheet.getRange(ADDRESS_LINE_1_CELL).setValue(orderData['Customer Address Line 1'] || '');
    newSheet.getRange(ADDRESS_LINE_2_CELL).setValue(orderData['Customer Address Line 2'] || '');
    
    const cityStateZip = `${orderData['Customer Address City'] || ''}${orderData['Customer Address City'] && (orderData['Customer Address State'] || orderData['Customer Address ZIP']) ? ', ' : ''}${orderData['Customer Address State'] || ''} ${orderData['Customer Address ZIP'] || ''}`.trim();
    newSheet.getRange(CITY_STATE_ZIP_CELL).setValue(cityStateZip);
    
    let deliveryDateFormatted = orderData['Delivery Date'];
    if (deliveryDateFormatted && typeof deliveryDateFormatted === 'string' && !deliveryDateFormatted.includes('/')) { 
        deliveryDateFormatted = Utilities.formatDate(new Date(parseInt(deliveryDateFormatted)), Session.getScriptTimeZone(), "MM/dd/yyyy");
    } else if (deliveryDateFormatted && typeof deliveryDateFormatted === 'string' && deliveryDateFormatted.match(/^\d{4}-\d{2}-\d{2}/)) {
        deliveryDateFormatted = Utilities.formatDate(new Date(deliveryDateFormatted.replace(/-/g, '/')), Session.getScriptTimeZone(), "MM/dd/yyyy");
    }
    newSheet.getRange(DELIVERY_DATE_CELL_INVOICE).setValue(deliveryDateFormatted || '');
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

    // Add Tip and Other Charges to Invoice Sheet
    let tipAmount = orderData['TipAmount'] || 0;
    let otherChargesAmount = orderData['OtherChargesAmount'] || 0;
    let otherChargesDescription = orderData['OtherChargesDescription'] || "Other Charges";

    if (tipAmount > 0) {
        newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow).setValue("Tip").setWrap(false);
        newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setValue(tipAmount).setNumberFormat("$#,##0.00");
        grandTotal += tipAmount;
        currentRow++;
    }
    if (otherChargesAmount > 0) {
        newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow).setValue(otherChargesDescription).setWrap(false);
        newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setValue(otherChargesAmount).setNumberFormat("$#,##0.00");
        grandTotal += otherChargesAmount;
        currentRow++;
    }
    
    // Delivery Fee
    let deliveryFee = BASE_DELIVERY_FEE;
    if (orderData['master_delivery_time_ms']) {
        const deliveryHour = new Date(orderData['master_delivery_time_ms']).getHours();
        if (deliveryHour >= DELIVERY_FEE_CUTOFF_HOUR) {
            deliveryFee = AFTER_4PM_DELIVERY_FEE;
        }
    }
    newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow).setValue("Delivery Fee").setWrap(false);
    newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setValue(deliveryFee).setNumberFormat("$#,##0.00");
    grandTotal += deliveryFee;
    currentRow++;

    // Utensil Costs
    if (orderData['Include Utensils?'] === 'Yes') {
        const numUtensils = parseInt(orderData['If yes: how many?']) || 0;
        if (numUtensils > 0) {
            const utensilTotalCost = numUtensils * COST_PER_UTENSIL_SET;
            newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow).setValue(`Utensils (${numUtensils} sets)`).setWrap(false);
            newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setValue(utensilTotalCost).setNumberFormat("$#,##0.00");
            grandTotal += utensilTotalCost;
            currentRow++;
        }
    }

    const grandTotalDescCell = newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow);
    grandTotalDescCell.setValue("Grand Total:").setFontWeight("bold").setWrap(false);
    const grandTotalValueCell = newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow);
    grandTotalValueCell.setValue(grandTotal).setNumberFormat("$#,##0.00").setFontWeight("bold").setHorizontalAlignment("right").setWrap(false);
    newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow + ":" + ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setBorder(true, null, null, null, null, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    SpreadsheetApp.flush();
    return { id: newSheet.getSheetId(), url: newSheet.getParent().getUrl() + '#gid=' + newSheet.getSheetId(), name: newSheetName };
  } catch (e) { console.error("Error in populateInvoiceSheet for order " + orderNum + ": " + e.toString() + (e.stack ? ("\nStack: " + e.stack) : "")); throw new Error("Failed to populate invoice sheet: " + e.message); }
}

/**
 * Creates a PDF of the populated invoice sheet and prepares a draft email reply with it attached.
 * @param {string} orderNum The order number.
 * @param {string} populatedSheetSpreadsheetId The ID of the spreadsheet containing the populated invoice sheet.
 * @param {string} populatedSheetName The name of the populated invoice sheet.
 * @returns {{pdfBlob: GoogleAppsScript.Base.Blob, draft: GoogleAppsScript.Gmail.GmailDraft, draftId: string}} Object containing PDF blob, draft object, and draft ID.
 * @throws {Error} If order data or sheets are not found, or PDF/email creation fails.
 */
function createPdfAndPrepareEmailReply(orderNum, populatedSheetSpreadsheetId, populatedSheetName) {
  const userProps = PropertiesService.getUserProperties(); const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { console.error("PDFEmail: Order data for " + orderNum + " not found."); throw new Error("Order data not found for PDF/email creation.");}
  const orderData = JSON.parse(orderDataString);
  try {
    console.log(`PDFEmail: Opening spreadsheet ID: ${populatedSheetSpreadsheetId}, Sheet: ${populatedSheetName}`);
    const spreadsheet = SpreadsheetApp.openById(populatedSheetSpreadsheetId); const sheetToExport = spreadsheet.getSheetByName(populatedSheetName);
    if (!sheetToExport) { console.error(`PDFEmail: Sheet "${populatedSheetName}" not found.`); throw new Error("Populated invoice sheet not found for PDF generation."); }
    
    console.log(`PDFEmail: Sheet "${populatedSheetName}" found. Hiding others temporarily for PDF generation.`);
    const allSheets = spreadsheet.getSheets(); const hiddenSheetIds = [];
    allSheets.forEach(sheet => { if (sheet.getSheetId() !== sheetToExport.getSheetId()) { sheet.hideSheet(); hiddenSheetIds.push(sheet.getSheetId()); }});
    SpreadsheetApp.flush(); // Ensure changes are applied before PDF generation
    
    // Export only the specific sheet
    const pdfBlob = sheetToExport.getAs('application/pdf').setName(`${populatedSheetName}.pdf`); 
    console.log(`PDFEmail: PDF blob created: ${pdfBlob.getName()}`);
    
    // Unhide sheets immediately after PDF generation
    hiddenSheetIds.forEach(id => { const sheet = allSheets.find(s => s.getSheetId() === id); if (sheet) sheet.showSheet(); });
    SpreadsheetApp.flush(); // Ensure changes are applied
    console.log("PDFEmail: Sheets unhidden.");
    
    const threadId = orderData.threadId;
    if (!threadId) { console.error("PDFEmail: Thread ID missing for order " + orderNum); throw new Error("Original email thread ID not found."); }
    const thread = GmailApp.getThreadById(threadId);
    if (!thread) { console.error("PDFEmail: Could not retrieve Gmail thread ID: " + threadId); throw new Error("Could not retrieve original email thread."); }
    const messages = thread.getMessages(); const messageToReplyTo = messages[messages.length - 1]; // Reply to the last message in the thread
    
    const recipient = orderData['Contact Email'] || orderData['Customer Address Email'] || orderData['Ordering Person Email']; // Prioritize Contact Email if explicitly set
    const subject = `Re: Catering Order Confirmation - ${orderNum}`;
    const body = `Dear ${orderData['Contact Person'] || orderData['Customer Name'] || 'Valued Customer'},\n\nPlease find attached the invoice for your recent catering order (${orderNum}).\n\nDelivery is scheduled for ${orderData['Delivery Date']} around ${data['Delivery Time']}.\n\nThank you for your business!\n\nBest regards,\n[Your Company Name]`;
    
    console.log(`PDFEmail: Preparing draft reply to: ${recipient}`);
    const draft = messageToReplyTo.createDraftReply(body, { htmlBody: body.replace(/\n/g, '<br>'), attachments: [pdfBlob], to: recipient });
    console.log("PDFEmail: Draft email created. ID: " + draft.getId());
    return { pdfBlob: pdfBlob, draft: draft, draftId: draft.getId() };
  } catch (e) { console.error("Error in createPdfAndPrepareEmailReply for order " + orderNum + ": " + e.toString() + (e.stack ? ("\nStack: " + e.stack) : "")); throw new Error("Failed to create PDF or prepare email: " + e.message); }
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