// V13

/**
 * Parses a JSON string, handling common formatting issues (like markdown code blocks).
 * @param {string} raw The raw string potentially containing JSON.
 * @returns {object|Array} The parsed JSON object or array.
 * @throws {Error} If the string cannot be parsed as JSON.
 */
function _parseJson(raw) {
  if (!raw || typeof raw !== 'string') {
    console.error("Error in _parseJson: Input is not a valid string or is empty. Input: " + raw);
    return {};
  }
  try {
    const cleanedJsonString = raw.replace(/^```json\s*([\s\S]*?)\s*```$/, '$1').trim();
    return JSON.parse(cleanedJsonString);
  } catch (e) {
    console.error("Error in _parseJson (first attempt): " + e.toString() + ". Raw input: " + raw);
    const match = raw.match(/\{[\s\S]*\}|\[[\s\S]*\]/); // Fallback to find any JSON object/array
    if (match && match[0]) {
      try {
        return JSON.parse(match[0]);
      } catch (e2) {
        console.error("Error in _parseJson (fallback attempt): " + e2.toString() + ". Matched string: " + match[0]);
        throw new Error('Invalid JSON response from AI after attempting to clean: ' + raw);
      }
    }
    throw new Error('Invalid JSON response from AI, and no object/array found: ' + raw);
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
 * Parses a date string into milliseconds since epoch. Handles MM/DD/YYYY, YYYY-MM-DD, and MM/DD.
 * Defaults to current date if parsing fails.
 * @param {string} dateString The date string to parse.
 * @returns {number} The date in milliseconds since epoch.
 */
function _parseDateToMsEpoch(dateString) {
  if (!dateString || typeof dateString !== 'string') { return new Date().getTime(); }
  let date; const currentYear = new Date().getFullYear();
  if (dateString.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) { date = new Date(dateString); } // MM/DD/YYYY
  else if (dateString.match(/^\d{4}-\d{2}-\d{2}$/)) { date = new Date(dateString.replace(/-/g, '/')); } //YYYY-MM-DD
  else if (dateString.match(/^\d{1,2}\/\d{1,2}$/)) { // MM/DD (assume current year)
    const parts = dateString.split('/');
    date = new Date(currentYear, parseInt(parts[0]) - 1, parseInt(parts[1]));
  } else { // Try direct parsing for other formats (e.g., "May 20", "May 20 2025")
    date = new Date(dateString);
    // If direct parsing results in a year far in the past (e.g. 1970 for "May 20"), set to current year
    if (date && date.getFullYear() < 2000 && (date.getMonth() < new Date().getMonth() || (date.getMonth() === new Date().getMonth() && date.getDate() < new Date().getDate()))) {
        // If the parsed date is in the current month/day but far past, or past month/day, assume next year
        date.setFullYear(currentYear + 1);
    } else if (date && date.getFullYear() < 2000) {
        // Otherwise, assume current year
        date.setFullYear(currentYear);
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