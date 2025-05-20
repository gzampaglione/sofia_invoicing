// Prompts.gs
// This file contains all Gemini AI prompts used for data extraction and matching.

/**
 * Builds the Gemini prompt for extracting contact and order details from an email body.
 * Differentiates between 'Customer Name' (the main entity for billing) and 'Delivery Contact Person' (who receives the delivery).
 * Ensures delivery date/time are from the body, and asks for all relevant contact info fields.
 * Explicitly distinguishes customer's delivery address from sender's (El Merkury) address.
 * @param {string} body The plain text body of the email.
 * @returns {string} The formatted prompt.
 */
function _buildContactInfoPrompt(body) {
  return 'I will present you with an email or email chain. The subject matter of the email is always the same: it is always a potential customer requesting a catering order to be delivered in the near future from a Philadelphia restaurant named El Merkury, which is owned by Sofia Deleon (also spelled de Leon). Your job is to assist the catering manager of El Merkury with automatically extracting and categorizing information from this email, and then to generate generate an invoice, a kitchen sheet, and a response back to the customer. \n So, lets start with the first step, of extracting information from the email. Please extract the following fields from that email and return ONLY a JSON object with no extra commentary:\n' +
         'Customer Name: (The primary person or organization placing the order for billing/invoice, usually the *individual\'s name* like "Ashley Duchi", not the institution like "University of Pennsylvania". This should NOT be related to elmerkury.com sender.)\n' +
         'Customer Address Line 1: (The street address for DELIVERY. Do NOT use 2104 Chestnut St or any El Merkury address for delivery.)\n' +
         'Customer Address Line 2: (Apt, Suite, Floor for DELIVERY. Do NOT use 2104 Chestnut St or any El Merkury address for delivery.)\n' +
         'Customer Address City: (The city for DELIVERY. Default to Philadelphia unless otherwise specified.)\n' +
         'Customer Address State: (The state for DELIVERY. Default to PA unless otherwise specified.)\n' +
         'Customer Address ZIP: (The ZIP code for DELIVERY. Default to 19103 unless otherwise specified.)\n' +
         'Customer Address Phone: (This is the primary phone number for the *customer placing the order*, typically found in their signature. Format: (XXX) XXX-XXXX if possible)\n' +
         'Customer Address Email: (This is the primary email for the *customer placing the order*, typically found in their signature.)\n' +
         'Delivery Date: (Must be from the email body, not the email header. Extract it from the body, then format as MM/DD/YYYY)\n' +
         'Delivery Time: (Must be from the email body, not the email header. Extract it from the body, then format it as HH:MM AM/PM, e.g., "7:00 PM", "9:30 AM")\n' +
         'Include Utensils?: (If it is clearly specified in the body of the email that the customer wishes utensils to be included, then mark Yes. Otherwise, return Unknown)\n' +
         'If yes: how many?: (If earlier it is clearly specified in the body of the email that the customer wishes utensils to be included, then note how many should be included. Otherwise, return 0)\n' +
         'Delivery Contact Person: (The specific person who will receive the delivery, if different from Customer Name, e.g., "Romina")\n' +
         'Delivery Contact Phone: (The specific phone number for the delivery contact, e.g., Romina\'s direct line)\n' +
         'The email to be analyzed:\n' + body + '\n\n If the email is a chain, please look through all the emails sent by the customer (e.g., the sender of the email did NOT have @elmerkury.com in their domain name). Start with the most recent email as the sourth of truth, and then work backwards if there are missing values. If there are two values within the email chain for the same field, use the value from the more recent email';
}

/**
 * Builds the Gemini prompt for extracting ordered items from an email body.
 * Emphasizes extracting each distinct line as a separate item with quantity and full description,
 * including modifiers, flavors, and sub-quantities.
 * @param {string} body The plain text body of the email.
 * @returns {string} The formatted prompt.
 */
function _buildStructuredItemExtractionPrompt(body) {
  return 'From the email body provided, extract ONLY the ordered items. It is CRITICAL to treat each line in the order section of the email that appears to request a product as a SEPARATE item in the output array. For example, if the email says "1 Small Cheesy Rice" and on a new line "1 Small Cheesy Rice (Vegan)", these must be two distinct entries in the JSON. For each distinct item line, provide its "quantity" as a string and a "description" string that is the clean, full text from that line, including all modifiers, flavors, and sub-quantities mentioned for that specific line. Return as JSON in the format:\n' +
         '{ "Items Ordered": [ { "quantity": "1", "description": "Large Hilacha Chicken" }, { "quantity": "1", "description": "Small Cheesy Rice" }, { "quantity": "1", "description": "Small Cheesy Rice (Vegan)" }, {"quantity": "1", "description": "Large Taquitos Tray (12 Chile Chicken, 20 Chicken and Cheese, 8 Jackfruit)"} ] }\n\n' +
         'Only include food and tray items. Do not include headers, greetings, closings, or other conversational text from the email. Ensure quantities are extracted as strings.\n\n' +
         'Email:\n' + body;
}

/**
 * Builds the Gemini prompt for matching extracted email items to a master QuickBooks item list.
 * Provides detailed instructions for matching, confidence scoring, and handling flavors/details.
 * @param {Array<object>} emailItems An array of objects, each with 'description' and 'quantity' from the email.
 * @param {Array<object>} masterQBItems An array of master QuickBooks item objects.
 * @returns {string} The formatted prompt.
 */
function _buildItemMatchingPrompt(emailItems, masterQBItems) {
  if (!emailItems || emailItems.length === 0) return '';
  if (!masterQBItems || masterQBItems.length === 0) {
    console.warn("Master QB Items list is empty, cannot build detailed item matching prompt. Using fallback.");
    return `Match the following email items to the fallback SKU "${FALLBACK_CUSTOM_ITEM_SKU}":\n` +
           emailItems.map((item, index) => `${index + 1}. Email Line: "${item.description}" (Ordered Quantity: ${item.quantity})`).join('\n') +
           '\nReturn ONLY a JSON array of objects with the format: ' +
           `[ { "original_email_description": "...", "extracted_main_quantity": "...", "matched_qb_item_id": "${FALLBACK_CUSTOM_ITEM_SKU}", "matched_qb_item_name": "Custom Unmatched Item", "match_confidence": "Low", "parsed_flavors_or_details": "...", "identified_flavors": [] } ]`;
  }
  
  const masterItemDetailsForPrompt = masterQBItems.map(item => {
    let detailString = `- Name: "${item.Name}" (SKU: ${item.SKU}, Price: $${item.Price !== undefined ? item.Price.toFixed(2) : '0.00'})`;
    if (item.Category) detailString += ` [Category: ${item.Category}]`;
    if (item.Item && item.Item !== item.Name) detailString += ` [Base Item: ${item.Item}]`;
    if (item.Subtype) detailString += ` [Type: ${item.Subtype}]`;
    if (item.Size) detailString += ` [Size: ${item.Size}]`;
    if (item.Descriptor) detailString += ` [Details: ${item.Descriptor}]`;
    const flavors = [item['Flavor 1'], item['Flavor 2'], item['Flavor 3'], item['Flavor 4'], item['Flavor 5']].filter(f => f && f.toString().trim() !== "").join('; ');
    if (flavors) detailString += ` [Std Flavors: ${flavors}]`;
    return detailString;
  }).join('\n');

  const emailItemDetailsForPrompt = emailItems.map((item, index) => `${index + 1}. Email Line: "${item.description}" (Ordered Quantity: ${item.quantity})`).join('\n');
  
  const fallbackQBItem = masterQBItems.find(item => item.SKU === FALLBACK_CUSTOM_ITEM_SKU);
  const fallbackSkuForPrompt = fallbackQBItem ? fallbackQBItem.SKU : FALLBACK_CUSTOM_ITEM_SKU;
  const fallbackNameForPrompt = fallbackQBItem ? fallbackQBItem.Name : "Custom Unmatched Item";

  const prompt = `
    You are an item matching assistant for a catering business. Match items from an email order to a QuickBooks item list.
    Email Order Items (includes quantity ordered for the line, and full description from email):
    ${emailItemDetailsForPrompt}

    QuickBooks Master Item List (includes item name, SKU, price, category, and other descriptive details):
    ${masterItemDetailsForPrompt}

    For EACH "Email Line" provided:
    1. "original_email_description": The full description text from the email line.
    2. "extracted_main_quantity": The "Ordered Quantity" associated with that email line (as a string).
    3. "matched_qb_item_id": The "SKU" from the QuickBooks Master Item List that is the BEST match for the main item in the email line. This SKU will be used as the QuickBooks Item ID.
    4. "matched_qb_item_name": The "Name" from the QuickBooks Master Item List corresponding to the "matched_qb_item_id" (SKU).
    5. "match_confidence": Your confidence in this match (choose one: High, Medium, Low).
    6. "parsed_flavors_or_details": The full string of specific flavors, sub-quantities, modifiers, or special instructions mentioned in the email line for this item. If none, make this an empty string.
    7. "identified_flavors": An ARRAY of distinct flavor names parsed from "parsed_flavors_or_details". Example: ["3 Chile Chicken (spicy)", "Chicken and Cheese"]. If no distinct flavors are identifiable, return an empty array [].

    Important Matching Rules:
    - If an "Email Line" does not match any specific item well, use the fallback SKU "${fallbackSkuForPrompt}" and name "${fallbackNameForPrompt}" for "matched_qb_item_id" and "matched_qb_item_name", set "match_confidence" to "Low", and "identified_flavors" to an empty array.
    - Prioritize clear matches based on Name, Item, Size, and keywords in Descriptor.
    - Do NOT invent new QuickBooks items, SKUs, or names. Only use SKUs and names from the provided QuickBooks Master Item List.
    Return ONLY a valid JSON array of objects. Ensure "extracted_main_quantity" is a string.

    Example Output Format:
    [ { "original_email_description": "...", "extracted_main_quantity": "...", "matched_qb_item_id": "...", "matched_qb_item_name": "...", "match_confidence": "High", "parsed_flavors_or_details": "...", "identified_flavors": [] } ]`;

  return prompt;
}