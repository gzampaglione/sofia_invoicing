// Code.gs
// This is the main script file containing the core workflow logic for the Google Workspace Add-on.
// It orchestrates calls to functions defined in Constants.gs, Prompts.gs, and Utils.gs.

// === HOMEPAGE CARD (Primary Entry Point if manifest is updated) ===
/**
 * Creates the initial homepage card for the add-on, offering workflow options.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object.
 * @returns {GoogleAppsScript.Card_Service.Card} The constructed homepage card.
 */
function createHomepageCard(e) {
  console.log("createHomepageCard triggered. Event: " + JSON.stringify(e));
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Select Workflow"));
  const section = CardService.newCardSection();

  const processEmailAction = CardService.newAction()
    .setFunctionName("buildAddOnCard")
    .setParameters(e && e.messageMetadata ? {messageId: e.messageMetadata.messageId} : (e ? e.parameters : {}));


  section.addWidget(CardService.newTextButton().setText("Process Incoming Catering Email").setOnClickAction(processEmailAction).setTextButtonStyle(CardService.TextButtonStyle.FILLED));
  const pennInvoiceAction = CardService.newAction().setFunctionName("handlePennInvoiceWorkflowPlaceholder");
  section.addWidget(CardService.newTextButton().setText("Process Penn Invoice (Coming Soon)").setOnClickAction(pennInvoiceAction).setDisabled(true));
  card.addSection(section);
  return card.build();
}

/**
 * Placeholder function for the Penn Invoice workflow (feature not yet implemented).
 * @param {GoogleAppsScript.Addons.EventObject} e The event object.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse} An action response showing a placeholder card.
 */
function handlePennInvoiceWorkflowPlaceholder(e) {
  const card = CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("Penn Invoice Workflow"))
    .addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("This feature is under development.")))
    .build();
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(card)).build();
}

// === ENTRY POINT for Catering Email Workflow (can be called from homepage or directly) ===
/**
 * Builds the initial add-on card by parsing the current Gmail message.
 * Extracts contact information, items, and performs preliminary calculations using Gemini API.
 * MODIFIED: Uses AI-extracted customer email for client matching.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object from Gmail.
 * @returns {GoogleAppsScript.Card_Service.Card} The review contact card.
 */
function buildAddOnCard(e) {
  console.log("buildAddOnCard triggered. Event: " + JSON.stringify(e));
  let msgId;

  // Robustly determine the message ID
  if (e && e.messageMetadata && e.messageMetadata.messageId) {
    msgId = e.messageMetadata.messageId;
  } else if (e && e.gmail && e.gmail.messageId) {
    msgId = e.gmail.messageId;
  } else if (e && e.parameters && e.parameters.messageId) {
    msgId = e.parameters.messageId;
  } else {
    const currentEventObject = e.commonEventObject || e;
    if (currentEventObject && currentEventObject.messageMetadata && currentEventObject.messageMetadata.messageId) {
        msgId = currentEventObject.messageMetadata.messageId;
    } else {
        const currentMessage = GmailApp.getCurrentMessage();
        if (currentMessage) {
            msgId = currentMessage.getId();
            console.log("buildAddOnCard: Using current open message ID: " + msgId);
        } else {
            console.error("buildAddOnCard: Could not determine messageId. Event: " + JSON.stringify(e));
            return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("Error"))
              .addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Please open or select a catering order email to process."))).build();
        }
    }
  }
  console.log("Processing message ID: " + msgId);

  const message = GmailApp.getMessageById(msgId);
  const body = message.getPlainBody();
  const senderEmailFull = message.getFrom(); // This is the sender of the *currently viewed* email.
  console.log("buildAddOnCard: Email sender of current message (message.getFrom()):", senderEmailFull);

  // Call Gemini API to extract structured data
  const contactInfoParsed = _parseJson(callGemini(_buildContactInfoPrompt(body)));
  const itemsParsed = _parseJson(callGemini(_buildStructuredItemExtractionPrompt(body)));

  const timestampSuffix = Date.now().toString().slice(-5);
  const randomPrefix = Math.floor(Math.random() * 900 + 100).toString();
  const orderNum = (randomPrefix + timestampSuffix).slice(0,8).padStart(8,'0');

  // MODIFICATION: Determine the email to use for client matching.
  // Prioritize the 'Customer Address Email' extracted by AI from the email body/signature.
  let emailForClientMatching = contactInfoParsed['Customer Address Email'];
  if (!emailForClientMatching || emailForClientMatching.trim() === "") {
    console.warn("buildAddOnCard: 'Customer Address Email' (for client matching) was not found or is empty in AI parse. Client matching may result in 'Unknown'.");
    // If desired, a fallback could be senderEmailFull IF !senderEmailFull.includes('elmerkury.com'),
    // but this could be unreliable if Sofia is forwarding an email from a new client.
    // Sticking to AI-parsed customer email for now. If it's blank, _matchClient will handle it.
  }
  console.log("buildAddOnCard: Email address being used for client matching:", (emailForClientMatching || "''(empty)"));
  const client = _matchClient(emailForClientMatching); // Use the (potentially empty) AI-parsed customer email.
  // END MODIFICATION for client matching email source

  console.log("buildAddOnCard: Matched Client (result from _matchClient): " + client);

  const data = {};
  data['Internal Sender Name'] = _extractNameFromEmail(senderEmailFull); // Still useful for logging/context
  data['Internal Sender Email'] = _extractActualEmail(senderEmailFull);  // Still useful for logging/context

  data['Customer Name'] = contactInfoParsed['Customer Name'] || '';
  if (!data['Customer Name'] && !data['Internal Sender Email'].includes("elmerkury.com")) {
    data['Customer Name'] = data['Internal Sender Name'];
  }
  
  data['Contact Person'] = contactInfoParsed['Delivery Contact Person'] || data['Customer Name'] || data['Internal Sender Name'];
  
  data['Customer Address Phone'] = contactInfoParsed['Customer Address Phone'] || '';
  data['Contact Phone'] = contactInfoParsed['Delivery Contact Phone'] || data['Customer Address Phone'] || '';

  // This is the primary customer email for correspondence and was intended for client matching.
  data['Customer Address Email'] = contactInfoParsed['Customer Address Email'] || ''; 
  data['Contact Email'] = data['Customer Address Email'] || data['Internal Sender Email']; // Fallback for contact email

  data['Customer Address Line 1'] = contactInfoParsed['Customer Address Line 1'] || '';
  data['Customer Address Line 2'] = contactInfoParsed['Customer Address Line 2'] || '';
  data['Customer Address City'] = contactInfoParsed['Customer Address City'] || '';
  data['Customer Address State'] = contactInfoParsed['Customer Address State'] || '';
  data['Customer Address ZIP'] = contactInfoParsed['Customer Address ZIP'] || '';
  data['Delivery Date'] = contactInfoParsed['Delivery Date'] || ''; 
  data['Delivery Time'] = contactInfoParsed['Delivery Time'] || ''; 

  data['Client'] = client; // This will now reflect the match based on Customer Address Email
  data['orderNum'] = orderNum;
  data['messageId'] = msgId;
  data['threadId'] = message.getThread().getId();

  Object.assign(data, contactInfoParsed);
  data['Items Ordered'] = itemsParsed['Items Ordered'] || [];

  let preliminarySubtotal = 0;
  data['Items Ordered'].forEach(item => {
    const priceMatch = item.description.match(/\$(\d+(\.\d{2})?)/);
    const itemPrice = priceMatch ? parseFloat(priceMatch[1]) : 0;
    preliminarySubtotal += (parseInt(item.quantity) || 1) * itemPrice;
  });

  data['PreliminarySubtotalForTip'] = preliminarySubtotal; // Stored for tip explanation
  data['TipAmount'] = (preliminarySubtotal * 0.10);
  console.log("Preliminary Subtotal for Tip: $" + preliminarySubtotal.toFixed(2));
  console.log("Calculated Initial Tip Amount: $" + data['TipAmount'].toFixed(2));

  PropertiesService.getUserProperties().setProperty(orderNum, JSON.stringify(data));
  
  return buildReviewContactCard({ parameters: { orderNum: orderNum } });
}

// === BUILD CUSTOMER CONTACT REVIEW CARD ===
/**
 * Builds the card for reviewing and editing customer and delivery contact details.
 * MODIFIED: Reorganized with thematic bolded headers and dividers;
 * Customer Email/Phone moved under Customer Billing Information.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object containing the order number.
 * @returns {GoogleAppsScript.Card_Service.Card} The constructed review card.
 */
function buildReviewContactCard(e) {
  const orderNum = e.parameters.orderNum;
  const userProps = PropertiesService.getUserProperties();
  const dataString = userProps.getProperty(orderNum);
  if (!dataString) {
    console.error("Error in buildReviewContactCard: Order data for " + orderNum + " not found.");
    return CardService.newCardBuilder().addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Error: Order data not found. Please try again."))).build();
  }
  const data = JSON.parse(dataString); 

  const cardSection = CardService.newCardSection();

  // Order Number (at the top)
  cardSection.addWidget(CardService.newTextParagraph().setText('📋 Order #: <b>' + orderNum + '</b>'));
  cardSection.addWidget(CardService.newDivider());

  // --- Customer Billing Information (includes primary contact) ---
  cardSection.addWidget(CardService.newTextParagraph().setText("<b>Customer Billing & Contact Information:</b>"));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Name').setTitle('Customer Name (for Invoice)').setValue(data['Customer Name'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Client').setTitle('Client Account').setValue(data['Client'] || 'Unknown')); 
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address Email').setTitle('Customer Email').setValue(data['Customer Address Email'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address Phone').setTitle('Customer Phone').setValue(_formatPhone(data['Customer Address Phone'] || '')));
  cardSection.addWidget(CardService.newDivider());

  // --- Delivery Address ---
  cardSection.addWidget(CardService.newTextParagraph().setText("<b>Delivery Address:</b>"));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address Line 1').setTitle('Street Line 1').setValue(data['Customer Address Line 1'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address Line 2').setTitle('Street Line 2 (Apt, Floor, etc.)').setValue(data['Customer Address Line 2'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address City').setTitle('City').setValue(data['Customer Address City'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address State').setTitle('State').setValue(data['Customer Address State'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address ZIP').setTitle('ZIP Code').setValue(data['Customer Address ZIP'] || ''));
  cardSection.addWidget(CardService.newDivider());
  
  // --- Delivery Logistics ---
  cardSection.addWidget(CardService.newTextParagraph().setText("<b>Delivery Date & Time:</b>"));
  cardSection.addWidget(CardService.newDatePicker().setFieldName("Delivery Date").setTitle("Delivery Date").setValueInMsSinceEpoch(_parseDateToMsEpoch(data['Delivery Date'])));
  
  const deliveryTimeInput = CardService.newSelectionInput().setFieldName('Delivery Time').setTitle('Delivery Time');
  if (typeof deliveryTimeInput.setType === 'function') {
    deliveryTimeInput.setType(CardService.SelectionInputType.DROPDOWN); 
  } else {
    console.warn("setType method missing on deliveryTimeInput in buildReviewContactCard.");
  }
  const rawDeliveryTimeFromAI = data['Delivery Time'] || '';
  const normalizedDeliveryTime = _normalizeTimeFormat(rawDeliveryTimeFromAI);
  console.log(`ReviewCard - Order ${orderNum} - Original Delivery Time: "${rawDeliveryTimeFromAI}", Normalized: "${normalizedDeliveryTime}"`);
  const selectedTime = normalizedDeliveryTime; 

  const startHour = 5; const endHour = 23; 
  for (let h = startHour; h <= endHour; h++) { 
    for (let m = 0; m < 60; m += 15) {
      const timeValue = Utilities.formatDate(new Date(2000, 0, 1, h, m), Session.getScriptTimeZone(), 'h:mm a');
      deliveryTimeInput.addItem(timeValue, timeValue, selectedTime === timeValue); 
    }
  }
  cardSection.addWidget(deliveryTimeInput);
  cardSection.addWidget(CardService.newDivider());

  // --- On-Site Delivery Contact Person ---
  cardSection.addWidget(CardService.newTextParagraph().setText("<b>On-Site Delivery Contact Person:</b>"));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Contact Person').setTitle('Name').setValue(data['Contact Person'] || data['Customer Name'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Contact Phone').setTitle('Phone').setValue(_formatPhone(data['Contact Phone'] || data['Customer Address Phone'] || '')));
  cardSection.addWidget(CardService.newDivider());

  // --- Additional Order Options ---
  cardSection.addWidget(CardService.newTextParagraph().setText("<b>Additional Order Options:</b>"));
  const utensilsValue = data['Include Utensils?'] || 'Unknown';
  const utensilsInput = CardService.newSelectionInput().setFieldName('Include Utensils?').setTitle('Include Utensils?');
  if (typeof utensilsInput.setType === 'function') { 
    utensilsInput.setType(CardService.SelectionInputType.DROPDOWN); 
  } else {
    console.warn("setType method missing on utensilsInput in buildReviewContactCard.");
  }
  utensilsInput.addItem('Yes', 'Yes', utensilsValue === 'Yes').addItem('No', 'No', utensilsValue === 'No').addItem('Unknown', 'Unknown', utensilsValue !== 'Yes' && utensilsValue !== 'No');
  cardSection.addWidget(utensilsInput);
  
  let numUtensilsVal = ''; // Initialize value for utensil count
  // Check if 'Include Utensils?' is 'Yes' from current data OR from form input if this is a rebuild
  const showUtensilCount = (data['Include Utensils?'] === 'Yes') || 
                           (e && e.formInput && e.formInput['Include Utensils?'] === 'Yes');

  if (showUtensilCount) {
    if (data['If yes: how many?']) { // Prioritize stored data if available
        numUtensilsVal = data['If yes: how many?'];
    }
    cardSection.addWidget(CardService.newTextInput().setFieldName('If yes: how many?').setTitle('How many utensil sets?').setValue(numUtensilsVal));
  }

  // Footer and Action Button
  const action = CardService.newAction().setFunctionName('handleContactInfoSubmitWithValidation').setParameters({ orderNum: orderNum });
  const footer = CardService.newFixedFooter().setPrimaryButton(CardService.newTextButton().setText('Confirm Contact & Proceed to Items').setOnClickAction(action).setTextButtonStyle(CardService.TextButtonStyle.FILLED));
  
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Step 1: Customer & Order Details'))
    .addSection(cardSection)
    .setFixedFooter(footer)
    .build();
}

/**
 * Handles the submission of contact information with client-side validation.
 * Displays an error card if required fields are missing.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object from the form submission.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse} An action response to update the card or show an error.
 */
function handleContactInfoSubmitWithValidation(e) {
  const inputs = e.commonEventObject.formInputs;
  const orderNum = e.parameters.orderNum;

  // Retrieve values safely using the new helper function from Utils.gs
  const customerName = _getFormInputValue(inputs, 'Customer Name');
  const contactPerson = _getFormInputValue(inputs, 'Contact Person');
  // Use the new field names for retrieval here
  const customerPhone = _getFormInputValue(inputs, 'Customer Address Phone'); // Renamed field
  const customerAddressLine1 = _getFormInputValue(inputs, 'Customer Address Line 1');
  const deliveryDateMs = _getFormInputValue(inputs, 'Delivery Date', true); // Pass true for date inputs
  const deliveryTimeStr = _getFormInputValue(inputs, 'Delivery Time');

  const validationMessages = []; 

  // Validation rules
  if (!deliveryDateMs || !deliveryTimeStr) {
    validationMessages.push("• Delivery Date and Time are required.");
  }
  if (!customerName && !contactPerson) {
    validationMessages.push("• Either 'Customer Name' or 'Delivery Contact Name' is required.");
  }
  // Validate against the 'Customer Phone' field as it's now primary.
  if (!customerPhone && !contactPerson) { // Changed to check Customer Phone primarily
    validationMessages.push("• A 'Customer Phone' or 'Delivery Contact Name' (if no customer phone) is required.");
  }
  if (!customerAddressLine1) {
    validationMessages.push("• 'Delivery Address Line 1' is required.");
  }

  if (validationMessages.length > 0) {
    // If validation fails, build and return an error card
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Validation Error').setImageStyle(CardService.ImageStyle.CIRCLE).setImageUrl('https://fonts.gstatic.com/s/i/googlematerialicons/error/v15/gm_grey_24dp.png'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText('<b>Please correct the following issues:</b><br>' + validationMessages.join('<br>'))))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextButton().setText('Back to Details').setOnClickAction(CardService.newAction().setFunctionName('buildReviewContactCard').setParameters({ orderNum: orderNum })))
      ).build();
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(card)).build();
  }

  // If validation passes, proceed with actual submission
  return handleContactInfoSubmit(e);
}


// === SUBMIT CONTACT INFO & PROCEED TO ITEM MAPPING ===
/**
 * Processes the submitted contact information and moves to the item mapping stage.
 * This function is typically called by `handleContactInfoSubmitWithValidation` after validation passes.
 * MODIFIED: Robustly handles DatePicker output to ensure correct YYYY-MM-DD storage in script's timezone.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object from the form submission.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse} An action response to navigate to the next card.
 */
function handleContactInfoSubmit(e) {
  const inputs = e.commonEventObject.formInputs;
  const orderNum = e.parameters.orderNum;
  console.log(`handleContactInfoSubmit: Processing orderNum: ${orderNum}`);
  
  const newData = {}; 
  let deliveryDateMsFromForm = null; // Raw ms from DatePicker
  let deliveryTimeStrFromForm = '';

  const userProps = PropertiesService.getUserProperties();
  const existingRaw = userProps.getProperty(orderNum);
  if (!existingRaw) { 
    console.error(`Error in handleContactInfoSubmit: Original order data for ${orderNum} not found.`);
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Original order data missing.")).build();
  }
  const existing = JSON.parse(existingRaw);
  console.log(`handleContactInfoSubmit (Order: ${orderNum}): Existing - Delivery Date: "${existing['Delivery Date']}", Time: "${existing['Delivery Time']}", MasterMS: ${existing['master_delivery_time_ms']}`);

  for (const key in inputs) {
    if (key === "Delivery Date") { 
      deliveryDateMsFromForm = _getFormInputValue(inputs, key, true); 
      if (deliveryDateMsFromForm) {
        // The msFromForm often represents UTC midnight of the selected day, or local midnight depending on client.
        // To be safe, extract UTC components of the selected day.
        const datePickerDateObj = new Date(deliveryDateMsFromForm);
        const yearUtc = datePickerDateObj.getUTCFullYear();
        const monthUtc = datePickerDateObj.getUTCMonth(); // 0-11
        const dayUtc = datePickerDateObj.getUTCDate();

        // Create a new Date object using these UTC date components.
        // new Date(year, month, day) constructs the date at midnight in the SCRIPT's local timezone.
        const localDateAtMidnight = new Date(yearUtc, monthUtc, dayUtc);
        
        console.log(`handleContactInfoSubmit (Order: ${orderNum}): DatePicker raw ms: ${deliveryDateMsFromForm}. Interpreted as UTC Date: ${yearUtc}-${String(monthUtc + 1).padStart(2,'0')}-${String(dayUtc).padStart(2,'0')}. Constructed local Date obj for formatting: ${localDateAtMidnight}`);

        // Format this local-midnight Date object into YYYY-MM-DD using the script's timezone.
        newData[key] = Utilities.formatDate(localDateAtMidnight, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        newData[key] = ""; 
        console.warn(`handleContactInfoSubmit (Order: ${orderNum}): Delivery Date from form is null/empty.`);
      }
    } else if (key === "Delivery Time") {
      deliveryTimeStrFromForm = _getFormInputValue(inputs, key); 
      newData[key] = deliveryTimeStrFromForm;
    } else { 
      newData[key] = _getFormInputValue(inputs, key); 
    }
  }
  console.log(`handleContactInfoSubmit (Order: ${orderNum}): Values from Form (newData) - Delivery Date: "${newData['Delivery Date']}", Delivery Time: "${newData['Delivery Time']}"`);
  // deliveryDateMsFromForm here is the raw MS from the picker, which we've now processed above into newData['Delivery Date']
  // For master_delivery_time_ms, we need an epoch MS that represents local midnight + time.
  // So, use the _parseDateToMsEpoch on the newly formatted YYYY-MM-DD string.
  let epochMsForMasterCalc;
  if (newData['Delivery Date'] && newData['Delivery Date'].match(/^\d{4}-\d{2}-\d{2}$/)) {
      epochMsForMasterCalc = _parseDateToMsEpoch(newData['Delivery Date']);
  } else {
      epochMsForMasterCalc = deliveryDateMsFromForm; // Fallback to raw if format is unexpected
  }


  const merged = { ...existing, ...newData };

  if (epochMsForMasterCalc && deliveryTimeStrFromForm && deliveryTimeStrFromForm.trim() !== "") {
    merged['master_delivery_time_ms'] = _combineDateAndTime(epochMsForMasterCalc, deliveryTimeStrFromForm);
    console.log(`handleContactInfoSubmit (Order: ${orderNum}): Recalculated master_delivery_time_ms using epoch ${epochMsForMasterCalc} and time "${deliveryTimeStrFromForm}": ${merged['master_delivery_time_ms']}`);
  } else {
    console.warn(`handleContactInfoSubmit (Order: ${orderNum}): Could not recalculate master_delivery_time_ms. epochMsForMasterCalc: ${epochMsForMasterCalc}, TimeStr: "${deliveryTimeStrFromForm}".`);
 }
  console.log(`handleContactInfoSubmit (Order: ${orderNum}): Merged Data - Delivery Date: "${merged['Delivery Date']}", Time: "${merged['Delivery Time']}", MasterMS: ${merged['master_delivery_time_ms']}`);

  if (!merged['Items Ordered'] || !Array.isArray(merged['Items Ordered']) || (merged['Items Ordered'].length > 0 && typeof merged['Items Ordered'][0].description === 'undefined')) {
    console.log(`handleContactInfoSubmit (Order: ${orderNum}): Re-running item extraction.`);
    const message = GmailApp.getMessageById(merged.messageId);
    const body = message.getPlainBody();
    const itemsParsed = _parseJson(callGemini(_buildStructuredItemExtractionPrompt(body)));
    merged['Items Ordered'] = itemsParsed['Items Ordered'] || [];
  }

  PropertiesService.getUserProperties().setProperty(orderNum, JSON.stringify(merged));
  console.log(`handleContactInfoSubmit (Order: ${orderNum}): Data saved. Proceeding to item mapping.`);
  
  const itemMappingCard = buildItemMappingAndPricingCard({ parameters: { orderNum: orderNum } });
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(itemMappingCard)).build();
}

// === ITEM MAPPING AND PRICING CARD ===
/**
 * Builds the card for mapping extracted email items to QuickBooks items and reviewing pricing.
 * MODIFIED: "Customer Notes" field now pre-filled with item.parsed_flavors_or_details.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object containing the order number.
 * @returns {GoogleAppsScript.Card_Service.Card} The constructed item mapping card.
 */
function buildItemMappingAndPricingCard(e) {
  const orderNum = e.parameters.orderNum;
  const userProps = PropertiesService.getUserProperties();
  const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { 
    console.error("Error in buildItemMappingAndPricingCard: Order data for " + orderNum + " not found.");
    return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("Error")).addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Order data not found for item mapping. Please restart."))).build();
  }
  const orderData = JSON.parse(orderDataString);
  const emailItems = orderData['Items Ordered']; 
  if (!emailItems || !Array.isArray(emailItems) || (emailItems.length > 0 && (typeof emailItems[0] !== 'object' || typeof emailItems[0].description === 'undefined'))) { 
    console.error("Error: 'Items Ordered' for " + orderNum + " is missing or not in expected format: " + JSON.stringify(emailItems));
    return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("Item Loading Error")).addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Could not load items for matching. Please try reopening."))).build();
  }
  const masterQBItems = getMasterQBItems();
  if (masterQBItems.length === 0) { 
    console.error("Error: Master item list ('Item Lookup' sheet) is empty or could not be loaded.");
    return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("Configuration Error")).addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Master item list ('Item Lookup') could not be loaded. Check config."))).build();
  }

  const suggestedMatches = getGeminiItemMatches(emailItems, masterQBItems);
  
  orderData['tempSuggestedMatches'] = suggestedMatches; 
  PropertiesService.getUserProperties().setProperty(orderNum, JSON.stringify(orderData));

  const card = CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle(`Step 2: Map Items for Order ${orderNum}`));
  const headerSection = CardService.newCardSection();
  if (suggestedMatches.length > 0) {
    headerSection.addWidget(CardService.newTextParagraph().setText(`AI detected ${suggestedMatches.length} potential item(s) from the email.`));
  } else {
    headerSection.addWidget(CardService.newTextParagraph().setText("No items were automatically extracted by AI. Please add items manually."));
  }
  card.addSection(headerSection);

  const itemsDisplaySection = CardService.newCardSection();

  if (suggestedMatches.length > 0) {
    suggestedMatches.forEach((item, index) => {
      const masterMatchDetails = masterQBItems.find(master => master.SKU === item.matched_qb_item_id) || 
                               masterQBItems.find(master => master.SKU === FALLBACK_CUSTOM_ITEM_SKU);
      
      itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`<b>Email Line ${index + 1}:</b> "${item.original_email_description}"`));
      itemsDisplaySection.addWidget(CardService.newTextInput().setFieldName(`item_qty_${index}`).setTitle('Unit Quantity').setValue(item.extracted_main_quantity || '1'));
      itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`AI Match Confidence: ${item.match_confidence || 'N/A'}`));
      
      const qbItemDropdown = CardService.newSelectionInput().setFieldName(`item_qb_sku_${index}`).setTitle('QuickBooks Item');
      if (qbItemDropdown && typeof qbItemDropdown.setType === 'function') {
        qbItemDropdown.setType(CardService.SelectionInputType.DROPDOWN); 
      } else { console.warn(`setType method missing on qbItemDropdown for item ${index}.`); }
      
      let preSelected = false;
      masterQBItems.forEach(masterItem => {
        if (!masterItem.SKU) return;
        const isSelected = masterItem.SKU === item.matched_qb_item_id;
        if (isSelected) preSelected = true;
        qbItemDropdown.addItem(`${masterItem.Name} (SKU: ${masterItem.SKU})`, masterItem.SKU, isSelected);
      });
      if (!preSelected && item.matched_qb_item_id !== FALLBACK_CUSTOM_ITEM_SKU && masterQBItems.length > 0) {
        const fallbackInList = masterQBItems.find(mi => mi.SKU === FALLBACK_CUSTOM_ITEM_SKU);
        if (fallbackInList) { qbItemDropdown.addItem(`${fallbackInList.Name} (SKU: ${fallbackInList.SKU})`, fallbackInList.SKU, true); }
      }
      itemsDisplaySection.addWidget(qbItemDropdown);

      itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`<b>Customer Flavors (from email):</b><br><font color="#555555"><i>${item.parsed_flavors_or_details || "None specified"}</i></font>`));
      
      const aiIdentifiedFlavors = item.identified_flavors || [];
      let displayedFlavorDropdownCount = 0;
      let unmappedAiFlavorsForNotes = []; // Store AI flavors that didn't map to a dropdown

      if (masterMatchDetails && masterMatchDetails.SKU !== FALLBACK_CUSTOM_ITEM_SKU) {
        const itemStandardFlavorsArray = [
          masterMatchDetails['Flavor 1'], masterMatchDetails['Flavor 2'],
          masterMatchDetails['Flavor 3'], masterMatchDetails['Flavor 4'],
          masterMatchDetails['Flavor 5']
        ].filter(Boolean).map(f => f.toString().trim());

        if (itemStandardFlavorsArray.length > 0) {
          aiIdentifiedFlavors.forEach((aiFlavor, aiFlavorIndex) => { 
            const matchedStdFlavor = _findBestStandardFlavorMatch(aiFlavor, itemStandardFlavorsArray);
            if (matchedStdFlavor) {
              displayedFlavorDropdownCount++;
              const flavorDropdown = CardService.newSelectionInput()
                .setFieldName(`item_${index}_requested_flavor_dropdown_${aiFlavorIndex}`) 
                .setTitle(`Flavor Choice ${displayedFlavorDropdownCount} (AI found: "${aiFlavor}")`);
              if (typeof flavorDropdown.setType === 'function') {
                flavorDropdown.setType(CardService.SelectionInputType.DROPDOWN);
              } else { console.warn(`setType missing for flavor dropdown ${index}-${aiFlavorIndex}`); }
              flavorDropdown.addItem("-- Select a Standard Flavor --", "", !matchedStdFlavor); 
              itemStandardFlavorsArray.forEach(stdOpt => {
                flavorDropdown.addItem(stdOpt, stdOpt, stdOpt === matchedStdFlavor);
              });
              itemsDisplaySection.addWidget(flavorDropdown);
            } else {
              unmappedAiFlavorsForNotes.push(aiFlavor); // Collect unmapped AI flavors
            }
          });
        } else { 
          unmappedAiFlavorsForNotes = unmappedAiFlavorsForNotes.concat(aiIdentifiedFlavors);
        }
      } else { 
        unmappedAiFlavorsForNotes = unmappedAiFlavorsForNotes.concat(aiIdentifiedFlavors);
      }
      
      // MODIFICATION: Pre-fill "Customer Notes" with item.parsed_flavors_or_details
      // This contains the customer's original phrasing for flavors/notes for this item.
      // Any unmapped AI flavors are usually part of this string already if AI extracted them correctly.
      let customerNotesValue = item.parsed_flavors_or_details || '';
      // If parsed_flavors_or_details is empty but we had unmapped AI ones, use those as a fallback.
      if (!customerNotesValue && unmappedAiFlavorsForNotes.length > 0) {
          customerNotesValue = "Unmatched AI Flavors: " + unmappedAiFlavorsForNotes.join(', ');
      }

      itemsDisplaySection.addWidget(CardService.newTextInput()
        .setFieldName(`item_${index}_custom_notes`)
        .setTitle("Customer Notes (original details / additional requests)") // Updated title slightly
        .setValue(customerNotesValue) // Use the original parsed details
        .setMultiline(true));
      // END MODIFICATION

      const initialPrice = masterMatchDetails ? (masterMatchDetails.Price !== undefined ? masterMatchDetails.Price : 0) : 0;
      itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`<b>Unit Price:</b> $${initialPrice.toFixed(2)}`));
      
      const removeItemSwitch = CardService.newSwitch()
        .setFieldName(`item_remove_${index}`)
        .setValue("true") 
        .setControlType(CardService.SwitchControlType.SWITCH)
        .setSelected(false); 

      const removeItemDecoratedText = CardService.newDecoratedText()
        .setText("Remove this item?") 
        .setSwitchControl(removeItemSwitch);
      itemsDisplaySection.addWidget(removeItemDecoratedText);
      
      itemsDisplaySection.addWidget(CardService.newDivider());
    }); 
  } 
  card.addSection(itemsDisplaySection);

  const manualAddSection = CardService.newCardSection().setHeader("Manually Add New Item").setCollapsible(true);
  const newItemDropdown = CardService.newSelectionInput().setFieldName("new_item_qb_sku").setTitle("Select Item");
  if (newItemDropdown && typeof newItemDropdown.setType === 'function') { 
    newItemDropdown.setType(CardService.SelectionInputType.DROPDOWN); 
  } else { console.warn("setType method missing on newItemDropdown."); }
  newItemDropdown.addItem("--- Select Item to Add ---", "", true);
  masterQBItems.forEach(masterItem => {
    if (!masterItem.SKU) return;
    newItemDropdown.addItem(`${masterItem.Name} (SKU: ${masterItem.SKU})`, masterItem.SKU, false);
  });
  manualAddSection.addWidget(newItemDropdown);
  manualAddSection.addWidget(CardService.newTextInput().setFieldName("new_item_qty").setTitle("Unit Quantity").setValue("1"));
  manualAddSection.addWidget(CardService.newTextInput().setFieldName("new_item_kitchen_notes").setTitle("Flavors/Notes for Kitchen (Manual Add)").setMultiline(true)); 
  manualAddSection.addWidget(CardService.newTextParagraph().setText("Unit Price will be based on selected SKU."));
  card.addSection(manualAddSection);

  const additionalChargesSection = CardService.newCardSection().setHeader("Additional Charges"); 
  additionalChargesSection.addWidget(CardService.newTextInput().setFieldName("tip_amount").setTitle("Tip Amount ($)").setValue((orderData['TipAmount'] || 0).toFixed(2)));
  const initialTipAmount = parseFloat(orderData['TipAmount'] || 0);
  if (initialTipAmount > 0 && orderData['PreliminarySubtotalForTip'] !== undefined) {
    const prelimSubtotal = parseFloat(orderData['PreliminarySubtotalForTip']);
    if (!isNaN(prelimSubtotal)) { 
        additionalChargesSection.addWidget(CardService.newTextParagraph()
            .setText(`<font color="#555555" size="1"><i>(Initial 10% tip suggestion of $${initialTipAmount.toFixed(2)} was based on an email-parsed subtotal of $${prelimSubtotal.toFixed(2)}. You can adjust.)</i></font>`));
    }
  }
  additionalChargesSection.addWidget(CardService.newTextInput().setFieldName("other_charges_amount").setTitle("Other Charges Amount ($)").setValue("0.00"));
  additionalChargesSection.addWidget(CardService.newTextInput().setFieldName("other_charges_description").setTitle("Other Charges Description"));
  card.addSection(additionalChargesSection);

  const action = CardService.newAction().setFunctionName('handleItemMappingSubmit')
    .setParameters({ 
        orderNum: orderNum, 
        ai_item_count: suggestedMatches.length.toString()
    });
  card.setFixedFooter(CardService.newFixedFooter().setPrimaryButton(CardService.newTextButton().setText('Confirm All Items & Proceed').setOnClickAction(action).setTextButtonStyle(CardService.TextButtonStyle.FILLED)));
  return card.build();
}

/**
 * Handles the submission of item mapping and pricing, calculates final totals.
 * MODIFIED: Reads "remove item" from Switch; processes new flavor UI.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object from the form submission.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse} An action response to navigate to the next card.
 */
function handleItemMappingSubmit(e) {
  const formInputs = e.formInputs || (e.commonEventObject && e.commonEventObject.formInputs);
  if (!formInputs) { console.error("Error: formInputs is undefined in handleItemMappingSubmit."); return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Could not read form data.")).build(); }
  
  const orderNum = e.parameters.orderNum;
  const aiItemCount = parseInt(e.parameters.ai_item_count) || 0;

  if (!orderNum) { console.error("Error: Order number is missing in handleItemMappingSubmit."); return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Order number missing.")).build(); }
  
  const userProps = PropertiesService.getUserProperties();
  const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { console.error("Error: Original order data for " + orderNum + " not found in handleItemMappingSubmit."); return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Original order data not found.")).build(); }
  
  const orderData = JSON.parse(orderDataString);
  const masterQBItems = getMasterQBItems();
  const confirmedQuickBooksItems = [];

  const suggestedMatchesFromStorage = orderData['tempSuggestedMatches'] || [];

  for (let i = 0; i < aiItemCount; i++) { 
    // MODIFICATION: Reading value from a Switch.
    // A Switch that is "on" and has a value="true" will submit that value.
    // If it's "off", the fieldName might be absent or have a different value depending on how unselected state is handled.
    // If .setValue("true") is used, when ON, formInputs[`item_remove_${i}`] = ["true"]. If OFF, it might be absent.
    const removeSwitchFieldName = `item_remove_${i}`;
    const removeThisItem = (formInputs[removeSwitchFieldName] && formInputs[removeSwitchFieldName][0] === "true");
    // END MODIFICATION for remove item switch

    if (removeThisItem) { console.log(`Item ${i} (AI) marked for removal via switch.`); continue; }

    const qtyString = formInputs[`item_qty_${i}`] && formInputs[`item_qty_${i}`][0];
    const qbItemSKU = formInputs[`item_qb_sku_${i}`] && formInputs[`item_qb_sku_${i}`][0]; 
    
    const currentProcessedItem = suggestedMatchesFromStorage[i] || {}; 
    const originalDescription = currentProcessedItem.original_email_description || "N/A";
    const aiIdentifiedFlavorsForItem = currentProcessedItem.identified_flavors || [];

    let finalKitchenNotesParts = [];
    let selectedDropdownFlavorValues = [];

    aiIdentifiedFlavorsForItem.forEach((originalAiFlavor, aiFlavorIndex) => {
      const dropdownFieldName = `item_${i}_requested_flavor_dropdown_${aiFlavorIndex}`;
      // Check if the dropdown was actually rendered (i.e., AI flavor was matched to standard)
      // and if a value was selected from it.
      if (formInputs[dropdownFieldName] && formInputs[dropdownFieldName][0] && formInputs[dropdownFieldName][0] !== "") {
        selectedDropdownFlavorValues.push(formInputs[dropdownFieldName][0]);
      }
    });

    if (selectedDropdownFlavorValues.length > 0) {
      finalKitchenNotesParts.push("Selected Flavors: " + selectedDropdownFlavorValues.join(', '));
    }
    
    const customNotesFieldName = `item_${i}_custom_notes`;
    const customNotesValue = (formInputs[customNotesFieldName] && formInputs[customNotesFieldName][0]) ? formInputs[customNotesFieldName][0].trim() : "";
    if (customNotesValue !== "") {
      finalKitchenNotesParts.push("Customer Notes: " + customNotesValue);
    }
    
    if (finalKitchenNotesParts.length === 0 && currentProcessedItem.parsed_flavors_or_details) {
        finalKitchenNotesParts.push(currentProcessedItem.parsed_flavors_or_details);
    }
    let kitchenNotesAndFlavors = finalKitchenNotesParts.join('; ').trim();
    if (!kitchenNotesAndFlavors && currentProcessedItem.parsed_flavors_or_details) {
        kitchenNotesAndFlavors = currentProcessedItem.parsed_flavors_or_details;
    }

    if (qbItemSKU && qtyString) {
      const masterItemDetails = masterQBItems.find(master => master.SKU === qbItemSKU);
      let unitPrice = 0;
      let itemName = "Custom Item";
      let itemSKU = qbItemSKU; 

      if (masterItemDetails) {
        unitPrice = masterItemDetails.Price || 0;
        itemName = masterItemDetails.Name;
      } else { 
        console.warn("Master details not found for SKU: " + qbItemSKU + " in submit. Defaulting to fallback.");
        itemSKU = FALLBACK_CUSTOM_ITEM_SKU;
        const fallbackMasterItem = masterQBItems.find(master => master.SKU === FALLBACK_CUSTOM_ITEM_SKU);
        unitPrice = fallbackMasterItem ? (fallbackMasterItem.Price || 0) : 0;
        itemName = fallbackMasterItem ? fallbackMasterItem.Name : "Custom Unmatched Item";
      }
      confirmedQuickBooksItems.push({ 
          quickbooks_item_id: itemSKU, 
          quickbooks_item_name: itemName, 
          sku: itemSKU, 
          quantity: parseInt(qtyString) || 1, 
          unit_price: unitPrice, 
          kitchen_notes_and_flavors: kitchenNotesAndFlavors, 
          original_email_description: originalDescription 
      });
    }
  } 

  const newItemQbSKU = formInputs.new_item_qb_sku && formInputs.new_item_qb_sku[0];
  const newItemQtyString = formInputs.new_item_qty && formInputs.new_item_qty[0];
  if (newItemQbSKU && newItemQbSKU !== "" && newItemQtyString) {
    const manMasterItemDetails = masterQBItems.find(master => master.SKU === newItemQbSKU);
    if (manMasterItemDetails) {
      const unitPrice = manMasterItemDetails.Price || 0;
      const newItemKitchenNotes = (formInputs.new_item_kitchen_notes && formInputs.new_item_kitchen_notes[0]) || "";
      confirmedQuickBooksItems.push({ 
          quickbooks_item_id: newItemQbSKU, 
          quickbooks_item_name: manMasterItemDetails.Name, 
          sku: manMasterItemDetails.SKU, 
          quantity: parseInt(newItemQtyString) || 1, 
          unit_price: unitPrice, 
          kitchen_notes_and_flavors: newItemKitchenNotes, 
          original_email_description: null 
      });
    }
  }

  delete orderData['tempSuggestedMatches']; 

  orderData['ConfirmedQBItems'] = confirmedQuickBooksItems; 
  orderData['TipAmount'] = parseFloat(formInputs.tip_amount && formInputs.tip_amount[0]) || 0;
  orderData['OtherChargesAmount'] = parseFloat(formInputs.other_charges_amount && formInputs.other_charges_amount[0]) || 0;
  orderData['OtherChargesDescription'] = (formInputs.other_charges_description && formInputs.other_charges_description[0]) || "";

  PropertiesService.getUserProperties().setProperty(orderNum, JSON.stringify(orderData));
  console.log("Confirmed items and charges for order " + orderNum + ": " + JSON.stringify(orderData));
  const invoiceActionsCard = buildInvoiceActionsCard({ parameters: { orderNum: orderNum } });
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(invoiceActionsCard)).build();
}

/**
 * Builds the final review and actions card before document generation.
 * Summarizes order details, confirmed items, and charges.
 * MODIFIED: Updates button text and order.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object.
 * @returns {GoogleAppsScript.Card_Service.Card} The constructed card.
 */
function buildInvoiceActionsCard(e) {
  const orderNum = e.parameters.orderNum;
  console.log(`buildInvoiceActionsCard: Building for orderNum: ${orderNum}`);
  const card = CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle(`Step 3: Final Review & Actions for ${orderNum}`));
  const orderDetailsSection = CardService.newCardSection().setHeader("Order Details");
  const userProps = PropertiesService.getUserProperties(); 
  const orderDataString = userProps.getProperty(orderNum);

  if (orderDataString) {
    const orderData = JSON.parse(orderDataString);
    console.log(`buildInvoiceActionsCard (Order: ${orderNum}): Read from props - Stored Delivery Date: "${orderData['Delivery Date']}", Time: "${orderData['Delivery Time']}"`);

    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Customer Name (for Invoice):</b> " + (orderData['Customer Name'] || 'N/A') + "</i>"));
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Customer Phone:</b> " + _formatPhone(orderData['Customer Address Phone'] || 'N/A') + "</i>"));
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Customer Email:</b> " + (orderData['Customer Address Email'] || 'N/A') + "</i>"));
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Delivery Contact:</b> " + (orderData['Contact Person'] || 'N/A') + "</i>"));
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Delivery Phone:</b> " + _formatPhone(orderData['Contact Phone'] || 'N/A') + "</i>"));

    let address = "<i><b>Address:</b> ";
    if (orderData['Customer Address Line 1']) address += orderData['Customer Address Line 1'];
    if (orderData['Customer Address Line 2']) address += (address === "<i><b>Address:</b> " ? "" : ", ") + orderData['Customer Address Line 2'];
    if (orderData['Customer Address City']) address += (address === "<i><b>Address:</b> " ? "" : ", ") + orderData['Customer Address City'];
    if (orderData['Customer Address State']) address += (address === "<i><b>Address:</b> " ? "" : ", ") + orderData['Customer Address State'];
    if (orderData['Customer Address ZIP']) address += " " + orderData['Customer Address ZIP'];
    address += "</i>";
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText(address.length <= "<i><b>Address:</b> </i>".length + 1 ? "<i><b>Address:</b> N/A</i>" : address));
    
    let deliveryDateForDisplay = "N/A";
    if (orderData['Delivery Date'] && orderData['Delivery Date'].match(/^\d{4}-\d{2}-\d{2}$/)) {
        const parts = orderData['Delivery Date'].split('-');
        const dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        if (!isNaN(dateObj.getTime())) {
            deliveryDateForDisplay = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");
        } else {
            deliveryDateForDisplay = orderData['Delivery Date']; 
        }
    } else if (orderData['Delivery Date']) {
        deliveryDateForDisplay = orderData['Delivery Date']; 
    }
    const deliveryTimeForDisplay = orderData['Delivery Time'] ? _normalizeTimeFormat(orderData['Delivery Time']) : 'N/A';
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Delivery:</b> " + deliveryDateForDisplay + " at " + deliveryTimeForDisplay + "</i>"));
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Client:</b> " + (orderData['Client'] || 'N/A') + "</i>")); 
  } else { 
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("Could not retrieve order details.")); 
  }
  card.addSection(orderDetailsSection);

  const itemSummarySection = CardService.newCardSection().setHeader("Confirmed Items Summary");
  let subTotalForCharges = 0; 
  if (orderDataString) {
    const orderData = JSON.parse(orderDataString);
    const confirmedItems = orderData['ConfirmedQBItems'];
    if (confirmedItems && Array.isArray(confirmedItems) && confirmedItems.length > 0) {
      confirmedItems.forEach(item => {
        const itemTotal = (item.quantity || 0) * (item.unit_price || 0);
        subTotalForCharges += itemTotal;
        itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>${item.original_email_description || item.quickbooks_item_name}</b><br>Qty: ${item.quantity}, Unit Price: $${(item.unit_price || 0).toFixed(2)}, Total: $${itemTotal.toFixed(2)}${item.kitchen_notes_and_flavors ? '<br><font color="#666666"><i>Flavor/Custom Notes: ' + item.kitchen_notes_and_flavors + '</i></font>' : ''}`));
      });
      itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Items Subtotal: $${subTotalForCharges.toFixed(2)}</b>`));
    }
    
    if (confirmedItems && confirmedItems.length > 0 && (orderData['TipAmount'] > 0 || orderData['OtherChargesAmount'] > 0 || (orderData['Include Utensils?'] === 'Yes' && parseInt(orderData['If yes: how many?']) > 0 ) || BASE_DELIVERY_FEE > 0 )) { // Added check for utensil number and delivery fee
        itemSummarySection.addWidget(CardService.newDivider());
    }
    
    let grandTotal = subTotalForCharges; 

    if (orderData['TipAmount'] > 0) {
        itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Tip:</b> $${orderData['TipAmount'].toFixed(2)}`));
        grandTotal += orderData['TipAmount'];
    }
    if (orderData['OtherChargesAmount'] > 0) {
        itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>${orderData['OtherChargesDescription'] || 'Other Charges'}:</b> $${orderData['OtherChargesAmount'].toFixed(2)}`));
        grandTotal += orderData['OtherChargesAmount'];
    }
    
    let deliveryFee = BASE_DELIVERY_FEE;
    if (orderData['master_delivery_time_ms']) { 
        const deliveryHour = new Date(orderData['master_delivery_time_ms']).getHours();
        console.log(`buildInvoiceActionsCard (Order: ${orderNum}): Delivery hour for fee calc: ${deliveryHour} from master_ms: ${orderData['master_delivery_time_ms']}`);
        if (deliveryHour >= DELIVERY_FEE_CUTOFF_HOUR) {
            deliveryFee = AFTER_4PM_DELIVERY_FEE;
        }
    } else {
        console.warn(`buildInvoiceActionsCard (Order: ${orderNum}): master_delivery_time_ms not found for fee calculation. Using base fee.`);
    }
    // Only add delivery fee to summary if it's greater than 0 (or always show it)
    if (deliveryFee > 0) { 
      itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Delivery Fee:</b> $${deliveryFee.toFixed(2)}`));
      grandTotal += deliveryFee;
    }


    if (orderData['Include Utensils?'] === 'Yes') {
        const numUtensils = parseInt(orderData['If yes: how many?']) || 0;
        if (numUtensils > 0) {
            const utensilTotalCost = numUtensils * COST_PER_UTENSIL_SET;
            itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Utensils (${numUtensils} sets):</b> $${utensilTotalCost.toFixed(2)}`));
            grandTotal += utensilTotalCost;
        }
    }

    if (grandTotal > 0 || subTotalForCharges > 0) { 
      itemSummarySection.addWidget(CardService.newDivider()); 
      itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Estimated Grand Total: $${grandTotal.toFixed(2)}</b>`));
    } else if (!(confirmedItems && confirmedItems.length > 0)) { 
      itemSummarySection.addWidget(CardService.newTextParagraph().setText("No items or charges confirmed for this order.")); 
    }
  } else { 
    itemSummarySection.addWidget(CardService.newTextParagraph().setText("Could not retrieve item data for summary.")); 
  }
  card.addSection(itemSummarySection);

  const actionsSection = CardService.newCardSection();

  // MODIFICATION: Re-ordered and re-named buttons
  const backToCustomerDetailsAction = CardService.newAction()
    .setFunctionName('buildReviewContactCard') 
    .setParameters({ orderNum: orderNum });
  actionsSection.addWidget(CardService.newTextButton()
    .setText("Edit Customer") // New Text
    .setIcon(CardService.Icon.PERSON) // Example Icon
    .setOnClickAction(backToCustomerDetailsAction));

  const backToItemsAction = CardService.newAction().setFunctionName('buildItemMappingAndPricingCard').setParameters({ orderNum: orderNum }); 
  actionsSection.addWidget(CardService.newTextButton()
    .setText("Edit Items") // New Text
    .setIcon(CardService.Icon.EDIT) // Example Icon
    .setOnClickAction(backToItemsAction));

  const generateDocsAndEmailAction = CardService.newAction().setFunctionName('handleGenerateInvoiceAndEmail').setParameters({ orderNum: orderNum });
  actionsSection.addWidget(CardService.newTextButton()
    .setText("Generate Documents") // New Text
    .setIcon(CardService.Icon.DESCRIPTION) // Example Icon (was 📄厨️)
    .setOnClickAction(generateDocsAndEmailAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED));
  // END MODIFICATION
  
  card.addSection(actionsSection);
  return card.build();
}

/**
 * Handles the generation of invoice and kitchen documents and prepares the email reply.
 * MODIFIED: Added debug console.log statements.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse} An action response.
 */
function handleGenerateInvoiceAndEmail(e) {
  const orderNum = e.parameters.orderNum;
  console.log(`handleGenerateInvoiceAndEmail: Started for orderNum: ${orderNum}`);
  let invoiceSheetInfo = null;
  let kitchenSheetInfo = null;
  
  try {
    console.log(`handleGenerateInvoiceAndEmail: Attempting to retrieve orderData for orderNum: ${orderNum}`);
    const userProps = PropertiesService.getUserProperties();
    const orderDataString = userProps.getProperty(orderNum);
    if (!orderDataString) {
      throw new Error(`Order data not found in UserProperties for orderNum: ${orderNum}`);
    }
    const orderData = JSON.parse(orderDataString); // For logging/checking if needed
    console.log(`handleGenerateInvoiceAndEmail: Successfully retrieved orderData for ${orderNum}. Customer: ${orderData['Customer Name']}`);

    console.log(`handleGenerateInvoiceAndEmail: Calling populateInvoiceSheet for orderNum: ${orderNum}`);
    invoiceSheetInfo = populateInvoiceSheet(orderNum);
    if (!invoiceSheetInfo || !invoiceSheetInfo.id || !invoiceSheetInfo.url || !invoiceSheetInfo.name) {
      console.error(`handleGenerateInvoiceAndEmail: populateInvoiceSheet failed or returned invalid data for orderNum: ${orderNum}. Response: ${JSON.stringify(invoiceSheetInfo)}`);
      throw new Error("Failed to populate invoice sheet or retrieve its details.");
    }
    console.log(`handleGenerateInvoiceAndEmail: Invoice sheet populated: ${invoiceSheetInfo.name} (ID: ${invoiceSheetInfo.id}), URL: ${invoiceSheetInfo.url}`);
    
    console.log(`handleGenerateInvoiceAndEmail: Calling populateKitchenSheet for orderNum: ${orderNum}`);
    kitchenSheetInfo = populateKitchenSheet(orderNum); 
    if (!kitchenSheetInfo || !kitchenSheetInfo.id || !kitchenSheetInfo.url || !kitchenSheetInfo.name) {
      console.warn(`handleGenerateInvoiceAndEmail: populateKitchenSheet failed or returned invalid data for order ${orderNum}. Response: ${JSON.stringify(kitchenSheetInfo)}. Continuing with invoice PDF and email.`);
    } else {
      console.log(`handleGenerateInvoiceAndEmail: Kitchen sheet populated: ${kitchenSheetInfo.name} (ID: ${kitchenSheetInfo.id}), URL: ${kitchenSheetInfo.url}`);
    }
    
    console.log(`handleGenerateInvoiceAndEmail: Calling createPdfAndPrepareEmailReply for orderNum: ${orderNum}, invoiceSheetId: ${invoiceSheetInfo.id}, invoiceSheetName: ${invoiceSheetInfo.name}`);
    console.log(`handleGenerateInvoiceAndEmail: Calling createPdfAndPrepareEmailReply for orderNum: ${orderNum}, using global SHEET_ID: ${SHEET_ID}, and invoiceSheetName: ${invoiceSheetInfo.name}`);
    
    const emailInfo = createPdfAndPrepareEmailReply(orderNum, SHEET_ID, invoiceSheetInfo.name);

    if (!emailInfo || !emailInfo.draftId) {
      console.error(`handleGenerateInvoiceAndEmail: createPdfAndPrepareEmailReply failed or returned invalid data for orderNum: ${orderNum}. Response: ${JSON.stringify(emailInfo)}`);
      throw new Error("Failed to create Invoice PDF or prepare email draft.");
    }
    console.log(`handleGenerateInvoiceAndEmail: PDF and Email draft created successfully for orderNum: ${orderNum}. Draft ID: ${emailInfo.draftId}`);
    
    const successCard = CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("✅ Success! Documents Prepared"));
    const successSection = CardService.newCardSection();
    successSection.addWidget(CardService.newTextParagraph().setText(`Invoice sheet "<b>${invoiceSheetInfo.name}</b>" has been created.`));
    successSection.addWidget(CardService.newButtonSet().addButton(CardService.newTextButton().setText("Open Invoice Sheet").setOpenLink(CardService.newOpenLink().setUrl(invoiceSheetInfo.url))));
    
    if (kitchenSheetInfo && kitchenSheetInfo.url) {
        successSection.addWidget(CardService.newTextParagraph().setText(`Kitchen sheet "<b>${kitchenSheetInfo.name}</b>" has also been created.`));
        successSection.addWidget(CardService.newButtonSet().addButton(CardService.newTextButton().setText("Open Kitchen Sheet").setOpenLink(CardService.newOpenLink().setUrl(kitchenSheetInfo.url))));
    }
    
    successSection.addWidget(CardService.newTextParagraph().setText("A draft email with the invoice PDF attached has been created."));
    successSection.addWidget(CardService.newButtonSet().addButton(CardService.newTextButton().setText("Open Draft Email").setOpenLink(CardService.newOpenLink().setUrl(`https://mail.google.com/mail/u/0/#drafts?compose=${emailInfo.draftId}`))));
    
    const clearAction = CardService.newAction().setFunctionName("handleClearAndClose").setParameters({orderNum: orderNum});
    successSection.addWidget(CardService.newTextButton().setText("Done (Clear & Close Sidebar)").setOnClickAction(clearAction));
    
    successCard.addSection(successSection);
    console.log(`handleGenerateInvoiceAndEmail: Successfully built success card for orderNum: ${orderNum}.`);
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(successCard.build())).build();

  } catch (err) {
    console.error(`Error in handleGenerateInvoiceAndEmail for orderNum ${orderNum}: ${err.toString()}${(err.stack ? ("\nStack: " + err.stack) : "")}`);
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error generating documents/email: " + err.message)).build();
  }
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