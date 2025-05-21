// Code.gs
// Main script file for El Merkury Catering Add-on.

// === HOMEPAGE CARD ===
/**
 * Creates the initial homepage card for the add-on.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object.
 * @returns {GoogleAppsScript.Card_Service.Card} The homepage card.
 */
function createHomepageCard(e) {
  console.log("createHomepageCard triggered. Event: " + JSON.stringify(e));
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("El Merkury Catering Assistant"));
  const section = CardService.newCardSection();

  const processEmailAction = CardService.newAction()
    .setFunctionName("buildAddOnCard")
    .setParameters(e && e.messageMetadata ? {messageId: e.messageMetadata.messageId} : (e ? e.parameters : {}));

  section.addWidget(CardService.newTextButton().setText("Process Incoming Catering Email").setOnClickAction(processEmailAction).setTextButtonStyle(CardService.TextButtonStyle.FILLED));
  // Placeholder for future functionality
  const pennInvoiceAction = CardService.newAction().setFunctionName("handlePennInvoiceWorkflowPlaceholder");
  section.addWidget(CardService.newTextButton().setText("Process Penn Invoice (Coming Soon)").setOnClickAction(pennInvoiceAction).setDisabled(true));
  card.addSection(section);
  return card.build();
}

/**
 * Placeholder for Penn Invoice workflow.
 */
function handlePennInvoiceWorkflowPlaceholder(e) {
  const card = CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("Penn Invoice Workflow"))
    .addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("This feature is under development.")))
    .build();
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(card)).build();
}

// === ENTRY POINT - STEP 0: INITIAL EMAIL PARSING ===
/**
 * Builds the initial add-on card by parsing the current Gmail message.
 * Extracts data using Gemini and sets up the order.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object.
 * @returns {GoogleAppsScript.Card_Service.Card} The review contact card (Step 1).
 */
function buildAddOnCard(e) {
  console.log("buildAddOnCard triggered. Event: " + JSON.stringify(e));
  let msgId;

  if (e && e.messageMetadata && e.messageMetadata.messageId) { msgId = e.messageMetadata.messageId; }
  else if (e && e.gmail && e.gmail.messageId) { msgId = e.gmail.messageId; }
  else if (e && e.parameters && e.parameters.messageId) { msgId = e.parameters.messageId; }
  else {
    const currentEventObject = e.commonEventObject || e;
    if (currentEventObject && currentEventObject.messageMetadata && currentEventObject.messageMetadata.messageId) {
      msgId = currentEventObject.messageMetadata.messageId;
    } else {
      const currentMessage = GmailApp.getCurrentMessage();
      if (currentMessage) {
        msgId = currentMessage.getId();
        console.log("buildAddOnCard: Using current open message ID: " + msgId);
      } else {
        console.error("buildAddOnCard: Could not determine messageId.");
        return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("Error"))
          .addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Please open a catering order email to process."))).build();
      }
    }
  }
  console.log("Processing message ID: " + msgId);

  const message = GmailApp.getMessageById(msgId);
  const body = message.getPlainBody();
  const senderEmailFull = message.getFrom();
  console.log("buildAddOnCard: Sender of current message (message.getFrom()):", senderEmailFull);

  const contactInfoParsed = _parseJson(callGemini(_buildContactInfoPrompt(body)));
  const itemsParsed = _parseJson(callGemini(_buildStructuredItemExtractionPrompt(body)));

  const timestampSuffix = Date.now().toString().slice(-5);
  const randomPrefix = Math.floor(Math.random() * 900 + 100).toString();
  const orderNum = (randomPrefix + timestampSuffix).slice(0,8).padStart(8,'0');

  let emailForClientMatching = contactInfoParsed['Customer Address Email'];
  if (!emailForClientMatching || emailForClientMatching.trim() === "") {
    console.warn("buildAddOnCard: 'Customer Address Email' (for client matching) not found/empty in AI parse.");
  }
  console.log("buildAddOnCard: Email address being used for client matching:", (emailForClientMatching || "''(empty)"));
  const client = _matchClient(emailForClientMatching); 
  console.log("buildAddOnCard: Matched Client: " + client);

  const data = {};
  data['Internal Sender Name'] = _extractNameFromEmail(senderEmailFull);
  data['Internal Sender Email'] = _extractActualEmail(senderEmailFull);
  data['Customer Name'] = contactInfoParsed['Customer Name'] || '';
  if (!data['Customer Name'] && !data['Internal Sender Email'].includes("elmerkury.com")) {
    data['Customer Name'] = data['Internal Sender Name'];
  }
  data['Contact Person'] = contactInfoParsed['Delivery Contact Person'] || data['Customer Name'] || data['Internal Sender Name'];
  data['Customer Address Phone'] = contactInfoParsed['Customer Address Phone'] || '';
  data['Contact Phone'] = contactInfoParsed['Delivery Contact Phone'] || data['Customer Address Phone'] || '';
  data['Customer Address Email'] = contactInfoParsed['Customer Address Email'] || ''; 
  data['Contact Email'] = data['Customer Address Email'] || data['Internal Sender Email'];
  data['Customer Address Line 1'] = contactInfoParsed['Customer Address Line 1'] || '';
  data['Customer Address Line 2'] = contactInfoParsed['Customer Address Line 2'] || '';
  data['Customer Address City'] = contactInfoParsed['Customer Address City'] || '';
  data['Customer Address State'] = contactInfoParsed['Customer Address State'] || '';
  data['Customer Address ZIP'] = contactInfoParsed['Customer Address ZIP'] || '';
  data['Delivery Date'] = contactInfoParsed['Delivery Date'] || ''; // Expecting YYYY-MM-DD from AI
  data['Delivery Time'] = contactInfoParsed['Delivery Time'] || ''; 
  data['Include Utensils?'] = contactInfoParsed['Include Utensils?'] || 'Unknown';
  data['If yes: how many?'] = contactInfoParsed['If yes: how many?'] || (contactInfoParsed['Include Utensils?'] === 'Yes' ? '0' : '');


  data['Client'] = client; 
  data['orderNum'] = orderNum;
  data['messageId'] = msgId;
  data['threadId'] = message.getThread().getId();
  Object.assign(data, contactInfoParsed); // Merge any other fields AI found
  data['Items Ordered'] = itemsParsed['Items Ordered'] || [];

  let preliminarySubtotal = 0;
  data['Items Ordered'].forEach(item => {
    const priceMatch = item.description.match(/\$(\d+(\.\d{2})?)/);
    preliminarySubtotal += (parseInt(item.quantity) || 1) * (priceMatch ? parseFloat(priceMatch[1]) : 0);
  });
  data['PreliminarySubtotalForTip'] = preliminarySubtotal; 
  data['TipAmount'] = (preliminarySubtotal * 0.10); // Default 10% tip suggestion
  console.log(`Initial parse for order ${orderNum}: Tip base subtotal $${preliminarySubtotal.toFixed(2)}, Tip $${data['TipAmount'].toFixed(2)}`);

  PropertiesService.getUserProperties().setProperty(orderNum, JSON.stringify(data));
  return buildReviewContactCard({ parameters: { orderNum: orderNum } });
}

// === STEP 1 CARD: CUSTOMER & DELIVERY DETAILS ===
/**
 * Builds the card for reviewing and editing customer and delivery contact details.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object.
 * @returns {GoogleAppsScript.Card_Service.Card} The review card.
 */
function buildReviewContactCard(e) {
  const orderNum = e.parameters.orderNum;
  const userProps = PropertiesService.getUserProperties();
  const dataString = userProps.getProperty(orderNum);
  if (!dataString) {
    console.error("buildReviewContactCard: Order data for " + orderNum + " not found.");
    return CardService.newCardBuilder().addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Error: Order data missing."))).build();
  }
  const data = JSON.parse(dataString); 

  const cardSection = CardService.newCardSection();
  cardSection.addWidget(CardService.newTextParagraph().setText('📋 Order #: <b>' + orderNum + '</b>'));
  cardSection.addWidget(CardService.newDivider());

  cardSection.addWidget(CardService.newTextParagraph().setText("<b>Customer Billing & Contact Information:</b>"));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Name').setTitle('Customer Name (for Invoice)').setValue(data['Customer Name'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Client').setTitle('Client Account').setValue(data['Client'] || 'Unknown')); 
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address Email').setTitle('Customer Email').setValue(data['Customer Address Email'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address Phone').setTitle('Customer Phone').setValue(_formatPhone(data['Customer Address Phone'] || '')));
  cardSection.addWidget(CardService.newDivider());

  cardSection.addWidget(CardService.newTextParagraph().setText("<b>Delivery Address:</b>"));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address Line 1').setTitle('Street Line 1').setValue(data['Customer Address Line 1'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address Line 2').setTitle('Street Line 2 (Apt, Floor, etc.)').setValue(data['Customer Address Line 2'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address City').setTitle('City').setValue(data['Customer Address City'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address State').setTitle('State').setValue(data['Customer Address State'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Customer Address ZIP').setTitle('ZIP Code').setValue(data['Customer Address ZIP'] || ''));
  cardSection.addWidget(CardService.newDivider());
  
  cardSection.addWidget(CardService.newTextParagraph().setText("<b>Delivery Date & Time:</b>"));
  cardSection.addWidget(CardService.newDatePicker().setFieldName("Delivery Date").setTitle("Delivery Date").setValueInMsSinceEpoch(_parseDateToMsEpoch(data['Delivery Date'])));
  
  const deliveryTimeInput = CardService.newSelectionInput().setFieldName('Delivery Time').setTitle('Delivery Time');
  if (typeof deliveryTimeInput.setType === 'function') {
    deliveryTimeInput.setType(CardService.SelectionInputType.DROPDOWN); 
  } else { console.warn("setType missing on deliveryTimeInput."); }
  const rawDeliveryTimeFromAI = data['Delivery Time'] || '';
  const normalizedDeliveryTime = _normalizeTimeFormat(rawDeliveryTimeFromAI);
  console.log(`ReviewCard - Order ${orderNum} - Original Time: "${rawDeliveryTimeFromAI}", Normalized: "${normalizedDeliveryTime}"`);
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

  cardSection.addWidget(CardService.newTextParagraph().setText("<b>On-Site Delivery Contact Person:</b>"));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Contact Person').setTitle('Name').setValue(data['Contact Person'] || data['Customer Name'] || ''));
  cardSection.addWidget(CardService.newTextInput().setFieldName('Contact Phone').setTitle('Phone').setValue(_formatPhone(data['Contact Phone'] || data['Customer Address Phone'] || '')));
  cardSection.addWidget(CardService.newDivider());

  cardSection.addWidget(CardService.newTextParagraph().setText("<b>Additional Order Options:</b>"));
  const utensilsValue = data['Include Utensils?'] || 'Unknown';
  const utensilsInput = CardService.newSelectionInput().setFieldName('Include Utensils?').setTitle('Include Utensils?');
  if (typeof utensilsInput.setType === 'function') { 
    utensilsInput.setType(CardService.SelectionInputType.DROPDOWN); 
  } else { console.warn("setType missing on utensilsInput."); }
  utensilsInput.addItem('Yes', 'Yes', utensilsValue === 'Yes').addItem('No', 'No', utensilsValue === 'No').addItem('Unknown', 'Unknown', utensilsValue !== 'Yes' && utensilsValue !== 'No');
  cardSection.addWidget(utensilsInput);
  
  let numUtensilsVal = ''; 
  const showUtensilCount = (data['Include Utensils?'] === 'Yes') || 
                           (e && e.commonEventObject && e.commonEventObject.formInputs && e.commonEventObject.formInputs['Include Utensils?'] && e.commonEventObject.formInputs['Include Utensils?'].stringInputs && e.commonEventObject.formInputs['Include Utensils?'].stringInputs.value[0] === 'Yes');
  if (showUtensilCount) {
    if (data['If yes: how many?'] !== undefined && data['If yes: how many?'] !== null && data['If yes: how many?'] !== "") {
        numUtensilsVal = data['If yes: how many?'].toString();
    }
    cardSection.addWidget(CardService.newTextInput().setFieldName('If yes: how many?').setTitle('How many utensil sets?').setValue(numUtensilsVal));
  }

  const action = CardService.newAction().setFunctionName('handleContactInfoSubmitWithValidation').setParameters({ orderNum: orderNum });
  const footer = CardService.newFixedFooter().setPrimaryButton(CardService.newTextButton().setText('Confirm Contact & Proceed').setOnClickAction(action).setTextButtonStyle(CardService.TextButtonStyle.FILLED));
  
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Step 1: Customer & Order Details'))
    .addSection(cardSection)
    .setFixedFooter(footer)
    .build();
}

// === SUBMIT STEP 1 & VALIDATE ===
/**
 * Handles submission of contact info with validation.
 */
function handleContactInfoSubmitWithValidation(e) {
  const inputs = e.commonEventObject.formInputs;
  const orderNum = e.parameters.orderNum;
  const customerName = _getFormInputValue(inputs, 'Customer Name');
  const contactPerson = _getFormInputValue(inputs, 'Contact Person');
  const customerPhone = _getFormInputValue(inputs, 'Customer Address Phone');
  const customerAddressLine1 = _getFormInputValue(inputs, 'Customer Address Line 1');
  const deliveryDateMs = _getFormInputValue(inputs, 'Delivery Date', true); 
  const deliveryTimeStr = _getFormInputValue(inputs, 'Delivery Time');
  const validationMessages = []; 

  if (!deliveryDateMs || !deliveryTimeStr) { validationMessages.push("• Delivery Date and Time are required."); }
  if (!customerName && !contactPerson) { validationMessages.push("• Either 'Customer Name' or 'Delivery Contact Name' is required."); }
  if (!customerPhone && !contactPerson) { validationMessages.push("• A 'Customer Phone' or 'Delivery Contact Name' (if no customer phone) is required.");}
  if (!customerAddressLine1) { validationMessages.push("• 'Delivery Address Line 1' is required.");}

  if (validationMessages.length > 0) {
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Validation Error').setImageUrl('https://fonts.gstatic.com/s/i/googlematerialicons/error/v15/gm_grey_24dp.png'))
      .addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText('<b>Please correct the following:</b><br>' + validationMessages.join('<br>'))))
      .addSection(CardService.newCardSection().addWidget(CardService.newTextButton().setText('Back to Details').setOnClickAction(CardService.newAction().setFunctionName('buildReviewContactCard').setParameters({ orderNum: orderNum }))))
      .build();
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(card)).build();
  }
  return handleContactInfoSubmit(e); // Proceed if validation passes
}

/**
 * Processes submitted contact info and saves it.
 */
function handleContactInfoSubmit(e) {
  const inputs = e.commonEventObject.formInputs;
  const orderNum = e.parameters.orderNum;
  console.log(`handleContactInfoSubmit: Processing orderNum: ${orderNum}`);
  const newData = {}; 
  let deliveryDateMsFromForm = null; 
  let deliveryTimeStrFromForm = '';

  const userProps = PropertiesService.getUserProperties();
  const existingRaw = userProps.getProperty(orderNum);
  if (!existingRaw) { 
    console.error(`Error: Original order data for ${orderNum} not found in handleContactInfoSubmit.`);
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Original order data missing.")).build();
  }
  const existing = JSON.parse(existingRaw);
  console.log(`handleContactInfoSubmit (Order: ${orderNum}): Existing - Date: "${existing['Delivery Date']}", Time: "${existing['Delivery Time']}", MasterMS: ${existing['master_delivery_time_ms']}`);

  for (const key in inputs) {
    if (key === "Delivery Date") { 
      deliveryDateMsFromForm = _getFormInputValue(inputs, key, true); 
      if (deliveryDateMsFromForm) {
        const datePickerDateObj = new Date(deliveryDateMsFromForm);
        const yearUtc = datePickerDateObj.getUTCFullYear();
        const monthUtc = datePickerDateObj.getUTCMonth(); 
        const dayUtc = datePickerDateObj.getUTCDate();
        const localDateAtMidnight = new Date(yearUtc, monthUtc, dayUtc);
        console.log(`handleContactInfoSubmit (Order: ${orderNum}): DatePicker ms: ${deliveryDateMsFromForm}. UTC Date: ${yearUtc}-${String(monthUtc + 1).padStart(2,'0')}-${String(dayUtc).padStart(2,'0')}. Local obj: ${localDateAtMidnight}`);
        newData[key] = Utilities.formatDate(localDateAtMidnight, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else { newData[key] = ""; console.warn(`handleContactInfoSubmit (Order: ${orderNum}): Delivery Date from form is null.`); }
    } else if (key === "Delivery Time") {
      deliveryTimeStrFromForm = _getFormInputValue(inputs, key); 
      newData[key] = deliveryTimeStrFromForm;
    } else { newData[key] = _getFormInputValue(inputs, key); }
  }
  console.log(`handleContactInfoSubmit (Order: ${orderNum}): Form (newData) - Date: "${newData['Delivery Date']}", Time: "${newData['Delivery Time']}"`);
  
  let epochMsForMasterCalc;
  if (newData['Delivery Date'] && newData['Delivery Date'].match(/^\d{4}-\d{2}-\d{2}$/)) {
    epochMsForMasterCalc = _parseDateToMsEpoch(newData['Delivery Date']);
  } else if (deliveryDateMsFromForm) { 
    console.warn(`handleContactInfoSubmit (Order: ${orderNum}): newData['Delivery Date'] not yyyy-MM-dd. Using raw ms for master calc.`);
    epochMsForMasterCalc = deliveryDateMsFromForm; 
  } else { epochMsForMasterCalc = null; }
  console.log(`handleContactInfoSubmit (Order: ${orderNum}): epochForMasterCalc: ${epochMsForMasterCalc}, timeStrFromForm: "${deliveryTimeStrFromForm}"`);

  const merged = { ...existing, ...newData };

  if (epochMsForMasterCalc && deliveryTimeStrFromForm && deliveryTimeStrFromForm.trim() !== "") {
    merged['master_delivery_time_ms'] = _combineDateAndTime(epochMsForMasterCalc, deliveryTimeStrFromForm);
    console.log(`handleContactInfoSubmit (Order: ${orderNum}): Recalculated master_delivery_time_ms: ${merged['master_delivery_time_ms']}`);
  } else {
    console.warn(`handleContactInfoSubmit (Order: ${orderNum}): Could not recalc master_delivery_time_ms. epoch: ${epochMsForMasterCalc}, timeStr: "${deliveryTimeStrFromForm}".`);
    if (!(epochMsForMasterCalc && deliveryTimeStrFromForm && deliveryTimeStrFromForm.trim() !== "")) {
        delete merged['master_delivery_time_ms']; 
        console.log(`handleContactInfoSubmit (Order: ${orderNum}): master_delivery_time_ms cleared.`);
    }
 }
  console.log(`handleContactInfoSubmit (Order: ${orderNum}): Merged - Date: "${merged['Delivery Date']}", Time: "${merged['Delivery Time']}", MasterMS: ${merged['master_delivery_time_ms']}`);

  if (!merged['Items Ordered'] || !Array.isArray(merged['Items Ordered']) || (merged['Items Ordered'].length > 0 && typeof merged['Items Ordered'][0].description === 'undefined')) {
    console.log(`handleContactInfoSubmit (Order: ${orderNum}): Re-running item extraction.`);
    const message = GmailApp.getMessageById(merged.messageId);
    merged['Items Ordered'] = _parseJson(callGemini(_buildStructuredItemExtractionPrompt(message.getPlainBody())))['Items Ordered'] || [];
  }

  PropertiesService.getUserProperties().setProperty(orderNum, JSON.stringify(merged));
  console.log(`handleContactInfoSubmit (Order: ${orderNum}): Data saved. Proceeding to item mapping.`);
  
  const itemMappingCard = buildItemMappingAndPricingCard({ parameters: { orderNum: orderNum } });
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(itemMappingCard)).build();
}

// === ITEM MAPPING AND PRICING CARD ===
/**
 * Builds the card for mapping extracted email items to QuickBooks items and reviewing pricing.
 * MODIFIED: "Customer Notes" field pre-filled with item.parsed_flavors_or_details.
 * Uses Switch for "Remove item", "Additional Charges" section expanded.
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
      let customNotesFromUnmatchedAIFlavors = []; // This is the correct variable name declared here
      let displayedFlavorDropdownCount = 0;

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
              customNotesFromUnmatchedAIFlavors.push(aiFlavor); 
            }
          });
        } else { 
          customNotesFromUnmatchedAIFlavors = customNotesFromUnmatchedAIFlavors.concat(aiIdentifiedFlavors);
        }
      } else { 
        customNotesFromUnmatchedAIFlavors = customNotesFromUnmatchedAIFlavors.concat(aiIdentifiedFlavors);
      }
      
      let customerNotesValue = item.parsed_flavors_or_details || '';
      // FIX: Use the correctly declared variable name 'customNotesFromUnmatchedAIFlavors'
      if (!customerNotesValue && customNotesFromUnmatchedAIFlavors.length > 0) { // Line 413 was here
          customerNotesValue = "Unmatched AI Flavors: " + customNotesFromUnmatchedAIFlavors.join(', ');
      }

      itemsDisplaySection.addWidget(CardService.newTextInput()
        .setFieldName(`item_${index}_custom_notes`)
        .setTitle("Customer Notes (original details / additional requests)")
        .setValue(customerNotesValue) 
        .setMultiline(true));

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

// === SUBMIT STEP 2 & BUILD STEP 3 (FINAL REVIEW) ===
/**
 * Handles submission of item mapping, saves choices, and builds the final review card.
 */
function handleItemMappingSubmit(e) {
  const formInputs = e.formInputs || (e.commonEventObject && e.commonEventObject.formInputs);
  if (!formInputs) { console.error("Error: formInputs undefined in handleItemMappingSubmit."); return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Form data unreadable.")).build(); }
  
  const orderNum = e.parameters.orderNum;
  const aiItemCount = parseInt(e.parameters.ai_item_count) || 0;
  if (!orderNum) { console.error("Error: Order number missing in handleItemMappingSubmit."); return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Order number missing.")).build(); }
  
  const userProps = PropertiesService.getUserProperties();
  const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { console.error("Error: Original order data for " + orderNum + " not found."); return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Original order data not found.")).build(); }
  
  const orderData = JSON.parse(orderDataString);
  const masterQBItems = getMasterQBItems();
  const confirmedQuickBooksItems = [];
  const suggestedMatchesFromStorage = orderData['tempSuggestedMatches'] || [];

  for (let i = 0; i < aiItemCount; i++) { 
    const removeSwitchFieldName = `item_remove_${i}`;
    const removeThisItem = (formInputs[removeSwitchFieldName] && formInputs[removeSwitchFieldName][0] === "true");
    if (removeThisItem) { console.log(`Item ${i} (AI) marked for removal.`); continue; }

    const qtyString = formInputs[`item_qty_${i}`] && formInputs[`item_qty_${i}`][0];
    const qbItemSKU = formInputs[`item_qb_sku_${i}`] && formInputs[`item_qb_sku_${i}`][0]; 
    const currentProcessedItem = suggestedMatchesFromStorage[i] || {}; 
    const originalDescription = currentProcessedItem.original_email_description || "N/A";
    const aiIdentifiedFlavorsForItem = currentProcessedItem.identified_flavors || [];

    let finalKitchenNotesParts = [];
    let selectedDropdownFlavorValues = [];
    aiIdentifiedFlavorsForItem.forEach((originalAiFlavor, aiFlavorIndex) => {
      const dropdownFieldName = `item_${i}_requested_flavor_dropdown_${aiFlavorIndex}`;
      if (formInputs[dropdownFieldName] && formInputs[dropdownFieldName][0] && formInputs[dropdownFieldName][0] !== "") {
        selectedDropdownFlavorValues.push(formInputs[dropdownFieldName][0]);
      }
    });
    if (selectedDropdownFlavorValues.length > 0) { finalKitchenNotesParts.push("Selected Flavors: " + selectedDropdownFlavorValues.join(', ')); }
    
    const customNotesFieldName = `item_${i}_custom_notes`;
    const customNotesValue = (formInputs[customNotesFieldName] && formInputs[customNotesFieldName][0]) ? formInputs[customNotesFieldName][0].trim() : "";
    if (customNotesValue !== "") {
      // Prepend "Customer Notes: " only if it's not already the only content and not redundant
      if (finalKitchenNotesParts.length > 0 || !customNotesValue.startsWith("Customer Notes:")) {
         finalKitchenNotesParts.push("Customer Notes: " + customNotesValue);
      } else if (finalKitchenNotesParts.length === 0) {
         finalKitchenNotesParts.push(customNotesValue);
      }
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
      let unitPrice = 0; let itemName = "Custom Item"; let itemSKU = qbItemSKU; 
      if (masterItemDetails) {
        unitPrice = masterItemDetails.Price || 0; itemName = masterItemDetails.Name;
      } else { 
        console.warn("Master details not found for SKU: " + qbItemSKU + ". Defaulting."); itemSKU = FALLBACK_CUSTOM_ITEM_SKU;
        const fbMaster = masterQBItems.find(master => master.SKU === FALLBACK_CUSTOM_ITEM_SKU);
        unitPrice = fbMaster ? (fbMaster.Price || 0) : 0; itemName = fbMaster ? fbMaster.Name : "Custom Item";
      }
      confirmedQuickBooksItems.push({ quickbooks_item_id: itemSKU, quickbooks_item_name: itemName, sku: itemSKU, quantity: parseInt(qtyString) || 1, unit_price: unitPrice, kitchen_notes_and_flavors: kitchenNotesAndFlavors, original_email_description: originalDescription });
    }
  } 

  const newItemQbSKU = formInputs.new_item_qb_sku && formInputs.new_item_qb_sku[0];
  const newItemQtyString = formInputs.new_item_qty && formInputs.new_item_qty[0];
  if (newItemQbSKU && newItemQbSKU !== "" && newItemQtyString) {
    const manMasterItemDetails = masterQBItems.find(master => master.SKU === newItemQbSKU);
    if (manMasterItemDetails) {
      confirmedQuickBooksItems.push({ quickbooks_item_id: newItemQbSKU, quickbooks_item_name: manMasterItemDetails.Name, sku: manMasterItemDetails.SKU, quantity: parseInt(newItemQtyString) || 1, unit_price: manMasterItemDetails.Price || 0, kitchen_notes_and_flavors: (formInputs.new_item_kitchen_notes && formInputs.new_item_kitchen_notes[0]) || "", original_email_description: null });
    }
  }
  delete orderData['tempSuggestedMatches']; 
  orderData['ConfirmedQBItems'] = confirmedQuickBooksItems; 
  orderData['TipAmount'] = parseFloat(formInputs.tip_amount && formInputs.tip_amount[0]) || 0;
  orderData['OtherChargesAmount'] = parseFloat(formInputs.other_charges_amount && formInputs.other_charges_amount[0]) || 0;
  orderData['OtherChargesDescription'] = (formInputs.other_charges_description && formInputs.other_charges_description[0]) || "";

  PropertiesService.getUserProperties().setProperty(orderNum, JSON.stringify(orderData));
  console.log("Confirmed items and charges for order " + orderNum + ": " + JSON.stringify(orderData));
  
  // Now proceed to the Final Review Card which has the "Finalize Sheets" button
  const finalReviewCard = buildFinalReviewCard({ parameters: { orderNum: orderNum } });
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(finalReviewCard)).build();
}

// === STEP 3 CARD: FINAL REVIEW & ACTIONS (before sheet/PDF generation) ===
/**
 * Builds the final review card before any documents are generated.
 * From here, user can finalize sheets, or go back to edit customer/items.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object.
 * @returns {GoogleAppsScript.Card_Service.Card} The final review card.
 */
function buildFinalReviewCard(e) { // Renamed from buildInvoiceActionsCard
  const orderNum = e.parameters.orderNum;
  console.log(`buildFinalReviewCard: Building for orderNum: ${orderNum}`);
  const card = CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle(`Step 3: Final Review for Order ${orderNum}`));
  
  const userProps = PropertiesService.getUserProperties(); 
  const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) {
    console.error(`buildFinalReviewCard: Order data for ${orderNum} not found.`);
    return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("Error")).addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Order data could not be loaded."))).build();
  }
  const orderData = JSON.parse(orderDataString);

  // Section 1: Order Details Summary
  const orderDetailsSection = CardService.newCardSection().setHeader("Order Details Summary");
  console.log(`buildFinalReviewCard (Order: ${orderNum}): Displaying Date: "${orderData['Delivery Date']}", Time: "${orderData['Delivery Time']}"`);
  orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Customer Name:</b> " + (orderData['Customer Name'] || 'N/A') + "</i>"));
  orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Client:</b> " + (orderData['Client'] || 'N/A') + "</i>")); 
  // ... (add more key customer/delivery details similar to old buildInvoiceActionsCard display if needed for review)

  let deliveryDateForDisplay = "N/A";
  if (orderData['Delivery Date'] && orderData['Delivery Date'].match(/^\d{4}-\d{2}-\d{2}$/)) {
      const parts = orderData['Delivery Date'].split('-');
      const dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      if (!isNaN(dateObj.getTime())) {
          deliveryDateForDisplay = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");
      } else { deliveryDateForDisplay = orderData['Delivery Date']; }
  } else if (orderData['Delivery Date']) { deliveryDateForDisplay = orderData['Delivery Date']; }
  const deliveryTimeForDisplay = orderData['Delivery Time'] ? _normalizeTimeFormat(orderData['Delivery Time']) : 'N/A';
  orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Delivery:</b> " + deliveryDateForDisplay + " at " + deliveryTimeForDisplay + "</i>"));
  card.addSection(orderDetailsSection);

  // Section 2: Item Summary (from old buildInvoiceActionsCard)
  const itemSummarySection = CardService.newCardSection().setHeader("Confirmed Items & Charges Summary");
  let subTotalForCharges = 0; 
  const confirmedItems = orderData['ConfirmedQBItems'] || [];
  if (confirmedItems.length > 0) {
    confirmedItems.forEach(item => {
      const itemTotal = (item.quantity || 0) * (item.unit_price || 0);
      subTotalForCharges += itemTotal;
      itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>${item.original_email_description || item.quickbooks_item_name}</b><br>Qty: ${item.quantity}, Unit Price: $${(item.unit_price || 0).toFixed(2)}, Total: $${itemTotal.toFixed(2)}${item.kitchen_notes_and_flavors ? '<br><font color="#666666"><i>Flavor/Custom Notes: ' + item.kitchen_notes_and_flavors + '</i></font>' : ''}`));
    });
    itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Items Subtotal: $${subTotalForCharges.toFixed(2)}</b>`));
  }
  
  const hasAdditionalChargesOrFees = (orderData['TipAmount'] > 0 || orderData['OtherChargesAmount'] > 0 || (orderData['Include Utensils?'] === 'Yes' && parseInt(orderData['If yes: how many?']) > 0 ) || BASE_DELIVERY_FEE > 0);
  if (subTotalForCharges > 0 && hasAdditionalChargesOrFees) {
    itemSummarySection.addWidget(CardService.newDivider());
  }
  let grandTotal = subTotalForCharges; 
  // ... (Tip, Other Charges, Delivery Fee, Utensils, Grand Total - same logic as old buildInvoiceActionsCard) ...
  if (orderData['TipAmount'] > 0) { itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Tip:</b> $${orderData['TipAmount'].toFixed(2)}`)); grandTotal += orderData['TipAmount']; }
  if (orderData['OtherChargesAmount'] > 0) { itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>${orderData['OtherChargesDescription'] || 'Other Charges'}:</b> $${orderData['OtherChargesAmount'].toFixed(2)}`)); grandTotal += orderData['OtherChargesAmount']; }
  let deliveryFee = BASE_DELIVERY_FEE;
  if (orderData['master_delivery_time_ms']) { 
    const deliveryHour = new Date(orderData['master_delivery_time_ms']).getHours();
    if (deliveryHour >= DELIVERY_FEE_CUTOFF_HOUR) { deliveryFee = AFTER_4PM_DELIVERY_FEE; }
  }
  if (deliveryFee > 0) { itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Delivery Fee:</b> $${deliveryFee.toFixed(2)}`)); grandTotal += deliveryFee; }
  if (orderData['Include Utensils?'] === 'Yes') {
    const numUtensils = parseInt(orderData['If yes: how many?']) || 0;
    if (numUtensils > 0) { const utensilCost = numUtensils * COST_PER_UTENSIL_SET; itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Utensils (${numUtensils} sets):</b> $${utensilCost.toFixed(2)}`)); grandTotal += utensilCost; }
  }
  if (grandTotal > 0 || subTotalForCharges > 0) { 
    if (hasAdditionalChargesOrFees) { itemSummarySection.addWidget(CardService.newDivider()); }
    itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Estimated Grand Total: $${grandTotal.toFixed(2)}</b>`));
  } else if (confirmedItems.length === 0) { itemSummarySection.addWidget(CardService.newTextParagraph().setText("No items confirmed.")); }
  card.addSection(itemSummarySection);

  // Section 3: Actions
  const actionsSection = CardService.newCardSection();
  const backToCustomerDetailsAction = CardService.newAction().setFunctionName('buildReviewContactCard').setParameters({ orderNum: orderNum });
  actionsSection.addWidget(CardService.newTextButton().setText("Edit Customer").setIcon(CardService.Icon.PERSON).setOnClickAction(backToCustomerDetailsAction));

  const backToItemsAction = CardService.newAction().setFunctionName('buildItemMappingAndPricingCard').setParameters({ orderNum: orderNum }); 
  actionsSection.addWidget(CardService.newTextButton().setText("Edit Items").setIcon(CardService.Icon.EDIT).setOnClickAction(backToItemsAction));
  
  // MODIFIED: Primary action button
  const finalizeSheetsAction = CardService.newAction().setFunctionName('handleFinalizeSheets').setParameters({ orderNum: orderNum });
  actionsSection.addWidget(CardService.newTextButton()
    .setText("✅ Finalize Sheets & Prepare for PDF/Email") 
    .setOnClickAction(finalizeSheetsAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED));
  
  card.addSection(actionsSection);
  return card.build();
}

// === STEP 3.5: FINALIZE SHEETS ===
/**
 * Populates the Invoice and Kitchen Google Sheets.
 * Then builds the card for PDF generation and email options.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse} ActionResponse to show the PDF/Email options card.
 */
function handleFinalizeSheets(e) {
  const orderNum = e.parameters.orderNum;
  console.log(`handleFinalizeSheets: Started for orderNum: ${orderNum}`);
  let invoiceSheetInfo = null;
  let kitchenSheetInfo = null;
  
  try {
    const userProps = PropertiesService.getUserProperties();
    let orderDataString = userProps.getProperty(orderNum); // Get current orderData
    if (!orderDataString) { throw new Error(`Order data not found for ${orderNum}`); }
    let orderData = JSON.parse(orderDataString);

    console.log(`handleFinalizeSheets: Populating Invoice Sheet for ${orderNum}`);
    invoiceSheetInfo = populateInvoiceSheet(orderNum); // Assumes this function is in Utils.gs or Code.gs
    if (!invoiceSheetInfo || !invoiceSheetInfo.name || !invoiceSheetInfo.url) {
      throw new Error("Failed to populate invoice Google Sheet or get its info.");
    }
    console.log(`handleFinalizeSheets: Invoice Sheet "${invoiceSheetInfo.name}" populated.`);
    // Store sheet info in orderData to pass to next card
    orderData.invoiceSheetName = invoiceSheetInfo.name;
    orderData.invoiceSheetUrl = invoiceSheetInfo.url;
    // invoiceSheetInfo.id is the GID, might not be needed by card if URL is enough
    
    console.log(`handleFinalizeSheets: Populating Kitchen Sheet for ${orderNum}`);
    kitchenSheetInfo = populateKitchenSheet(orderNum); // Assumes this function is in Utils.gs or Code.gs
    if (kitchenSheetInfo && kitchenSheetInfo.name) {
      console.log(`handleFinalizeSheets: Kitchen Sheet "${kitchenSheetInfo.name}" populated.`);
      orderData.kitchenSheetName = kitchenSheetInfo.name;
      orderData.kitchenSheetUrl = kitchenSheetInfo.url;
    } else {
      console.warn(`handleFinalizeSheets: Kitchen Sheet population failed or returned no info for order ${orderNum}.`);
      orderData.kitchenSheetName = null; // Ensure it's explicitly null if failed
      orderData.kitchenSheetUrl = null;
    }

    // Save updated orderData (with sheet names/URLs)
    PropertiesService.getUserProperties().setProperty(orderNum, JSON.stringify(orderData));
    
    console.log(`handleFinalizeSheets: Sheets populated. Building PDF/Email options card for ${orderNum}`);
    const optionsCard = buildPdfAndEmailOptionsCard({ parameters: { orderNum: orderNum }}); // Pass orderNum
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(optionsCard)).build();

  } catch (err) {
    console.error(`Error in handleFinalizeSheets for orderNum ${orderNum}: ${err.toString()}${(err.stack ? ("\nStack: " + err.stack) : "")}`);
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error finalizing sheets: " + err.message)).build();
  }
}

// === STEP 4 CARD: PDF & EMAIL OPTIONS ===
/**
 * Builds the card displayed after sheets are generated.
 * Offers options to open generated sheets, view PDF (if generated), and compose email from templates.
 * @param {GoogleAppsScript.Addons.EventObject} e Event object containing orderNum.
 * @return {GoogleAppsScript.Card_Service.Card} The PDF and Email options card.
 */
function buildPdfAndEmailOptionsCard(e) { // Renamed from buildSuccessCardWithTemplates for clarity
  const orderNum = e.parameters.orderNum;
  console.log(`buildPdfAndEmailOptionsCard: Building for orderNum ${orderNum}`);

  const userProps = PropertiesService.getUserProperties();
  const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) {
    console.error(`buildPdfAndEmailOptionsCard: Order data for ${orderNum} not found.`);
    return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("Error"))
        .addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Order data missing. Please restart process.")))
        .build();
  }
  const orderData = JSON.parse(orderDataString);

  const card = CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("✅ Sheets Finalized! Order: " + orderNum));
  
  // Section for links to generated sheets
  const generatedDocsSection = CardService.newCardSection().setHeader("View Generated Sheets");
  if (orderData.invoiceSheetName && orderData.invoiceSheetUrl) {
    generatedDocsSection.addWidget(CardService.newTextParagraph().setText(`Invoice Sheet: "<b>${orderData.invoiceSheetName}</b>"`));
    generatedDocsSection.addWidget(CardService.newButtonSet().addButton(
      CardService.newTextButton().setText("Open Invoice Sheet").setOpenLink(CardService.newOpenLink().setUrl(orderData.invoiceSheetUrl))
    ));
  } else {
    generatedDocsSection.addWidget(CardService.newTextParagraph().setText("Invoice Google Sheet link not available."));
  }
  if (orderData.kitchenSheetName && orderData.kitchenSheetUrl) {
    generatedDocsSection.addWidget(CardService.newTextParagraph().setText(`Kitchen Sheet: "<b>${orderData.kitchenSheetName}</b>"`));
    generatedDocsSection.addWidget(CardService.newButtonSet().addButton(
      CardService.newTextButton().setText("Open Kitchen Sheet").setOpenLink(CardService.newOpenLink().setUrl(orderData.kitchenSheetUrl))
    ));
  }
  card.addSection(generatedDocsSection);

  // Section for PDF and Email actions
  const pdfEmailSection = CardService.newCardSection().setHeader("Generate PDF & Compose Email");
  
  // Display link to PDF if it was already generated and saved in a previous step (e.g., user clicked button twice)
  if (orderData.pdfUrl && orderData.pdfName) {
      pdfEmailSection.addWidget(CardService.newTextParagraph().setText(`Previously generated PDF: "<b>${orderData.pdfName}</b>"`));
      pdfEmailSection.addWidget(CardService.newButtonSet().addButton(
          CardService.newTextButton().setText("Open Generated PDF").setOpenLink(CardService.newOpenLink().setUrl(orderData.pdfUrl))
      ));
      pdfEmailSection.addWidget(CardService.newDivider());
  }

  const emailTemplates = getEmailTemplateList(); // From EmailTemplates.gs
  const templateDropdown = CardService.newSelectionInput()
    .setFieldName("selected_email_template")
    .setTitle("Select Email Template");
  
  // MODIFICATION: Explicitly set type to DROPDOWN
  if (templateDropdown && typeof templateDropdown.setType === 'function') {
    templateDropdown.setType(CardService.SelectionInputType.DROPDOWN);
  } else {
    console.warn("buildPdfAndEmailOptionsCard: setType method missing on templateDropdown.");
  }
  // END MODIFICATION
  
  if (emailTemplates.length > 0) {
    templateDropdown.addItem("--- Choose a template ---", "", true);
    emailTemplates.forEach(template => {
      templateDropdown.addItem(template.name, template.id, false);
    });
  } else {
    templateDropdown.addItem("No email templates found", "", true);
  }
  pdfEmailSection.addWidget(templateDropdown);

  // Future: Placeholder for "Use HTML for PDF?" switch
  // const useHtmlPdfSwitch = CardService.newSwitch().setFieldName("use_html_pdf_flag").setValue("true").setSelected(false);
  // pdfEmailSection.addWidget(CardService.newDecoratedText().setText("Use Custom HTML Invoice for PDF?").setSwitchControl(useHtmlPdfSwitch));


  const generatePdfAndEmailAction = CardService.newAction()
    .setFunctionName("handleGeneratePdfAndComposeEmail") 
    .setParameters({ orderNum: orderNum }); // pdfFileId will be handled inside the handler now
  pdfEmailSection.addWidget(CardService.newTextButton()
    .setText("📨 Generate PDF & Create Draft Email")
    .setOnClickAction(generatePdfAndEmailAction)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED));
  
  pdfEmailSection.addWidget(CardService.newTextParagraph()
    .setText("<font size='1' color='#555555'><i>If draft opens in wrong Google account, manually open 'Drafts' in correct Gmail tab.</i></font>"));
  card.addSection(pdfEmailSection);

  // Final Action Section
  const finalActionsSection = CardService.newCardSection();
  const clearAction = CardService.newAction().setFunctionName("handleClearAndClose").setParameters({orderNum: orderNum});
  finalActionsSection.addWidget(CardService.newTextButton().setText("Done (Clear & Close Sidebar)").setOnClickAction(clearAction));
  card.addSection(finalActionsSection);
  
  return card.build();
}

// === STEP 4.5: GENERATE PDF & COMPOSE EMAIL DRAFT ===
/**
 * Handles PDF generation (from HTML), saving to Drive, and composing an email from a template.
 * MODIFIED: Simplifies Drive saving attempt to focus on basic root folder creation.
 * @param {GoogleAppsScript.Addons.EventObject} e Event object from button click.
 * @return {GoogleAppsScript.Card_Service.ActionResponse} ActionResponse to open draft or show notification/updated card.
 */
function handleGeneratePdfAndComposeEmail(e) {
  const orderNum = e.parameters.orderNum;
  const selectedTemplateId = e.formInputs && e.formInputs.selected_email_template && e.formInputs.selected_email_template[0];

  console.log(`handleGeneratePdfAndComposeEmail (Simplified Drive Test): Order: ${orderNum}, Template ID: "${selectedTemplateId}"`);

  if (!selectedTemplateId || selectedTemplateId === "") {
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Please select an email template first.")).build();
  }

  const userProps = PropertiesService.getUserProperties();
  const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) {
    console.error(`handleGeneratePdfAndComposeEmail: Order data for ${orderNum} not found.`);
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Order data not found.")).build();
  }
  let orderData = JSON.parse(orderDataString); 

  const template = getEmailTemplateById(selectedTemplateId); 
  if (!template) {
    console.error(`handleGeneratePdfAndComposeEmail: Email template "${selectedTemplateId}" not found.`);
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Selected template not found.")).build();
  }

  let pdfBlob;
  let pdfName = `Invoice-${orderNum}-${(orderData['Customer Name'] || 'Unknown').replace(/[^a-zA-Z0-9\s]/g, "_").replace(/\s+/g, "_")}.pdf`;
  let pdfFileId = null;
  let pdfViewUrl = null;

  try {
    console.log(`handleGeneratePdfAndComposeEmail: Generating PDF from HTML for order ${orderNum}`);
    pdfBlob = generateInvoicePdfFromHtml(orderNum, orderData); 

    if (!pdfBlob) {
      throw new Error("HTML PDF Blob generation returned null.");
    }
    pdfName = pdfBlob.getName(); 
    console.log(`handleGeneratePdfAndComposeEmail: PDF blob "${pdfName}" created from HTML.`);

    // --- Simplified Drive Save Attempt ---
    console.log(`handleGeneratePdfAndComposeEmail: Attempting to save PDF to Drive root folder.`);
    let folder;
    try {
        folder = DriveApp.getRootFolder(); // Directly attempt to get root folder
        console.log(`handleGeneratePdfAndComposeEmail: Accessed root Drive folder: "${folder.getName()}".`);
    } catch (rootFolderError) {
        console.error(`handleGeneratePdfAndComposeEmail: CRITICAL - Error accessing root Drive folder: ${rootFolderError.message}. Stack: ${rootFolderError.stack}`);
        throw new Error(`Failed to access root Drive folder: ${rootFolderError.message}`);
    }
    
    const pdfFile = folder.createFile(pdfBlob);
    pdfFileId = pdfFile.getId();
    pdfViewUrl = pdfFile.getUrl();
    // --- End Simplified Drive Save Attempt ---
    
    orderData.pdfFileId = pdfFileId; 
    orderData.pdfUrl = pdfViewUrl;   
    orderData.pdfName = pdfName;     
    PropertiesService.getUserProperties().setProperty(orderNum, JSON.stringify(orderData)); 
    console.log(`handleGeneratePdfAndComposeEmail: PDF "${pdfName}" saved. ID: ${pdfFileId}, URL: ${pdfViewUrl}`);

  } catch (pdfOrDriveError) {
    console.error(`handleGeneratePdfAndComposeEmail: Error during PDF generation or Drive save for order ${orderNum}: ${pdfOrDriveError.toString()}`);
    const errorCard = buildPdfAndEmailOptionsCard({parameters: {orderNum: orderNum}}); 
    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(`Error creating/saving PDF: ${pdfOrDriveError.message}. Please try again or check logs.`))
        .setNavigation(CardService.newNavigation().updateCard(errorCard)) 
        .build();
  }
  
  const { subject, body } = populateEmailTemplate(template, orderData, pdfFileId); 

  try {
    const thread = GmailApp.getThreadById(orderData.threadId);
    if (!thread) { throw new Error("Original email thread not found for reply."); }
    const messageToReplyTo = thread.getMessages()[thread.getMessages().length - 1];
    
    const draftOptions = {
      htmlBody: body.replace(/\n/g, '<br>'),
      to: orderData['Contact Email'] || orderData['Customer Address Email'] || messageToReplyTo.getFrom(),
    };
    if (pdfBlob) { 
      draftOptions.attachments = [pdfBlob];
    }

    const draft = messageToReplyTo.createDraftReply("", draftOptions);
    draft.update(draftOptions.to, subject, body, draftOptions); 

    console.log(`handleGeneratePdfAndComposeEmail: Draft created. ID: ${draft.getId()}`);
    const draftUrl = `https://mail.google.com/mail/u/0/#drafts?compose=${draft.getId()}`;
    
    const successCardWithPdfLink = buildPdfAndEmailOptionsCard({parameters: {orderNum: orderNum}});

    return CardService.newActionResponseBuilder()
      .setOpenLink(CardService.newOpenLink().setUrl(draftUrl))
      .setNotification(CardService.newNotification().setText("PDF generated and email draft prepared!"))
      .setNavigation(CardService.newNavigation().updateCard(successCardWithPdfLink)) 
      .build();

  } catch (draftErr) {
    console.error(`handleGeneratePdfAndComposeEmail: Error creating draft for ${orderNum}: ${draftErr.toString()}`);
    const errorCard = buildPdfAndEmailOptionsCard({parameters: {orderNum: orderNum}});
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Error creating email draft: " + draftErr.message + (pdfBlob ? " PDF was generated." : "")))
      .setNavigation(CardService.newNavigation().updateCard(errorCard))
      .build();
  }
}


/**
 * Clears user properties for the current order and closes the sidebar.
 * Also attempts to trash the temporary PDF from Drive.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse} An action response to close the add-on.
 */
function handleClearAndClose(e) {
    const orderNum = e.parameters.orderNum;
    if (orderNum) { 
      const userProps = PropertiesService.getUserProperties();
      const orderDataString = userProps.getProperty(orderNum);
      if (orderDataString) {
        const orderData = JSON.parse(orderDataString);
        if (orderData.pdfFileId) { // Check if PDF ID was stored
          try {
            DriveApp.getFileById(orderData.pdfFileId).setTrashed(true);
            console.log(`handleClearAndClose: Temporary PDF file ${orderData.pdfFileId} for order ${orderNum} moved to trash.`);
          } catch (err) {
            console.warn(`handleClearAndClose: Could not trash PDF file ${orderData.pdfFileId} for order ${orderNum}: ${err.message}`);
          }
        }
      }
      userProps.deleteProperty(orderNum); 
      console.log("Cleared data for orderNum: " + orderNum); 
    }
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().popToRoot()).setNotification(CardService.newNotification().setText("Order data cleared. Add-on is ready for the next email.")).build();
}