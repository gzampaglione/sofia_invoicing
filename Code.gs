// Code.gs - V15
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
 * @param {GoogleAppsScript.Addons.EventObject} e The event object from Gmail.
 * @returns {GoogleAppsScript.Card_Service.Card} The review contact card.
 */
function buildAddOnCard(e) {
  console.log("buildAddOnCard triggered. Event: " + JSON.stringify(e));
  let msgId;

  // Robustly determine the message ID from various event object structures
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
  const senderEmailFull = message.getFrom();
  console.log("Sender Email for matching: " + senderEmailFull);

  // Call Gemini API to extract structured data
  const contactInfoParsed = _parseJson(callGemini(_buildContactInfoPrompt(body)));
  const itemsParsed = _parseJson(callGemini(_buildStructuredItemExtractionPrompt(body)));

  const orderingPersonName = _extractNameFromEmail(senderEmailFull); // This is the sender of the email (Sofia Deleon)
  const orderingPersonEmail = _extractActualEmail(senderEmailFull);

  // Generate a unique order number
  const timestampSuffix = Date.now().toString().slice(-5);
  const randomPrefix = Math.floor(Math.random() * 900 + 100).toString();
  const orderNum = (randomPrefix + timestampSuffix).slice(0,8).padStart(8,'0');

  // Determine client based on email domain rules
  const client = _matchClient(orderingPersonEmail);
  console.log("Matched Client: " + client);

  const data = {};
  // Store sender's info separately (Sofia Deleon / elmerkury.com)
  data['Ordering Person Name'] = orderingPersonName;
  data['Ordering Person Email'] = orderingPersonEmail;

  // Set Customer Name (e.g., Ashley Duchi), defaulting to empty. It MUST NOT be elmerkury.com sender.
  data['Customer Name'] = contactInfoParsed['Customer Name'] || '';
  // If Gemini failed to find a customer name AND the sender is NOT El Merkury, fallback to sender's name.
  // This ensures 'Customer Name' is never Sofia Deleon if it's an external order.
  if (!data['Customer Name'] && !orderingPersonEmail.includes("elmerkury.com")) {
      data['Customer Name'] = orderingPersonName;
  }
  
  // Set Delivery Contact Person (e.g., Romina), defaulting to Customer Name or sender if not found by Gemini
  data['Contact Person'] = contactInfoParsed['Delivery Contact Person'] || data['Customer Name'] || orderingPersonName;
  
  // Set Delivery Contact Phone, prioritizing from Gemini's extraction, then Customer Address Phone, then empty string.
  data['Contact Phone'] = contactInfoParsed['Delivery Contact Phone'] || contactInfoParsed['Customer Address Phone'] || '';
  // Set Delivery Contact Email, prioritizing from Gemini's extraction, then Customer Address Email, then Ordering Person's email.
  data['Contact Email'] = contactInfoParsed['Delivery Contact Email'] || contactInfoParsed['Customer Address Email'] || orderingPersonEmail;

  // Invoice recipient email, primarily Customer Address Email, falls back to ordering person
  data['Customer Address Email'] = contactInfoParsed['Customer Address Email'] || orderingPersonEmail;


  data['Client'] = client;
  data['orderNum'] = orderNum;
  data['messageId'] = msgId;
  data['threadId'] = message.getThread().getId();

  // Merge remaining extracted contact info (address, dates, utensils) directly into data object.
  // This will apply values from contactInfoParsed if they exist, possibly overwriting initial defaults.
  Object.assign(data, contactInfoParsed);

  data['Items Ordered'] = itemsParsed['Items Ordered'] || [];

  // Calculate preliminary grand total for initial tip display.
  // This is a basic calculation; the final calculation uses actual QB item prices later.
  let preliminarySubtotal = 0;
  data['Items Ordered'].forEach(item => {
    // Attempt to parse price from description if explicitly mentioned (e.g., "Pupusas Tray – Large ($140)")
    const priceMatch = item.description.match(/\$(\d+(\.\d{2})?)/);
    const itemPrice = priceMatch ? parseFloat(priceMatch[1]) : 0;
    preliminarySubtotal += (parseInt(item.quantity) || 1) * itemPrice;
  });

  // Calculate 10% tip from preliminary subtotal as per request
  data['TipAmount'] = (preliminarySubtotal * 0.10);
  console.log("Preliminary Subtotal for Tip: $" + preliminarySubtotal.toFixed(2));
  console.log("Calculated Initial Tip Amount: $" + data['TipAmount'].toFixed(2));

  // Store the compiled data in UserProperties for persistence across card interactions
  PropertiesService.getUserProperties().setProperty(orderNum, JSON.stringify(data));
  
  // Build and return the first review card
  return buildReviewContactCard({ parameters: { orderNum: orderNum } });
}


// === BUILD CUSTOMER CONTACT REVIEW CARD ===
/**
 * Builds the card for reviewing and editing customer and delivery contact details.
 * Includes client-side validation requirements.
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

  const section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph().setText('📋 Order #: <b>' + orderNum + '</b>'));
  
  // Display Ordering Person (Sender) - this is Sofia Deleon
  section.addWidget(CardService.newTextParagraph().setText("<b>Ordering Person (Sender):</b> " + (data['Ordering Person Name'] || 'N/A') + " (" + (data['Ordering Person Email'] || 'N/A') + ")"));

  // Editable fields for Customer and Delivery Contact
  // Removed .setHint() as it's static instructional text, not dynamic pre-fill. setValue handles pre-fill.
  // .setRequired() is NOT for TextInput, handled by handleContactInfoSubmitWithValidation.
  section.addWidget(CardService.newTextInput().setFieldName('Customer Name').setTitle('Customer Name (for invoice)').setValue(data['Customer Name'] || ''));
  section.addWidget(CardService.newTextInput().setFieldName('Contact Person').setTitle('Delivery Contact Person').setValue(data['Contact Person'] || data['Customer Name'] || ''));
  
  section.addWidget(CardService.newTextInput().setFieldName('Contact Phone').setTitle('Delivery Contact Phone').setValue(_formatPhone(data['Contact Phone'] || data['Customer Address Phone'] || '')));
  section.addWidget(CardService.newTextInput().setFieldName('Contact Email').setTitle('Delivery Contact Email').setValue(data['Contact Email'] || data['Customer Address Email'] || ''));
  
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address Line 1').setTitle('Delivery Address Line 1').setValue(data['Customer Address Line 1'] || ''));
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address Line 2').setTitle('Delivery Address Line 2').setValue(data['Customer Address Line 2'] || ''));
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address City').setTitle('Delivery City').setValue(data['Customer Address City'] || ''));
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address State').setTitle('Delivery State').setValue(data['Customer Address State'] || ''));
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address ZIP').setTitle('Delivery ZIP').setValue(data['Customer Address ZIP'] || ''));
  
  // Date and Time fields with .setRequired(true) as supported
  const deliveryDatePicker = CardService.newDatePicker().setFieldName("Delivery Date").setTitle("Delivery Date").setValueInMsSinceEpoch(_parseDateToMsEpoch(data['Delivery Date'])).setRequired(true);
  section.addWidget(deliveryDatePicker);

  const deliveryTimeInput = CardService.newSelectionInput().setFieldName('Delivery Time').setTitle('Delivery Time').setRequired(true);
  deliveryTimeInput.setType(CardService.SelectionInputType.DROPDOWN); // Set type directly
  const selectedTime = data['Delivery Time'] || '';
  const startHour = 5; const endHour = 23; // 5 AM to 11 PM
  for (let h = startHour; h <= endHour; h++) { // Loop includes endHour for 11 PM
    for (let m = 0; m < 60; m += 15) {
      const time = Utilities.formatDate(new Date(2000, 0, 1, h, m), Session.getScriptTimeZone(), 'h:mm a');
      deliveryTimeInput.addItem(time, time, selectedTime === time);
    }
  }
  section.addWidget(deliveryTimeInput);

  const utensilsValue = data['Include Utensils?'] || 'Unknown';
  const utensilsInput = CardService.newSelectionInput().setFieldName('Include Utensils?').setTitle('Include Utensils?');
  utensilsInput.setType(CardService.SelectionInputType.DROPDOWN); // Set type directly
  utensilsInput.addItem('Yes', 'Yes', utensilsValue === 'Yes').addItem('No', 'No', utensilsValue === 'No').addItem('Unknown', 'Unknown', utensilsValue !== 'Yes' && utensilsValue !== 'No');
  section.addWidget(utensilsInput);
  if (utensilsValue === 'Yes') {
    section.addWidget(CardService.newTextInput().setFieldName('If yes: how many?').setTitle('How many utensils?').setValue(data['If yes: how many?'] || ''));
  }

  section.addWidget(CardService.newTextInput().setFieldName('Client').setTitle('Client (based on email domain)').setValue(data['Client'] || 'Unknown'));
  
  // Action with validation
  const action = CardService.newAction().setFunctionName('handleContactInfoSubmitWithValidation').setParameters({ orderNum: orderNum });
  const footer = CardService.newFixedFooter().setPrimaryButton(CardService.newTextButton().setText('Confirm Contact & Proceed to Items').setOnClickAction(action).setTextButtonStyle(CardService.TextButtonStyle.FILLED));
  return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle('Step 1: Customer & Order Details')).addSection(section).setFixedFooter(footer).build();
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
  const contactPhone = _getFormInputValue(inputs, 'Contact Phone');
  const customerAddressLine1 = _getFormInputValue(inputs, 'Customer Address Line 1');
  const deliveryDateMs = _getFormInputValue(inputs, 'Delivery Date', true); // Pass true for date inputs
  const deliveryTimeStr = _getFormInputValue(inputs, 'Delivery Time');

  const validationMessages = []; 

  // Validation rules
  if (!deliveryDateMs || !deliveryTimeStr) {
    validationMessages.push("• Delivery Date and Time are required.");
  }
  if (!customerName && !contactPerson) {
    validationMessages.push("• Either 'Customer Name (for invoice)' or 'Delivery Contact Person' is required.");
  }
  if (!contactPhone) { // Ensure a phone number is provided for delivery contact
    validationMessages.push("• A 'Delivery Contact Phone' number is required.");
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
 * @param {GoogleAppsScript.Addons.EventObject} e The event object from the form submission.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse} An action response to navigate to the next card.
 */
function handleContactInfoSubmit(e) {
  const inputs = e.commonEventObject.formInputs;
  const orderNum = e.parameters.orderNum;
  const newData = {};
  let deliveryDateMs = null;
  let deliveryTimeStr = '';

  // Use the helper to safely retrieve all form input values
  for (const key in inputs) {
    if (key === "Delivery Date") { 
      deliveryDateMs = _getFormInputValue(inputs, key, true);
      if (deliveryDateMs) {
          newData[key] = Utilities.formatDate(new Date(deliveryDateMs), Session.getScriptTimeZone(), "MM/dd/yyyy");
      }
    } else if (key === "Delivery Time") {
      newData[key] = _getFormInputValue(inputs, key);
      deliveryTimeStr = newData[key]; // Keep deliveryTimeStr updated for _combineDateAndTime
    }
    else { 
      newData[key] = _getFormInputValue(inputs, key); 
    }
  }

  const userProps = PropertiesService.getUserProperties();
  const existingRaw = userProps.getProperty(orderNum);
  if (!existingRaw) { 
    console.error("Error in handleContactInfoSubmit: Original order data for " + orderNum + " not found.");
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Original order data missing.")).build();
  }
  const existing = JSON.parse(existingRaw);
  // Merge new (edited) data with existing data, prioritizing new data
  const merged = { ...existing, ...newData };

  // Combine date and time into a single timestamp for easier manipulation later
  if (deliveryDateMs && deliveryTimeStr) {
    merged['master_delivery_time_ms'] = _combineDateAndTime(deliveryDateMs, deliveryTimeStr);
    console.log("Master Delivery Time (ms): " + merged['master_delivery_time_ms']);
  }

  // Re-run item extraction if not already present or if it's in a bad format.
  // This is a safety net in case the initial extraction failed or was incomplete.
  if (!merged['Items Ordered'] || !Array.isArray(merged['Items Ordered']) || (merged['Items Ordered'].length > 0 && typeof merged['Items Ordered'][0].description === 'undefined')) {
    const message = GmailApp.getMessageById(merged.messageId);
    const body = message.getPlainBody();
    const itemsParsed = _parseJson(callGemini(_buildStructuredItemExtractionPrompt(body)));
    merged['Items Ordered'] = itemsParsed['Items Ordered'] || [];
  }

  // Save the updated data and proceed to item mapping
  PropertiesService.getUserProperties().setProperty(orderNum, JSON.stringify(merged));
  const itemMappingCard = buildItemMappingAndPricingCard({ parameters: { orderNum: orderNum } });
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(itemMappingCard)).build();
}

// === ITEM MAPPING AND PRICING CARD ===
/**
 * Builds the card for mapping extracted email items to QuickBooks items and reviewing pricing.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object containing the order number.
 * @returns {GoogleAppsScript.Card_Service.Card} The constructed item mapping card.
 */
function buildItemMappingAndPricingCard(e) {
  const orderNum = e.parameters.orderNum;
  const userProps = PropertiesService.getUserProperties();
  const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { 
    console.error("Error in buildItemMappingAndPricingCard: Order data for " + orderNum + " not found.");
    return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("Error")).addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Order data not found for item mapping. Please restart from the email."))).build();
  }
  const orderData = JSON.parse(orderDataString);
  const emailItems = orderData['Items Ordered'];
  if (!emailItems || !Array.isArray(emailItems) || (emailItems.length > 0 && (typeof emailItems[0] !== 'object' || typeof emailItems[0].description === 'undefined'))) { 
    console.error("Error in buildItemMappingAndPricingCard: 'Items Ordered' for " + orderNum + " is missing or not in expected format: " + JSON.stringify(emailItems));
    return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("Item Loading Error")).addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Could not load items for matching. Please try reopening the add-on for this email."))).build();
  }
  const masterQBItems = getMasterQBItems();
  if (masterQBItems.length === 0) { 
    console.error("Error in buildItemMappingAndPricingCard: Master item list ('Item Lookup' sheet) is empty or could not be loaded.");
    return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("Configuration Error")).addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Master item list ('Item Lookup' sheet) could not be loaded. Please check sheet configuration and permissions."))).build();
  }

  const suggestedMatches = getGeminiItemMatches(emailItems, masterQBItems);
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
      const masterMatchDetails = masterQBItems.find(master => master.SKU === item.matched_qb_item_id) || masterQBItems.find(master => master.SKU === FALLBACK_CUSTOM_ITEM_SKU);
      itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`<b>Email Line ${index + 1}:</b> "${item.original_email_description}"`));
      itemsDisplaySection.addWidget(CardService.newTextInput().setFieldName(`item_qty_${index}`).setTitle('Unit Quantity').setValue(item.extracted_main_quantity || '1'));
      itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`AI Match Confidence: ${item.match_confidence || 'N/A'}`));
      const qbItemDropdown = CardService.newSelectionInput().setFieldName(`item_qb_sku_${index}`).setTitle('Item');
      qbItemDropdown.setType(CardService.SelectionInputType.DROPDOWN); // Set type directly
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
      itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`Selected SKU: ${masterMatchDetails ? masterMatchDetails.SKU : 'N/A'}`));
      if (masterMatchDetails && masterMatchDetails.SKU !== FALLBACK_CUSTOM_ITEM_SKU) {
        itemsDisplaySection.addWidget(CardService.newTextParagraph().setText("--- Item Details ---"));
        itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`<b>Name:</b> ${masterMatchDetails.Name}`));
        itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`<b>Category:</b> ${masterMatchDetails.Category || 'N/A'}`));
        itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`<b>Size:</b> ${masterMatchDetails.Size || 'N/A'}`));
        const flavorsFromLookup = [masterMatchDetails['Flavor 1'], masterMatchDetails['Flavor 2'], masterMatchDetails['Flavor 3'], masterMatchDetails['Flavor 4'], masterMatchDetails['Flavor 5']].filter(Boolean).join(', ');
        itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`<b>Std. Flavors:</b> ${flavorsFromLookup || 'N/A'}`));
        itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`<b>Descriptor:</b> ${masterMatchDetails.Descriptor || 'N/A'}`));
        itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`<b>Base Price:</b> $${(masterMatchDetails.Price || 0).toFixed(2)}`));
        itemsDisplaySection.addWidget(CardService.newTextParagraph().setText("--- End Item Details ---"));
      }
      itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`<i>AI Parsed Flavors/Details: ${item.parsed_flavors_or_details || "None"}</i>`));
      if(item.identified_flavors && item.identified_flavors.length > 0){
        itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`<i>AI Identified Flavors: ${item.identified_flavors.join(", ")}</i>`));
      }
      itemsDisplaySection.addWidget(CardService.newTextInput().setFieldName(`item_kitchen_notes_${index}`).setTitle('Confirmed Flavors/Notes for Kitchen').setValue(item.parsed_flavors_or_details || '').setMultiline(true));
      const initialPrice = masterMatchDetails ? (masterMatchDetails.Price !== undefined ? masterMatchDetails.Price : 0) : 0;
      itemsDisplaySection.addWidget(CardService.newTextParagraph().setText(`<b>Unit Price:</b> $${initialPrice.toFixed(2)}`));
      itemsDisplaySection.addWidget(CardService.newSelectionInput().setFieldName(`item_remove_${index}`).setTitle("Remove this item?").setType(CardService.SelectionInputType.CHECKBOX).addItem("Yes, remove", "true", false));
      itemsDisplaySection.addWidget(CardService.newDivider());
    });
  }
  card.addSection(itemsDisplaySection);

  const manualAddSection = CardService.newCardSection().setHeader("Manually Add New Item").setCollapsible(true);
  const newItemDropdown = CardService.newSelectionInput().setFieldName("new_item_qb_sku").setTitle("Select Item");
  newItemDropdown.setType(CardService.SelectionInputType.DROPDOWN); // Set type directly
  newItemDropdown.addItem("--- Select Item to Add ---", "", true);
  masterQBItems.forEach(masterItem => {
    if (!masterItem.SKU) return;
    newItemDropdown.addItem(`${masterItem.Name} (SKU: ${masterItem.SKU})`, masterItem.SKU, false);
  });
  manualAddSection.addWidget(newItemDropdown);
  manualAddSection.addWidget(CardService.newTextInput().setFieldName("new_item_qty").setTitle("Unit Quantity").setValue("1"));
  manualAddSection.addWidget(CardService.newTextInput().setFieldName("new_item_kitchen_notes").setTitle("Flavors/Notes for Kitchen").setMultiline(true));
  manualAddSection.addWidget(CardService.newTextParagraph().setText("Unit Price will be based on selected SKU."));
  card.addSection(manualAddSection);

  // Additional Charges Section
  const additionalChargesSection = CardService.newCardSection().setHeader("Additional Charges (Optional)").setCollapsible(true);
  // Pre-populate tip amount calculated earlier
  additionalChargesSection.addWidget(CardService.newTextInput().setFieldName("tip_amount").setTitle("Tip Amount ($)").setValue((orderData['TipAmount'] || 0).toFixed(2)));
  additionalChargesSection.addWidget(CardService.newTextInput().setFieldName("other_charges_amount").setTitle("Other Charges Amount ($)").setValue("0.00"));
  additionalChargesSection.addWidget(CardService.newTextInput().setFieldName("other_charges_description").setTitle("Other Charges Description"));
  card.addSection(additionalChargesSection);

  const action = CardService.newAction().setFunctionName('handleItemMappingSubmit')
    .setParameters({ orderNum: orderNum, ai_item_count: suggestedMatches.length.toString() });
  card.setFixedFooter(CardService.newFixedFooter().setPrimaryButton(CardService.newTextButton().setText('Confirm All Items & Proceed').setOnClickAction(action).setTextButtonStyle(CardService.TextButtonStyle.FILLED)));
  return card.build();
}

/**
 * Handles the submission of item mapping and pricing, calculates final totals.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object from the form submission.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse} An action response to navigate to the next card.
 */
function handleItemMappingSubmit(e) {
  const formInputs = e.formInputs || (e.commonEventObject && e.commonEventObject.formInputs);
  if (!formInputs) { console.error("Error in handleItemMappingSubmit: formInputs is undefined."); return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Could not read form data.")).build(); }
  const orderNum = e.parameters.orderNum; const aiItemCount = parseInt(e.parameters.ai_item_count) || 0;
  if (!orderNum) { console.error("Error in handleItemMappingSubmit: Order number is missing."); return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Order number missing.")).build(); }
  const userProps = PropertiesService.getUserProperties(); const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { console.error("Error in handleItemMappingSubmit: Original order data for " + orderNum + " not found."); return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Original order data not found.")).build(); }
  const orderData = JSON.parse(orderDataString); const masterQBItems = getMasterQBItems(); const confirmedQuickBooksItems = [];
  
  for (let i = 0; i < aiItemCount; i++) {
    const removeThisItem = formInputs[`item_remove_${i}`] && formInputs[`item_remove_${i}`][0] === "true";
    if (removeThisItem) { console.log(`Item ${i} marked for removal.`); continue; }

    const qtyString = formInputs[`item_qty_${i}`] && formInputs[`item_qty_${i}`][0];
    const qbItemSKU = formInputs[`item_qb_sku_${i}`] && formInputs[`item_qb_sku_${i}`][0]; 
    const kitchenNotes = formInputs[`item_kitchen_notes_${i}`] && formInputs[`item_kitchen_notes_${i}`][0] || "";
    const originalEmailItem = orderData['Items Ordered'] && orderData['Items Ordered'][i];
    const originalDescription = originalEmailItem ? originalEmailItem.description : "N/A";

    if (qbItemSKU && qtyString) {
      const masterItemDetails = masterQBItems.find(master => master.SKU === qbItemSKU);
      let unitPrice = 0;
      let itemName = "Custom Item";
      let itemSKU = qbItemSKU;

      if (masterItemDetails) {
        unitPrice = masterItemDetails.Price || 0;
        itemName = masterItemDetails.Name;
      } else if (qbItemSKU === FALLBACK_CUSTOM_ITEM_SKU) {
        unitPrice = 0; // Fallback custom item can have 0 price initially
      } else {
        console.warn("Master details not found for SKU: " + qbItemSKU + ". Defaulting to fallback.");
        itemSKU = FALLBACK_CUSTOM_ITEM_SKU;
        unitPrice = 0; // Set to 0 if SKU not found and not the designated fallback
      }
      confirmedQuickBooksItems.push({ quickbooks_item_id: itemSKU, quickbooks_item_name: itemName, sku: itemSKU, quantity: parseInt(qtyString) || 1, unit_price: unitPrice, kitchen_notes_and_flavors: kitchenNotes, original_email_description: originalDescription });
    }
  }

  // Handle manually added new item
  const newItemQbSKU = formInputs.new_item_qb_sku && formInputs.new_item_qb_sku[0];
  const newItemQtyString = formInputs.new_item_qty && formInputs.new_item_qty[0];
  if (newItemQbSKU && newItemQbSKU !== "" && newItemQtyString) {
    const masterItemDetails = masterQBItems.find(master => master.SKU === newItemQbSKU);
    if (masterItemDetails) {
      const unitPrice = masterItemDetails.Price || 0;
      const newItemKitchenNotes = formInputs.new_item_kitchen_notes && formInputs.new_item_kitchen_notes[0] || "";
      confirmedQuickBooksItems.push({ quickbooks_item_id: newItemQbSKU, quickbooks_item_name: masterItemDetails.Name, sku: masterItemDetails.SKU, quantity: parseInt(newItemQtyString) || 1, unit_price: unitPrice, kitchen_notes_and_flavors: newItemKitchenNotes, original_email_description: newItemKitchenNotes || "Manually Added: " + masterItemDetails.Name });
    }
  }

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
 * @param {GoogleAppsScript.Addons.EventObject} e The event object.
 * @returns {GoogleAppsScript.Card_Service.Card} The constructed card.
 */
function buildInvoiceActionsCard(e) {
  const orderNum = e.parameters.orderNum;
  const card = CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle(`Step 3: Final Review & Actions for ${orderNum}`));
  const orderDetailsSection = CardService.newCardSection().setHeader("Order Details");
  const userProps = PropertiesService.getUserProperties(); const orderDataString = userProps.getProperty(orderNum);
  if (orderDataString) {
    const orderData = JSON.parse(orderDataString);
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Ordering Person:</b> " + (orderData['Ordering Person Name'] || 'N/A') + " (" + (orderData['Ordering Person Email'] || 'N/A') + ")</i>"));
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Customer Name (for Invoice):</b> " + (orderData['Customer Name'] || 'N/A') + "</i>"));
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Delivery Contact:</b> " + (orderData['Contact Person'] || 'N/A') + "</i>"));
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Delivery Phone:</b> " + _formatPhone(orderData['Contact Phone']) || 'N/A' + "</i>"));
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Delivery Email:</b> " + (orderData['Contact Email'] || 'N/A') + "</i>"));

    let address = "<i><b>Address:</b> ";
    if (orderData['Customer Address Line 1']) address += orderData['Customer Address Line 1'];
    if (orderData['Customer Address Line 2']) address += (address === "<i><b>Address:</b> " ? "" : ", ") + orderData['Customer Address Line 2'];
    if (orderData['Customer Address City']) address += (address === "<i><b>Address:</b> " ? "" : ", ") + orderData['Customer Address City'];
    if (orderData['Customer Address State']) address += (address === "<i><b>Address:</b> " ? "" : ", ") + orderData['Customer Address State'];
    if (orderData['Customer Address ZIP']) address += " " + orderData['Customer Address ZIP'];
    address += "</i>";
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText(address.length <= "<i><b>Address:</b> </i>".length + 1 ? "<i><b>Address:</b> N/A</i>" : address));
    
    let deliveryDateFormatted = orderData['Delivery Date'];
    if (deliveryDateFormatted && typeof deliveryDateFormatted === 'string' && !deliveryDateFormatted.includes('/')) {
        deliveryDateFormatted = Utilities.formatDate(new Date(parseInt(deliveryDateFormatted)), Session.getScriptTimeZone(), "MM/dd/yyyy");
    } else if (deliveryDateFormatted && typeof deliveryDateFormatted === 'string' && deliveryDateFormatted.match(/^\d{4}-\d{2}-\d{2}/)) {
        deliveryDateFormatted = Utilities.formatDate(new Date(deliveryDateFormatted.replace(/-/g, '/')), Session.getScriptTimeZone(), "MM/dd/yyyy");
    }
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Delivery:</b> " + (deliveryDateFormatted || 'N/A') + " at " + (orderData['Delivery Time'] || 'N/A') + "</i>"));
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Client:</b> " + (orderData['Client'] || 'N/A') + "</i>"));
  } else { orderDetailsSection.addWidget(CardService.newTextParagraph().setText("Could not retrieve order details.")); }
  card.addSection(orderDetailsSection);

  const itemSummarySection = CardService.newCardSection().setHeader("Confirmed Items Summary");
  let grandTotal = 0;
  if (orderDataString) {
    const orderData = JSON.parse(orderDataString);
    const confirmedItems = orderData['ConfirmedQBItems'];
    if (confirmedItems && Array.isArray(confirmedItems) && confirmedItems.length > 0) {
      confirmedItems.forEach(item => {
        const itemTotal = (item.quantity || 0) * (item.unit_price || 0);
        grandTotal += itemTotal;
        itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>${item.original_email_description || item.quickbooks_item_name}</b><br>Qty: ${item.quantity}, Unit Price: $${(item.unit_price || 0).toFixed(2)}, Total: $${itemTotal.toFixed(2)}${item.kitchen_notes_and_flavors ? '<br><font color="#666666"><i>Kitchen Notes: ' + item.kitchen_notes_and_flavors + '</i></font>' : ''}`));
      });
    }
    
    if (orderData['TipAmount'] > 0) {
        itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Tip:</b> $${orderData['TipAmount'].toFixed(2)}`));
        grandTotal += orderData['TipAmount'];
    }
    if (orderData['OtherChargesAmount'] > 0) {
        itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>${orderData['OtherChargesDescription'] || 'Other Charges'}:</b> $${orderData['OtherChargesAmount'].toFixed(2)}`));
        grandTotal += orderData['OtherChargesAmount'];
    }
    
    // Delivery Fee - always added now
    let deliveryFee = BASE_DELIVERY_FEE;
    if (orderData['master_delivery_time_ms']) {
        const deliveryHour = new Date(orderData['master_delivery_time_ms']).getHours();
        if (deliveryHour >= DELIVERY_FEE_CUTOFF_HOUR) {
            deliveryFee = AFTER_4PM_DELIVERY_FEE;
        }
    }
    itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Delivery Fee:</b> $${deliveryFee.toFixed(2)}`));
    grandTotal += deliveryFee;

    // Utensil Costs
    if (orderData['Include Utensils?'] === 'Yes') {
        const numUtensils = parseInt(orderData['If yes: how many?']) || 0;
        if (numUtensils > 0) {
            const utensilTotalCost = numUtensils * COST_PER_UTENSIL_SET;
            itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Utensils (${numUtensils} sets):</b> $${utensilTotalCost.toFixed(2)}`));
            grandTotal += utensilTotalCost;
        }
    }

    if (confirmedItems && confirmedItems.length > 0 || orderData['TipAmount'] > 0 || orderData['OtherChargesAmount'] > 0 || deliveryFee > 0 || (orderData['Include Utensils?'] === 'Yes' && parseInt(orderData['If yes: how many?']) > 0)) {
      itemSummarySection.addWidget(CardService.newDivider()); 
      itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Estimated Grand Total: $${grandTotal.toFixed(2)}</b>`));
    } else { 
      itemSummarySection.addWidget(CardService.newTextParagraph().setText("No items or charges confirmed for this order.")); 
    }
  } else { itemSummarySection.addWidget(CardService.newTextParagraph().setText("Could not retrieve item data for summary.")); }
  card.addSection(itemSummarySection);

  const actionsSection = CardService.newCardSection();
  const generateDocsAndEmailAction = CardService.newAction().setFunctionName('handleGenerateInvoiceAndEmail').setParameters({ orderNum: orderNum });
  actionsSection.addWidget(CardService.newTextButton().setText("📄厨️ Generate Documents & Prepare Email").setOnClickAction(generateDocsAndEmailAction).setTextButtonStyle(CardService.TextButtonStyle.FILLED));
  const backToItemsAction = CardService.newAction().setFunctionName('buildItemMappingAndPricingCard').setParameters({ orderNum: orderNum }); 
  actionsSection.addWidget(CardService.newTextButton().setText("✏️ Review/Edit Mapped Items Again").setOnClickAction(backToItemsAction));
  card.addSection(actionsSection);
  return card.build();
}

/**
 * Handles the generation of invoice and kitchen documents and prepares the email reply.
 * @param {GoogleAppsScript.Addons.EventObject} e The event object.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse} An action response.
 */
function handleGenerateInvoiceAndEmail(e) {
  const orderNum = e.parameters.orderNum;
  let invoiceSheetInfo = null;
  let kitchenSheetInfo = null;
  try {
    invoiceSheetInfo = populateInvoiceSheet(orderNum);
    if (!invoiceSheetInfo || !invoiceSheetInfo.id || !invoiceSheetInfo.url || !invoiceSheetInfo.name) { throw new Error("Failed to populate invoice sheet or retrieve its details."); }
    console.log("Invoice sheet populated: " + invoiceSheetInfo.name + " (ID: " + invoiceSheetInfo.id + ")");
    
    kitchenSheetInfo = populateKitchenSheet(orderNum); // Always generate kitchen sheet
    if (!kitchenSheetInfo || !kitchenSheetInfo.id || !kitchenSheetInfo.url || !kitchenSheetInfo.name) { console.warn("Warning: Failed to populate kitchen sheet for order " + orderNum + ". Continuing with invoice PDF and email."); }
    else { console.log("Kitchen sheet populated: " + kitchenSheetInfo.name + " (ID: " + kitchenSheetInfo.id + ")"); }
    
    const emailInfo = createPdfAndPrepareEmailReply(orderNum, invoiceSheetInfo.id, invoiceSheetInfo.name);
    if (!emailInfo || !emailInfo.draftId) { throw new Error("Failed to create Invoice PDF or prepare email draft."); }
    
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
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(successCard.build())).build();
  } catch (err) {
    console.error("Error in handleGenerateInvoiceAndEmail: " + err.toString() + (err.stack ? ("\nStack: " + err.stack) : ""));
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