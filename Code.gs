// V13
// This script contains the core workflow logic for the Google Workspace Add-on.
// Configuration constants are in Constants.gs.
// Utility functions are in Utils.gs.
// Gemini prompts are in Prompts.gs.

// === HOMEPAGE CARD (Primary Entry Point if manifest is updated) ===
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

// Placeholder for Penn Invoice workflow
function handlePennInvoiceWorkflowPlaceholder(e) {
  const card = CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle("Penn Invoice Workflow"))
    .addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("This feature is under development.")))
    .build();
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(card)).build();
}


// === ENTRY POINT for Catering Email Workflow (can be called from homepage or directly) ===
function buildAddOnCard(e) {
  console.log("buildAddOnCard triggered. Event: " + JSON.stringify(e));
  let msgId;

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

  // Gemini prompts for contact info and item extraction
  const contactInfoParsed = _parseJson(callGemini(_buildContactInfoPrompt(body)));
  const itemsParsed = _parseJson(callGemini(_buildStructuredItemExtractionPrompt(body)));

  const orderingPersonName = _extractNameFromEmail(senderEmailFull); // This is the sender
  const orderingPersonEmail = _extractActualEmail(senderEmailFull);

  const timestampSuffix = Date.now().toString().slice(-5);
  const randomPrefix = Math.floor(Math.random() * 900 + 100).toString();
  const orderNum = (randomPrefix + timestampSuffix).slice(0,8).padStart(8,'0');

  const client = _matchClient(orderingPersonEmail);
  console.log("Matched Client: " + client);

  const data = {};
  // Store sender's info separately (Sofia Deleon / elmerkury.com)
  data['Ordering Person Name'] = orderingPersonName;
  data['Ordering Person Email'] = orderingPersonEmail;

  // Set Customer Name (e.g., Ashley Duchi), defaulting to sender if not found by Gemini
  data['Customer Name'] = contactInfoParsed['Customer Name'] || orderingPersonName;
  
  // Set Delivery Contact Person (e.g., Romina), defaulting to Customer Name or sender if not found
  data['Contact Person'] = contactInfoParsed['Delivery Contact Person'] || data['Customer Name'] || orderingPersonName;
  
  // Prioritize Delivery Contact Phone, then Customer Address Phone, then Ordering Person's (sender's) phone if available, else empty string.
  data['Contact Phone'] = contactInfoParsed['Delivery Contact Phone'] || contactInfoParsed['Customer Address Phone'] || '';
  // Prioritize Delivery Contact Email, then Customer Address Email, then Ordering Person's email.
  data['Contact Email'] = contactInfoParsed['Delivery Contact Email'] || contactInfoParsed['Customer Address Email'] || orderingPersonEmail;

  // Invoice recipient email
  data['Customer Address Email'] = contactInfoParsed['Customer Address Email'] || orderingPersonEmail;


  data['Client'] = client;
  data['orderNum'] = orderNum;
  data['messageId'] = msgId;
  data['threadId'] = message.getThread().getId();

  // Merge remaining extracted contact info (address, dates, utensils)
  // These will override if Gemini provided them for specific fields.
  Object.assign(data, contactInfoParsed);

  data['Items Ordered'] = itemsParsed['Items Ordered'] || [];

  // Calculate preliminary grand total and tip
  let preliminarySubtotal = 0;
  // This preliminary calculation is for the initial tip display.
  // The full calculation will occur in handleItemMappingSubmit with actual QB item prices.
  data['Items Ordered'].forEach(item => {
    // Attempt to parse price from description if present (e.g., "Pupusas Tray – Large (<span class="math-inline">140\)"\)
const priceMatch \= item\.description\.match\(/\\</span>(\d+(\.\d{2})?)/);
    const itemPrice = priceMatch ? parseFloat(priceMatch[1]) : 0;
    preliminarySubtotal += (parseInt(item.quantity) || 1) * itemPrice;
  });

  // Calculate 10% tip from preliminary subtotal as per request
  data['TipAmount'] = (preliminarySubtotal * 0.10);
  console.log("Preliminary Subtotal for Tip: $" + preliminarySubtotal.toFixed(2));
  console.log("Calculated Initial Tip Amount: $" + data['TipAmount'].toFixed(2));

  PropertiesService.getUserProperties().setProperty(orderNum, JSON.stringify(data));
  return buildReviewContactCard({ parameters: { orderNum: orderNum } });
}


// === BUILD CUSTOMER CONTACT REVIEW CARD ===
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
  
  // Display Ordering Person (Sender)
  section.addWidget(CardService.newTextParagraph().setText("<b>Ordering Person (Sender):</b> " + (data['Ordering Person Name'] || 'N/A') + " (" + (data['Ordering Person Email'] || 'N/A') + ")"));

  // Editable fields for Customer and Delivery Contact
  section.addWidget(CardService.newTextInput().setFieldName('Customer Name').setTitle('Customer Name (for invoice)').setValue(data['Customer Name'] || '').setHint('e.g., Penn Medicine, Ashley Duchi').setRequired(true));
  section.addWidget(CardService.newTextInput().setFieldName('Contact Person').setTitle('Delivery Contact Person').setValue(data['Contact Person'] || data['Customer Name'] || '').setHint('e.g., Romina').setRequired(true));
  
  // The phone number for delivery contact should be used for validation and pre-filled with customer address phone if not found
  section.addWidget(CardService.newTextInput().setFieldName('Contact Phone').setTitle('Delivery Contact Phone').setValue(_formatPhone(data['Contact Phone'] || data['Customer Address Phone'] || '')).setHint('(XXX) XXX-XXXX').setRequired(true));
  section.addWidget(CardService.newTextInput().setFieldName('Contact Email').setTitle('Delivery Contact Email').setValue(data['Contact Email'] || data['Customer Address Email'] || '').setHint('e.g., delivery@example.com'));
  
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address Line 1').setTitle('Delivery Address Line 1').setValue(data['Customer Address Line 1'] || '').setRequired(true));
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address Line 2').setTitle('Delivery Address Line 2').setValue(data['Customer Address Line 2'] || ''));
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address City').setTitle('Delivery City').setValue(data['Customer Address City'] || ''));
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address State').setTitle('Delivery State').setValue(data['Customer Address State'] || ''));
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address ZIP').setTitle('Delivery ZIP').setValue(data['Customer Address ZIP'] || ''));
  
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

  section.addWidget(CardService.newTextInput().setFieldName('Client').setTitle('Client (based on email domain)').setValue(data['Client'] || 'Unknown').setHint('e.g., University of Pennsylvania'));
  
  // Action with validation
  const action = CardService.newAction().setFunctionName('handleContactInfoSubmitWithValidation').setParameters({ orderNum: orderNum });
  const footer = CardService.newFixedFooter().setPrimaryButton(CardService.newTextButton().setText('Confirm Contact & Proceed to Items').setOnClickAction(action).setTextButtonStyle(CardService.TextButtonStyle.FILLED));
  return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle('Step 1: Customer & Order Details')).addSection(section).setFixedFooter(footer).build();
}

/**
 * Handles the submission of contact information with validation.
 * This is the function linked to the "Confirm Contact & Proceed" button.
 */
function handleContactInfoSubmitWithValidation(e) {
  const inputs = e.commonEventObject.formInputs;
  const orderNum = e.parameters.orderNum;

  // Client-side validation
  const customerName = inputs['Customer Name']?.stringInputs?.value?.[0];
  const contactPerson = inputs['Contact Person']?.stringInputs?.value?.[0];
  const contactPhone = inputs['Contact Phone']?.stringInputs?.value?.[0];
  const customerAddressLine1 = inputs['Customer Address Line 1']?.stringInputs?.value?.[0];
  const deliveryDateMs = inputs['Delivery Date']?.dateInput?.msSinceEpoch;
  const deliveryTimeStr = inputs['Delivery Time']?.stringInputs?.value?.[0];

  const validationMessages = [];

  if (!deliveryDateMs || !deliveryTimeStr) {
    validationMessages.push("• Delivery Date and Time are required.");
  }
  if (!customerName && !contactPerson) {
    validationMessages.push("• Either Customer Name or Delivery Contact Person is required.");
  }
  if (!contactPhone) { // Only check Contact Phone as it covers both
    validationMessages.push("• A Delivery Contact Phone number is required.");
  }
  if (!customerAddressLine1) {
    validationMessages.push("• Delivery Address Line 1 is required.");
  }

  if (validationMessages.length > 0) {
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
// This function is now called by handleContactInfoSubmitWithValidation AFTER validation passes.
function handleContactInfoSubmit(e) {
  const inputs = e.commonEventObject.formInputs;
  const orderNum = e.parameters.orderNum;
  const newData = {};
  let deliveryDateMs = null;
  let deliveryTimeStr = '';

  for (const key in inputs) {
    if (key === "Delivery Date") { 
      deliveryDateMs = inputs[key].dateInput.msSinceEpoch;
      newData[key] = Utilities.formatDate(new Date(deliveryDateMs), Session.getScriptTimeZone(), "MM/dd/yyyy");
    } else if (key === "Delivery Time") {
      deliveryTimeStr = inputs[key].stringInputs?.value?.[0] || '';
      newData[key] = deliveryTimeStr;
    }
    else { 
      newData[key] = inputs[key].stringInputs?.value?.[0] || ''; 
    }
  }

  const userProps = PropertiesService.getUserProperties();
  const existingRaw = userProps.getProperty(orderNum);
  if (!existingRaw) { 
    console.error("Error in handleContactInfoSubmit: Original order data for " + orderNum + " not found.");
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Original order data missing.")).build();
  }
  const existing = JSON.parse(existingRaw);
  // When saving, prioritize the edited fields from the form.
  const merged = { ...existing, ...newData };

  if (deliveryDateMs && deliveryTimeStr) {
    merged['master_delivery_time_ms'] = _combineDateAndTime(deliveryDateMs, deliveryTimeStr);
    console.log("Master Delivery Time (ms): " + merged['master_delivery_time_ms']);
  }

  // Re-run item extraction if not already present or if it's in a bad format
  // This is a safety net in case the initial extraction failed or was incomplete.
  if (!merged['Items Ordered'] || !Array.isArray(merged['Items Ordered']) || (merged['Items Ordered'].length > 0 && typeof merged['Items Ordered'][0].description === 'undefined')) {
    const message = GmailApp.getMessageById(merged.messageId);
    const body = message.getPlainBody();
    const itemsParsed = _parseJson(callGemini(_buildStructuredItemExtractionPrompt(body)));
    merged['Items Ordered'] = itemsParsed['Items Ordered'] || [];
  }

  userProps.setProperty(orderNum, JSON.stringify(merged));
  const itemMappingCard = buildItemMappingAndPricingCard({ parameters: { orderNum: orderNum } });
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(itemMappingCard)).build();
}

// === ITEM MAPPING AND PRICING CARD ===
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
  additionalChargesSection.addWidget(CardService.newTextInput().setFieldName("tip_amount").setTitle("Tip Amount (<span class="math-inline">\)"\)\.setValue\(\(orderData\['TipAmount'\] \|\| 0\)\.toFixed\(2\)\)\);
additionalChargesSection\.addWidget\(CardService\.newTextInput\(\)\.setFieldName\("other\_charges\_amount"\)\.setTitle\("Other Charges Amount \(</span>)").setValue("0.00"));
  additionalChargesSection.addWidget(CardService.newTextInput().setFieldName("other_charges_description").setTitle("Other Charges Description"));
  card.addSection(additionalChargesSection);

  const action = CardService.newAction().setFunctionName('handleItemMappingSubmit')
    .setParameters({ orderNum: orderNum, ai_item_count: suggestedMatches.length.toString() });
  card.setFixedFooter(CardService.newFixedFooter().setPrimaryButton(CardService.newTextButton().setText('Confirm All Items & Proceed').setOnClickAction(action).setTextButtonStyle(CardService.TextButtonStyle.FILLED)));
  return card.build();
}

/**
 * Handles the submission of item mapping and pricing, calculates final totals.
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

  userProps.setProperty(orderNum, JSON.stringify(orderData));
  console.log("Confirmed items and charges for order " + orderNum + ": " + JSON.stringify(orderData));
  const invoiceActionsCard = buildInvoiceActionsCard({ parameters: { orderNum: orderNum } });
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(invoiceActionsCard)).build();
}

/**
 * Builds the final review and actions card before document generation.
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
      newSheet.getRange(ITEM_UNIT_PRICE_COL_INVOICE + currentRow).setValue(item.unit_price).setNumberFormat("<span class="math-inline">\#,\#\#0\.00"\); 
const lineTotal \= \(item\.quantity \|\| 0\) \* \(item\.unit\_price \|\| 0\);
newSheet\.getRange\(ITEM\_TOTAL\_PRICE\_COL\_INVOICE \+ currentRow\)\.setValue\(lineTotal\)\.setNumberFormat\("</span>#,##0.00"); 
      grandTotal += lineTotal; currentRow++;
    });

    // Add Tip and Other Charges to Invoice Sheet
    let tipAmount = orderData['TipAmount'] || 0;
    let otherChargesAmount = orderData['OtherChargesAmount'] || 0;
    let otherChargesDescription = orderData['OtherChargesDescription'] || "Other Charges";

    if (tipAmount > 0) {
        newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow).setValue("Tip").setWrap(false);
        newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setValue(tipAmount).setNumberFormat("<span class="math-inline">\#,\#\#0\.00"\);
grandTotal \+\= tipAmount;
currentRow\+\+;
\}
if \(otherChargesAmount \> 0\) \{
newSheet\.getRange\(ITEM\_DESCRIPTION\_COL\_INVOICE \+ currentRow\)\.setValue\(otherChargesDescription\)\.setWrap\(false\);
newSheet\.getRange\(ITEM\_TOTAL\_PRICE\_COL\_INVOICE \+ currentRow\)\.setValue\(otherChargesAmount\)\.setNumberFormat\("</span>#,##0.00");
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
            newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setValue(utensilTotalCost).setNumberFormat("<span class="math-inline">\#,\#\#0\.00"\);
grandTotal \+\= utensilTotalCost;
currentRow\+\+;
\}
\}
const grandTotalDescCell \= newSheet\.getRange\(ITEM\_DESCRIPTION\_COL\_INVOICE \+ currentRow\);
grandTotalDescCell\.setValue\("Grand Total\:"\)\.setFontWeight\("bold"\)\.setWrap\(false\);
const grandTotalValueCell \= newSheet\.getRange\(ITEM\_TOTAL\_PRICE\_COL\_INVOICE \+ currentRow\);
grandTotalValueCell\.setValue\(grandTotal\)\.setNumberFormat\("</span>#,##0.00").setFontWeight("bold").setHorizontalAlignment("right").setWrap(false);
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
    
    const pdfBlob = sheetToExport.getAs('application/pdf').setName(`${populatedSheetName}.pdf`); // Export only the specific sheet
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
    const body = `Dear ${orderData['Contact Person'] || orderData['Customer Name'] || 'Valued Customer'},\n\nPlease find attached the invoice for your recent catering order (${orderNum}).\n\nDelivery is scheduled for ${orderData['Delivery Date']} around ${orderData['Delivery Time']}.\n\nThank you for your business!\n\nBest regards,\n[Your Company Name]`;
    
    console.log(`PDFEmail: Preparing draft reply to: ${recipient}`);
    const draft = messageToReplyTo.createDraftReply(body, { htmlBody: body.replace(/\n/g, '<br>'), attachments: [pdfBlob], to: recipient });
    console.log("PDFEmail: Draft email created. ID: " + draft.getId());
    return { pdfBlob: pdfBlob, draft: draft, draftId: draft.getId() };
  } catch (e) { console.error("Error in createPdfAndPrepareEmailReply for order " + orderNum + ": " + e.toString() + (e.stack ? ("\nStack: " + e.stack) : "")); throw new Error("Failed to create PDF or prepare email: " + e.message); }
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