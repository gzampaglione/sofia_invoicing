// === USER INSTRUCTIONS FOR REVIEWING REGENERATED CODE ===
// 1. Manifest Update: If you implement the createHomepageCard, update `onTriggerFunction` in appsscript.json.
//    For now, the primary entry point for the catering workflow is buildAddOnCard.
// 2. API Key: Ensure 'GL_API_KEY' is correctly set in Script Properties.
// 3. Sheet Configuration:
//    - SHEET_ID: Verify this points to your main Google Sheet.
//    - ITEM_LOOKUP_SHEET_NAME: Verify tab name for your master item list.
//    - INVOICE_TEMPLATE_SHEET_NAME: Verify tab name for your invoice template.
//    - KITCHEN_SHEET_TEMPLATE_NAME: Verify tab name for your kitchen sheet template.
//    - FALLBACK_CUSTOM_ITEM_SKU: Ensure an item with this SKU exists in your "Item Lookup" sheet.
// 4. Client Rules: The CLIENT_RULES_LOOKUP constant now holds client matching rules. Update it as needed.
// 5. Permissions (OAuth Scopes): This script interacts with Gmail, Sheets, and Drive (for PDFs).
//    Ensure your appsscript.json manifest includes necessary scopes (see previous discussion for examples).
// 6. Testing: Test incrementally. Use console.log() (for Cloud Logging) or Logger.log() for debugging.
// 7. Prompts: Review Gemini prompts to ensure they align with your data extraction and matching needs.
// 8. Invoice & Kitchen Template Cell Mapping: Carefully review all cell/column mapping constants
//    and adjust them to EXACTLY match your respective sheet layouts.
// === END USER INSTRUCTIONS ===

// === CONFIGURATION ===
const API_KEY = PropertiesService.getScriptProperties().getProperty('GL_API_KEY');
const SHEET_ID = '1qlEw5k5K-Tqg0joxXeiKtpgr6x8BzclobI-E-_mORhY'; // Main Spreadsheet ID
const ITEM_LOOKUP_SHEET_NAME = "Item Lookup";
const INVOICE_TEMPLATE_SHEET_NAME = "INVOICE_TEMPLATE";
const KITCHEN_SHEET_TEMPLATE_NAME = "KITCHEN_SHEET_TEMPLATE";

const FALLBACK_CUSTOM_ITEM_SKU = "CUSTOM_SKU"; // Ensure this SKU exists in your Item Lookup sheet

const CLIENT_RULES_LOOKUP = [
  { rule: "@wharton.upenn.edu", clientName: "University of Pennsylvania - Wharton" },
  { rule: "law.upenn.edu", clientName: "University of Pennsylvania - Law School"},
  { rule: "@upenn.edu", clientName: "University of Pennsylvania" }
].sort((a, b) => b.rule.length - a.rule.length); // Sort by rule length, descending


// === INVOICE TEMPLATE CELL MAPPING CONSTANTS ===
const ORDER_NUM_CELL = "D7";
const CUSTOMER_NAME_CELL = "B12";
const ADDRESS_LINE_1_CELL = "B13";
const ADDRESS_LINE_2_CELL = "B14";
const CITY_STATE_ZIP_CELL = "B15";
const DELIVERY_DATE_CELL_INVOICE = "E15";
const DELIVERY_TIME_CELL_INVOICE = "G15";

const ITEM_START_ROW_INVOICE = 19;
const ITEM_DESCRIPTION_COL_INVOICE = "B";
const ITEM_QTY_COL_INVOICE = "E";
const ITEM_UNIT_PRICE_COL_INVOICE = "F";
const ITEM_TOTAL_PRICE_COL_INVOICE = "G";

// === KITCHEN SHEET TEMPLATE CELL MAPPING CONSTANTS ===
const KITCHEN_CUSTOMER_PHONE_CELL = "B1";
const KITCHEN_DELIVERY_DATE_CELL = "C9";
const KITCHEN_DELIVERY_TIME_CELL = "F9";

const KITCHEN_ITEM_START_ROW = 12;
const KITCHEN_QTY_COL = "B";
const KITCHEN_SIZE_COL = "C";
const KITCHEN_ITEM_NAME_COL = "D";
const KITCHEN_FILLING_COL = "E";
const KITCHEN_NOTES_COL = "F";


// === HOMEPAGE CARD (Primary Entry Point if manifest is updated) ===
function createHomepageCard(e) {
  console.log("createHomepageCard triggered. Event: " + JSON.stringify(e));
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Select Workflow"));
  const section = CardService.newCardSection();

  const processEmailAction = CardService.newAction()
    .setFunctionName("buildAddOnCard")
    // Pass the original event parameters, which might include messageId if triggered contextually
    // or if the homepage itself was triggered by opening a message.
    .setParameters(e && e.parameters ? e.parameters : (e ? e : {}));


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

  // Try to get messageId from various possible event structures
  if (e && e.messageMetadata && e.messageMetadata.messageId) {
    msgId = e.messageMetadata.messageId;
  } else if (e && e.gmail && e.gmail.messageId) {
    msgId = e.gmail.messageId;
  } else if (e && e.parameters && e.parameters.messageId) { // If passed from homepage via parameters
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
  const senderEmail = message.getFrom();
  console.log("Sender Email for matching: " + senderEmail);

  const contactInfoParsed = _parseJson(callGemini(_buildContactInfoPrompt(body)));
  const itemsParsed = _parseJson(callGemini(_buildStructuredItemExtractionPrompt(body)));

  const senderNameFromEmail = _extractNameFromEmail(senderEmail);
  const orderNum = Date.now().toString(); // Original V12 order number format

  const client = _matchClient(senderEmail);
  console.log("Matched Client: " + client);

  const data = {};
  // Initialize with sender's name and email
  data['Customer Name'] = senderNameFromEmail;
  data['Customer Address Email'] = _extractActualEmail(senderEmail);
  data['Client'] = client;
  data['orderNum'] = orderNum;
  data['messageId'] = msgId;
  data['threadId'] = message.getThread().getId();

  // Merge Gemini's extracted contact info. This can overwrite 'Customer Name'
  // and add 'Contact Person', 'Contact Phone', 'Contact Email' if found by Gemini.
  Object.assign(data, contactInfoParsed);

  data['Items Ordered'] = itemsParsed['Items Ordered'] || [];

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

  // Field for 'Contact Person' - defaults to 'Customer Name' if 'Contact Person' isn't specifically extracted
  section.addWidget(CardService.newTextInput().setFieldName('Contact Person').setTitle('Contact Person').setValue(data['Contact Person'] || data['Customer Name'] || ''));
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address Line 1').setTitle('Address Line 1').setValue(data['Customer Address Line 1'] || ''));
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address Line 2').setTitle('Address Line 2').setValue(data['Customer Address Line 2'] || ''));
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address City').setTitle('City').setValue(data['Customer Address City'] || ''));
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address State').setTitle('State').setValue(data['Customer Address State'] || ''));
  section.addWidget(CardService.newTextInput().setFieldName('Customer Address ZIP').setTitle('ZIP').setValue(data['Customer Address ZIP'] || ''));
  // Field for 'Contact Phone' - defaults to 'Customer Address Phone' if 'Contact Phone' isn't specifically extracted
  section.addWidget(CardService.newTextInput().setFieldName('Contact Phone').setTitle('Contact Phone').setValue(_formatPhone(data['Contact Phone'] || data['Customer Address Phone'] || '')));
  
  const deliveryDatePicker = CardService.newDatePicker().setFieldName("Delivery Date").setTitle("Delivery Date").setValueInMsSinceEpoch(_parseDateToMsEpoch(data['Delivery Date']));
  section.addWidget(deliveryDatePicker);

  const deliveryTimeInput = CardService.newSelectionInput().setFieldName('Delivery Time').setTitle('Delivery Time');
  _setSelectionInputTypeSafely(deliveryTimeInput, CardService.SelectionInputType.DROPDOWN, "Delivery Time Dropdown");
  const selectedTime = data['Delivery Time'] || '';
  const startHour = 10; const endHour = 18;
  for (let h = startHour; h < endHour; h++) {
    for (let m = 0; m < 60; m += 15) {
      const time = Utilities.formatDate(new Date(2000, 0, 1, h, m), Session.getScriptTimeZone(), 'h:mm a');
      deliveryTimeInput.addItem(time, time, selectedTime === time);
    }
  }
  section.addWidget(deliveryTimeInput);

  const utensilsValue = data['Include Utensils?'] || 'Unknown';
  const utensilsInput = CardService.newSelectionInput().setFieldName('Include Utensils?').setTitle('Include Utensils?');
  _setSelectionInputTypeSafely(utensilsInput, CardService.SelectionInputType.DROPDOWN, "Utensils Dropdown");
  utensilsInput.addItem('Yes', 'Yes', utensilsValue === 'Yes').addItem('No', 'No', utensilsValue === 'No').addItem('Unknown', 'Unknown', utensilsValue !== 'Yes' && utensilsValue !== 'No');
  section.addWidget(utensilsInput);
  if (utensilsValue === 'Yes') {
    section.addWidget(CardService.newTextInput().setFieldName('If yes: how many?').setTitle('How many utensils?').setValue(data['If yes: how many?'] || ''));
  }

  section.addWidget(CardService.newTextInput().setFieldName('Client').setTitle('Client (based on email domain)').setValue(data['Client'] || 'Unknown'));
  // Field for 'Contact Email' - defaults to 'Customer Address Email' (sender's email) if 'Contact Email' isn't specifically extracted
  section.addWidget(CardService.newTextInput().setFieldName('Contact Email').setTitle('Contact Email').setValue(data['Contact Email'] || data['Customer Address Email'] || ''));

  const action = CardService.newAction().setFunctionName('handleContactInfoSubmit').setParameters({ orderNum: orderNum });
  const footer = CardService.newFixedFooter().setPrimaryButton(CardService.newTextButton().setText('Confirm Contact & Proceed to Items').setOnClickAction(action).setTextButtonStyle(CardService.TextButtonStyle.FILLED));
  return CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle('Step 1: Customer & Order Details')).addSection(section).setFixedFooter(footer).build();
}

// === SUBMIT CONTACT INFO & PROCEED TO ITEM MAPPING ===
function handleContactInfoSubmit(e) {
  const inputs = e.commonEventObject.formInputs;
  const orderNum = e.parameters.orderNum;
  const newData = {};
  for (const key in inputs) {
    if (key === "Delivery Date") { newData[key] = inputs[key].dateInput.msSinceEpoch; }
    else { newData[key] = inputs[key].stringInputs?.value?.[0] || ''; }
  }

  const userProps = PropertiesService.getUserProperties();
  const existingRaw = userProps.getProperty(orderNum);
  if (!existingRaw) {
    console.error("Error in handleContactInfoSubmit: Original order data for " + orderNum + " not found.");
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Original order data missing.")).build();
  }
  const existing = JSON.parse(existingRaw);
  const merged = { ...existing, ...newData };

  if (merged['Delivery Date'] && typeof merged['Delivery Date'] === 'number') {
      merged['Delivery Date'] = Utilities.formatDate(new Date(merged['Delivery Date']), Session.getScriptTimeZone(), "MM/dd/yyyy");
  }

  if (!merged['Items Ordered'] || !Array.isArray(merged['Items Ordered']) || (merged['Items Ordered'].length > 0 && typeof merged['Items Ordered'][0].description === 'undefined')) {
    console.log("Re-extracting items in handleContactInfoSubmit for order: " + orderNum);
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
      _setSelectionInputTypeSafely(qbItemDropdown, CardService.SelectionInputType.DROPDOWN, `Item Dropdown ${index}`);
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
      // Removed hidden item_price input
      itemsDisplaySection.addWidget(CardService.newSelectionInput().setFieldName(`item_remove_${index}`).setTitle("Remove this suggested item?").setType(CardService.SelectionInputType.CHECKBOX).addItem("Yes, remove", "true", false));
      itemsDisplaySection.addWidget(CardService.newDivider());
    });
  }
  card.addSection(itemsDisplaySection);

  const manualAddSection = CardService.newCardSection().setHeader("Manually Add New Item").setCollapsible(true);
  const newItemDropdown = CardService.newSelectionInput().setFieldName("new_item_qb_sku").setTitle("Select Item");
   _setSelectionInputTypeSafely(newItemDropdown, CardService.SelectionInputType.DROPDOWN, "New Item Manual Add Dropdown");
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

  const action = CardService.newAction().setFunctionName('handleItemMappingSubmit')
    .setParameters({ orderNum: orderNum, ai_item_count: suggestedMatches.length.toString() });
  card.setFixedFooter(CardService.newFixedFooter().setPrimaryButton(CardService.newTextButton().setText('Confirm All Items & Proceed').setOnClickAction(action).setTextButtonStyle(CardService.TextButtonStyle.FILLED)));
  return card.build();
}

function handleItemMappingSubmit(e) {
  const formInputs = e.formInputs || (e.commonEventObject && e.commonEventObject.formInputs);
  if (!formInputs) { console.error("Error in handleItemMappingSubmit: formInputs is undefined."); return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Could not read form data.")).build(); }
  const orderNum = e.parameters.orderNum; const aiItemCount = parseInt(e.parameters.ai_item_count) || 0;
  if (!orderNum) { console.error("Error in handleItemMappingSubmit: Order number is missing."); return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Order number missing.")).build(); }
  const userProps = PropertiesService.getUserProperties(); const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { console.error("Error in handleItemMappingSubmit: Original order data for " + orderNum + " not found."); return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText("Error: Original order data not found.")).build(); }
  const orderData = JSON.parse(orderDataString); const masterQBItems = getMasterQBItems(); const confirmedQuickBooksItems = [];
  for (let i = 0; i < aiItemCount; i++) {
    const removeThisItem = formInputs[`item_remove_${i}`] && formInputs[`item_remove_${i}`][0] === "true"; if (removeThisItem) { console.log(`Item ${i} marked for removal.`); continue; }
    const qtyString = formInputs[`item_qty_${i}`] && formInputs[`item_qty_${i}`][0]; const qbItemSKU = formInputs[`item_qb_sku_${i}`] && formInputs[`item_qb_sku_${i}`][0]; 
    const kitchenNotes = formInputs[`item_kitchen_notes_${i}`] && formInputs[`item_kitchen_notes_${i}`][0] || "";
    const originalEmailItem = orderData['Items Ordered'] && orderData['Items Ordered'][i]; const originalDescription = originalEmailItem ? originalEmailItem.description : "N/A";
    if (qbItemSKU && qtyString) {
      const masterItemDetails = masterQBItems.find(master => master.SKU === qbItemSKU); let unitPrice = 0; let itemName = "Custom Item"; let itemSKU = qbItemSKU;
      if (masterItemDetails) { unitPrice = masterItemDetails.Price || 0; itemName = masterItemDetails.Name; }
      else if (qbItemSKU === FALLBACK_CUSTOM_ITEM_SKU) { unitPrice = 0; } 
      else { console.warn("Master details not found for SKU: " + qbItemSKU); itemSKU = FALLBACK_CUSTOM_ITEM_SKU; unitPrice = 0; }
      confirmedQuickBooksItems.push({ quickbooks_item_id: itemSKU, quickbooks_item_name: itemName, sku: itemSKU, quantity: parseInt(qtyString) || 1, unit_price: unitPrice, kitchen_notes_and_flavors: kitchenNotes, original_email_description: originalDescription });
    }
  }
  const newItemQbSKU = formInputs.new_item_qb_sku && formInputs.new_item_qb_sku[0]; const newItemQtyString = formInputs.new_item_qty && formInputs.new_item_qty[0];
  if (newItemQbSKU && newItemQbSKU !== "" && newItemQtyString) {
    const masterItemDetails = masterQBItems.find(master => master.SKU === newItemQbSKU);
    if (masterItemDetails) {
      const unitPrice = masterItemDetails.Price || 0; const newItemKitchenNotes = formInputs.new_item_kitchen_notes && formInputs.new_item_kitchen_notes[0] || "";
      confirmedQuickBooksItems.push({ quickbooks_item_id: newItemQbSKU, quickbooks_item_name: masterItemDetails.Name, sku: masterItemDetails.SKU, quantity: parseInt(newItemQtyString) || 1, unit_price: unitPrice, kitchen_notes_and_flavors: newItemKitchenNotes, original_email_description: newItemKitchenNotes || "Manually Added: " + masterItemDetails.Name });
    }
  }
  orderData['ConfirmedQBItems'] = confirmedQuickBooksItems; userProps.setProperty(orderNum, JSON.stringify(orderData)); console.log("Confirmed items for order " + orderNum + ": " + JSON.stringify(confirmedQuickBooksItems));
  const invoiceActionsCard = buildInvoiceActionsCard({ parameters: { orderNum: orderNum } });
  return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(invoiceActionsCard)).build();
}

function buildInvoiceActionsCard(e) {
  const orderNum = e.parameters.orderNum;
  const card = CardService.newCardBuilder().setHeader(CardService.newCardHeader().setTitle(`Step 3: Final Review & Actions for ${orderNum}`));
  const orderDetailsSection = CardService.newCardSection().setHeader("Order Details");
  const userProps = PropertiesService.getUserProperties(); const orderDataString = userProps.getProperty(orderNum);
  if (orderDataString) {
    const orderData = JSON.parse(orderDataString);
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Customer:</b> " + (orderData['Contact Person'] || orderData['Customer Name'] || 'N/A') + "</i>"));
    let address = "<i><b>Address:</b> "; if (orderData['Customer Address Line 1']) address += orderData['Customer Address Line 1']; if (orderData['Customer Address Line 2']) address += (address === "<i><b>Address:</b> " ? "" : ", ") + orderData['Customer Address Line 2']; if (orderData['Customer Address City']) address += (address === "<i><b>Address:</b> " ? "" : ", ") + orderData['Customer Address City']; if (orderData['Customer Address State']) address += (address === "<i><b>Address:</b> " ? "" : ", ") + orderData['Customer Address State']; if (orderData['Customer Address ZIP']) address += " " + orderData['Customer Address ZIP']; address += "</i>";
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText(address.length <= "<i><b>Address:</b> </i>".length + 1 ? "<i><b>Address:</b> N/A</i>" : address));
    let deliveryDateFormatted = orderData['Delivery Date'];
    if (deliveryDateFormatted && typeof deliveryDateFormatted === 'string' && !deliveryDateFormatted.includes('/')) { deliveryDateFormatted = Utilities.formatDate(new Date(parseInt(deliveryDateFormatted)), Session.getScriptTimeZone(), "MM/dd/yyyy"); }
    else if (deliveryDateFormatted && typeof deliveryDateFormatted === 'string' && deliveryDateFormatted.match(/^\d{4}-\d{2}-\d{2}/)) { deliveryDateFormatted = Utilities.formatDate(new Date(deliveryDateFormatted.replace(/-/g, '/')), Session.getScriptTimeZone(), "MM/dd/yyyy"); }
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Delivery:</b> " + (deliveryDateFormatted || 'N/A') + " at " + (orderData['Delivery Time'] || 'N/A') + "</i>"));
    orderDetailsSection.addWidget(CardService.newTextParagraph().setText("<i><b>Client:</b> " + (orderData['Client'] || 'N/A') + "</i>"));
  } else { orderDetailsSection.addWidget(CardService.newTextParagraph().setText("Could not retrieve order details.")); }
  card.addSection(orderDetailsSection);
  const itemSummarySection = CardService.newCardSection().setHeader("Confirmed Items Summary"); let grandTotal = 0;
  if (orderDataString) {
    const orderData = JSON.parse(orderDataString); const confirmedItems = orderData['ConfirmedQBItems'];
    if (confirmedItems && Array.isArray(confirmedItems) && confirmedItems.length > 0) {
      confirmedItems.forEach(item => { const itemTotal = (item.quantity || 0) * (item.unit_price || 0); grandTotal += itemTotal; itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>${item.original_email_description || item.quickbooks_item_name}</b><br>Qty: ${item.quantity}, Unit Price: $${(item.unit_price || 0).toFixed(2)}, Total: $${itemTotal.toFixed(2)}${item.kitchen_notes_and_flavors ? '<br><font color="#666666"><i>Kitchen Notes: ' + item.kitchen_notes_and_flavors + '</i></font>' : ''}`)); });
      itemSummarySection.addWidget(CardService.newDivider()); itemSummarySection.addWidget(CardService.newTextParagraph().setText(`<b>Estimated Grand Total: $${grandTotal.toFixed(2)}</b>`));
    } else { itemSummarySection.addWidget(CardService.newTextParagraph().setText("No items confirmed for this order.")); }
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

function handleGenerateInvoiceAndEmail(e) {
  const orderNum = e.parameters.orderNum;
  let invoiceSheetInfo = null; let kitchenSheetInfo = null;
  try {
    invoiceSheetInfo = populateInvoiceSheet(orderNum);
    if (!invoiceSheetInfo || !invoiceSheetInfo.id || !invoiceSheetInfo.url || !invoiceSheetInfo.name) { throw new Error("Failed to populate invoice sheet or retrieve its details."); }
    console.log("Invoice sheet populated: " + invoiceSheetInfo.name + " (ID: " + invoiceSheetInfo.id + ")");
    kitchenSheetInfo = populateKitchenSheet(orderNum);
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

function handleClearAndClose(e) {
    const orderNum = e.parameters.orderNum;
    if (orderNum) { PropertiesService.getUserProperties().deleteProperty(orderNum); console.log("Cleared data for orderNum: " + orderNum); }
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().popToRoot()).setNotification(CardService.newNotification().setText("Order data cleared. Add-on is ready for the next email.")).build();
}

function populateKitchenSheet(orderNum) {
  const userProps = PropertiesService.getUserProperties(); const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { console.error("KitchenSheet: Order data for " + orderNum + " not found."); throw new Error("Order data not found for kitchen sheet: " + orderNum); }
  const orderData = JSON.parse(orderDataString); const confirmedItems = orderData['ConfirmedQBItems']; const masterAllItems = getMasterQBItems();
  if (!confirmedItems || !Array.isArray(confirmedItems)) { console.error("KitchenSheet: Confirmed items not found for order " + orderNum); throw new Error("Confirmed items not found for kitchen sheet generation."); }
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID); const templateSheet = spreadsheet.getSheetByName(KITCHEN_SHEET_TEMPLATE_NAME);
    if (!templateSheet) { console.error("Kitchen sheet template '" + KITCHEN_SHEET_TEMPLATE_NAME + "' not found."); throw new Error("Kitchen sheet template not found."); }
    const newSheetName = `Kitchen - ${orderNum} - ${orderData['Contact Person'] || orderData['Customer Name'] || 'Unknown'}`.substring(0,100); 
    const newSheet = templateSheet.copyTo(spreadsheet).setName(newSheetName);
    const customerNameForKitchen = orderData['Contact Person'] || orderData['Customer Name'] || '';
    const contactPhoneForKitchen = orderData['Contact Phone'] || orderData['Customer Address Phone'] || 'N/A';
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
    SpreadsheetApp.flush(); return { id: spreadsheet.getId(), url: newSheet.getParent().getUrl() + '#gid=' + newSheet.getSheetId(), name: newSheetName };
  } catch (e) { console.error("Error in populateKitchenSheet for order " + orderNum + ": " + e.toString() + (e.stack ? ("\nStack: " + e.stack) : "")); throw new Error("Failed to populate kitchen sheet: " + e.message); }
}

function populateInvoiceSheet(orderNum) {
  const userProps = PropertiesService.getUserProperties(); const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { console.error("InvoiceSheet: Order data for " + orderNum + " not found."); throw new Error("Order data not found for " + orderNum); }
  const orderData = JSON.parse(orderDataString); const confirmedItems = orderData['ConfirmedQBItems'];
  if (!confirmedItems || !Array.isArray(confirmedItems)) { console.error("InvoiceSheet: Confirmed items not found for order " + orderNum); throw new Error("Confirmed items not found for invoice generation."); }
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID); const templateSheet = spreadsheet.getSheetByName(INVOICE_TEMPLATE_SHEET_NAME);
    if (!templateSheet) { console.error("Invoice template sheet '" + INVOICE_TEMPLATE_SHEET_NAME + "' not found."); throw new Error("Invoice template sheet not found."); }
    const newSheetName = `Invoice - ${orderNum} - ${orderData['Contact Person'] || orderData['Customer Name'] || 'Unknown'}`.substring(0,100); 
    const newSheet = templateSheet.copyTo(spreadsheet).setName(newSheetName);
    newSheet.getRange(ORDER_NUM_CELL).setValue(orderData.orderNum);
    newSheet.getRange(CUSTOMER_NAME_CELL).setValue(orderData['Contact Person'] || orderData['Customer Name'] || ''); 
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
    let currentRow = ITEM_START_ROW_INVOICE; let grandTotal = 0;
    confirmedItems.forEach(item => {
      const descriptionCell = newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow);
      descriptionCell.setValue(item.original_email_description || item.quickbooks_item_name).setWrap(false); 
      newSheet.getRange(ITEM_QTY_COL_INVOICE + currentRow).setValue(item.quantity);
      newSheet.getRange(ITEM_UNIT_PRICE_COL_INVOICE + currentRow).setValue(item.unit_price).setNumberFormat("$#,##0.00"); 
      const lineTotal = (item.quantity || 0) * (item.unit_price || 0);
      newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setValue(lineTotal).setNumberFormat("$#,##0.00"); 
      grandTotal += lineTotal; currentRow++;
    });
    const grandTotalDescCell = newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow);
    grandTotalDescCell.setValue("Grand Total:").setFontWeight("bold").setWrap(false);
    const grandTotalValueCell = newSheet.getRange(ITEM_TOTAL_PRICE_COL_INVOICE + currentRow);
    grandTotalValueCell.setValue(grandTotal).setNumberFormat("$#,##0.00").setFontWeight("bold").setHorizontalAlignment("right").setWrap(false);
    newSheet.getRange(ITEM_DESCRIPTION_COL_INVOICE + currentRow + ":" + ITEM_TOTAL_PRICE_COL_INVOICE + currentRow).setBorder(true, null, null, null, null, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    SpreadsheetApp.flush(); return { id: spreadsheet.getId(), url: newSheet.getParent().getUrl() + '#gid=' + newSheet.getSheetId(), name: newSheetName };
  } catch (e) { console.error("Error in populateInvoiceSheet for order " + orderNum + ": " + e.toString() + (e.stack ? ("\nStack: " + e.stack) : "")); throw new Error("Failed to populate invoice sheet: " + e.message); }
}

function createPdfAndPrepareEmailReply(orderNum, populatedSheetSpreadsheetId, populatedSheetName) {
  const userProps = PropertiesService.getUserProperties(); const orderDataString = userProps.getProperty(orderNum);
  if (!orderDataString) { console.error("PDFEmail: Order data for " + orderNum + " not found."); throw new Error("Order data not found for PDF/email creation.");}
  const orderData = JSON.parse(orderDataString);
  try {
    console.log(`PDFEmail: Opening spreadsheet ID: ${populatedSheetSpreadsheetId}, Sheet: ${populatedSheetName}`);
    const spreadsheet = SpreadsheetApp.openById(populatedSheetSpreadsheetId); const sheetToExport = spreadsheet.getSheetByName(populatedSheetName);
    if (!sheetToExport) { console.error(`PDFEmail: Sheet "${populatedSheetName}" not found.`); throw new Error("Populated invoice sheet not found for PDF generation."); }
    console.log(`PDFEmail: Sheet "${populatedSheetName}" found. Hiding others.`);
    const allSheets = spreadsheet.getSheets(); const hiddenSheetIds = [];
    allSheets.forEach(sheet => { if (sheet.getSheetId() !== sheetToExport.getSheetId()) { sheet.hideSheet(); hiddenSheetIds.push(sheet.getSheetId()); }});
    SpreadsheetApp.flush(); console.log("PDFEmail: Other sheets hidden.");
    const pdfBlob = spreadsheet.getAs('application/pdf').setName(`${populatedSheetName}.pdf`); console.log(`PDFEmail: PDF blob created: ${pdfBlob.getName()}`);
    hiddenSheetIds.forEach(id => { const sheet = allSheets.find(s => s.getSheetId() === id); if (sheet) sheet.showSheet(); });
    SpreadsheetApp.flush(); console.log("PDFEmail: Sheets unhidden.");
    const threadId = orderData.threadId;
    if (!threadId) { console.error("PDFEmail: Thread ID missing for order " + orderNum); throw new Error("Original email thread ID not found."); }
    const thread = GmailApp.getThreadById(threadId);
    if (!thread) { console.error("PDFEmail: Could not retrieve Gmail thread ID: " + threadId); throw new Error("Could not retrieve original email thread."); }
    const messages = thread.getMessages(); const messageToReplyTo = messages[messages.length - 1];
    const recipient = orderData['Contact Email'] || orderData['Customer Address Email'];
    const subject = `Re: Catering Order Confirmation - ${orderNum}`;
    const body = `Dear ${orderData['Contact Person'] || orderData['Customer Name'] || 'Valued Customer'},\n\nPlease find attached the invoice for your recent catering order (${orderNum}).\n\nDelivery is scheduled for ${orderData['Delivery Date']} around ${orderData['Delivery Time']}.\n\nThank you for your business!\n\nBest regards,\n[Your Company Name]`;
    console.log(`PDFEmail: Preparing draft reply to: ${recipient}`);
    const draft = messageToReplyTo.createDraftReply(body, { htmlBody: body.replace(/\n/g, '<br>'), attachments: [pdfBlob], to: recipient });
    console.log("PDFEmail: Draft email created. ID: " + draft.getId());
    return { pdfBlob: pdfBlob, draft: draft, draftId: draft.getId() };
  } catch (e) { console.error("Error in createPdfAndPrepareEmailReply for order " + orderNum + ": " + e.toString() + (e.stack ? ("\nStack: " + e.stack) : "")); throw new Error("Failed to create PDF or prepare email: " + e.message); }
}

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

function getGeminiItemMatches(emailItems, masterQBItems) {
  if (!emailItems || emailItems.length === 0) return [];
  if (!masterQBItems || masterQBItems.length === 0) { 
    console.warn("Master QB Items list is empty in getGeminiItemMatches.");
    return emailItems.map(item => ({
      original_email_description: item.description, extracted_main_quantity: item.quantity,
      matched_qb_item_id: FALLBACK_CUSTOM_ITEM_SKU, matched_qb_item_name: "Custom Item (No Master List)",
      match_confidence: "Low", parsed_flavors_or_details: item.description,
      identified_flavors: [] 
    }));
  }
  const masterItemDetailsForPrompt = masterQBItems.map(item => {
    let detailString = `- Name: "${item.Name}" (SKU: ${item.SKU}, Price: $${item.Price !== undefined ? item.Price.toFixed(2) : '0.00'})`; 
    if (item.Category) detailString += ` [Category: ${item.Category}]`; if (item.Item && item.Item !== item.Name) detailString += ` [Base Item: ${item.Item}]`;
    if (item.Subtype) detailString += ` [Type: ${item.Subtype}]`; if (item.Size) detailString += ` [Size: ${item.Size}]`; if (item.Descriptor) detailString += ` [Details: ${item.Descriptor}]`;
    const flavors = [item['Flavor 1'], item['Flavor 2'], item['Flavor 3'], item['Flavor 4'], item['Flavor 5']].filter(f => f && f.toString().trim() !== "").join('; ');
    if (flavors) detailString += ` [Std Flavors: ${flavors}]`; return detailString;
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
    1. "original_email_description": The full description text from the "Email Line".
    2. "extracted_main_quantity": The "Ordered Quantity" associated with that "Email Line" (as a string).
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
    [ { "original_email_description": "...", "extracted_main_quantity": "...", "matched_qb_item_id": "...", "matched_qb_item_name": "...", "match_confidence": "High", "parsed_flavors_or_details": "...", "identified_flavors": [] } ]`;
  console.log("Prompt for getGeminiItemMatches (first 500 chars):\n" + prompt.substring(0, 500)); 
  const geminiResponseText = callGemini(prompt); console.log("Raw response from getGeminiItemMatches: " + geminiResponseText);
  try {
    const parsedResponse = _parseJson(geminiResponseText);
    if (Array.isArray(parsedResponse)) { return parsedResponse.map(item => ({ ...item, identified_flavors: item.identified_flavors || [] })) ; } 
    else { console.error("Parsed Gemini response is not an array: " + JSON.stringify(parsedResponse)); throw new Error("Parsed response is not an array."); }
  } catch (e) {
    console.error("Error parsing Gemini item matching response: " + e.toString() + " Raw response for parse error: " + geminiResponseText);
    return emailItems.map(item => ({ 
      original_email_description: item.description, extracted_main_quantity: item.quantity,
      matched_qb_item_id: fallbackSkuForPrompt, matched_qb_item_name: fallbackNameForPrompt + " (AI Error)",
      match_confidence: "Low", parsed_flavors_or_details: "AI matching error occurred.", identified_flavors: []
    }));
  }
}

function _formatPhone(phone) { if (!phone) return ''; const digits = phone.replace(/\D/g, ''); if (digits.length === 10) { return `(${digits.substr(0, 3)}) ${digits.substr(3, 3)}-${digits.substr(6)}`; } return phone; }
function _buildContactInfoPrompt(body) { return 'Extract the following fields from this email and return ONLY a JSON object with no extra commentary:\nCustomer Name\nCustomer Address Line 1\nCustomer Address Line 2\nCustomer Address City\nCustomer Address State\nCustomer Address ZIP\nCustomer Address Phone\nDelivery Date\nDelivery Time\nInclude Utensils?\nIf yes: how many?\nContact Person\nContact Phone\nContact Email\n\nEmail:\n' + body; }
function _buildStructuredItemExtractionPrompt(body) { return 'From the email body provided, extract ONLY the ordered items. It is CRITICAL to treat each line in the order section of the email that appears to request a product as a SEPARATE item in the output array. For example, if the email says "1 Small Cheesy Rice" and on a new line "1 Small Cheesy Rice (Vegan)", these must be two distinct entries in the JSON. For each distinct item line, provide its "quantity" as a string and a "description" string that is the clean, full text from that line, including all modifiers, flavors, and sub-quantities mentioned for that specific line. Return as JSON in the format:\n{ "Items Ordered": [ { "quantity": "1", "description": "Large Hilacha Chicken" }, { "quantity": "1", "description": "Small Cheesy Rice" }, { "quantity": "1", "description": "Small Cheesy Rice (Vegan)" }, {"quantity": "1", "description": "Large Taquitos Tray (12 Chile Chicken, 20 Chicken and Cheese, 8 Jackfruit)"} ] }\n\nOnly include food and tray items. Do not include headers, greetings, closings, or other conversational text from the email. Ensure quantities are extracted as strings.\n\nEmail:\n' + body; }
function callGemini(prompt) { if (!API_KEY) { console.error("Error: GL_API_KEY is not set."); throw new Error("No API key set."); } const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + encodeURIComponent(API_KEY); const payload = { contents: [{ role: "user", parts: [{ text: prompt }] }] }; const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true }; console.log("Calling Gemini. Prompt length: " + prompt.length); const response = UrlFetchApp.fetch(url, options); const responseCode = response.getResponseCode(); const responseBody = response.getContentText(); if (responseCode === 200) { const json = JSON.parse(responseBody); if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts.length > 0) { return json.candidates[0].content.parts[0].text; } else { console.error("Gemini response missing expected structure: " + responseBody); if (json.candidates && json.candidates[0] && json.candidates[0].finishReason) { console.error("Gemini finishReason: " + json.candidates[0].finishReason); if (json.candidates[0].safetyRatings) { console.error("SafetyRatings: " + JSON.stringify(json.candidates[0].safetyRatings)); } if (json.candidates[0].finishReason === "SAFETY" || json.candidates[0].finishReason === "OTHER") { throw new Error("Gemini request blocked due to: " + json.candidates[0].finishReason + ". Check safety ratings in log."); } } throw new Error("Invalid Gemini response structure. See logs."); } } else { console.error("Gemini API Error - Code: " + responseCode + " Body: " + responseBody); throw new Error("Gemini API request failed. Code: " + responseCode + ". See logs for details."); } }
function _parseJson(raw) { if (!raw || typeof raw !== 'string') { console.error("Error in _parseJson: Input is not a valid string or is empty. Input: " + raw); return {}; } try { const cleanedJsonString = raw.replace(/^```json\s*([\s\S]*?)\s*```$/, '$1').trim(); return JSON.parse(cleanedJsonString); } catch (e) { console.error("Error in _parseJson (first attempt): " + e.toString() + ". Raw input: " + raw); const match = raw.match(/\{[\s\S]*\}|\[[\s\S]*\]/); if (match && match[0]) { try { return JSON.parse(match[0]); } catch (e2) { console.error("Error in _parseJson (fallback attempt): " + e2.toString() + ". Matched string: " + match[0]); throw new Error('Invalid JSON response from AI after attempting to clean: ' + raw); } } throw new Error('Invalid JSON response from AI, and no object/array found: ' + raw); } }
function _extractNameFromEmail(email) { if (!email) return ''; const match = email.match(/^(.*?)</); if (match && match[1]) return match[1].trim(); const namePart = email.split('@')[0]; return namePart.replace(/[._\d]+$/, '').replace(/[._]/g, ' ').trim();  }

function _matchClient(senderEmailField) {
  if (!senderEmailField) return 'Unknown';
  let emailAddress = senderEmailField;
  const emailMatch = senderEmailField.match(/<([^>]+)>/);
  if (emailMatch && emailMatch[1]) {
    emailAddress = emailMatch[1];
  }
  const emailLower = emailAddress.toLowerCase().trim();
  console.log("Matching client for extracted email: " + emailLower);
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

function _extractActualEmail(senderEmailField) {
    if (!senderEmailField) return '';
    const emailMatch = senderEmailField.match(/<([^>]+)>/);
    if (emailMatch && emailMatch[1]) {
        return emailMatch[1].trim();
    }
    return senderEmailField.trim();
}

function _parseDateToMsEpoch(dateString) {
  if (!dateString || typeof dateString !== 'string') { return new Date().getTime(); }
  let date; const currentYear = new Date().getFullYear();
  if (dateString.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) { date = new Date(dateString); }
  else if (dateString.match(/^\d{4}-\d{2}-\d{2}$/)) { date = new Date(dateString.replace(/-/g, '/')); }
  else if (dateString.match(/^\d{1,2}\/\d{1,2}$/)) { const parts = dateString.split('/'); date = new Date(currentYear, parseInt(parts[0]) - 1, parseInt(parts[1])); }
  else { date = new Date(dateString); if (date && date.getFullYear() < 2000) { date.setFullYear(currentYear); } }
  if (isNaN(date.getTime())) { return new Date().getTime(); }
  return date.getTime();
}

// Helper function to safely set SelectionInput type
function _setSelectionInputTypeSafely(selectionInput, typeEnum, widgetNameForLog) {
    try {
        if (typeof CardService === 'undefined' || CardService === null || 
            typeof CardService.SelectionInputType === 'undefined' || CardService.SelectionInputType === null) {
            console.error(`CardService or CardService.SelectionInputType is undefined. Cannot set type for ${widgetNameForLog}.`);
            return; // Critical CardService component missing
        }

        let typeValue;
        // Check if 'typeEnum' is the actual enum value or its string key
        if (typeof typeEnum === 'string' && CardService.SelectionInputType[typeEnum] !== undefined) {
            typeValue = CardService.SelectionInputType[typeEnum];
        } else if (Object.values(CardService.SelectionInputType).includes(typeEnum)) {
            typeValue = typeEnum;
        } else {
            console.error(`SelectionInputType "${typeEnum}" is not a valid key or value in CardService.SelectionInputType for ${widgetNameForLog}.`);
            return;
        }
        
        if (selectionInput && typeof selectionInput.setType === 'function') {
            selectionInput.setType(typeValue);
        } else {
            console.error(`Cannot call setType on ${widgetNameForLog}. Object is not a valid SelectionInput or setType method is missing. Object type: ${typeof selectionInput}`);
        }
    } catch (err) {
        console.error(`Error setting type for ${widgetNameForLog}: ${err.toString()}`);
    }
}