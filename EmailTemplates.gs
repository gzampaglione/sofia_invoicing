// EmailTemplates.gs

/**
 * Defines the email templates.
 * Uses template literals for multi-line string definitions within this function for readability.
 * Placeholders should be normal, e.g., ${placeholderName}, as they are part of strings
 * defined within this function's scope. They are not evaluated as variables until
 * populateEmailTemplate is called with actual data.
 * @return {object} The object containing all defined email templates.
 */
function getDefinedEmailTemplates() {
    // Define template bodies using backtick template literals for easy multi-line editing.
    // The ${...} here will be treated as part of the string when this function executes
    // and returns the strings. They are NOT evaluated as undefined global variables at script load time.
    const defaultConfirmationBody = `Dear \${customerName},

Thank you for your recent catering order (#\${orderNum}) with El Merkury! We're excited to prepare it for you.

Your invoice is attached to this email for your review.

Delivery Details:
Date: \${deliveryDateFormatted}
Time: \${deliveryTimeFormatted}
Address:
\${deliveryAddressLine1}
\${deliveryAddressLine2OrEmpty}\${deliveryCityStateZip}

Order Grand Total: $\${grandTotalFormatted}

Please review the attached invoice for full details. If you have any questions or need to make changes, please reply to this email or call us.

We look forward to serving you!

Best regards,
Sofia & The El Merkury Team
catering@elmerkury.com`;

    const pennConfirmationPOBody = `Dear \${customerName},

Thank you for placing your catering order (#\${orderNum}) with El Merkury for your event at \${clientName}.

We have attached a preliminary invoice for your records and for submission to your department for PO generation.
As per University of Pennsylvania procedures, please process this to obtain a Purchase Order (PO) number. 
Once you provide us with the PO number, we will issue an updated, official invoice reflecting it.

Your current order details:
Delivery Date: \${deliveryDateFormatted}
Time: \${deliveryTimeFormatted}
Address:
\${deliveryAddressLine1}
\${deliveryAddressLine2OrEmpty}\${deliveryCityStateZip}

Estimated Grand Total: $\${grandTotalFormatted}

Please let us know if there are any immediate changes to your order or once you have the PO number.

Best regards,
Sofia & The El Merkury Team
catering@elmerkury.com`;

    const requestMoreInfoBody = `Dear \${customerName},

Thank you for your catering order (#\${orderNum}) with El Merkury! We've received your request, and your invoice is attached.

Delivery Details:
Date: \${deliveryDateFormatted}
Time: \${deliveryTimeFormatted}
Address:
\${deliveryAddressLine1}
\${deliveryAddressLine2OrEmpty}\${deliveryCityStateZip}

Estimated Grand Total: $\${grandTotalFormatted}

Before we can fully finalize your order, could you please provide us with the following information:
[INSERT SPECIFIC QUESTIONS HERE - e.g., final guest count, specific dietary needs, clarification on an item, preferred on-site contact if different, PO number if applicable]

Please reply to this email with the details at your earliest convenience.

Best regards,
Sofia & The El Merkury Team
catering@elmerkury.com`;

    const simpleThankYouBody = `Dear \${customerName},

Thank you for choosing El Merkury!

Please find your invoice attached for order #\${orderNum}, scheduled for \${deliveryDateFormatted} at \${deliveryTimeFormatted}.

Let us know if you have any questions.

Best,
The El Merkury Team`;

    // Return the structured template object
    return {
        'defaultConfirmation': {
            name: 'Standard Order Confirmation',
            subject: 'Your El Merkury Catering Order #${orderNum} is Confirmed!', // Use ${...} directly
            body: defaultConfirmationBody.trim() 
        },
        'pennConfirmationPO': {
            name: 'Penn Order Confirmation (PO Follow-up)',
            subject: 'El Merkury Catering Order #${orderNum} Received - Penn PO Invoice to Follow',
            body: pennConfirmationPOBody.trim()
        },
        'requestMoreInfo': {
            name: 'Confirmation & Request for More Info',
            subject: 'El Merkury Catering Order #${orderNum} - Action Required: More Information Needed',
            body: requestMoreInfoBody.trim()
        },
        'simpleThankYou': {
            name: 'Simple Thank You & Invoice Attached',
            subject: 'Invoice for your El Merkury Catering Order #${orderNum}',
            body: simpleThankYouBody.trim()
        }
        // Add more templates here following the same pattern
    };
}

// Initialize EMAIL_TEMPLATES by calling the function.
// This assigns the object returned by getDefinedEmailTemplates() to the global constant.
const EMAIL_TEMPLATES = getDefinedEmailTemplates();

/**
 * Retrieves all available email templates for display in a dropdown.
 * @return {Array<object>} An array of objects, each with 'id' (template key) and 'name' (display name).
 */
function getEmailTemplateList() {
  const templateList = [];
  for (const templateId in EMAIL_TEMPLATES) {
    // Ensure it's a direct property and not from the prototype chain
    if (Object.prototype.hasOwnProperty.call(EMAIL_TEMPLATES, templateId)) {
      templateList.push({
        id: templateId, 
        name: EMAIL_TEMPLATES[templateId].name
      });
    }
  }
  return templateList;
}

/**
 * Retrieves a specific email template definition by its ID.
 * @param {string} templateId The ID (key) of the template.
 * @return {object|null} The template object {name, subject, body} or null if not found.
 */
function getEmailTemplateById(templateId) {
  return EMAIL_TEMPLATES[templateId] || null;
}

/**
 * Populates an email template's subject and body with order data.
 * Handles placeholders in the format ${placeholderName}.
 * @param {object} template The email template object, containing 'subject' and 'body' with placeholders.
 * @param {object} orderData The order data object containing values for the placeholders.
 * @param {string} [pdfFileIdForLinkResolution] Optional. The Google Drive File ID of the PDF invoice. (Used by orderData.pdfUrl)
 * @return {{subject: string, body: string}} The populated subject and body.
 */
function populateEmailTemplate(template, orderData, pdfFileIdForLinkResolution) {
  let populatedSubject = template.subject;
  let populatedBody = template.body;

  // Helper for safe replacement of placeholders to avoid 'undefined' or 'null' literal strings
  const F = (val) => (val !== undefined && val !== null) ? val.toString() : ''; 

  // Generic placeholder replacement function
  const replacePlaceholders = (text) => {
    if (typeof text !== 'string') return '';
    // Regex to find ${placeholderName} directly
    return text.replace(/\$\{(\w+)\}/g, (match, placeholderName) => {
      switch (placeholderName) {
        case 'orderNum': return F(orderData.orderNum);
        case 'customerName': return F(orderData['Contact Person'] || orderData['Customer Name']);
        case 'clientName': return F(orderData['Client']);
        case 'deliveryDateFormatted':
          let df = "Not specified";
          if (orderData['Delivery Date']) { // Expects YYYY-MM-DD
            try {
              const ds = orderData['Delivery Date'];
              let td;
              if (ds.match(/^\d{4}-\d{2}-\d{2}$/)) { 
                const p = ds.split('-'); td = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2])); 
              } else if (ds.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) { // Fallback for MM/DD/YYYY
                const p = ds.split('/'); td = new Date(parseInt(p[2]), parseInt(p[0]) - 1, parseInt(p[1])); 
              } else { 
                td = new Date(_parseDateToMsEpoch(ds)); // Robust fallback
              }
              if (td && !isNaN(td.getTime())) {
                 df = Utilities.formatDate(td, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy"); // Corrected 'GGGG' to 'yyyy'
              } else { 
                console.warn(`populateEmailTemplate: Could not parse date "${ds}" into valid Date object for order ${F(orderData.orderNum)}.`);
                df = ds; 
              }
            } catch (e) { 
                console.warn(`populateEmailTemplate: Error formatting deliveryDate for order ${F(orderData.orderNum)}: "${orderData['Delivery Date']}". Error: ${e.toString()}`);
                df = orderData['Delivery Date'] || "Not specified"; 
            }
          }
          return df;
        case 'deliveryTimeFormatted': 
            return orderData['Delivery Time'] ? _normalizeTimeFormat(orderData['Delivery Time']) : "Not specified";
        case 'deliveryAddressLine1': return F(orderData['Customer Address Line 1']);
        case 'deliveryAddressLine2OrEmpty': 
          const adl2 = F(orderData['Customer Address Line 2']);
          return adl2 ? adl2 + '\n' : ''; // Adds newline only if content exists
        case 'deliveryCityStateZip':
          const city = F(orderData['Customer Address City']);
          const state = F(orderData['Customer Address State']);
          const zip = F(orderData['Customer Address ZIP']);
          let csz = `${city}${city && (state || zip) ? ', ' : ''}${state} ${zip}`.trim();
          return (csz === ',' || csz.trim() === '') ? '' : csz;
        case 'grandTotalFormatted':
          let gt = 0, is = 0;
          (orderData['ConfirmedQBItems'] || []).forEach(item => { is += (parseFloat(item.quantity) || 0) * (parseFloat(item.unit_price) || 0); });
          gt = is;
          gt += parseFloat(orderData['TipAmount'] || 0);
          gt += parseFloat(orderData['OtherChargesAmount'] || 0);
          
          // Access constants safely (assuming they are global or passed via orderData if not)
          const effBaseDelFee = typeof BASE_DELIVERY_FEE !== 'undefined' ? BASE_DELIVERY_FEE : 0;
          const effCutoff = typeof DELIVERY_FEE_CUTOFF_HOUR !== 'undefined' ? DELIVERY_FEE_CUTOFF_HOUR : 16;
          const effAfter4Fee = typeof AFTER_4PM_DELIVERY_FEE !== 'undefined' ? AFTER_4PM_DELIVERY_FEE : 0;
          const effCostU = typeof COST_PER_UTENSIL_SET !== 'undefined' ? COST_PER_UTENSIL_SET : 0;

          let delFee = effBaseDelFee;
          if (orderData['master_delivery_time_ms']) {
            if (new Date(orderData['master_delivery_time_ms']).getHours() >= effCutoff) delFee = effAfter4Fee;
          }
          if (delFee > 0) gt += delFee;
          if (orderData['Include Utensils?'] === 'Yes') {
            const nu = parseInt(orderData['If yes: how many?']) || 0;
            if (nu > 0) gt += nu * effCostU;
          }
          return gt.toFixed(2);
        case 'invoicePdfLink':
          let pdfLink = "(Invoice is attached)"; 
          if (orderData.pdfUrl) { 
            pdfLink = `You can also view your invoice here: ${orderData.pdfUrl}`; 
          } else if (pdfFileIdForLinkResolution) { // pdfFileIdForLinkResolution is passed to this function
            // This fallback is less ideal as it requires another DriveApp call if orderData.pdfUrl wasn't pre-populated
            console.warn("populateEmailTemplate: orderData.pdfUrl was missing; pdfFileIdForLinkResolution was provided. Ideally, set orderData.pdfUrl beforehand.");
             try {
                const file = DriveApp.getFileById(pdfFileIdForLinkResolution);
                pdfLink = `You can also view your invoice here: ${file.getUrl()}`;
            } catch (linkErr) {
                console.warn("Could not generate PDF view link from pdfFileIdForLinkResolution '" + pdfFileIdForLinkResolution + "': " + linkErr.toString());
            }
          }
          return pdfLink;
        default: 
          console.warn(`populateEmailTemplate: Unknown placeholder encountered: \${${placeholderName}} in template for order ${F(orderData.orderNum)}`);
          return match; // Return the original match (e.g., "${unknown}") if not recognized
      }
    });
  };

  populatedSubject = replacePlaceholders(populatedSubject);
  populatedBody = replacePlaceholders(populatedBody);

  return { subject: populatedSubject, body: populatedBody };
}