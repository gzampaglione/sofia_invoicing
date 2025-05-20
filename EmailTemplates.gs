// EmailTemplates.gs

// This object will store our email templates.
// Each key can be an internal template ID/name, and the value will be an object
// with 'name' (for display in dropdown), 'subject', and 'body'.
// Placeholders like ${customerName}, ${orderNum}, ${deliveryDate}, ${deliveryTime},
// ${grandTotal}, ${invoiceLink} (optional if we provide a link in email too)
// will be replaced by actual data.

const EMAIL_TEMPLATES = {
  'defaultConfirmation': {
    name: 'Standard Order Confirmation',
    subject: 'Your El Merkury Catering Order #${orderNum} is Confirmed!',
    body:
`Dear ${customerName},

Thank you for your recent catering order (#${orderNum}) with El Merkury! We're excited to prepare it for you.

Your invoice is attached to this email.

Delivery Details:
Date: ${deliveryDate}
Time: ${deliveryTime}
Address: 
${deliveryAddressLine1}
${deliveryAddressLine2OrEmpty}
${deliveryCityStateZip}

Order Grand Total: $${grandTotal}

Please review the attached invoice for full details. If you have any questions or need to make changes, please reply to this email or call us.

We look forward to serving you!

Best regards,
Sofia & The El Merkury Team
catering@elmerkury.com`
  },
  'pennConfirmationPO': {
    name: 'Penn Order Confirmation (PO Follow-up)',
    subject: 'El Merkury Catering Order #${orderNum} Received - PO Invoice to Follow',
    body:
`Dear ${customerName},

Thank you for placing your catering order (#${orderNum}) with El Merkury for your event at ${clientName}.

We have attached a preliminary invoice for your records. 
As per University of Pennsylvania procedures, we will submit this invoice through the Penn purchasing system to obtain a Purchase Order (PO) number. Once the PO is approved and assigned, we will send you an updated, official invoice reflecting the PO number.

Your current order details:
Delivery Date: ${deliveryDate}
Time: ${deliveryTime}
Address: 
${deliveryAddressLine1}
${deliveryAddressLine2OrEmpty}
${deliveryCityStateZip}

Estimated Grand Total: $${grandTotal}

Please let us know if there are any immediate changes to your order.

Best regards,
Sofia & The El Merkury Team
catering@elmerkury.com`
  },
  'requestMoreInfo': {
    name: 'Confirmation & Request for More Info',
    subject: 'El Merkury Catering Order #${orderNum} - Action Required: More Information Needed',
    body:
`Dear ${customerName},

Thank you for your catering order (#${orderNum}) with El Merkury! We've received your request.

Your invoice is attached for your review.

Delivery Details:
Date: ${deliveryDate}
Time: ${deliveryTime}
Address: 
${deliveryAddressLine1}
${deliveryAddressLine2OrEmpty}
${deliveryCityStateZip}

Estimated Grand Total: $${grandTotal}

Before we can fully confirm and finalize your order, we need a little more information regarding:
[PLEASE SPECIFY WHAT INFO IS NEEDED HERE - e.g., final guest count, specific dietary restrictions not clear, PO number if applicable, preferred payment method]

Please reply to this email with the details at your earliest convenience.

Best regards,
Sofia & The El Merkury Team
catering@elmerkury.com`
  }
  // Add more templates here as needed
};

/**
 * Retrieves all available email templates for display.
 * @return {Array<object>} An array of objects, each with 'id' and 'name'.
 */
function getEmailTemplateList() {
  const templateList = [];
  for (const templateId in EMAIL_TEMPLATES) {
    templateList.push({
      id: templateId,
      name: EMAIL_TEMPLATES[templateId].name
    });
  }
  return templateList;
}

/**
 * Retrieves a specific email template by its ID.
 * @param {string} templateId The ID of the template.
 * @return {object|null} The template object (name, subject, body) or null if not found.
 */
function getEmailTemplateById(templateId) {
  return EMAIL_TEMPLATES[templateId] || null;
}

/**
 * Populates an email template with order data.
 * @param {object} template The email template object {subject, body}.
 * @param {object} orderData The order data containing placeholder values.
 * @param {string} pdfViewLink (Optional) A direct link to view the PDF.
 * @return {{subject: string, body: string}} The populated subject and body.
 */
function populateEmailTemplate(template, orderData, pdfViewLink) {
  let subject = template.subject;
  let body = template.body;

  // Define a helper for safe replacement (handles null/undefined)
  const F = (val) => val || ''; 

  // Common placeholders
  subject = subject.replace(/\$\{orderNum\}/g, F(orderData.orderNum));
  body = body.replace(/\$\{orderNum\}/g, F(orderData.orderNum));
  body = body.replace(/\$\{customerName\}/g, F(orderData['Contact Person'] || orderData['Customer Name']));
  
  let deliveryDateForEmail = "Not specified";
  if (orderData['Delivery Date']) { // Expects YYYY-MM-DD
      try {
          const dateStr = orderData['Delivery Date'];
          let tempDate;
          if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
              const parts = dateStr.split('-');
              tempDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
          } else if (dateStr.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
              const parts = dateStr.split('/');
              tempDate = new Date(parseInt(parts[2]), parseInt(parts[0]) - 1, parseInt(parts[1]));
          } else {
              tempDate = new Date(_parseDateToMsEpoch(dateStr)); // Fallback to full parse
          }
          if (!isNaN(tempDate.getTime())) {
               deliveryDateForEmail = Utilities.formatDate(tempDate, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy");
          } else { deliveryDateForEmail = orderData['Delivery Date']; }
      } catch (e) { deliveryDateForEmail = orderData['Delivery Date']; }
  }
  body = body.replace(/\$\{deliveryDate\}/g, deliveryDateForEmail);
  
  const deliveryTimeForEmail = orderData['Delivery Time'] ? _normalizeTimeFormat(orderData['Delivery Time']) : "Not specified";
  body = body.replace(/\$\{deliveryTime\}/g, deliveryTimeForEmail);

  body = body.replace(/\$\{deliveryAddressLine1\}/g, F(orderData['Customer Address Line 1']));
  body = body.replace(/\$\{deliveryAddressLine2OrEmpty\}/g, F(orderData['Customer Address Line 2'])); // Handles if L2 is empty

  const city = F(orderData['Customer Address City']);
  const state = F(orderData['Customer Address State']);
  const zip = F(orderData['Customer Address ZIP']);
  let cityStateZip = `${city}${city && (state || zip) ? ', ' : ''}${state} ${zip}`.trim();
  if (cityStateZip === ',') cityStateZip = ''; // Clean up if only comma
  body = body.replace(/\$\{deliveryCityStateZip\}/g, cityStateZip);
  
  body = body.replace(/\$\{clientName\}/g, F(orderData['Client']));
  
  // Calculate Grand Total again for the email (as it was in buildInvoiceActionsCard)
  // This ensures accuracy based on final confirmed items and charges.
  let grandTotal = 0;
  let subTotalForCharges = 0;
  const confirmedItems = orderData['ConfirmedQBItems'] || [];
  confirmedItems.forEach(item => {
      subTotalForCharges += (item.quantity || 0) * (item.unit_price || 0);
  });
  grandTotal = subTotalForCharges;
  if (orderData['TipAmount'] > 0) grandTotal += orderData['TipAmount'];
  if (orderData['OtherChargesAmount'] > 0) grandTotal += orderData['OtherChargesAmount'];
  let deliveryFee = BASE_DELIVERY_FEE; // Assumes BASE_DELIVERY_FEE is accessible or pass it
  if (orderData['master_delivery_time_ms']) {
      const deliveryHour = new Date(orderData['master_delivery_time_ms']).getHours();
      if (deliveryHour >= DELIVERY_FEE_CUTOFF_HOUR) deliveryFee = AFTER_4PM_DELIVERY_FEE;
  }
  if (deliveryFee > 0) grandTotal += deliveryFee;
  if (orderData['Include Utensils?'] === 'Yes') {
      const numUtensils = parseInt(orderData['If yes: how many?']) || 0;
      if (numUtensils > 0) grandTotal += numUtensils * COST_PER_UTENSIL_SET;
  }
  body = body.replace(/\$\{grandTotal\}/g, grandTotal.toFixed(2));

  if (pdfViewLink) {
    body = body.replace(/\$\{invoiceLink\}/g, pdfViewLink); // Optional: if you want a direct link in the email too
  } else {
    body = body.replace(/\$\{invoiceLink\}/g, "(See attached PDF)");
  }
  
  // Add any other placeholders you need

  return { subject: subject, body: body };
}