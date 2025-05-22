# El Merkury Invoicing - Gmail Add-on
## About
This is an app script to parse emails/email chains, extract content in standardized form, and perform a some actions based on the content of those emails. The email is always a potential customer requesting a catering order to be delivered in the near future from a Philadelphia restaurant named El Merkury, which is owned by Sofia Deleon (also spelled de Leon). The script helps the catering manager of El Merkury with automatically extracting and categorizing information from this email, and then to generate generate an invoice, a kitchen sheet, and a response back to the customer.

## Current Status (as of 22 May 25)
- Live on sofia@elmerkury.com with relatively few bugs
- Github includes Sofia codebase
- Insert into catering@elmerkury.com

## Codebases
| Name    | GL_API_KEY      | SHEET_ID         |
|---------|------------------|------------------|
| El Merkury   | AIzaSyCl56gLtcTXtrXHaQEEqQbMgn7ObVp3ABo   | 1QEnOr1KCu42w36BcIjYuYA9Hol8rPj-uVH8m_244BX8  |
| Gerardo | AIzaSyAGLdXjwawqNJGQmYaqB_Umvv8w4Dg7188   | 1qlEw5k5K-Tqg0joxXeiKtpgr6x8BzclobI-E-_mORhY  |


## Bugs
- If there is a request to separate meat from vegetables, then that's a $5 fee
- Move Utensils to second screen and ensure that the script scans for any indication of number of people, as well as an explicit request for utensils
- Generating PDF from Google Doc doesn't work, only uses invoice (Google server-side bug)

## Tweaks
- Stylistic revisions to invoicing HTML
- Remove "selected flavors" from Google Doc kitchen sheet, only keep customer notes
- Revising template emails (tabs, right customer name instead of delivery contact)

## Future Builds
- Redo the kitchen sheet in HTML, save into system somewhere (Toast integration is bad, Square integration is good)
- Penn PO workflow
- Tax calculation on invoice (tax exempt = Penn)
- QuickBooks integration --> save invoice
- Google Calendar integration of order

## Future evolution, Square + KDS
Email Received
  ↓
Google Apps Script
  - Parse email
  - Look up SKUs in Square
  - Create $0 Order via Square API
      - Tag as "Manual Release"
      - Include item list and delivery info
  - Push structured order to Firebase
      ↓
Square Dashboard
  - Staff manually reviews order
  - Clicks “Mark as Ready”
      ↓
Tablet KDS App
  - Shows all pending orders
  - Staff marks "Ready" / "Completed"
