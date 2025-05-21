# El Merkury Invoicing - Gmail Add-on
## About
This is an app script to parse emails/email chains, extract content in standardized form, and perform a some actions based on the content of those emails. The email is always a potential customer requesting a catering order to be delivered in the near future from a Philadelphia restaurant named El Merkury, which is owned by Sofia Deleon (also spelled de Leon). The script helps the catering manager of El Merkury with automatically extracting and categorizing information from this email, and then to generate generate an invoice, a kitchen sheet, and a response back to the customer.

## Bugs
- Generating PDF from Google Doc doesn't work, only uses invoice (Google server-side bug)

## Tweaks
- Stylistic revisions to invoicing HTML
- Formatting / ensuring proper information of kitchen sheet. Suggest HTML
- Revising template emails (tabs, right customer name instead of delivery contact)

## Future Builds
- Penn PO workflow
- QuickBooks integration