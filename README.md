# invoice-to-csv

Builds Xero-ready invoice CSVs from Google Sheets, for bulk import instead
of manual entry.

Built for [Care Culture](https://careculture.com.au), an NDIS provider,
where invoicing meant batching many similar invoices at once. Xero's
invoice UI is fine for one invoice and slow for forty: every line item is
several clicks and a page wait. Xero accepts CSV imports, so this tool does
the assembly in a spreadsheet and produces an upload-ready CSV.

## How it works

- Enter invoice data in the sheet, one row per line item. Columns:
  Participant Name, Item/Program, Date of Service, Qty (hours/kms),
  Account, Reference, Issue Date, Due Date, InvoiceNumber, Unit Amount,
  Account Number, Description, Item Code, and Cancellation.
- The columns are NDIS-flavoured (participants, hours/kms, item codes),
  but it's just Xero contact and line-item data underneath; rename to suit.
- The script groups line items by invoice and writes a CSV matching Xero's
  sales invoice import format.
- Import the CSV in Xero (Business → Invoices → Import) and approve the
  drafts there.

Xero's import creates draft invoices, so nothing is sent to a customer
without review.

## Setup

1. [Copy the template sheet](https://docs.google.com/spreadsheets/d/1_kaGakTYlVhIAVzrF_eI92yMmj5GDcITP4J7gHkHbTA)
   (File → Make a copy). The script is bound to the sheet, so the copy
   includes it.
2. Use the button on the Instructions sheet to generate the CSV.
3. On first run, Google will ask you to authorise the script. Read
   `code.js` first so you know what you're authorising; it's short.

To use it in an existing sheet instead, create a bound script
(Extensions → Apps Script, or script.google.com) and paste in the contents
of `code.js`.

## Notes

- Built against Xero's CSV import format as of 2021; if Xero changes the
  format, the column mapping in `code.js` is the place to update.
- This is a small internal tool, published as-is. It has run our invoicing
  since 2021, but read the code before trusting it with yours.

MIT.
