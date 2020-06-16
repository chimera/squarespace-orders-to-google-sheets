# Squarespace Orders to Google Sheets
A Google Script that accesses the Squarespace Orders API and fills in a spreadsheet with details

## Installation
- Create a Google Spreadsheet. Name the sheet something memorable.
- Either manually set the following values for your header in Row 1, or uncomment lines 293-294 to have the script automatically overwrite Row 1 every time it runs.
```
    "Imported E-Mail Data",
    "Attendee Name",
    "Purchased",
    "Class Date",
    "Order placed on",
    "Qty",
    'Unit Price',
    "Phone",
    "E-mail",
    "Additional Info",
    "Member?"
```
- Click **Tools > Script Editor**.
- Copy paste the script.js into the code window.
- Change the `API_KEY` and `SHEET_NAME` and other variables as desired.
- Click Save, then choose the `importOrderData` function in the dropdown, then click Run. Data should appear on the last row in the file, presuming that Row 1 has some sort of headers filled in.
- After you've verified it works, go to *Edit > Current Project's Triggers*
  and set up a trigger to run `importOrderData` on a regular basis.

## Notes
Currently, the last row of data will be continually replaced/updated
every run (and used as a cursor to remember the last order) so if you're
going to edit the data make sure you're not editing the last row cuz it'll
either be overwritten or cause duplication.

Also, currently there is no provision for orders with multiple quantities.
Ideally the order would be somewhat duplicated and items given unique IDs
to give each item its own row.
