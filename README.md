# Google Spreadsheet Addon Gmail Label Email Username Filter

## Usage

1. Create a new Google Spreadsheet.
2. Click *Extension > Apps Script* to open the Google Sheet Script Editor.
3. Paste the code of `extract-email-by-label.js` into edit area.
4. Config the `EXCEPTION_EMAIL_1` to `EXCEPTION_EMAIL_3` variable if you want to ignore some e-mail. If you need more, just union them by `&&` operator.
5. Select function to `onOpen()`.
6. Back to Spreadsheet, in each sheet, input the label name in `B1`. If the label is nested, connect the parent label via `-`. For example, `Shopping-Amazon-Receipt`.
7. Click the *Extract Emails > Extract Emails...*
8. The result will show up since forth row.
9. If the sender doesn't has name, it will display "N/A"

## Note

1. Due to the Gmail's restriction, the maximum number of crawling mail is 500.
2. Each time we click *Extract Emails*, it will clear the current sheet. If the data is needed, please backup or create a new sheet before using this add-on.
3. If there is an error, please check the label name (missing blank? wrong spelling?) or nested relationship.

