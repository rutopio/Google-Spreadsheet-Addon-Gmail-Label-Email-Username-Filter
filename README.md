# Google Spreadsheet Addon Gmail Label Email Username Filter


A Google Spreadsheet add-on that crawls usernames and emails based on a Gmail specific label. 
For instance, if you want to send a newsletter to someone labeled `Participant-2022`, you can use this add-on that easily crawls sender information in just one click.

## Usage

1. Create a new Google Spreadsheet.

2. Click *Extension > Apps Script* to open the Google Sheet Script Editor.

3. Paste the code of `extract-email-by-label.js` into the edit area.

4. Configure the `EXCEPTION_EMAIL_1` to `EXCEPTION_EMAIL_3` variable if you want to ignore some e-mail addresses. If you need more exception rules, just union them with the `&&` operator.

5. Select the function `onOpen()`.

6. Back to the spreadsheet, in each sheet, input the label name in `B1`. If the label is nested, connect the parent (and even grandparent) label name via `-`. For example, `Shopping-Amazon-Receipt`.

7. Click the *Extract Emails > Extract Emails...*

8. The result will show up since the forth row.

9. If the sender doesn't has name, it will display "N/A".


## Note

1. Due to the Gmail's restriction, the maximum number of crawling mail is **500**.

2. Each time we click Extract Emails, it will clear the current sheet and re-generate the data. If the former data is needed, please backup or create a new sheet before using this add-on.

3. If there is an error, please check the label name (missing blank? wrong spelling?) or nested relationship.

4. Google might change the position of `Apps Script` menu (it was at *Tools > Script* editor before), if you can't find it, please read the latest Google Spreadsheet Guide.