## Configuration
To make this project run on a Google Sheet, you will need to open a new Google Sheet > Extensions > Apps Script > paste all functions into the document.

To make the detectSheetChanges function run, you will need to go to the Apps Script > Triggers > Add Trigger > fill out information as follows:
Choose which function to run: detectSheetChanges
Choose which deployment should run: Head
Select event source: From spreadsheet
Select event type: On change
Failure notification settings: user's choice
