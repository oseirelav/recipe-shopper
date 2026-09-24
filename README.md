## Configuration
To make this project run on a Google Sheet, you will need to open a new Google Sheet, create all necessary sheets and then go to Extensions > Apps Script > paste all functions into the document

Necessary sheets are as follows:
- Recipe List
- Shopping List
- Owned Items List
- Substitutions List
- Available Recipes List

After pasting all functions, you will want to delete all sheets to get the proper formatting

Then, you can type in a recipe name in Recipe List.

To make the detectSheetChanges function run, you will need to go to the Apps Script > Triggers > Add Trigger > fill out information as follows:
- Choose which function to run: detectSheetChanges
- Choose which deployment should run: Head
- Select event source: From spreadsheet
- Select event type: On change
- Failure notification settings: user's choice

You may need to press save multiple times and go through Google's warning.

Alternatively, you can just use the following link to make a copy: https://docs.google.com/spreadsheets/d/1mwQxNJO-tsN8jSlOwO1Ge9gL8TPf4-bnhfRF5PetBqc/copy

Please note you will still need to go into Apps Script and set up the trigger for detectSheetChanges.
