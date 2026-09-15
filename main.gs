const workbook = SpreadsheetApp.getActiveSpreadsheet();
const recipeList = workbook.getSheetByName("Recipe List");
const shoppingList = workbook.getSheetByName("Shopping List");

function setCell(sheet, row, column, value) {
  const cell = sheet.getRange(row, column);
  cell.setValue(value);
}

function onEdit(e) {
  if (!e) {
    return;
  }
  const range = e.range;
  const sheet = range.getSheet();

  const isRightSheet = sheet.getName() === "Recipe List"
  const isRightColumn = range.getColumn() === 1;

  if (!isRightSheet || !isRightColumn) {
    return;
  }

  const row = range.getRow();
  const cell = sheet.getRange(row,2);
  const name = cell.getValue();

  if (range.getValue()) {
    `checked`
    AddIngredients(name, row);
  }
  else {
    `unchecked`
    sheet.getRange(range.getRow(),2);
    SubtractIngredients(name, row);
  }
}