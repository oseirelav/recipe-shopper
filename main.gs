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
  const row = range.getRow();

  if (sheet.getName() === "Recipe List" && range.getColumn() === 1) {
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
  if (sheet.getName() != "Recipe List" && range.getColumn() === 2) {
    const oldValue = e.oldValue || "";
    const unitCell = sheet.getRange(row,2);
    const newValue = unitCell.getValue();
    const numberCell = sheet.getRange(row,1);
    const number = numberCell.getValue();
    if (oldValue) {
      if (!isUnit(oldValue) || !equivalentUnits(oldValue, newValue)) {
        console.log('not a valid unit conversion');
        unitCell.setValue(oldValue);
      }
      else {
        conversion = convert(number, getConversionRate(oldValue,newValue));
        numberCell.setValue(conversion);
      }
    }
  }
}
