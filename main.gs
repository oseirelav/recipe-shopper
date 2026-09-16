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

  if (sheet.getName() === "Recipe List") {
    if (range.getColumn() === 1) {
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
    if (range.getColumn() === 3) {
      if (sheet.getRange(row,1).getValue()) {
        `checked`
        const cell = sheet.getRange(row,2);
        const name = cell.getValue();
        const oldValue = e.oldValue || 0;
        const newValue = range.getValue();
        const isNumber = typeof newValue === 'number' && Number.isFinite(newValue);
        if (!isNumber) {
          console.log('not a valid number');
          range.setValue(oldValue);
        }
        else {
          SubtractIngredients(name, row, oldValue);
          AddIngredients(name,row);
        }
      }
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
