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
  const sheetName = sheet.getName();
  const newValue = range.getValue();
  const column = range.getColumn();

  if (sheetName === "Recipe List") {
    if (column === 1) {
      const cell = sheet.getRange(row,2);
      const name = cell.getValue();
      if (newValue) {
        `checked`
        AddIngredients(name, row);
      }
      else {
        `unchecked`
        sheet.getRange(range.getRow(),2);
        SubtractIngredients(name, row);
      }
    }
    if (column === 3) {
      const cell = sheet.getRange(row,2);
      const name = cell.getValue();
      const oldValue = e.oldValue || 0;
      const isNumber = typeof newValue === 'number' && Number.isFinite(newValue);
      if (!isNumber) {
        console.log('not a valid number');
        range.setValue(oldValue);
      }
      if (sheet.getRange(row,1).getValue()) {
        `checked`
        SubtractIngredients(name, row, oldValue);
        AddIngredients(name,row);
      }
    }
  }
  if (sheetName === "Shopping List") {
    if (row === 1) {
      if (column === 6) {
        if (newValue) {
          clearCart();
        }
      }
      else if (column === 8) {
        if (newValue) {
          undoClearCart();
        }
        else {
          shoppingList.getRange(1,8).setValue('TRUE');
        }
      }
    }
  }
  if (sheetName != "Shopping List" || row != 1 || column != 6) {
    const backup = workbook.getSheetByName('Undo Clear Cart');
    if (backup) {
      workbook.deleteSheet(backup);
      shoppingList.getRange(1,8).setValue('TRUE');
    }
  }
  if ((sheetName != "Recipe List" || "Owned Items List") && column === 2 && row != 1) {
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
