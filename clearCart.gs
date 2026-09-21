function clearCart(sheet=shoppingList) {
  sheet.getRange(1,7).setValue('FALSE');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }
  let name = '_System_Cart_Data';
  if (sheet.getName() === 'Owned Items List') {
    name = '_System_Owned_Data';
  }
  let backup = workbook.getSheetByName(name);
  if (!backup) {
    backup = workbook.insertSheet(name);
  }
  backup.clear();
  backup.hideSheet();
  let i = 2;
  for (const [number, unit, item] of getIngredients(sheet.getName())) {
    const currentMaxRows = backup.getMaxRows();
    if (i > currentMaxRows) {
      backup.insertRowsAfter(currentMaxRows,1000);
    }
    backup.getRange(i,1).setValue(number);
    backup.getRange(i,2).setValue(unit);
    backup.getRange(i,3).setValue(item);
    i++;
  }
  sheet.deleteRows(2,lastRow-1);
  sheet.getRange(1,9).setValue('FALSE');
  if (sheet.getName() === "Shopping List") {
    const lastRowInColumn = getLastRowInColumn(recipeList, 2);
    const checked = recipeList.getRange(2,1,lastRowInColumn-1).getValues();
    const checkedRows = [];
    const values = [];
    for (let i = 2; i < lastRowInColumn+1; i++) {
      if (checked[i-2][0]) {
        checkedRows.push([i]);
      }
      values.push(['FALSE']);
    }
    const numRows = values.length;
    if (numRows > 0) {
      recipeList.getRange(2,1,numRows).setValues(values);
    }
    const numCheckedRows = checkedRows.length;
    if (numCheckedRows > 0) {
      backup.getRange(1,4,numCheckedRows).setValues(checkedRows);
    }
  }
  else if (sheet.getName() === "Owned Items List") {
    if (shoppingList.getRange(1,5).getValue()) {
      resetShoppingList();
    }
  }
}

function undoClearCart(sheet=shoppingList) {
  console.log(sheet.getName(), "name");
  let name = '_System_Cart_Data';
  if (sheet.getName() === 'Owned Items List') {
    name = '_System_Owned_Data'
  }
  const backup = workbook.getSheetByName(name);
  if (!backup) {
    console.log(name);
    return;
  }
  console.log("test")
  let i = 2;
  for (const [number, unit, item] of getIngredients(name)) {
    const currentMaxRows = backup.getMaxRows();
    if (i > currentMaxRows) {
      backup.insertRowsAfter(currentMaxRows,1000);
    }
    sheet.getRange(i,1).setValue(number);
    sheet.getRange(i,2).setValue(unit);
    sheet.getRange(i,3).setValue(item);
    i++;
  }
  if (sheet.getName() === "Shopping List") {
    const vals = backup.getRange(1,4, backup.getLastRow()).getValues();
    for (const val of vals) {
      if (Number(val)) {
        recipeList.getRange(Number(val),1).setValue('TRUE');
      }
    }
  }
  else if (sheet.getName() === "Owned Items List") {
    SubtractIngredients(shoppingList, '_System_Owned_Data');
  }
}