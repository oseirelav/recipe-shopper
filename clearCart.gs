function clearCart() {
  shoppingList.getRange(1,7).setValue('FALSE');
  const lastRow = shoppingList.getLastRow();
  if (lastRow < 2) {
    return;
  }
  let backup = workbook.getSheetByName('_System_Cart_Data');
  if (!backup) {
    backup = workbook.insertSheet('_System_Cart_Data');
  }
  backup.clear();
  backup.hideSheet();
  let i = 2;
  for (const [number, unit, item] of getIngredients("Shopping List")) {
    backup.getRange(i,1).setValue(number);
    backup.getRange(i,2).setValue(unit);
    backup.getRange(i,3).setValue(item);
    i++;
  }
  shoppingList.deleteRows(2,lastRow-1);
  shoppingList.getRange(1,9).setValue('FALSE');
    
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

function undoClearCart() {
  const backup = workbook.getSheetByName('_System_Cart_Data');
  if (!backup) {
    return;
  }
  let i = 2;
  for (const [number, unit, item] of getIngredients('_System_Cart_Data')) {
    shoppingList.getRange(i,1).setValue(number);
    shoppingList.getRange(i,2).setValue(unit);
    shoppingList.getRange(i,3).setValue(item);
    i++;
  }
  const vals = backup.getRange(1,4, backup.getLastRow()).getValues();
  for (const val of vals) {
    if (Number(val)) {
      recipeList.getRange(Number(val),1).setValue('TRUE');
    }
  }
}