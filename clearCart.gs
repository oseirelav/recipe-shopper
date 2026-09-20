function clearCart() {
  shoppingList.getRange(1,7).setValue('FALSE');
  const lastRow = shoppingList.getLastRow();
  if (lastRow < 2) {
    return;
  }
  let backup = workbook.getSheetByName('Undo Clear Cart');
  if (!backup) {
    backup = workbook.insertSheet('Undo Clear Cart');
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
  shoppingList.deleteRows(2,lastRow);
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
  recipeList.getRange(2,1,values.length).setValues(values);
  backup.getRange(1,4,checkedRows.length).setValues(checkedRows);
}

function undoClearCart() {
  const backup = workbook.getSheetByName('Undo Clear Cart');
  if (!backup) {
    return;
  }
  let i = 2;
  for (const [number, unit, item] of getIngredients('Undo Clear Cart')) {
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