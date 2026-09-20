function clearCart() {
  shoppingList.getRange(1,6).setValue('FALSE');
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
  shoppingList.getRange(1,8).setValue('FALSE');
  const numRows = recipeList.getLastRow()-1;
  const values = [];
  for (let i = 0; i < numRows; i++) {
    values.push(['FALSE']);
  }
  recipeList.getRange(2,1,numRows).setValues(values);
}

function undoClearCart() {
  if (!(workbook.getSheetByName('Undo Clear Cart'))) {
    return;
  }
  let i = 2;
  for (const [number, unit, item] of getIngredients('Undo Clear Cart')) {
    shoppingList.getRange(i,1).setValue(number);
    shoppingList.getRange(i,2).setValue(unit);
    shoppingList.getRange(i,3).setValue(item);
    i++;
  }
}