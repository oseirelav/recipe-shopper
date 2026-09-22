function isPossibleRecipe(recipe, ownedPairs, minServings) {
  try {
    const recipePairs = getIngredients(recipe.getName());
    const numServings = recipe.getRange(1,5).getValue();
    let maxServings = Infinity;
    for (const [number, unit, item] of recipePairs) {
      let hasItem = false;
      let ownedNumber = 0;
      let ownedUnit = "";
      for (const [num, un, it] of ownedPairs) {
        if (it === item) {
          hasItem = true;
          ownedNumber = num;
          ownedUnit = un;
          break;
        }
      }
      if (!hasItem) {
        return [false, null];
      }
      if (equivalentUnits(unit, ownedUnit)) {
        if (number/numServings*minServings > ownedNumber) {
          return [false, null];
        }
      }
      const newNum = convert(ownedNumber,getConversionRate(ownedUnit, unit));
      if (number/numServings*minServings > newNum) {
        return [false, null];
      }
      maxServings = Math.min(maxServings,Math.floor(newNum*numServings/number));
    }
    return [true,maxServings];
  }
  catch (error) {
    console.log(error);
    return [false, null];
  }
}

function updateAvailableRecipes() {
  const numRows = availableRecipes.getLastRow()-1;
  if (numRows > 0) {
    availableRecipes.deleteRows(2,numRows);
  }
  const allSheets = workbook.getSheets();
  const recipeSheets = [];
  for (const sheet of allSheets) {
    if (sheet && !sheet.getName().toLowerCase().includes("list") && !sheet.getName().includes("_System") && sheet.getRange(1,10).getValue()) {
      recipeSheets.push(sheet);
    }
  }
  const minServings = availableRecipes.getRange(1,5).getValue();
  let nextRow = 2;
  const ownedPairs = getIngredients("Owned Items List");
  for (const recipe of recipeSheets) {
    const [possible, numServings] = isPossibleRecipe(recipe, ownedPairs, minServings);
    if (possible) {
      const currentMaxRows = availableRecipes.getMaxRows();
      if (nextRow > currentMaxRows) {
        availableRecipes.insertRowsAfter(currentMaxRows,1000);
      }
      availableRecipes.getRange(nextRow,1).setValue(recipe.getName());
      availableRecipes.getRange(nextRow,2).setValue(numServings);
      availableRecipes.getRange(nextRow,3).insertCheckboxes();
      nextRow++;
    }
  }
}
