const workbook = SpreadsheetApp.getActiveSpreadsheet();
const recipeList = workbook.getSheetByName("Recipe List");
const shoppingList = workbook.getSheetByName("Shopping List");

function setCell(sheet, row, column, value) {
  const cell = sheet.getRange(row, column);
  cell.setValue(value);
}

function getNumServings(name) {
  console.log("get # servings");
  const recipe = workbook.getSheetByName(name);
  const servingRange = recipe.getRange(1,5);
  const servings = servingRange.getValue();
  return servings;
}

function getTotalServings(row) {
  console.log("get total servings");
  const servingRange = recipeList.getRange(row,3);
  const servings = servingRange.getValue();
  return servings;
}

function getConversionRate(unit, newUnit) {
  console.log('need to code conversion rates')
  return 1;
}

function convert(number, conversionRate) {
  console.log("convert to different units");
  return number*conversionRate;
}

function getNewUnit(number, unit) {
  console.log("choose conversion");
  newUnit = "test"
  console.log('need to code how to get new unit')
  newNum = convert(number, getConversionRate(unit, newUnit));
  return (newNum,newUnit);
}

function getIngredients(name) {
  console.log("get ingredients")
  const recipe = workbook.getSheetByName(name);
  const lastRow = recipe.getLastRow();
  if (lastRow < 2) {
    console.log("no ingredients listed")
    return;
  }
  const amountRange = recipe.getRange(2,1,lastRow-1);
  const unitRange = recipe.getRange(2,2,lastRow-1);
  const itemRange = recipe.getRange(2,3,lastRow-1);
  const rawAmountData = amountRange.getValues();
  const rawUnitData = unitRange.getValues();
  const rawItemData = itemRange.getValues();
  const amountList = rawAmountData.flat().filter(item => item !== "");
  const unitList = rawUnitData.flat().filter(item => item !== "");
  const itemList = rawItemData.flat().filter(item => item !== "");
  const paired = amountList.map((element, index) => [element, unitList[index], itemList[index]]);
  return paired;
}

function AddIngredients(name, row) {
  console.log("add ingredients")
  if (!name) {
    console.log("no recipe")
    return;
  }
  const timestampCell = recipeList.getRange(2, 5);
  timestampCell.setValue(new Date());
  const ingredientPairs = getIngredients(name);

  const servings = getNumServings(name);
  const isNum = typeof servings === 'number' && Number.isFinite(servings);
  if (!isNum) {
    servings = 1;
  }
  const totalServings = getTotalServings(row);
  console.log(servings,totalServings);
  const isNumber = typeof totalServings === 'number' && Number.isFinite(totalServings);
  if (!isNumber) {
    totalServings = servings;
  }
  
  const lastRow = shoppingList.getLastRow();
  let nextRow = lastRow + 1;
  if (lastRow < 2) {
    console.log("list empty");
    let i = 2;
    for (const [number, unit, item] of ingredientPairs) {
      setCell(shoppingList, i, 1, number*(totalServings/servings));
      setCell(shoppingList, i, 2, unit);
      setCell(shoppingList, i, 3, item);
      i++;
    }
  }
  else {
    console.log("list not empty");
    shopPairs = getIngredients("Shopping List");
    for (const [number, unit, item] of ingredientPairs) {
      let add = true;
      for (let i = 2; i <= shopPairs.length+1; i++) {
        const [num, un, it] = shopPairs[i-2];
        if (it == item) {
          console.log("item exists");
          let newNum = 0;
          let newUnit = un;   
          if (un == unit) {
            console.log("add same unit");
            newNum = num+number*(totalServings/servings)
          }
          else {
            console.log("add not the same unit");
            unitChange = convert(number, getConversionRate(unit, un));
            newNum = num+unitChange*(totalServings/servings)
          }
          newNum, newUnit = getNewUnit(newNum, newUnit);
          console.log("new num", newNum);
          setCell(shoppingList, i, 1, newNum);
          setCell(shoppingList, i, 2, newUnit);
          add = false;
        }
      }
      if (add) {
        console.log("item does not exist");
        setCell(shoppingList, nextRow, 1, number*(totalServings/servings));
        setCell(shoppingList, nextRow, 2, unit);
        setCell(shoppingList, nextRow, 3, item);
        nextRow++;
      }
    }
  }
}

function SubtractIngredients(name, row) {
  console.log("subtract ingredients")
  if (!name) {
    console.log("no recipe")
    return;
  }
  const timestampCell = recipeList.getRange(2, 5);
  timestampCell.clearContent();
  const ingredientPairs = getIngredients(name);
  
  const servings = getNumServings(name);
  const isNum = typeof servings === 'number' && Number.isFinite(servings);
  if (!isNum) {
    servings = 1;
  }
  const totalServings = getTotalServings(row);
  console.log(servings,totalServings);
  const isNumber = typeof totalServings === 'number' && Number.isFinite(totalServings);
  if (!isNumber) {
    totalServings = servings;
  }
  
  const lastRow = shoppingList.getLastRow();
  if (lastRow < 2) {
    console.log("nothing to remove");
  }
  else {
    console.log("possibly items to remove");
    shopPairs = getIngredients("Shopping List");
    for (const [number, unit, item] of ingredientPairs) {
      for (let i = 2; i <= shopPairs.length+1; i++) {
        const [num, un, it] = shopPairs[i-2];
        if (it == item) {
          console.log("item is there", i, shopPairs.length+1);
          if (un == unit) {
            console.log("subtract same unit");
            const newNum = Math.max(0,num-number*(totalServings/servings));
            setCell(shoppingList, i, 1, newNum);
          }
          else {
            console.log("subtract not the same unit");
          }
        }
      }
    }
    const amountRange = shoppingList.getRange(2,1,lastRow-1);
    const rawAmountData = amountRange.getValues();
    const amountList = rawAmountData.flat().filter(item => item !== "");
    for (let i = amountList.length+1; i >= 2;i--) {
      if (amountList[i-2] == 0) {
        shoppingList.deleteRow(i);
      }
    }
  }
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
    console.log("checked")
    AddIngredients(name, row);
  }
  else {
    console.log("unchecked")
    sheet.getRange(range.getRow(),2);
    SubtractIngredients(name, row);
  }
}


