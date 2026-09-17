function getNumServings(name) {
  `get # servings`
  const recipe = workbook.getSheetByName(name);
  const servingRange = recipe.getRange(1,5);
  const servings = servingRange.getValue();
  return servings;
}

function getTotalServings(row) {
  `get total servings`
  const servingRange = recipeList.getRange(row,3);
  const servings = servingRange.getValue();
  return servings;
}

function getIngredients(name) {
  `get ingredients`
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

function AddIngredients(sheet, name, row, oldValue = null) {
  `add ingredients`
  if (!name) {
    console.log("no recipe")
    return;
  }
  const ingredientPairs = getIngredients(name);

  let servings = 1;
  let totalServings = 0;
  if (sheet === shoppingList) {
    servings = getNumServings(name);
    totalServings = getTotalServings(row);
  }
  else {
    totalServings = getNumServings(name);
    oldValue = Number(oldValue);
    if (oldValue != 0) {
      console.log("test");
      servings = oldValue;
    }
  }
  const isNum = typeof servings === 'number' && Number.isFinite(servings);
  if (!isNum) {
    servings = 1;
  }
  const isNumber = typeof totalServings === 'number' && Number.isFinite(totalServings);
  if (!isNumber) {
    totalServings = servings;
  }
  console.log(servings, totalServings, name, oldValue);
  const lastRow = sheet.getLastRow();
  let nextRow = lastRow + 1;
  if (lastRow < 2) {
    `list empty`
    let i = 2;
    for (const [number, unit, item] of ingredientPairs) {
      const [newNum, newUnit] = getNewUnit(number*(totalServings/servings), unit);   
      setCell(sheet, i, 1, newNum);
      setCell(sheet, i, 2, newUnit);
      setCell(sheet, i, 3, item);
      i++;
    }
  }
  else {
    `list not empty`
    console.log(sheet.getName(), "hello");
    shopPairs = getIngredients(sheet.getName());
    for (const [number, unit, item] of ingredientPairs) {
      let add = true;
      for (let i = 2; i <= shopPairs.length+1; i++) {
        const [num, un, it] = shopPairs[i-2];
        if (it == item) {
          `item exists`
          let newNum = 0;
          let newUnit = un;   
          if (un == unit) {
            `add same unit`
            console.log(sheet.getName(), "same");
            if (sheet === shoppingList) {
              newNum = num+number*(totalServings/servings);
            }
            else {
              newNum = number*(totalServings/servings);
            }
          }
          else {
            `add not the same unit`
            console.log(sheet.getName(), "not same")
            unitChange = convert(number, getConversionRate(unit, un));
            if (sheet === shoppingList) {
              newNum = num+unitChange*(totalServings/servings)
            }
            else {
              newNum = number*(totalServings/servings);
            }
          }
          const [newNumber, newUn] = getNewUnit(newNum, newUnit);
          setCell(sheet, i, 1, newNumber);
          setCell(sheet, i, 2, newUn);
          add = false;
          console.log(newNumber, newUn, newNum, newUnit);
        }
      }
      if (add) {
        `item does not exist`
        const [newNum, newUnit] = getNewUnit(number*(totalServings/servings), unit);   
        setCell(sheet, nextRow, 1, newNum);
        setCell(sheet, nextRow, 2, newUnit);
        setCell(sheet, nextRow, 3, item);
        nextRow++;
      }
    }
  }
}

function SubtractIngredients(sheet, name, row, oldValue=null) {
  `subtract ingredients`
  if (!name) {
    console.log("no recipe")
    return;
  }
  const ingredientPairs = getIngredients(name);
  
  const servings = getNumServings(name);
  const isNum = typeof servings === 'number' && Number.isFinite(servings);
  if (!isNum) {
    servings = 1;
  }
  let totalServings = getTotalServings(row);
  if(!oldValue) {
    const isNumber = typeof totalServings === 'number' && Number.isFinite(totalServings);
    if (!isNumber) {
      totalServings = servings;
    }
  }
  else {
    oldValue = Number(oldValue);
    if (!Number.isFinite(oldValue)) {
      totalServings = servings;
    }
    else {
      totalServings = oldValue;
    }
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    console.log("nothing to remove");
  }
  else {
    `possibly items to remove`
    shopPairs = getIngredients(sheet.getName());
    for (const [number, unit, item] of ingredientPairs) {
      for (let i = 2; i <= shopPairs.length+1; i++) {
        const [num, un, it] = shopPairs[i-2];
        if (it == item) {
          let newNum = 0;
          let newUnit = un;   
          if (un == unit) {
            `subtract same unit`
            newNum = Math.max(0,num-number*(totalServings/servings));
          }
          else {
            `subtract not the same unit`
            unitChange = convert(number, getConversionRate(unit, un));
            newNum = num-unitChange*(totalServings/servings);

          }
          const [newNumber, newUn] = getNewUnit(newNum, newUnit);
          setCell(sheet, i, 1, Math.max(newNumber,0));
          setCell(sheet, i, 2, newUn); 
        }
      }
    }
    const amountRange = sheet.getRange(2,1,lastRow-1);
    const rawAmountData = amountRange.getValues();
    const amountList = rawAmountData.flat().filter(item => item !== "");
    for (let i = amountList.length+1; i >= 2;i--) {
      if (amountList[i-2] == 0) {
        sheet.deleteRow(i);
      }
    }
  }
}