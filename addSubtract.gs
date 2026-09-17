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

function AddIngredients(name, row) {
  `add ingredients`
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
  const totalServings = getTotalServings(row);
  const isNumber = typeof totalServings === 'number' && Number.isFinite(totalServings);
  if (!isNumber) {
    totalServings = servings;
  }
  
  const lastRow = shoppingList.getLastRow();
  let nextRow = lastRow + 1;
  if (lastRow < 2) {
    `list empty`
    let i = 2;
    for (const [number, unit, item] of ingredientPairs) {
      const [newNum, newUnit] = getNewUnit(number*(totalServings/servings), unit);   
      setCell(shoppingList, i, 1, newNum);
      setCell(shoppingList, i, 2, newUnit);
      setCell(shoppingList, i, 3, item);
      i++;
    }
  }
  else {
    `list not empty`
    shopPairs = getIngredients("Shopping List");
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
            newNum = num+number*(totalServings/servings)
          }
          else {
            `add not the same unit`
            unitChange = convert(number, getConversionRate(unit, un));
            newNum = num+unitChange*(totalServings/servings)
          }
          const [newNumber, newUn] = getNewUnit(newNum, newUnit);
          setCell(shoppingList, i, 1, newNumber);
          setCell(shoppingList, i, 2, newUn);
          add = false;
        }
      }
      if (add) {
        `item does not exist`
        const [newNum, newUnit] = getNewUnit(number*(totalServings/servings), unit);   
        setCell(shoppingList, nextRow, 1, newNum);
        setCell(shoppingList, nextRow, 2, newUnit);
        setCell(shoppingList, nextRow, 3, item);
        nextRow++;
      }
    }
  }
}

function SubtractIngredients(name, row, oldValue=null) {
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
  
  const lastRow = shoppingList.getLastRow();
  if (lastRow < 2) {
    console.log("nothing to remove");
  }
  else {
    `possibly items to remove`
    shopPairs = getIngredients("Shopping List");
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
          setCell(shoppingList, i, 1, newNumber);
          setCell(shoppingList, i, 2, newUn); 
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