function getNumServings(name) {
  `get # servings`
  if (!name) {
    const dataTrackerSheet = workbook.getSheetByName("_System_Sheet_Data");
    const lastRow = dataTrackerSheet.getLastRow();
    for (let i = 1; i <= lastRow; i++) {
      if (dataTrackerSheet.getRange(i,1).getValue() === -1) {
        return dataTrackerSheet.getRange(i,6).getValue();
      }
    }
    return -1;
  }
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
  if (!name) {
    const dataTrackerSheet = workbook.getSheetByName("_System_Sheet_Data");
    const lastRow = dataTrackerSheet.getLastRow();
    for (let i = 1; i <= lastRow; i++) {
      if (dataTrackerSheet.getRange(i,1).getValue() === -1) {
        const startRow = i;
        let rowCount = 1;
        if (startRow != 0) {
          const numRows = dataTrackerSheet.getLastRow()-startRow
          if (numRows >= 1) {
            const data = dataTrackerSheet.getRange(startRow+1,1, dataTrackerSheet.getLastRow()-startRow).getValues();
            for (const row of data) {
              const cellValue = row[0];
              if (cellValue === "") {
                rowCount++;
              }
              else {
                break;
              }
            } 
          }
        }
        const amountRange = dataTrackerSheet.getRange(startRow,2,rowCount);
        const unitRange = dataTrackerSheet.getRange(startRow,3,rowCount);
        const itemRange = dataTrackerSheet.getRange(startRow,4,rowCount);
        const rawAmountData = amountRange.getValues();
        const rawUnitData = unitRange.getValues();
        const rawItemData = itemRange.getValues();
        const amountList = rawAmountData.flat().filter(item => item !== "");
        const unitList = rawUnitData.flat().filter(item => item !== "");
        const itemList = rawItemData.flat().filter(item => item !== "");
        const paired = amountList.map((element, index) => [element, unitList[index], itemList[index]]);
        return paired;
      }
    }
    return;
  }
  const recipe = workbook.getSheetByName(name);
  if (!recipe) {
    console.log(name);
    return;
  }
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
    return -1;
  }
  const ingredientPairs = getIngredients(name);

  let servings = 1;
  let totalServings = 0;
  if (sheet === shoppingList || sheet.getName().includes("_System")) {
    servings = getNumServings(name);
    totalServings = getTotalServings(row);
  }
  else {
    totalServings = getNumServings(name);
    oldValue = Number(oldValue);
    if (oldValue != 0) {
      servings = oldValue;
    }
  }
  if (servings === 0) {
    return -1;
  }
  const isNum = typeof servings === 'number' && Number.isFinite(servings);
  if (!isNum) {
    servings = 1;
  }
  const isNumber = typeof totalServings === 'number' && Number.isFinite(totalServings);
  if (!isNumber) {
    totalServings = servings;
  }
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
            if (sheet === shoppingList) {
              newNum = num+number*(totalServings/servings);
            }
            else {
              newNum = number*(totalServings/servings);
            }
          }
          else {
            `add not the same unit`
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

function SubtractIngredients(sheet, name=null, row=null, oldValue=null) {
  `subtract ingredients`
  const ingredientPairs = getIngredients(name);
  if (ingredientPairs === null) {
    console.log(null);
    return;
  }
  let servings = 1;
  let totalServings = 1;
  if (name === null || (name != 'Owned Items List' && !name.includes("_System"))) {
    servings = getNumServings(name);
    totalServings = getTotalServings(row);
  }
  const isNum = typeof servings === 'number' && Number.isFinite(servings);
  if (servings === 0) {
    return -1;
  }
  if (!isNum) {
    servings = 1;
  }
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
    const shopPairs = getIngredients(sheet.getName());
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