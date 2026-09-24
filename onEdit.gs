const workbook = SpreadsheetApp.getActiveSpreadsheet();
const recipeList = workbook.getSheetByName("Recipe List");
const shoppingList = workbook.getSheetByName("Shopping List");
const availableRecipes = workbook.getSheetByName("Available Recipes List");

function setCell(sheet, row, column, value) {
  const currentMaxRows = sheet.getMaxRows();
  if (row > currentMaxRows) {
    sheet.insertRowsAfter(currentMaxRows,1000);
  }
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
  const oldValue = e.oldValue || "";
    
  if (sheetName.includes("_System")) {
    range.setValue(oldValue);
    return;
  }

  const currentMaxRows = sheet.getMaxRows();

  if (row === 1 && oldValue != "" && sheetName.toLowerCase().includes("list") && column != 5 && range.isChecked() == null) {
    range.setValue(oldValue);
    return;
  }

  if (sheetName === "Recipe List") {
    if (recipeList.getRange(row,2).getValue() != "") {
      if (column === 1) {
        const cell = sheet.getRange(row,2);
        const name = cell.getValue();
        if (newValue) {
          `checked`
          const success = AddIngredients(shoppingList, name, row);
          if (success === -1) {
            sheet.getRange(row,1).setValue('FALSE');
          }
        }
        else {
          `unchecked`
          const success = SubtractIngredients(shoppingList, name, row);
          if (success === -1) {
            sheet.getRange(row,1).setValue('TRUE');
          }
        }
      }
      else if (column === 2 && row != 1) {
        recipeList.getRange(row,1).clearContent();
        recipeList.getRange(row,1).insertCheckboxes();
        recipeList.getRange(row,4).clearContent();
        recipeList.getRange(row,4).insertCheckboxes();
        if (newValue === "") {
          console.log('not a valid name');
          range.setValue(oldValue);
        }
        else {
          if (oldValue === "") {
            let recipeSheet = workbook.getSheetByName(String(newValue));
            if (!recipeSheet) {
              workbook.insertSheet(String(newValue),workbook.getNumSheets());
            }
            else {
              let i = 1;
              while(true) {
                const newName = String(newValue) + " " + String(i);
                const newSheet = workbook.getSheetByName(newName);
                if (!newSheet) {
                  workbook.insertSheet(newName, workbook.getNumSheets());
                  range.setValue(newName);
                  break; 
                }
                else {
                  i++;
                }
              }
            }
            const newSheet = workbook.getSheetByName(range.getValue());
            formatNewRecipeSheet(newSheet);
            updateSheetAddition(newSheet.getSheetId());
          }
          else {
            console.log("else");
            let newSheet = workbook.getSheetByName(oldValue);
            if (!newSheet) {
              workbook.insertSheet(oldValue,workbook.getNumSheets());
              newSheet = workbook.getSheetByName(oldValue);
              formatNewRecipeSheet(newSheet);
              updateSheetAddition(newSheet.getSheetId());
            }
            try {  
              newSheet.setName(newValue);
            } 
            catch (error) {
              let i = 1;
              while(true) {
                const newName = String(newValue) + " " + String(i);
                try {
                  newSheet.setName(newName);
                  range.setValue(newName);
                  break;
                }
                catch (error) {
                  i++;
                }
              }
            }
          }
          updateSheetNameMemory();
        }
      }
      else if (column === 3) {
        const cell = sheet.getRange(row,2);
        const name = cell.getValue();
        const isNumber = typeof newValue === 'number' && Number.isFinite(newValue);
        if (!isNumber || Number(newValue) < 0) {
          console.log('not a valid number');
          oldValue = Number(oldValue);
          if (Number.isFinite(oldValue)) {
            range.setValue(oldValue);
          }
          else {
            range.setValue(0);
          }
        }
        if (sheet.getRange(row,1).getValue()) {
          `checked`
          SubtractIngredients(shoppingList, name, row, oldValue);
          AddIngredients(shoppingList, name,row);
        }
      }
      else if (column === 4) {
        if (row != 1) {
          const newTabName = recipeList.getRange(row,2).getValue();
          transferTabs(newTabName,row,column);
        }
      }
      else if (column === 5) {
        const recipeName = recipeList.getRange(row,2).getValue();
        const recipePage = workbook.getSheetByName(recipeName);
        recipePage.getRange(1,13).setValue(newValue);
      }
    }
    else {
      range.setValue(oldValue);
    }
  }
  else if (sheetName === "Shopping List") {
    if (row === 1) {
      if (column === 5) {
        if (newValue) {
          SubtractIngredients(shoppingList, 'Owned Items List');
        }
        else {
          resetShoppingList()
        }
      }
      else if (column === 7) {
        if (newValue) {
          clearCart();
        }
      }
      else if (column === 9) {
        if (newValue) {
          undoClearCart();
        }
        else {
          shoppingList.getRange(1,9).setValue('TRUE');
        }
      }
    }
  }
  else if (sheetName === "Owned Items List") {
    if (row === 1) {
      if (column === 7) {
        if (newValue) {
          clearCart(sheet);
        }
      }
      else if (column === 9) {
        if (newValue) {
          undoClearCart(sheet);
        }
      }
    }
    updateAvailableRecipes()
  }
  else if (sheetName === "Available Recipes List") {
    if (column === 1) {
      if (oldValue != "") {
        range.setValue(oldValue);
      }
    }
    else if (column === 2 || (row === 1 && column === 5)) {
      const isNumber = typeof newValue === 'number' && Number.isFinite(newValue);
      if (!isNumber || Number(newValue) < 0) {
        console.log('not a valid number');
        oldValue = Number(oldValue);
        if (Number.isFinite(oldValue)) {
          range.setValue(oldValue);
        }
        else {
          range.setValue(0);
        }
      }
      if (column === 5) {
        updateAvailableRecipes();
      }
    }
    else if (column === 3) {
      if (row != 1) {
        const newTabName = sheet.getRange(row,1).getValue();
        transferTabs(newTabName,row,column);
      }
    }
  }
  if (sheetName === "Substitutions List") {
    if (column === 6 && row != 1) {
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
    else if (column == 5 && row != 1) {
      if (!Number(newValue) || Number(newValue) < 0) {
        range.setValue(oldValue);
      }
    }
  }
  if (sheetName != "Shopping List" || row != 1 || column != 7) {
    const backup = workbook.getSheetByName('_System_Cart_Data');
    if (backup) {
      try {
        workbook.deleteSheet(backup);
      }
      catch (error) {
        console.log(error);
      }
      shoppingList.getRange(1,9).setValue('TRUE');
    }
  }
  if (sheetName != "Owned Items List" || row != 1 || column != 7) {
    const backup = workbook.getSheetByName('_System_Owned_Data');
    if (backup) {
      try {
        workbook.deleteSheet(backup);
      }
      catch (error) {
        console.log(error);
      }
      workbook.getSheetByName("Owned Items List").getRange(1,9).setValue(true);
    }
  }
  if (sheetName != "Recipe List" && !sheetName.includes("_System") && sheetName != "Available Recipes List") {
    if (column === 2 && row != 1) {
      const unitCell = sheet.getRange(row,2);
      const newValue = unitCell.getValue();
      const numberCell = sheet.getRange(row,1);
      const number = numberCell.getValue();
      if (oldValue) {
        if ((!isUnit(newValue) && isUnit(oldValue)) || !equivalentUnits(oldValue, newValue)) {
          console.log('not a valid unit conversion');
          unitCell.setValue(oldValue);
        }
        else {
          conversion = convert(number, getConversionRate(oldValue,newValue));
          numberCell.setValue(conversion);
        }
      }
    }
    else if (column == 1 && row != 1) {
      if (!Number(newValue) || Number(newValue) < 0) {
        range.setValue(oldValue);
      }
    }
  }
  if (!sheetName.toLowerCase().includes("list")) {
    if (row === 1) {    
      if (column === 13) {
        if (sheet.getRange(1,11).getValue()) {
          recipeList.getRange(findRowWithValue(recipeList,2,sheetName),5).setValue(newValue);
        }
      }
      else if (column === 11) {
        if (newValue) {
          `checked`
          if (Number(sheet.getRange(1,6).getValue()) <= 0) {
            range.setValue(oldValue);
          }
          else {
            let newRow = findRowWithValue(recipeList,2,sheetName);
            if (newRow === 0) {
              newRow = getLastRowInColumn(recipeList,2)+1;
              if (newRow > currentMaxRows) {
                sheet.insertRowsAfter(currentMaxRows,1000);
              }
              recipeList.getRange(newRow, 2).setValue(sheetName);
            }
            if (newRow > currentMaxRows) {
              sheet.insertRowsAfter(currentMaxRows,1000);
            }
            const servings = sheet.getRange(1,6).getValue();
            recipeList.getRange(newRow, 3).setValue(servings);
            recipeList.getRange(newRow,1).clearContent();
            recipeList.getRange(newRow,1).insertCheckboxes();
            recipeList.getRange(newRow,4).clearContent();
            recipeList.getRange(newRow,4).insertCheckboxes();
            recipeList.getRange(newRow,5).setValue(sheet.getRange(1,13).getValue());
            const [possible, numServings] = isPossibleRecipe(sheet,getIngredients("Owned Items List"), Number(availableRecipes.getRange(1,5).getValue()));
            if (possible) {
              const insertRow = availableRecipes.getLastRow()+1;
              const availableRecipesMaxRows = availableRecipes.getMaxRows();
              if (insertRow > availableRecipesMaxRows) {
                availableRecipes.insertRowsAfter(availableRecipesMaxRows,1000);
              }
              availableRecipes.getRange(insertRow,1).setValue(sheetName);
              availableRecipes.getRange(insertRow,2).setValue(numServings);
              availableRecipes.getRange(insertRow,3).insertCheckboxes();
            }
          }
        }
        else {
          `unchecked`
          const newRow = findRowWithValue(recipeList,2, sheetName);
          if (newRow != 0 && newRow != 1) {
            if (recipeList.getRange(newRow,1)) {
              SubtractIngredients(shoppingList,sheetName,newRow);
            }
            recipeList.deleteRow(newRow);
          }
          const rowToDeleteAvailableRecipes = findRowWithValue(availableRecipes, 1, sheetName);
          if (rowToDeleteAvailableRecipes != 0 && rowToDeleteAvailableRecipes != 1) {
            availableRecipes.deleteRow(rowToDeleteAvailableRecipes);
          }
        }
      }
      else if (column === 9) {
        if (newValue) {
          `checked`
          transferTabs("Recipe List",row,column);
        }
        else {
          `unchecked`
          console.log("unchecked somehow");
        }
      }
      else if (column === 6) {
        const isNumber = typeof newValue === 'number' && Number.isFinite(newValue);
        if (!isNumber || Number(newValue) < 0) {
          console.log('not a valid number');
          oldValue = Number(oldValue);
          if (Number.isFinite(oldValue)) {
            range.setValue(oldValue);
          }
          else {
            range.setValue(0);
          }
        }
        else {
          if (sheet.getRange(1,11).getValue()) { 
            if (Number(newValue) === 0) {
              range.setValue(oldValue);
            }
            else {
              AddIngredients(sheet, sheetName, null, oldValue);
              recipeList.getRange(findRowWithValue(recipeList,2,sheetName),3).setValue(newValue);
            }
          }
        }
      }
      else {
        range.setValue(oldValue);
      }
    }
  }
  updateSheetEdit(sheet.getSheetId());
}