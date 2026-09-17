const workbook = SpreadsheetApp.getActiveSpreadsheet();
const recipeList = workbook.getSheetByName("Recipe List");
const shoppingList = workbook.getSheetByName("Shopping List");

function setCell(sheet, row, column, value) {
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

  if (sheetName === "Recipe List") {
    if (column === 1) {
      const cell = sheet.getRange(row,2);
      const name = cell.getValue();
      if (newValue) {
        `checked`
        AddIngredients(name, row);
      }
      else {
        `unchecked`
        sheet.getRange(range.getRow(),2);
        SubtractIngredients(name, row);
      }
    }
    else if (column === 2) {
      const oldValue = e.oldValue || "";
      
      if (newValue === "") {
        console.log('not a valid name');
        range.setValue(oldValue);
      }
      else {
        if (oldValue === "") {
          let recipeSheet = workbook.getSheetByName(String(newValue));
          if (!recipeSheet) {
            workbook.insertSheet(String(newValue));
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
        }
        else {
          console.log("else");
          let newSheet = workbook.getSheetByName(oldValue);
          if (!newSheet) {
            workbook.insertSheet(oldValue,workbook.getNumSheets());
            newSheet = workbook.getSheetByName(oldValue);
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
      }
    }
    else if (column === 3) {
      const cell = sheet.getRange(row,2);
      const name = cell.getValue();
      const oldValue = e.oldValue || 0;
      const isNumber = typeof newValue === 'number' && Number.isFinite(newValue);
      if (!isNumber) {
        console.log('not a valid number');
        range.setValue(oldValue);
      }
      if (sheet.getRange(row,1).getValue()) {
        `checked`
        SubtractIngredients(name, row, oldValue);
        AddIngredients(name,row);
      }
    }
    else if (column === 4) {
      if (row != 1) {
        const newTabName = recipeList.getRange(row,2).getValue();
        transferTabs(newTabName,row,column);
      }
    }
  }
  if (sheetName === "Shopping List") {
    if (row === 1) {
      if (column === 6) {
        if (newValue) {
          clearCart();
        }
      }
      else if (column === 8) {
        if (newValue) {
          undoClearCart();
        }
        else {
          shoppingList.getRange(1,8).setValue('TRUE');
        }
      }
    }
  }
  if (sheetName != "Shopping List" || row != 1 || column != 6) {
    const backup = workbook.getSheetByName('Undo Clear Cart');
    if (backup) {
      workbook.deleteSheet(backup);
      shoppingList.getRange(1,8).setValue('TRUE');
    }
  }
  if ((sheetName != "Recipe List" && sheetName != "Owned Items List") && column === 2 && row != 1) {
    const oldValue = e.oldValue || "";
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
  if (!sheetName.toLowerCase().includes("list")) {
    if (row === 1) {    
      if (column === 10) {
        if (newValue) {
          `checked`
          const newRow = getLastRowInColumn(recipeList,2)+1
          const servings = sheet.getRange(1,5).getValue();
          
          recipeList.getRange(newRow, 2).setValue(sheetName);
          recipeList.getRange(newRow, 3).setValue(servings);
        }
        else {
          `unchecked`
          const newRow = findRowWithValue(recipeList,2, sheetName);
          if (newRow != 0 && newRow != 1) {
            recipeList.deleteRow(newRow);
          }
        }
      }
      else if (column === 8) {
        if (newValue) {
          `checked`
          transferTabs("Recipe List",row,column);
        }
        else {
          `unchecked`
          console.log("unchecked somehow");
        }
      }
    }
  }
}