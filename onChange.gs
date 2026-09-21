
function detectSheetChanges(e) {
  if (!e) {
    return;
  }
  const sheets = workbook.getSheets();
  if (e.changeType === "INSERT_GRID") {
    const trackerSheet = workbook.getSheetByName("_System_Sheet_Names");
    const dataTrackerSheet = workbook.getSheetByName("_System_Sheet_Data");
    if (!trackerSheet || !dataTrackerSheet) {
      updateSheetNameMemory();
      updateSheetMemory();
      return;
    }
    const oldData = trackerSheet.getDataRange().getValues();

    const ids = [];

    oldData.forEach(row => {
      ids.push(row[0]);
    });
    let newSheet = null;
    for (const sheet of sheets) {
      const currentId = sheet.getSheetId();
      if (!ids.includes(currentId)) {
        newSheet = getSheetById(workbook,currentId);
      }
    }

    if (newSheet != null) {
      try {
        newSheet.setFrozenRows(1);
        const firstRow = newSheet.getRange("1:1");
        firstRow.setBackground("#cfe2f3");
        firstRow.setFontWeight("bold");
        const index = newSheet.getIndex();
        const prevSheetName = getSheetByIndex(index-1).getName();
        if (prevSheetName && !prevSheetName.includes("List")) {
          const newItems = [["#",	"Unit",	"Item",	"Serves:",	0,	"Recipe",	"Go To Recipe List"]]
          newSheet.getRange(1,1,1,7).setValues(newItems);
          newSheet.getRange("F:F").setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
          newSheet.setColumnWidth(7, 120);
          newSheet.getRange("H1:H1").insertCheckboxes();
          newSheet.setColumnWidth(8, 50);
          newSheet.getRange(1,9).setValue("Completed?")
          newSheet.getRange("J1:J1").insertCheckboxes();
          newSheet.setColumnWidth(10, 50);
        }
        updateSheetAddition(newSheet.getSheetId());
      }
      catch (error) {
        console.log(error);
      }
    }
    else {
      console.log("null");
    }
  }
  else if (e.changeType === "OTHER") {
    const trackerSheet = workbook.getSheetByName("_System_Sheet_Names");
    const dataTrackerSheet = workbook.getSheetByName("_System_Sheet_Data");
    if (!trackerSheet || !dataTrackerSheet) {
      updateSheetNameMemory();
      updateSheetMemory();
      return;
    }
    updateSheetMemory();
    const oldData = trackerSheet.getDataRange().getValues();
    const oldMemory = {};

    oldData.forEach(row => {
      const id = row[0];
      const name = row[1];
      oldMemory[id] = name;
    });
    for (const sheet of sheets) {
      const id = sheet.getSheetId();
      const currentName = sheet.getName();
      const oldName = oldMemory[id];

      if (oldName && oldName !== currentName) {
        if (oldName.toLowerCase().includes("list")) {
          sheet.setName(oldName);
          break;
        }
        const names = recipeList.getRange(2,2,getLastRowInColumn(recipeList,2)-1).getValues();
        for (let i = 0; i < names.length; i++) {
          if (names[i] == oldName) {
            const currentMaxRows = recipeList.getMaxRows();
            if (i+2 > currentMaxRows) {
              recipeList.insertRowsAfter(currentMaxRows,1000);
            }
            recipeList.getRange(i+2,2).setValue(currentName);
            break;
          }
        }
      }
    }
  }
  else if (e.changeType === "REMOVE_GRID") {
    const trackerSheet = workbook.getSheetByName("_System_Sheet_Names");
    const dataTrackerSheet = workbook.getSheetByName("_System_Sheet_Data");
    if (!trackerSheet || !dataTrackerSheet) {
      updateSheetNameMemory();
      updateSheetMemory();
      return;
    }
    const [id, name] = findMissingId(); 
    if (name.toLowerCase().includes("list")) {
      const startRow = findRowWithValue(dataTrackerSheet,1,id);
      if (startRow != 0) {
        let rowCount = 1;
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
        let index = 0;
        switch(name) {
          case 'Recipe List':
            index = 2;
            break;
          case 'Shopping List':
            index = 3;
            break;
          case 'Owned Items List':
            index = 4;
            break;
          default:
            break;
        }
        workbook.insertSheet(name,index);
        const newSheet = workbook.getSheetByName(name);
        const values = dataTrackerSheet.getRange(startRow,2,rowCount,dataTrackerSheet.getLastColumn()).getValues();
        console.log(values);
        newSheet.setFrozenRows(1);
        const firstRow = newSheet.getRange("1:1");
        firstRow.setBackground("#cfe2f3");
        firstRow.setFontWeight("bold");
        switch(name) {
          case 'Recipe List':
            newSheet.getRange("A2:A").insertCheckboxes();
            newSheet.getRange("D2:D").insertCheckboxes();
            break;
          case 'Shopping List':
            newSheet.getRange(1,5).insertCheckboxes();
            newSheet.getRange(1,7).insertCheckboxes();
            newSheet.getRange(1,9).insertCheckboxes();
            newSheet.setColumnWidth(4,150);
            break;
          case 'Owned Items List':
            newSheet.getRange(1,7).insertCheckboxes();
            newSheet.getRange(1,9).insertCheckboxes();
            break;
          default: 
            break;
        }
        const numberRows = values.length;
        const numberCols = values[0].length;
        if (numberRows > 0 && numberCols > 0) {
          const currentMaxRows = newSheet.getMaxRows();
          if (numberRows > currentMaxRows) {
            newSheet.insertRowsAfter(currentMaxRows,1000);
          }
          newSheet.getRange(1,1,numberRows,numberCols).setValues(values);
        }
        updateSheetAddition(newSheet.getSheetId())
        dataTrackerSheet.deleteRows(startRow, rowCount);
      }
    }
    else {
      const rowToDelete = findRowWithValue(recipeList,2,name);
      if (rowToDelete != 0 && rowToDelete != 1) {
        const rowWithId = findRowWithValue(dataTrackerSheet,1,id);
        if (rowWithId >= 1) {
          dataTrackerSheet.getRange(rowWithId,1).setValue(-1);
        }
        if (recipeList.getRange(rowToDelete,1)) {
          SubtractIngredients(shoppingList, null, rowToDelete);
        }
        recipeList.deleteRow(rowToDelete);
        updateSheetDeletion(-1);
      }
    }
  }
  updateSheetNameMemory();
}