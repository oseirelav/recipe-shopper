
function detectSheetChanges(e) {
  if (!e) {
    return;
  }
  const sheets = workbook.getSheets();
  if (e.changeType === "INSERT_GRID") {
    const trackerSheet = workbook.getSheetByName("_System_Sheet_Names");
    if (!trackerSheet) {
      updateSheetNameMemory();
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
    }
    else {
      console.log("null");
    }
  }
  else if (e.changeType === "OTHER") {
    const trackerSheet = workbook.getSheetByName("_System_Sheet_Names");
    if (!trackerSheet) {
      updateSheetNameMemory();
      return;
    }
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
            recipeList.getRange(i+2,2).setValue(currentName);
            break;
          }
        }
      }
    }
  }
  else if (e.changeType === "REMOVE_GRID") {
    console.log("in statement");
    const trackerSheet = workbook.getSheetByName("_System_Sheet_Names");
    if (!trackerSheet) {
      updateSheetNameMemory();
      return;
    }
    const oldData = trackerSheet.getDataRange().getValues();

    let ids = [];
    const oldMemory = {};
    oldData.forEach(row => {
      const id = row[0];
      ids.push(row[0]);
      const name = row[1];
      oldMemory[id] = name;
    });
    console.log(oldMemory);
    for (const sheet of sheets) {
      const currentId = sheet.getSheetId();
      if (ids.includes(currentId)) {
        ids = ids.filter(item => item != currentId);
      }
    }
    console.log(ids, ids[0], oldMemory[ids[0]], "testtttt");
    const prevSheetName = oldMemory[ids[0]];
    const rowToDelete = findRowWithValue(recipeList,2,prevSheetName);
    if (rowToDelete != 0 && rowToDelete != 1) {
      console.log(rowToDelete, "delete row");
      recipeList.deleteRow(rowToDelete);
    }
  }
  updateSheetNameMemory();
}