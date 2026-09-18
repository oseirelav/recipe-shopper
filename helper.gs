function getLastRowInColumn(sheet, column) {
  const maxRow = sheet.getLastRow();
  if (maxRow === 0) {
    return 0;
  }
  const columnData = sheet.getRange(1,column,maxRow,1).getValues();
  for (let i = columnData.length - 1; i >= 0; i--) {
    if (columnData[i][0] !== "") {
      return i + 1;
    }
  }
  return 0;
}

function findRowWithValue(sheet,column, value) {
  const maxRow = sheet.getLastRow();
  if (maxRow === 0) {
    return 0;
  }
  const columnData = sheet.getRange(1,column,maxRow,1).getValues();
  for (let i = 0; i < columnData.length; i++) {
    if (columnData[i] == value) {
      return i+1;
    }
  }
  return 0;
}

function updateSheetNameMemory() {
  let trackerSheet = workbook.getSheetByName("_System_Sheet_Names");
  if (!trackerSheet) {
    trackerSheet = workbook.insertSheet("_System_Sheet_Names");
    trackerSheet.hideSheet();
  }
  const sheets = workbook.getSheets();
  const data = [];

  sheets.forEach(sheet => {
    data.push([sheet.getSheetId(), sheet.getName()]);
  })

  trackerSheet.clear();
  trackerSheet.getRange(1,1, data.length,2).setValues(data);
}

function updateSheetMemory() {
  let trackerSheet = workbook.getSheetByName("_System_Sheet_Data");
  if (!trackerSheet) {
    trackerSheet = workbook.insertSheet("_System_Sheet_Data");
    trackerSheet.hideSheet();
  }
  const sheets = workbook.getSheets();
  const data = [];

  sheets.forEach(sheet => {
    if (sheet.getName() != "_System_Sheet_Data") {
      data.push([sheet.getSheetId(), sheet.getDataRange().getValues()]);
    }
  })
  trackerSheet.clear();
  let row = 1;
  for (const [id, sheetData] of data) {
    trackerSheet.getRange(row,1).setValue(id);
    if (getSheetById(workbook,id).getLastRow() === 0) {
      break;
    }
    trackerSheet.getRange(row,2, sheetData.length,sheetData[0].length).setValues(sheetData);
    row += sheetData.length;
  }
}

function findMissingId() {
  const trackerSheet = workbook.getSheetByName("_System_Sheet_Names");
  const sheets = workbook.getSheets();
  const oldData = trackerSheet.getDataRange().getValues();

  let ids = [];
  const oldMemory = {};
  oldData.forEach(row => {
    const id = row[0];
    ids.push(row[0]);
    const name = row[1];
    oldMemory[id] = name;
  });

  for (const sheet of sheets) {
    const currentId = sheet.getSheetId();
    if (ids.includes(currentId)) {
      ids = ids.filter(item => item != currentId);
    }
  }
  return [ids[0],oldMemory[ids[0]]];
}

function updateSheetDeletion(id) {
  const dataTrackerSheet = workbook.getSheetByName("_System_Sheet_Data");
  if (!dataTrackerSheet) {
    updateSheetMemory();
    return;
  }
  const startRow = findRowWithValue(dataTrackerSheet,1,id);

  if (startRow != 0) {
    let deleteCount = 1;
    const numRows = dataTrackerSheet.getLastRow()-startRow
    if (numRows >= 1) {
      const data = dataTrackerSheet.getRange(startRow+1,1, dataTrackerSheet.getLastRow()-startRow).getValues();
      for (const row of data) {
        const cellValue = row[0];
        if (cellValue === "") {
          deleteCount++;
        }
        else {
          break;
        }
      } 
    }
    dataTrackerSheet.deleteRows(startRow, deleteCount);
  }
}

function updateSheetAddition(id) {
  const data = workbook.getSheetById(id).getDataRange().getValues();

  const dataTrackerSheet = workbook.getSheetByName("_System_Sheet_Data");

  if (!dataTrackerSheet) {
    updateSheetMemory();
    return;
  }

  const startRow = dataTrackerSheet.getLastRow()+1;
  dataTrackerSheet.getRange(startRow,1).setValue(id);
  dataTrackerSheet.getRange(startRow,2, data.length,data[0].length).setValues(data);
  console.log("testing", startRow, data);
}

function updateSheetEdit(id) {
  updateSheetDeletion(id);
  updateSheetAddition(id);
}

function getSheetById(workbook,id) {
  return workbook.getSheets().find(sheet => sheet.getSheetId() == id) || null;
}

function getSheetByIndex(index) {
  const sheets = workbook.getSheets();
  return sheets[index-1] || null;
}