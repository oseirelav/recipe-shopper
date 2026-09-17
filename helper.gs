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

function getSheetById(workbook,id) {
  return workbook.getSheets().find(sheet => sheet.getSheetId() == id) || null;
}

function getSheetByIndex(index) {
  const sheets = workbook.getSheets();
  return sheets[index-1] || null;
}