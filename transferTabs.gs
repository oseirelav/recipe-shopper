function transferTabs(newName, row, column) {
  const currentSheet = workbook.getActiveSheet();
  currentSheet.getRange(row,column).setValue('FALSE');
  goToTab(newName);
}

function goToTab(name) {
  if (name) {
    workbook.getSheetByName(name).activate();
  }
}
