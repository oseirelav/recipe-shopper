function formatAllNewSheets(newSheet) {
  newSheet.setFrozenRows(1);
  newSheet.getRange(1, 1, newSheet.getMaxRows(), newSheet.getMaxColumns()).setVerticalAlignment("middle");
  const firstRow = newSheet.getRange("1:1");
  firstRow.setBackground("#cfe2f3");
  firstRow.setFontWeight("bold");
}

function formatNewRecipeSheet(newSheet) {
  formatAllNewSheets(newSheet);
  const newItems = [["#",	"Unit",	"Item",	"Specifications", "Serves:",	0,	"Recipe",	"Go To Recipe List"]]
  newSheet.getRange(1,1,1,8).setValues(newItems);
  newSheet.setColumnWidth(1,50);
  newSheet.setColumnWidth(2,50);
  newSheet.setColumnWidth(3,150);
  newSheet.setColumnWidth(4,150);
  newSheet.getRange("D:D").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  newSheet.setColumnWidth(5,60);
  newSheet.setColumnWidth(6,50);
  newSheet.getRange("C:C").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  newSheet.getRange("G:G").setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  newSheet.setColumnWidth(8, 120);
  newSheet.getRange("I1:I1").insertCheckboxes();
  newSheet.setColumnWidth(9, 50);
  newSheet.getRange(1,10).setValue("Completed?");
  newSheet.getRange("K1:K1").insertCheckboxes();
  newSheet.setColumnWidth(11, 50);
  newSheet.getRange(1,12).setValue("Source:");
  newSheet.setColumnWidth(12, 60);
}
