/** @OnlyCurrentDoc */


function P() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D7:D11').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Purchase Stock'), true);
  spreadsheet.getRange('B5').activate();
  spreadsheet.getRange('\'Order interface\'!D7:D11').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);
  spreadsheet.getRange('B5').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Order interface'), true);
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('D4').activate();
};



function S() {
   var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D7:D11').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Sell Products '), true);
  spreadsheet.getRange('B5').activate();
  spreadsheet.getRange('\'Order interface\'!D7:D11').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true);
  spreadsheet.getRange('B5').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Order interface'), true);
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('D4').activate();
};
