function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Maintenance')
        .addItem('Sort Sheets', 'sortSheets').addToUi();
}

function sortSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var sheetNames = [];
    for (var sheet in sheets) {
        sheetNames.push(sheets[sheet].getName());
    }
    sheetNames.sort();
    SpreadsheetApp.getUi().alert(sheetNames); return;
    x = 1;
    for (var sheet in sheetNames) {
        ss.setActiveSheet(ss.getSheetByName(sheetNames[sheet]));
        ss.moveActiveSheet(x);
        x++;
    }
    SpreadsheetApp.getUi().alert('Done sorting sheets.');
}