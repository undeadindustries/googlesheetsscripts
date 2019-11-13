/**
*If you have many sheets/tabs in a Google sheet, you can create a "Summary" sheet. 
*This script will populate all of column A to a horizontal row in the Summary sheet. 
*This way, you have one summary you can reference. It helps to lock the Summary sheet so no one edits it on purpose. 
*You must edit the sheets it references. 
*This also assumes you have a tab called "template" that you can copy from to create new tabs. It skips that tab.
*Tab names are case sensitive. ie: Summary needs the capital S. template needs the lowercase t. 
*template must have the same number of rows with data as the rest of your sheets. That is how the script knows how many rows to copy.
*/
var numHeaders = 0
function onOpen() {
  numHeaders = getNumHeaders();
  autoPopulateHeader();
  autoPopulateCells();
}

//this function only works up to column AZ. I didn't need more then that.
function nextChar(c) {
    if (c.length < 2 && c != 'z') {
     return String.fromCharCode(c.charCodeAt(0) + 1);
    }
    if (c == 'z') {
     return 'aa'
    }
    if (c.length > 1) {
      a = c.split();
      var lastLetter = a[a.length - 1];
      lastLetter = String.fromCharCode(lastLetter.charCodeAt(0) + 1);
      return c.replace(/.$/,lastLetter)
    }
    showAlert('error in nextChar function')
    return 'a'
  }
  
  function getNumHeaders() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cellLen = 1;
    var row = 0;
    var templateSheet = ss.getSheetByName('template');
    while (cellLen > 0) {
      row++
      var cellVal = templateSheet.getRange('a'+row).getValue();
      cellLen = cellVal.length;
    }
    return row-1
  }
  
  function autoPopulateHeader() {
    if (numHeaders == 0) {
     numHeaders = getNumHeaders();
    }
    
    var col = 'a';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var template = ss.getSheetByName('template'); 
    var templateCells = template.getRange('a1:a'+numHeaders).getValues();
    var summary = ss.getSheetByName('Summary');
    for (i = 0; i < numHeaders; i++) {
      summary.getRange(col+'2').setValue( templateCells[i] );
      col = nextChar(col)
    }
  }
  
  
  
  function autoPopulateCells() {
    if (numHeaders == 0) {
     numHeaders = getNumHeaders();
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var summary = ss.getSheetByName('Summary');
    var row = 3;
    var col = 'a';
    
    for (var sheet in sheets) {
      if (sheets[sheet].getName() == 'Summary' || sheets[sheet].getName() == 'template') {
       continue;
      }
       vals = sheets[sheet].getRange('b1:b'+ numHeaders).getValues()
      for (i = 0; i < numHeaders; i++) {
        try {
         summary.getRange(col+row).setValue( vals[i] );
        } catch (err) {
           showAlert(err);
            var str = 'Col: '+ col +', row: '+ row +', i: '+i+', sheet: '+sheet;
            showAlert(str);
        }
       col = nextChar(col)
      }
      row++;
      col = 'a';
    }
    
  }
      
  function showAlert(alertText) {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(alertText);
  }
 