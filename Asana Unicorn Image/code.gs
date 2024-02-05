function onEdit() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Specify the range of cells you want to sort
  var rangeToSort = sheet.getRange("A2:Z100"); // Change "A1:Z100" to your desired range
  
  // Sort the range in ascending order based on the first column
  rangeToSort.sort({column: 2, ascending: true});
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CRISP Tools')
      .addItem('Celebrate', 'startMovingImage')
      .addItem('Celebrate 2', 'startMovingImage2')
      .addItem('Celebrate 3', 'startMovingImage3')

      .addToUi();
}


function startMovingImage() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('MovingImage')
      .setWidth(500)
      .setHeight(500);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'Way To Go!');
}

function startMovingImage2() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('MovingImage2')
      .setWidth(500)
      .setHeight(500);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'You Did It!');
}

function startMovingImage3() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('MovingImage3')
      .setWidth(500)
      .setHeight(500);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'Amazing!');
}
