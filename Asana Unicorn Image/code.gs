/* working - filters
function onEdit() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Specify the range of cells you want to sort
  var rangeToSort = sheet.getRange("A2:Z100"); // Change "A1:Z100" to your desired range
  
  // Sort the range in ascending order based on the first column
  rangeToSort.sort({column: 2, ascending: true});
}
*/

/* working - sorts, only deletes row when column J is checked

function onEdit(e) {
  var sheet = e && e.source ? e.source.getActiveSheet() : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Check if the event object and the range are defined
  if (e && e.range) {
    var range = e.range;

    // Check if the edited range is within column J and a checkbox is checked
    if (range.getColumn() == 10 && range.isChecked()) { // Column J is the 10th column
      var row = range.getRow();
      sheet.deleteRow(row);
    } else if (range.getColumn() == 2) { // Column B is the 2nd column
      // Sort the range in ascending order based on column B
      var rangeToSort = sheet.getRange("A2:Z100"); // Change "A2:Z100" to your desired range
      rangeToSort.sort({column: 2, ascending: true});
    }
  }
}

*/

/* working - launches image on checkbox
function onEdit(e) {
  const range = e.range;
  if (range.isChecked() && range.getValue() === true) {
    startMovingImage();
  }
}

*/


function onEdit(e) {
  var sheet = e && e.source ? e.source.getActiveSheet() : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (e && e.range) {
    var range = e.range;

    // Handle checkbox selection in column J
    if (range.isChecked() && range.getValue() === true && range.getColumn() == 10) {  // Check for column J

      startMovingImage(range.getRow()); // Call startMovingImage only for column J checkboxes

      // Delayed row deletion with 5 second delay
      var rowToDelete = range.getRow();
      Utilities.sleep(3000); // Delay for 3 seconds
      sheet.deleteRow(rowToDelete);
    }

    // Handle sorting (unchanged)
    else if (range.getColumn() == 2) {
      var rangeToSort = sheet.getRange("A2:Z100"); // Adjust range if needed
      rangeToSort.sort({column: 2, ascending: true});
    }
  }
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




/* filters, cuts row and adds to Closed tab, random celebration

function onEdit(e) {
  var sheet = e && e.source ? e.source.getActiveSheet() : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (e && e.range) {
    var range = e.range;

    // Handle checkbox selection in column J
    if (range.isChecked() && range.getValue() === true && range.getColumn() == 10) {  // Check for column J

      // Randomly choose one of the three functions
      var randomFunction = Math.floor(Math.random() * 3); // Generate a random number between 0 and 2

      switch (randomFunction) {
        case 0:
          startMovingImage(range.getRow());
          break;
        case 1:
          startMovingImage2(range.getRow());
          break;
        case 2:
          startMovingImage3(range.getRow());
          break;
      }

      // Cut the row
      var rowToCut = range.getRow();
      var rowData = sheet.getRange(rowToCut, 1, 1, sheet.getLastColumn()).getValues()[0]; // Get the row data
      sheet.deleteRow(rowToCut);

      // Delayed row deletion with 5 second delay
      Utilities.sleep(2000); // Delay for 2 seconds

      // Insert the row into the "Closed" tab at row 2
      var closedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Closed");
      if (!closedSheet) {
        closedSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Closed");
      }
      
      // Get active user's name
      var activeUser = Session.getActiveUser().getEmail();

      // Append active user's name to the row data
      rowData.push(activeUser);

      closedSheet.insertRowBefore(2); // Insert a new row at row 2
      var closedRange = closedSheet.getRange(2, 1, 1, rowData.length);
      closedRange.setValues([rowData]); // Set the row data in row 2 of "Closed"

      // Add current date to column L (12th column) in the "Closed" tab
      var currentDate = new Date();
      var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
      closedSheet.getRange(2, 12).setValue(formattedDate);
    }

    // Handle sorting (unchanged)
    else if (range.getColumn() == 2) {
      var rangeToSort = sheet.getRange("A2:Z100"); // Adjust range if needed
      rangeToSort.sort({column: 2, ascending: true});
    }
  }
}

*/
