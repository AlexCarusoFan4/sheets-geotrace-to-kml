// onOpen executes everytime a user opens the document
function onOpen() {
  // Get User Interface
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ODK Utilities')
      .addItem('Convert sheet lines to KML','LinesToKML')
      .addSeparator()
      .addItem('Convert sheet polygons to KML','PolygonsToKML')
      .addToUi();
      
}

function LinesToKML(){
 // Display a dialog box with a title, message, input field, and "Yes" and "No" buttons. The
 // user can also close the dialog by clicking the close button in its title bar.
 var ui = SpreadsheetApp.getUi();
 var responseColumnSelected = ui.prompt('Column name', 'What is the name of the column with the "geotrace" data?', ui.ButtonSet.OK_CANCEL);

 // In the case user hits OK
 if (responseColumnSelected.getSelectedButton() == ui.Button.OK) {
   // Get value from the textbox
   var columnName = responseColumnSelected.getResponseText();
   var sheet = SpreadsheetApp.getActiveSheet();
   try{
     // Findout the column Number
     // Function found in: https://gist.github.com/russenreaktor/5520691
     var columnIndex = getColumnNrByName(sheet, columnName);
     // Get the letter of that column
     var columnID = String.fromCharCode(65+columnIndex);
     var alertMsg = "The script will read column '" + columnID + "' and stop when it finds any blank cell.";
     var alertScriptStop = ui.alert(alertMsg);
     // Call the function that makes the magic
     convertColumnToLine(columnIndex);
   } catch(err){
     // If name of the column provided doesn't exist
     var alertColumnNotFound = ui.alert(err);
   }
 }
}

function PolygonsToKML(){
 // Display a dialog box with a title, message, input field, and "Yes" and "No" buttons. The
 // user can also close the dialog by clicking the close button in its title bar.
 var ui = SpreadsheetApp.getUi();
 var responseColumnSelected = ui.prompt('Column name', 'What is the name of the column with the "geotrace" data?', ui.ButtonSet.OK_CANCEL);

 // In the case user hits OK
 if (responseColumnSelected.getSelectedButton() == ui.Button.OK) {
   // Get value from the textbox
   var columnName = responseColumnSelected.getResponseText();
   var sheet = SpreadsheetApp.getActiveSheet();
   try{
     // Findout the column Number
     // Function found in: https://gist.github.com/russenreaktor/5520691
     var columnIndex = getColumnNrByName(sheet, columnName);
     // Get the letter of that column
     var columnID = String.fromCharCode(65+columnIndex);
     var alertMsg = "The script will read column '" + columnID + "' and stop when it finds any blank cell.";
     var alertScriptStop = ui.alert(alertMsg);
     // Call the function that makes the magic
     convertColumntoPolygon(columnIndex);
   } catch(err){
     // If name of the column provided doesn't exist
     var alertColumnNotFound = ui.alert(err);
   }
 }
}



function createGoogleDriveTextFile() {
  var content,fileName,mimeType,newFile;//Declare variable names
  
 // mimeType = "kml";
  fileName = "KML Export " + new Date().toString().slice(0,15);//Create a new file name with date on end
  content = "This is the file Content";

  newFile = DriveApp.createFile(fileName,content,mimeType);//Create a new KML file in the root folder
}
