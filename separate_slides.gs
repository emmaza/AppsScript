/**
 * Creates a slide for each row in the Google Sheet.
 */
function mailMergeSlidesFromSheets() {
  // Load data from the spreadsheet
  var dataRange = SpreadsheetApp.getActive().getDataRange();
  var sheetContents = dataRange.getValues();

  // Save the header in a variable called header
  var header = sheetContents.shift();

  // Create an array to save the data to be written back to the sheet.
  // We'll use this array to save links to Google Slides.
  var updatedContents = [];

  // Add the header to the array that will be written back
  // to the sheet.
  updatedContents.push(header);

  // For each row, see if the 4th column is empty.
  // If it is empty, it means that a slide deck hasn't been
  // created yet.
  sheetContents.forEach(function(row) {
    if(row[3] === "") {
      // Create a Google Slides presentation using
      // information from the row.
      var slides = createSlidesFromRow(row);
      var slidesId = slides.getId();
   
      // Create the Google Slides' URL using its Id.
      var slidesUrl = `https://docs.google.com/presentation/d/${slidesId}/edit`;

      // Add this URL to the 4th column of the row and add this row
      // to the updatedContents array to be written back to the sheet.
      row[3] = slidesUrl;
      updatedContents.push(row);
    }
  });

  // Write the updated data back to the Google Sheets spreadsheet.
  dataRange.setValues(updatedContents);

}

function createSlidesFromRow(row) {
 // Create a copy of the Slides template
 var deck = createCopyOfSlidesTemplate();

 // Rename the deck using the firstname and lastname of the student
 deck.setName(row[0] + " " + row[1]);

 // Replace template variables using the student's information.
 deck.replaceAllText("{{firstName}}", row[0]);
 deck.replaceAllText("{{lastName}}", row[1]);
 deck.replaceAllText("{{grade}}", row[2]);
 deck.replaceAllText("{{artCategory}}", row[4]);


 return deck;
}

function createCopyOfSlidesTemplate() {
 //
 var TEMPLATE_ID = "1_nQhzo8N4q-C4fa_ayRrq2pA1caG_7jzqrZDg0FqyL4";

 // Create a copy of the file using DriveApp
 var copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy();

 // Load the copy using the SlidesApp.
 var slides = SlidesApp.openById(copy.getId());

 return slides;
}

function onOpen() {
 // Create a custom menu to make it easy to run the Mail Merge
 // script from the sheet.
 SpreadsheetApp.getUi().createMenu("reflections")
   .addItem("Create Slides", "mailMergeSlidesFromSheets")
   .addToUi();
}