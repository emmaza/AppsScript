/**
 * Creates a Google Slides presenation containing a slide for each row in a Google Sheet. The function reads the data from the active sheet and populates each slide in the presentation with the corresponding row's data. The function also updates the sheet with the links to the newly created Google Slides presentation.
 */
function mailMergeCertificatesFromSheets() {
  // Load data from the spreadsheet
  var dataSheet = SpreadsheetApp.getActiveSheet();
  var dataRange = dataSheet.getDataRange();
  var sheetContents = dataRange.getValues();

  // Save the header in a variable called header
  var header = sheetContents.shift();

  // Create an array to save the data to be written back to the sheet.
  // We'll use this array to save links to Google Slides.
  var updatedContents = [];

  // Create a new Google Slides presentation
  var presentation = createCopyOfSlidesTemplate();
  var slides = presentation.getSlides();
  var slide = slides[0];

  // Loop through each row in the sheet
  // for (var i = 0; i < sheetContents.length; i++) {
  // the output slides should be in the right order, not reverse
  for (var i = sheetContents.length-1; i >=0 ; i--) {
    var row = sheetContents[i];
    var firstName = row[0];
    var lastName = row[1];
    var artCategory = row[2];
    var slidesUrl = row[3];
    
    // Create a new slide in the presentation and populate it with data from the row
    var newSlide = slide.duplicate();
    newSlide.replaceAllText("{{firstName}}", firstName);
    newSlide.replaceAllText("{{lastName}}", lastName);
    newSlide.replaceAllText("{{artCategory}}", artCategory);

    // Generate a new Google Slides link for this specific slide
    // all slides are in the same presentation
    slidesUrl = `https://docs.google.com/presentation/d/${presentation.getId()}/edit#slide=id.${newSlide.getObjectId()}`;

    // Update the corresponding row in the sheet with the new Google Slides link
    row[3] = slidesUrl;  

    // Add the updated row to the array that will be written back to the sheet
    updatedContents.unshift(row);
  }

  // remove the template page in the presentation
  slide.remove();

  // Add the header to the array that will be written back to the sheet.
  updatedContents.unshift(header);

  // Write the updated data back to the Google Sheets spreadsheet.
  dataRange.setValues(updatedContents);
}


function createCopyOfSlidesTemplate() {
  // Change the TEMPLATE_ID to your template slide ID
  var TEMPLATE_ID = "1Y6X57unOYqkb46LkIs2X9pdKlnSPHD5Cz6qLzd3L9-g";

  // Create a copy of the file using DriveApp
  var copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy();

  // Load the copy using the SlidesApp.
  var slides = SlidesApp.openById(copy.getId());

  return slides;
}

function onOpen() {
 // Create a custom menu to make it easy to run the Mail Merge
 // script from the sheet.
 SpreadsheetApp.getUi().createMenu("Create_certificates")
   .addItem("Create certificates", "mailMergeCertificatesFromSheets")
   .addToUi();
}
