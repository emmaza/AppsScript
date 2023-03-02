// Creates a Google Slides presenation containing a slide for each row in a Google Sheet.
// Images will be inserted into the corresponding slides.
// For videos and audios, only file links are inserted.

function mailMergeShowcaseFromSheets() {
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

  // Loop through each row in the sheet in reverse order
  // so that the output slides are in the right order
  for (var i = sheetContents.length-1; i >=0 ; i--) {
    var row = sheetContents[i];
    var slidesUrl = row[0];
    var firstName = row[1];
    var lastName = row[2];
    var artCategory = row[3];
    var title = row[4];
    var statement = row[5];
    var artworkLink = row[6];  

    // Create a new slide in the presentation and populate it with data from the row
    var newSlide = slide.duplicate();
    newSlide.replaceAllText("{{firstName}}", firstName);
    // use last name initials for presenation
    newSlide.replaceAllText("{{lastName}}", lastName.charAt(0) + ".");
    newSlide.replaceAllText("{{artCategory}}", artCategory);
    newSlide.replaceAllText("{{title}}", title);
    newSlide.replaceAllText("{{statement}}", statement);

    // Add image or video to the slide depending on the art category
    if (artCategory === "Visual Arts" || artCategory === "Photography" || artCategory === "Literature") {
      putImage(newSlide, artworkLink);
    } else if (artworkLink != "") {
      putVideo(newSlide, artworkLink);
    }
   
    // Generate a new Google Slides link for this specific slide
    // all slides are in the same presentation
    slidesUrl = `https://docs.google.com/presentation/d/${presentation.getId()}/edit#slide=id.${newSlide.getObjectId()}`;

    // Update the corresponding row in the sheet with the new Google Slides link
    row[0] = slidesUrl;  

    // Add the updated row to the array that will be written back to the sheet
    updatedContents.unshift(row);
  }

  // remove the template page in the presentation
  slide.remove();

  // Add the header to the array that will be written back
  // to the sheet.
  updatedContents.unshift(header);

  // Write the updated data back to the Google Sheets spreadsheet.
  dataRange.setValues(updatedContents);

}

function putImage(newSlide, artworkLink){
  // Extract the file ID from the artwork link
  const [fileId, ] = artworkLink.match(/[-\w]{25,}/);

  // Get the image file from Google Drive
  var imageFile = DriveApp.getFileById(fileId);

  try {
    // Insert the image into the slide
    var image = newSlide.insertImage(imageFile.getBlob());

    // Define maximum width and height for the image
    const MAX_WIDTH = 400;
    const MAX_HEIGHT = 470;

    // Calculate the best ratio to resize the image
    const widthRatio = MAX_WIDTH / image.getWidth();
    const heightRatio = MAX_HEIGHT / image.getHeight();
    const ratio = Math.min(widthRatio, heightRatio);

    // Resize the image
    const newWidth = Math.floor(image.getWidth() * ratio);
    const newHeight = Math.floor(image.getHeight() * ratio);
    image.setWidth(newWidth);
    image.setHeight(newHeight);

    // Position the image on the right side of the slide
    image.setLeft(420);
    image.setTop(100);
  } catch (e) {
    Logger.log(`Error inserting image: ${e.message} ${artworkLink}`);
  }

}

function putVideo(newSlide, artworkLink){    
  // slide.inserVideo only works with YouTube uploads
  // has to manually insert google drive videos

  // Insert the artworkLink into the slide
  try {
    var textBox = newSlide.insertTextBox(artworkLink);
    var text = textBox.getText();
    text.getTextStyle().setLinkUrl(artworkLink);
  } catch (e) {
    Logger.log("Error with video/audio link: " + e.message + artworkLink);
  }

}

function createCopyOfSlidesTemplate() {
  // Change the TEMPLATE_ID to your slide ID
  const TEMPLATE_ID = "1775gRlIa9okVMyHOZz0BEOlhia1jIciGJJRgGukqaZQ";

  // Create a copy of the file using DriveApp
  var copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy();

  // Load the copy using the SlidesApp.
  var slides = SlidesApp.openById(copy.getId());

  return slides;
}

function onOpen() {
 // Create a custom menu to make it easy to run the Mail Merge
 // script from the sheet.
 SpreadsheetApp.getUi().createMenu("Create_showcase")
   .addItem("Create showcase", "mailMergeShowcaseFromSheets")
   .addToUi();
}
