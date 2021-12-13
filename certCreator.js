// *************************************   CertCreator   *************************************
/*******************************************************
 * Copyright (C) 2010-Present Avant Assessment
 * All Rights Reserved
 *******************************************************/
// Author: Hugh MacWilliams

// Function: onOpen()
// Purpose: runs automatically every time the presentation is opened, sets up the ui dropdown menu
function onOpen() {

  const ui = SlidesApp.getUi();
  const menu = ui.createMenu('Cert Creator');                       //  create menu/dropdown
  menu.addItem('Specify Range', 'getUserInput');                    //  each option and its corresponding function
  menu.addItem('Individual', 'indiv')
  menu.addItem('Problems?', 'problems');

  menu.addToUi();

}


// Function: indiv()
// Purpose: this function is the same as getUserInput, but passes in a bool that will
// modify function to pull data from the individual sheet instead
function indiv(){

  getUserInput(true)

}



// Function: createSlide(int row, string presid, int index)
// Purpose: this function takes a single row from the sheet as input, makes a copy of the master slide,
// changes placeholder values to the new ones from that row, and appends said slide to the end of the slide deck.
function createSlide(row, presentationid, index, indiv) {

  const pres = SlidesApp.openById(presentationid);                  //  opens presentation by passed in ID, returns presentation object
  var slides = pres.getSlides();                                    //  gets template slide from deck

  pres.appendSlide(slides[0]);                                      //  appends a copy of the template slide to the end of the deck
  var slides = pres.getSlides();                                    //  updates the slide object so that you get the deck with appended slide


  //  lines 45-51 search for values in the slide and replace them with matching values from the corresponding row on the sheet
  // the column numbers are slightly different on the indiv sheet, hence the conditional
  if(indiv){
    slides[slides.length - 1].replaceAllText("{{FIRST NAME}} {{LAST NAME}}", row[16] + " " + row[17]);
    slides[slides.length - 1].replaceAllText("Organization",  "");
    slides[slides.length - 1].replaceAllText("Location",  row[15]);
    slides[slides.length - 1].replaceAllText("Serial Number",  row[0]);
    slides[slides.length - 1].replaceAllText("AWARD EARNED",  row[18] );
    slides[slides.length - 1].replaceAllText("Test_scale",  row[33] );
    slides[slides.length - 1].replaceAllText("Languages",  row[19] + " & " + row[20] );
  }
  //  lines 61-67 execute iff the user specifies individual from the dropdown menu
  else{
    slides[slides.length - 1].replaceAllText("{{FIRST NAME}} {{LAST NAME}}", row[16] + " " + row[17]);
    slides[slides.length - 1].replaceAllText("Organization",  row[14]);
    slides[slides.length - 1].replaceAllText("Location",  row[15]);
    slides[slides.length - 1].replaceAllText("Serial Number",  row[0]);
    slides[slides.length - 1].replaceAllText("AWARD EARNED",  row[18] );
    slides[slides.length - 1].replaceAllText("Test_scale",  row[22] );
    slides[slides.length - 1].replaceAllText("Languages",  row[19] + " & " + row[20] );
  }



}



// Function: createSlides(int row, string pres_id)
// Purpose: this function takes in the array of rows that are to be printed, then just runs createSlide on each row.
function createSlides(rows, presentationid, indiv){

  var i;

  for(i = 0; i < rows.length; i++){

    createSlide(rows[i], presentationid, i + 1, indiv);

  }

  if(indiv){
    var file = DriveApp.getFileById("1hEi0x5kJYd4jDwTSZKt2u--MdxThZvfBSQa9T7fN2Ro");
  }
  else{
    var file = DriveApp.getFileById("1JHtsi_62eqfKc2_VXKUm44OwX6EQyOSem0VSQzrv8Nk");
  }

}


// Function: getdata(int bottomrow, int toprow)
// Purpose: this function opens the specified sheet, gets the data from the sheet, cuts it down to the specified data, and calls createSlides using that data
function getdata(bottomrow, toprow, indiv){

  //  getting the sheet object with openByUrl
  if(indiv){
    //  opens individual sheet if specified
    var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1hEi0x5kJYd4jDwTSZKt2u--MdxThZvfBSQa9T7fN2Ro/edit#gid=0');
  }
  else{
    // defaults to Master GSB Fill
    var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/18A0ttcoWK2o52GPHIZkUfoNhInGiTGnHvDxdAuoc1GA/edit#gid=1676112815');
  }

  //  getting all of the data from the sheet
  const rows = ss.getDataRange().getValues();

  //  cutting data into selected rows
  const selectedRows = rows.slice(bottomrow - 1, toprow);

  //  calling createSlides with selectedRows and slide deck id
  createSlides(selectedRows, SlidesApp.getActivePresentation().getId(), indiv);

}



// Function: problems()
// Purpose: runs when the "Problems?" option is selected on the dropdown menu, essentially just emails hughrm23@gmail.com with response to prompt
function problems(){

  const ui = SlidesApp.getUi();
  var response = ui.prompt(' Problems?', 'Please enter your issue below and we will get back to you. \n Alternatively, you could email hughrm23@gmail.com  with concerns or bugs. \n Thanks for using CertCreator!\n\n', ui.ButtonSet.OK_CANCEL);


  if(response == ui.Button.OK){
    MailApp.sendEmail('hughrm23@gmail.com', 'CertCreator Issue', response.getResponseText());
  }

  ui.alert('Thank You! ', ' We appreciate the feedback, your issue will be addressed shortly!', ui.ButtonSet.OK);

}

// Function: getUserInput()
// Purpose: the main function in the program, this calls all of helper functions
function getUserInput(indiv){


  // asks user what rows of the sheet they'd like to select
  const ui = SlidesApp.getUi();
  var result = ui.prompt("Cert Selection", "What rows would you like to create certificates from? (Input A:B)",  ui.ButtonSet.OK_CANCEL);
  var buttonpress = result.getSelectedButton();


  // splits response string using the delimiter ":" & split f'n, returns array response[]
  var response = result.getResponseText().split(":");


  // checks to see if delimiter exists/string was split
  // THIS WONT CATCH ERRORS OF TYPE "asdsd:asdij" as it only checks to see if the string was split so as long as the string contained ":" it will be seen as valid input
  if(response[0] == result.getResponseText()){
    ui.alert('Incorrect Format', ' Please enter valid rows in the format of "A:B", A being the bottom row of the sheet, and B being the top row. ', ui.ButtonSet.OK);
    return;
  }


  // checks to see if first number is larger than 2nd
  if(parseInt(response[0]) > parseInt(response[1])){
    ui.alert('Incorrect Format', ' The first number must be smaller than the second number. Try again ', ui.ButtonSet.OK);
    return;
  }

  var firstRow = response[0];
  var secondRow = response[1];

  getdata(parseInt(firstRow), parseInt(secondRow), indiv)


}



// Function: printNames()
// Purpose: a debugging function that prints the array to the console
function printNames(array){

  var i;
  for(i = 0 ; i < array.length(); i++){
    Logger.log(" Name:" + array[i][16] + array[i][17]);
  }

}
