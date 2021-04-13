// Author: Hugh MacWilliams 
// *************************************   CertCreator   *************************************

function createSlide(row, presentationid, index) {

  const pres = SlidesApp.openById(presentationid);
  var slides = pres.getSlides();

  pres.appendSlide(slides[0]);
  var slides = pres.getSlides();

  slides[slides.length - 1].replaceAllText("{{FIRST NAME}} {{LAST NAME}}", row[16] + " " + row[17]);
  slides[slides.length - 1].replaceAllText("Organization",  row[14]);
  slides[slides.length - 1].replaceAllText("Location",  row[15]);
  slides[slides.length - 1].replaceAllText("Serial Number",  row[0]);
  slides[slides.length - 1].replaceAllText("AWARD EARNED Test_scale*",  row[18] + " " + row[22]);

  Logger.log("inserted: " + row[16] + " at index " + slides.length - 1);


}


function createSlides(rows, presentationid){

  var i;

  Logger.log("rows.len: " + rows.length);
  for(i = 0; i < rows.length; i++){

    createSlide(rows[i], presentationid, i + 1);

  }

}

function getdata(bottomrow, toprow){

  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1jM31Y5qKby6MV6XQ9vXEipJs7Xd83s3feap4Hr9xmZU/edit#gid=1676112815');
  
  const rows = ss.getDataRange().getValues();


  const selectedRows = rows.slice(bottomrow, toprow + 1);
  Logger.log(rows);

  createSlides(selectedRows, "1F6X5SfcUrhJ_okFhSijtKp6o5U5-Hf5oBZU2rZR_HwA");




}


function onOpen() {

  const ui = SlidesApp.getUi();
  const menu = ui.createMenu('Cert Creator');
  menu.addItem('Create New Certs', 'getdata');
  menu.addItem('Specify Range', 'getUserInput');
  // menu.addItem('Choose Folders', null);

  menu.addToUi();



}


function getUserInput(){

  const ui = SlidesApp.getUi();
  var result = ui.prompt("Cert Selection", "What rows would you like to create certificates from? (Input A:B)",  ui.ButtonSet.OK_CANCEL);
  var buttonpress = result.getSelectedButton();


  if(buttonpress == ui.Button.OK){
    Logger.log("User selected OK button");
    ui.prompt
  }


  var response = result.getResponseText().split(":");
  var firstRow = response[0];
  var secondRow = response[1];

  // SlidesApp.getUi().alert(typeodf(secondRow));

  getdata(parseInt(firstRow), parseInt(secondRow))




}




