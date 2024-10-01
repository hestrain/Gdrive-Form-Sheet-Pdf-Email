//------------------------------------
//Welcome to the AppsScript for creating a PDF in google from a google form!
//Everything that is greyed out and has 2 slashes before it is comments explaining how to use this form.


//the function that is triggered when the form is submitted
function AfterForm(e) {


     //e.namedvalues is the information stored in the form.
      const formInfo = e.namedValues;
      //triggers the function that created the PDF & stores it as a variable, pdfReport
      const pdfReport = createReportPDF(formInfo);
      //stores what row the new form submission has gone ito
      const entryRow = e.range.getRow();
      //stores the link to the pdf in the spreadsheet.
      // IMPORTANT replace "Reports" with the name of the tab where you want to store the link to the pdf.
      const sheetName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reports");
      //IMPORTANT change the numbers below to be the correct COLUMNs for not overlapping any of the form info.
      //this one is the URL for the PDF
      sheetName.getRange(entryRow, 17).setValue(pdfReport.getUrl());
      //this one is the name of the PDF
      sheetName.getRange(entryRow, 18).setValue(pdfReport.getName());


//function that creates the PDF. This is triggered in the other function.
function createReportPDF(formInfo){


//IMPORTANT - replace all the values below as described by their individual comments
// replace these garbled letters and numbers with everything that comes after...


//https://drive.google.com/drive/u/0/folders/[insert these letters and numbers below]
//THIS is where you're going to store the PDFs that you create
 const pdfFolder = DriveApp.getFolderById("10gv2G7XyEGv67D_hlnWLLgLB39uWXmq5");


 //https://drive.google.com/drive/u/0/folders/[insert these letters and numbers below]
 //THIS is where you're going to store the temporary documents that are creted from the template
 const tempFolder = DriveApp.getFolderById("1wwJDYfMm2Yem-Gh6SrZHiQInZ2_3WOBq");


 // "https://docs.google.com/spreadsheets/d/[insert these letters and numbers below]"
 //this is a link to the template document
 const templateDoc = DriveApp.getFileById("1Vy4pg41mA2tHQWHcyLM5jrnS-dypMmgWGdd5DMkaH9g");


//creates the temporary google doc based off of the template
 const newTempFile = templateDoc.makeCopy(tempFolder); 


//open the document you just created
 const openDoc = DocumentApp.openById(newTempFile.getId());
 //body is the body of the document
 const body = openDoc.getBody();


 //everything below is replace the placeholder values in the template document with the real values from the form


 //IMPORTANT - change these values to line up with your form and your template
 body.replaceText("{{Show Name}}", formInfo['Show Name'][0]);
 body.replaceText("{{Show Date}}", formInfo['Show Date'][0]);
 body.replaceText("{{Show Number}}", formInfo['Show Number'][0]);
 body.replaceText("{{Doors Open}}", formInfo['Doors Open'][0]);
 body.replaceText("{{Start Time}}", formInfo['Start Time'][0]);
 body.replaceText("{{Intermission Start Time}}", formInfo['Intermission Start Time'][0]);
 body.replaceText("{{Intermission End Time}}", formInfo['Intermission End Time'][0]);
 body.replaceText("{{End Time}}", formInfo['End Time'][0]);
 body.replaceText("{{House Count}}", formInfo['House Count'][0]);
 body.replaceText("{{Weather}}", formInfo['Weather'][0]);
 body.replaceText("{{Total Show Time}}", formInfo['Total Show Time'][0]);
 body.replaceText("{{General Notes}}", formInfo['General Notes'][0]);                              
 body.replaceText("{{Technical Notes}}", formInfo['Technical Notes'][0]);                              
 body.replaceText("{{Audience or FOH Notes}}", formInfo['Audience or FOH Notes'][0]);                              
 body.replaceText("{{Other}}", formInfo['Other'][0]);                              


//this is to replace the information in the Header of the file
//IMPORTANT change these values to match your template and your form
 const header = openDoc.getHeader();
 header.replaceText("{{Show Name}}", formInfo['Show Name'][0]);


//save the doucment and close it
 openDoc.saveAndClose();


//create the pdf type document
 const blobPDF = newTempFile.getAs(MimeType.PDF);
 //set the file title of the pdf
 //IMPORTANT - change the naming convention here to match what you need it to be.
 const pdfReport = pdfFolder.createFile(blobPDF).setName(formInfo['Show Name'][0] + " " + formInfo['Show Date'][0] + " Show Report");


//and we're done!
 return pdfReport;
}}

