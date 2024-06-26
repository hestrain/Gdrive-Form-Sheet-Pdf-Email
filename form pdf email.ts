var FOHSUBJECT = "FOH Report"
var FOHBODY = "Attached is the FOH report from "
var SSSUBJECT = "SS Report"
var SSBODY = "Attached is the SS report from event"

function AfterForm(e) {
  const sh = e.range.getSheet(); 
  if (sh.getName() == "FOHREPORT") {
   //if its an foh report

       const infoFOH = e.namedValues;
       createPDFFOH(infoFOH);
       const pdfFileFOH = createPDFFOH(infoFOH);
       const entryRowFOH = e.range.getRow();
       const wsFOH = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("heather pdf links");
       wsFOH.getRange(entryRowFOH, 12).setValue(pdfFileFOH.getUrl());
       wsFOH.getRange(entryRowFOH, 13).setValue(pdfFileFOH.getName());

       sendEmailFOH(pdfFileFOH);
  
  } else if (sh.getName() == "SSREPORT") {
      //if form is an ss report
       const infoSS = e.namedValues;
       createPDFSS(infoSS);
       const pdfFileSS = createPDFSS(infoSS);
       const entryRowSS = e.range.getRow();
       const wsSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("heather pdf links");
       wsSS.getRange(entryRowSS, 5).setValue(pdfFileSS.getUrl());
       wsSS.getRange(entryRowSS, 6).setValue(pdfFileSS.getName());

       var reportDate = infoSS[3];
       var venue = infoSS[2];

       sendEmailSS(pdfFileSS, reportDate, venue);


  }
}

function sendEmailFOH(pdfFileFOH, infoFOH){

  //const venue = infoFOH['Venue'][0];
  //const day = infoFOH['Event Date'][0]

 GmailApp.sendEmail("hestrain@gmail.com", FOHSUBJECT,FOHBODY,{attachments:pdfFileFOH},"FOH Manager");

}

function sendEmailSS(pdfFileSS,venue, reportDate){

 
 GmailApp.sendEmail("hestrain@gmail.com",venue + reportDate + "SS Report",SSBODY,{attachments:pdfFileSS});

}


function createPDFFOH(infoFOH){

  const pdfFolder = DriveApp.getFolderById("1UC14S58yO5EOjI6uhQ85Amifs0CJmrQP");
  const tempFolder = DriveApp.getFolderById("1wapidvdOS2a1U0REuczZlJnz4fyGAry6");
  const templateDoc = DriveApp.getFileById("1H5Zbd8BxHKLQZ7GEWZ8iCj-ComSBViNETK2fx3viIoY");

  const newTempFile = templateDoc.makeCopy(tempFolder);  

  const openDoc = DocumentApp.openById(newTempFile.getId());
  const body = openDoc.getBody();
  body.replaceText("{{Venue}}", infoFOH['Venue'][0]);
  body.replaceText("{{Event Date}}", infoFOH['Event Date'][0]);
  body.replaceText("{{Event Name}}", infoFOH['Event Name'][0]);
  body.replaceText("{{Scheduled Show Time}}", infoFOH['Scheduled Show Time'][0]);
  body.replaceText("{{Show Start Time}}", infoFOH['Show Start Time'][0]);
  body.replaceText("{{Intermission Start Time}}", infoFOH['Intermission Start Time'][0]);
  body.replaceText("{{Intermission End Time}}", infoFOH['Intermission End Time'][0]);
  body.replaceText("{{Show End Time}}", infoFOH['Show End Time'][0]);
  body.replaceText("{{House Count}}", infoFOH['House Count'][0]);
  body.replaceText("{{Weather}}", infoFOH['Weather'][0]);
  body.replaceText("{{FOH Manager}}", infoFOH['FOH Manager'][0]);
  body.replaceText("{{Email Address}}", infoFOH['Email Address'][0]);
  body.replaceText("{{Usher 1}}", infoFOH['Usher 1'][0]);
  body.replaceText("{{Usher 2}}", infoFOH['Usher 2'][0]);
  body.replaceText("{{Usher 3}}", infoFOH['Usher 3'][0]);
  body.replaceText("{{Usher 4}}", infoFOH['Usher 4'][0]);
  body.replaceText("{{Shift Start Time}}", infoFOH['Shift Start time'][0]);
  body.replaceText("{{Shift End Time}}", infoFOH['Shift End time'][0]);
  body.replaceText("{{Meal Break Length}}", infoFOH['Meal Break Length'][0]);
  body.replaceText("{{Notes on Staffing}}", infoFOH['Notes on Staffing'][0]);
  body.replaceText("{{Client Experience}}", infoFOH['Client Experience'][0]);
  body.replaceText("{{Facilities}}", infoFOH['Facilities'][0]);
  body.replaceText("{{Lost and Found}}", infoFOH['Lost and Found'][0]);
  body.replaceText("{{Other}}", infoFOH['Other'][0]);                               

  const header = openDoc.getHeader();
  header.replaceText("{{Venue}}", infoFOH['Venue'][0]);
  header.replaceText("{{Event Date}}", infoFOH['Event Date'][0]);

  openDoc.saveAndClose();

  const blobPDF = newTempFile.getAs(MimeType.PDF);
  const pdfFileFOH = pdfFolder.createFile(blobPDF).setName(infoFOH['Venue'][0] + " " + infoFOH['Event Date'][0] + " FOH Report");

  return pdfFileFOH;
}

function createPDFSS(infoSS) {

  const pdfFolder = DriveApp.getFolderById("1spp8i-yLaLnjxv-ZNi18k_rMoaINkUFv");
  const tempFolder = DriveApp.getFolderById("1wapidvdOS2a1U0REuczZlJnz4fyGAry6");
  const templateDoc = DriveApp.getFileById("1fQvmI3liAYcHKxTiinb4fYHiPcEIs-GclVT_XxcsbxQ");

  const newTempFile = templateDoc.makeCopy(tempFolder);  

  const openDoc = DocumentApp.openById(newTempFile.getId());
  const body = openDoc.getBody();
 body.replaceText("{Venue}", infoSS['Venue'][0]);
  body.replaceText("{Event Date}", infoSS['Event Date'][0]);
  body.replaceText("{Event Name}", infoSS['Event Name'][0]);
  body.replaceText("{Scheduled Show Time}", infoSS['Scheduled Show Time'][0]);
  body.replaceText("{Show Start Time}", infoSS['Show Start Time'][0]);
  body.replaceText("{Intermission Start Time}", infoSS['Intermission Start Time'][0]);
  body.replaceText("{Intermission End Time}", infoSS['Intermission End Time'][0]);
  body.replaceText("{Show End Time}", infoSS['Show End Time'][0]);
  body.replaceText("{House Count}", infoSS['House Count'][0]);
  body.replaceText("{Weather}", infoSS['Weather'][0]);
  body.replaceText("{{SS}}", infoSS['Stage Supervisor'][0]);
  body.replaceText("{Email Address}", infoSS['Email Address'][0]);
  body.replaceText("{{SS Start}}", infoSS['Supervisor Shift Start Time'][0]);
  body.replaceText("{{SS End}}", infoSS['Supervisor Shift End Time'][0]);
  body.replaceText("{{SS End}}", infoSS['Supervisor Meal Break Length'][0]);
  body.replaceText("{HL}", infoSS['Head of Lights'][0]);
  body.replaceText("{HL Start}", infoSS['Head of Lights Shift Start Time'][0]);
  body.replaceText("{HL End}", infoSS['Head of Lights Shift End Time'][0]);
  body.replaceText("{HL Meal}", infoSS['Head of Lights Meal Break Length'][0]);
  body.replaceText("{HOS}", infoSS['Head of Sound'][0])
  body.replaceText("{{HOS Start}}", infoSS['Head of Sound Shift Start Time'][0]);
  body.replaceText("{{HOS End}}", infoSS['Head of Sound Shift End Time'][0]);
  body.replaceText("{{HOS Meal}}", infoSS['Head of Sound Meal Break Length'][0]);
  body.replaceText("{{T1}}", infoSS['Technician 1'][0]);
  body.replaceText("{{T1 Start}}", infoSS['Technician 1 Shift Start Time'][0]);
  body.replaceText("{{T1 End}", infoSS['Technician 1 Shift End Time'][0]);
  body.replaceText("{{T1 Meal}}", infoSS['Technician 1 Meal Break Length'][0]);
  body.replaceText("{{T2}}", infoSS['Technician 2'][0]);
  body.replaceText("{{T2 Start}}", infoSS['Technician 2 Shift Start Time'][0]);
  body.replaceText("{{T2 End}}", infoSS['Technician 2 Shift End Time'][0]);
  body.replaceText("{{T2 Meal}}", infoSS['Technician 2 Meal Break Length'][0]);
  body.replaceText("{{HOV}}}", infoSS['Head Video Technician'][0]);
  body.replaceText("{{HR}}", infoSS['Head Rigger'][0]);
  body.replaceText("{{T3}}", infoSS['Technician 3'][0]);
  body.replaceText("{{T4}}", infoSS['Technician 4'][0]);
  body.replaceText("{Notes on Staffing}", infoSS['Notes on Staffing'][0]);
  body.replaceText("{Technical}", infoSS['Technical'][0]);
  body.replaceText("{Scheduling}", infoSS['Scheduling'][0]);
  body.replaceText("{Client Experience}", infoSS['Client Experience'][0]);
  body.replaceText("{Facilities}", infoSS['Facilities'][0]);
  body.replaceText("{Additional sale or rental items}", infoSS['Additional sale or rental items'][0]);
  body.replaceText("{Lost and Found}", infoSS['Lost and Found'][0]);
  body.replaceText("{Other}", infoSS['Other'][0]);                               


 //not woeking
  const header = openDoc.getHeader();
  header.replaceText("{{Venue}}", infoSS['Venue'][0]);
  header.replaceText("{{Event Date}}", infoSS['Event Date'][0]);

  // Retrieves the headers's container element which is DOCUMENT
  const parent = openDoc.getHeader().getParent();

  for (let i = 0; i < parent.getNumChildren(); i += 1) {
    // Retrieves the child element at the specified child index
    const child = parent.getChild(i);

    // Determine the exact type of a given child element
    const childType = child.getType();

    if (childType === DocumentApp.ElementType.HEADER_SECTION) {
      // Replaces all occurrences of a given text in regex pattern
      child.asHeaderSection().replaceText('{{Venue}}', infoSS['Venue'][0]);
      child.asHeaderSection().replaceText('{{Event Date}}', infoSS['Event Date'][0]);

    }
  }


  openDoc.saveAndClose();

  const blobPDF = newTempFile.getAs(MimeType.PDF);
  const pdfFileSS = pdfFolder.createFile(blobPDF).setName(infoSS['Event Date'][0] + " " + infoSS['Venue'][0]);

  return pdfFileSS;
}
