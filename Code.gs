function createBulkPDFs() {

  const sheetID = "GOOGLE_SHEET_ID_HERE";
  const templateID = "GOOGLE_TEMPLATE_DOCUMENT_ID";
  const tempFolderID = "TEMPORARY_GOOGLE_DRIVE_FOLDER";
  const pdfFolderID = "PDF_DOCS_GOOGLE_DRIVE_FOLDER";

  const tempFolder = DriveApp.getFolderById(tempFolderID);
  const pdfFolder = DriveApp.getFolderById(pdfFolderID);
  const docFile = DriveApp.getFileById(templateID);

  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('People');

  const data = ws.getDataRange().getDisplayValues().slice(1);
  let errors = new Array();
  
  data.forEach(row => {
    try {
      createPDF(row[0], row[1], row[3], row[0] + " " + row[1], docFile, tempFolder, pdfFolder);
      errors.push(["PDF Created"]);
    }
    catch { error.push(["PDF Failed"]); }
  }) // end of forEach

  ws.getRange(2,5,ws.getLastRow()-1,1).setValue(errors);

} // end of createBulkPDFs()


function createPDF(firstName, lastName, balanceAmount, pdfName, docFile, tempFolder, pdfFolder) {

  const tempFile = docFile.makeCopy(tempFolder).setName("tempFile");

  const tempDocFile = DocumentApp.openById(tempFile.getId());
  const body = tempDocFile.getBody();

  body.replaceText("{first}", firstName);
  body.replaceText("{last}", lastName);
  body.replaceText("{balance}", balanceAmount);
  tempDocFile.saveAndClose();

  const pdfBlob = tempDocFile.getAs('application/pdf');
  pdfFolder.createFile(pdfBlob).setName(pdfName);
  tempFile.setTrashed(true);

} // end of do createPDF()
