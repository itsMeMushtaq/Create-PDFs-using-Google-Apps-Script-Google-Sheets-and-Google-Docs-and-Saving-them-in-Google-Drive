function createBulkPDFs() {

  const sheetID = "GOOGLE_SHEET_ID_HERE";
  const templateID = "GOOGLE_TEMPLATE_DOCUMENT_ID";
  const tempFolderID = "TEMPORARY_GOOGLE_DRIVE_FOLDER";
  const pdfFolderID = "PDF_DOCS_GOOGLE_DRIVE_FOLDER";

  const tempFolder = DriveApp.getFolderById(tempFolderID);
  const pdfFolder = DriveApp.getFolderById(pdfFolderID);
  const docFile = DriveApp.getFileById(templateID);

  const workSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('People');
  const data = workSheet.getDataRange().getDisplayValues().slice(1);

  var pdfLink = "";
  let errors = [];  

  data.forEach(row => {
    try {
      pdfLink = createPDF(row[0], row[1], row[3], row[0] + " " + row[1], docFile, tempFolder, pdfFolder);
      //console.log(pdfLink);
      errors.push(pdfLink);
    } catch { errors.push(["PDF Failed"]); }
  }) // end of forEach

  for (var i=0; i<workSheet.getLastRow()-1; i++) { workSheet.getRange(i+2,5).setValue(errors[i]); }
  
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
  const pdf = pdfFolder.createFile(pdfBlob).setName(pdfName).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var pdfLink = "https://drive.google.com/uc?export=view&id=" + pdf.getId();
  tempFile.setTrashed(true); return pdfLink;

} // end of do createPDF()
