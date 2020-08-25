function createBulkPdfs(){
  //Excel Sheet   openById("1hgh5JL75WNF01iDclwwuak3NnX9tcHzFE2dMPQuLSYw")
  const docFile = DriveApp.getFileById("1fnb0n47QPwXu4HB_387KfsfL5bM0kvoRwyOWa2YmB6U");
  const tempFolder = DriveApp.getFolderById("1_34WCmbNF1rIX_ruDW7uBXu5r1YJaDgc"); 
  const pdfFolder = DriveApp.getFolderById("1cINdb-JgZ5MOWLMVf4Z5iniB4xUbtw8d"); 
  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("people")
  
  const data = currentSheet.getRange(2, 1, currentSheet.getLastRow()-1,4).getValues();
  
  let status = [];
  
  data.forEach(row => {
        try{
           createPdf(row[0],row[1],row[2],row[0]+" "+row[3],docFile,tempFolder,pdfFolder);
           status.push(["Successfully Created!"]);
        }
        catch(err){
           status.push(["Failed to Create!"]);
        }
  });
  currentSheet.getRange(2, 7, currentSheet.getLastRow()-1,1).setValues(status);

}

function createPdf(name,email,phone,pdfName,docFile,tempFolder,pdfFolder) {
  const tempFile = docFile.makeCopy(tempFolder); 
  const tempDocFile = DocumentApp.openById(tempFile.getId());
  const body = tempDocFile.getBody(); 
  body.replaceText("{Name}", name);
  body.replaceText("{Email}", email);
  body.replaceText("{Phone}", phone);
  tempDocFile.saveAndClose();
  const pdfBlob = tempFile.getAs(MimeType.PDF);
  pdfFolder.createFile(pdfBlob).setName(pdfName);
  tempFolder.removeFile(tempFile)
}
