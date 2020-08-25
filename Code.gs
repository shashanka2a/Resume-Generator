function createBulkPdfs(){

  const docFile = DriveApp.getFileById("1fnb0n47QPwXu4HB_387KfsfL5bM0kvoRwyOWa2YmB6U");
  const tempFolder = DriveApp.getFolderById("1_34WCmbNF1rIX_ruDW7uBXu5r1YJaDgc"); 
  const pdfFolder = DriveApp.getFolderById("1cINdb-JgZ5MOWLMVf4Z5iniB4xUbtw8d"); 
  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("people")
  
  const data = currentSheet.getRange(2, 1, currentSheet.getLastRow()-1,4).getValues();
  
  let status = [];
  let urls = [];
  
  data.forEach(row => {
        try{
           createPdf(row[0],row[1],row[2],row[0]+" "+row[3],docFile,tempFolder,pdfFolder);
           status.push(["Successfully Created!"]);
           urls.push([pdfFile.getUrl()]);
           sendEmail(row[1],pdfFile);
        }
        catch(err){
           urls.push(["NA"]);
           status.push(["Failed to Create!"]);
        }
  });
  
  currentSheet.getRange(2, 7, currentSheet.getLastRow()-1,1).setValues(status);
  currentSheet.getRange(2, 5, currentSheet.getLastRow()-1,1).setValues(urls);

}


function sendEmail(email,pdfFile){
  GmailApp.sendEmail(email, "Check out this template for your resume!","Attached below", {
    attachments:[pdfFile],
    name:'This is automated email generated as a part of Project'
  })
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
