//Code for creating damage logs of equipment - Google Sheets

//Make a interactive button in Google Sheets that converts line to a Google Doc template

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('AutoFill Docs');
    menu.addItem('Create New Docs', 'createNewGoogleDocs');
    menu.addToUi();
  }
  
  // Create docs
  function createNewGoogleDocs() {
    const destinationFolder = DriveApp.getFolderById('DESTINATION FOLDER');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Answers');
    const rows = sheet.getDataRange().getValues();
  
    rows.forEach(function(row, index) {
      if (index === 0) return;
      if (row[10]) return;
  
      var templateId;
      if (row[5] === 'Norsk') {
        templateId = 'TEMPLATE ID NORWEGIAN';
      } else if (row[5] === 'English') {
        templateId = 'TEMPLATE ID ENGLISH';
      } else {
        return;
      }
  
      const googleDocTemplate = DriveApp.getFileById(templateId);
      const copy = googleDocTemplate.makeCopy(`${row[1]}, ${row[0]} Feedback Responses`, destinationFolder);
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();
      const friendlyDate = new Date(row[0]).toLocaleDateString("en-GB");
  
      // Replace the content based on the language
      
      body.replaceText('{{Timestamp}}', friendlyDate);
  
      if (row[5] === 'Norsk') {
        body.replaceText('{{Field1}}', row[1]);
        body.replaceText('{{Field2}}', row[2]);
        body.replaceText('{{Field3}}', row[3]);
        body.replaceText('{{Field4}}', row[4]);
      } else if (row[5] === 'English') {
        body.replaceText('{{Field6}}', row[6]);
        body.replaceText('{{Field7}}', row[7]);
        body.replaceText('{{Field8}}', row[8]);
        body.replaceText('{{Field9}}', row[9]);
      }
  
      const docContent = body.getText();
      doc.saveAndClose();
  
      const url = doc.getUrl();
      sheet.getRange(index + 1, 11).setValue(url);
  
      const rangeWithDocUrl = sheet.getRange(index + 1, 11);
      const rangeWithPdfUrl = sheet.getRange(index + 1, 12);
      const pdfUrls = convertDocsToPDF(url, "REPLY", index);
      rangeWithPdfUrl.setValues([[pdfUrls]]);
    });
  }
  //Adds a link to a PDF version of docs. Only one we need to save
  
  function convertDocsToPDF(docUrl) {
    var docId = docUrl.match(/[-\w]{25,}/)[0];
    var docFile = DriveApp.getFileById(docId);
    var folder = DriveApp.getFolderById('DRIVE FOLDER'); 
    var pdfFile = folder.createFile(docFile.getAs('application/pdf')); 
    pdfFile.setName(docFile.getName() + ".pdf");
    var pdfUrl = pdfFile.getUrl();
    return pdfUrl;
  }
  