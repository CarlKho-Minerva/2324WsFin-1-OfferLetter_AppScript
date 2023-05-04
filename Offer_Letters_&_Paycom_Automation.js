// IMPORTANT
// Get IDs from files/folders in this GDrive folder: https://drive.google.com/drive/u/0/folders/1MWRfrXhLclP6GF3mJUiR--w4toVsU5rw

var reportDataStartRow = 2; // First row of data to process
var reportDataStartColumn = 2; // First column of data to process

// ignore
var actionsByColumn = {
  101: "Doc_Unsigned_Offer_Link",
  102: "PDF_Unsigned_Offer_Link",
  103: "HelloSign_Offer_UUID",
  104: "Signed_Offer_Link",
  105: "Onboarding_Email_Sent"
};

// All values in this block is from this sheet: https://docs.google.com/spreadsheets/d/1gztYtxhcxddD8xIsiCEK8jnO6URRxZ2DUOdFTZv62EE/edit#gid=1722570541
var summerInternshipSheetId = "1gztYtxhcxddD8xIsiCEK8jnO6URRxZ2DUOdFTZv62EE"
var checklistSheetName = "Offers Checklist";
var dashboard = "2022 Internships"
var updatesSheetName = "Updates";
var courseTestersSheetName = "Faculty Test Class Interns";
var externalsSheetName = "External Students/Partners to Invoice";

// Offer letters drafts folder
    // I will be placing the folder links for future ID references 
    // Changing the folders' nesting or location DOES NOT change the ID
var docDraftsFolderId = "11mBnoi8PgMSxZdq0sIa4xqxw9CD1rLrA";                                    // https://drive.google.com/drive/u/0/folders/11mBnoi8PgMSxZdq0sIa4xqxw9CD1rLrA
var pdfDraftsFolderId = "1026V89neIDCOByXirz4nmEP6WRo39W0C";                                    // https://drive.google.com/drive/u/0/folders/1026V89neIDCOByXirz4nmEP6WRo39W0C
var courseTestersDraftsFolderId = "1MNhMuJuh-qTtyBSWHLWnLRwcM45Tasqv";                          // https://drive.google.com/drive/u/0/folders/1MNhMuJuh-qTtyBSWHLWnLRwcM45Tasqv

var contractorOfferTemplateFileId = "1aKtcrOMxnewCDAiyMguL8KRDFZ63ptZCaVlALB-3gcc";             // https://docs.google.com/document/d/1aKtcrOMxnewCDAiyMguL8KRDFZ63ptZCaVlALB-3gcc/edit
var standardOfferTemplateFileId = "1KcjsDskNrh35-RxEcx8w7bMJslaTXWuPERVNLrtCACQ";               // https://docs.google.com/document/d/1KcjsDskNrh35-RxEcx8w7bMJslaTXWuPERVNLrtCACQ/edit
var stanadrdSFBasedOfferTemplateFileId = "1pcCriW0z7CU0TrMR8MuQgwVSOi1x5GjQclq9HFCa4aQ";        // https://docs.google.com/document/d/1pcCriW0z7CU0TrMR8MuQgwVSOi1x5GjQclq9HFCa4aQ/edit
var courseTestersTemplateFileId = "1jnrBDfIudt9c--Ki5rUE7lRXZ0kS2KZDMHt8jTfLqUs";               // https://docs.google.com/document/d/1jnrBDfIudt9c--Ki5rUE7lRXZ0kS2KZDMHt8jTfLqUs/edit
var stipendOfferTemplateFileId = "1yzZpH1n6EIUj3ZrP5jM2UGRTKpaInSVQj1tzYNXeciM";                // https://docs.google.com/document/d/1yzZpH1n6EIUj3ZrP5jM2UGRTKpaInSVQj1tzYNXeciM/edit

function getSheetDataByFileIdAndSheetName(fileId, sheetName) {
  // Imports data from managers sheet
  var sheet = SpreadsheetApp.openById(fileId).getSheetByName(sheetName);
  var numRows = sheet.getLastRow(); // Number of rows to process
  var numColumns = sheet.getLastColumn(); // Numbers of columns to process
  var dataRange = sheet.getRange(reportDataStartRow, reportDataStartColumn, numRows, numColumns);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();

  return data
}

function createOfferDocs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var checklistData = getSheetDataByFileIdAndSheetName(summerInternshipSheetId, checklistSheetName);
  var dashboardData = getSheetDataByFileIdAndSheetName(summerInternshipSheetId, dashboard);
  var courseTestersData = getSheetDataByFileIdAndSheetName(summerInternshipSheetId, courseTestersSheetName);
  var externalsData = getSheetDataByFileIdAndSheetName(summerInternshipSheetId, externalsSheetName);
  
  var today = Utilities.formatDate(new Date(), "GMT-8", "MM/dd/yyyy")
  
  for (var i = 0; i < checklistData.length; ++i) {
    var studentData = checklistData[i],
        docUnsingedOfferLink = studentData[7];

    var hasNoOfferLetterYet = docUnsingedOfferLink == ""

    if (hasNoOfferLetterYet) {
      var offerType = studentData[6];
      var location = studentData[5];

      var offerInfo = dashboardData[i];

      if (offerType == 'Standard' && location != "San Francisco") {
        generateOfferLetter(sheet, dashboard, standardOfferTemplateFileId, offerInfo, i + 2, today);
      } else if (offerType == 'Standard' && location == "San Francisco") {
        generateOfferLetter(sheet, dashboard, stanadrdSFBasedOfferTemplateFileId, offerInfo, i + 2, today);
      } else if (offerType == 'Contractor') {
        generateOfferLetter(sheet, dashboard, contractorOfferTemplateFileId, offerInfo, i + 2, today);
      } else if (offerType == 'Stipend') {
        generateOfferLetter(sheet, dashboard, stipendOfferTemplateFileId, offerInfo, i + 2, today);
      };
    }
  }

  for (var i = 0; i < externalsData.length; ++i) {
    var externalData = externalsData[i],
        externalOfferType = externalData[15],
        location = externalData[11],
        docUnsingedOfferLink = externalData[20];

    var hasNoOfferLetterYet = docUnsingedOfferLink == "";

    if (hasNoOfferLetterYet) {
      if (externalOfferType == "Standard") {
        if (location == "San Francisco") {
          generateOfferLetter(sheet, externalsSheetName, stanadrdSFBasedOfferTemplateFileId, externalData, i + 2, today);
        } else {
          generateOfferLetter(sheet, externalsSheetName, standardOfferTemplateFileId, externalData, i + 2, today);
        }
      } else if (externalOfferType == "Contractor") {
        generateOfferLetter(sheet, externalsSheetName, contractorOfferTemplateFileId, externalData, i + 2, today);
      }
    }
  }

  for (var i = 0; i < courseTestersData.length; ++i) {
    var courseTesterData = courseTestersData[i],
        docUnsingedOfferLink = courseTesterData[20];

    var hasNoOfferLetterYet = docUnsingedOfferLink == "";

    if (hasNoOfferLetterYet) {
      generateCourseTesterOfferLetter(sheet, courseTestersSheetName, courseTestersTemplateFileId, courseTesterData, i + 2, today);
    }
  }
}

function generateOfferLetter(sheet, sheetName, templateID, studentData, rowIndex, today) {
  var newDocID = DriveApp.getFileById(templateID).makeCopy().getId();

  var name = studentData[1],
      expectedGraduationYear = studentData[2].toString(),
      title = studentData[3],
      manager = studentData[5],
      studentId = studentData[17],
      startDate = new Date(studentData[8]),
      endDate = new Date(studentData[9]),
      salary = studentData[10],
      firstPayDate = Utilities.formatDate(new Date(studentData[18]), "GMT+1", "MM/dd/yyyy"),
      fileName = name + ' ' + title;

  if(new Date() > startDate) {
    var date = new Date();
    let start = date.getDate() - date.getDay();
    if (date.getDay() !== 0) {
      start += 7
    };

    var newStartDate = new Date(date.setDate(start));
    var startDate = Utilities.formatDate(new Date(newStartDate), "GMT+1", "MM/dd/yyyy");

    if(new Date() > endDate) {
      var end = newStartDate.getDate() + 14;
      var endDate = Utilities.formatDate(new Date(date.setDate(end)), "GMT+1", "MM/dd/yyyy");
    } else {
      endDate = Utilities.formatDate(endDate, "GMT+1", "MM/dd/yyyy");
    };
  } else {
    startDate = Utilities.formatDate(startDate, "GMT+1", "MM/dd/yyyy");
    endDate = Utilities.formatDate(endDate, "GMT+1", "MM/dd/yyyy");
  };

  Logger.log(name);
  
  var body = DocumentApp.openById(newDocID).getBody();
  body.replaceText('##TodaysDate##', today)
  body.replaceText('##FullName##', name)
  body.replaceText('##GraduationYear##', expectedGraduationYear.slice(2))
  body.replaceText('##JobTitle##', title)
  body.replaceText('##ManagersName##', manager)
  body.replaceText('##StartDate##', startDate)
  body.replaceText('##EndDate##', endDate)
  body.replaceText('##StudentID##', studentId)
  body.replaceText('##Salary##', salary)
  body.replaceText('##FirstPayDate##', firstPayDate)
  
  DriveApp.getFileById(newDocID).setName(fileName);
  DocumentApp.openById(newDocID).saveAndClose();

  var newDocHyperLink = `https://docs.google.com/document/d/${newDocID}/edit`

  // Posts the update in the logs.
  postUpdatedStudentInfoAsUpdate(sheet, sheetName, rowIndex, 101, newDocHyperLink);

  DriveApp.getFileById(newDocID).moveTo(DriveApp.getFolderById(docDraftsFolderId));

  granularConvertPDF(sheet, sheetName, newDocID, rowIndex);

  SpreadsheetApp.flush();
}

function generateCourseTesterOfferLetter(sheet, sheetName, templateID, studentData, rowIndex, today) {
  var newDocID = DriveApp.getFileById(templateID).makeCopy().getId();

  var name = studentData[1],
      expectedGraduationYear = studentData[2].toString(),
      title = studentData[3],
      manager = studentData[5],
      studentId = studentData[15],
      startDate = studentData[8],
      endDate = studentData[9],
      fileName = name + ' ' + title;

  // if(today > startDate) {
  //   var date = new Date(),
  //       newStartDate = new Date(date.setTime(date.getTime() + 7*86400000)),
  //       startDate = Utilities.formatDate(newStartDate, "GMT+1", "MM/dd/yyyy") 
  // }
  
  var body = DocumentApp.openById(newDocID).getBody();
  body.replaceText('##TodaysDate##', today)
  body.replaceText('##FullName##', name)
  body.replaceText('##GraduationYear##', expectedGraduationYear.slice(2))
  body.replaceText('##JobTitle##', title)
  body.replaceText('##ManagersName##', manager)
  body.replaceText('##StartDate##', startDate)
  body.replaceText('##EndDate##', endDate)
  body.replaceText('##StudentID##', studentId)
  
  DriveApp.getFileById(newDocID).setName(fileName);
  DocumentApp.openById(newDocID).saveAndClose();

  var newDocHyperLink = `https://docs.google.com/document/d/${newDocID}/edit`

  // Posts the update in the logs.
  postUpdatedStudentInfoAsUpdate(sheet, sheetName, rowIndex, 101, newDocHyperLink);

  DriveApp.getFileById(newDocID).moveTo(DriveApp.getFolderById(courseTestersDraftsFolderId));

  granularConvertPDF(sheet, sheetName, newDocID, rowIndex);

  SpreadsheetApp.flush();
}

function granularConvertPDF(sheet, sheetName, newDocID, rowIndex) {          
  var newDoc = DriveApp.getFileById(newDocID);
  var docFolder = DriveApp.getFolderById(pdfDraftsFolderId);
  var docBlob = newDoc.getBlob().getAs('application/pdf');
  var newPDFFile = docFolder.createFile(docBlob);

  newPDFFile.setName(newDoc.getName() + ".pdf");
  
  var newPDFFileID = newPDFFile.getId();
  var newPDFHyperLink = `https://drive.google.com/file/d/${newPDFFileID}/view`;

  // Posts the update in the logs.
  postUpdatedStudentInfoAsUpdate(sheet, sheetName, rowIndex, 102, newPDFHyperLink);
  SpreadsheetApp.flush();
}

function postUpdatedStudentInfoAsUpdate(sheet, sheetName, updatedRow, modifiedField, newValue) {
  if (!!newValue) {
    var updatesSheet = sheet.getSheetByName(updatesSheetName);
    var { studentEmail, studentFullName, paycomID, positionTitle } = getStudentInfoForUpdate(sheet, sheetName, updatedRow);
    var newRowNum = updatesSheet.getLastRow() + 1
    var actionNum = Object.keys(actionsByColumn).find(code => {
      return modifiedField == code
    });
    var actionCode = actionsByColumn[actionNum];
    
    updatesSheet.getRange(newRowNum, 1).setValue(new Date());
    updatesSheet.getRange(newRowNum, 2).setValue(studentEmail);
    updatesSheet.getRange(newRowNum, 3).setValue(studentFullName);
    updatesSheet.getRange(newRowNum, 4).setValue(paycomID);
    updatesSheet.getRange(newRowNum, 5).setValue(positionTitle);
    updatesSheet.getRange(newRowNum, 6).setValue(actionCode);
    updatesSheet.getRange(newRowNum, 7).setValue(newValue);
  };
};

function getStudentInfoForUpdate(sheet, sheetName, updatedRow) {
  var contentSheet = sheet.getSheetByName(sheetName);
  // Gets first occurrence since the data is returned as a list of lists with only one elemnt (one sheet row)
  var studentInfo = contentSheet.getRange(updatedRow, 1, 1, contentSheet.getLastColumn()).getValues()[0];
  var studentEmail = studentInfo[1],
      studentFullName = studentInfo[2],
      positionTitle = studentInfo[4],
      paycomID = studentInfo[17];

  return { studentEmail, studentFullName, paycomID, positionTitle }
};

function testNewDate() {
  var date = new Date();
  let start = date.getDate() - date.getDay();
  if (date.getDay() !== 0) {
    start += 7
  }

  var newStartDate = new Date(date.setDate(start));

  Logger.log(start);
  var startDate = Utilities.formatDate(new Date(newStartDate), "GMT+1", "MM/dd/yyyy");
  Logger.log(startDate);

  var end = newStartDate.getDate() + 14;

  // var startDate = Utilities.formatDate(new Date(date.setDate(start)), "GMT+1", "MM/dd/yyyy")
  var endDate = Utilities.formatDate(new Date(date.setDate(end)), "GMT+1", "MM/dd/yyyy")

  Logger.log(startDate);
  Logger.log(endDate);
}
























