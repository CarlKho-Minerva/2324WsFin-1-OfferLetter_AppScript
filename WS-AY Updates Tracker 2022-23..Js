var fieldPositions = {
  studentEmail: 2,
  studentFullName: 3,
  citizenship: 4,
  wsEligible: 6,
  onRotation: 7,
  company: 8,
  role: 9,
  manager: 10,
  paycomDeptCode: 11,
  projectCode: 12,
  ctdSignOff: 13,
  positionUpdateBool: 14,
  paycomID: 15,
  managerEmail: 18,
  avgHours: 19
};

var actionsByColumn = {
  3: "StudentEmail",
  4: "StudentFullName",
  5: "Citizenship",
  7: "Eligibility",
  8: "Rotation",
  9: "Company",
  10: "Role",
  11: "Manager",
  12: "Dept Code",
  13: "Project Code",
  14: "CTD Sign Off",
  15: "Position Update Bool",
  16: "PaycomID",
  19: "Manager Email",
  20: "Average Hours",
  101: "Doc_Unsigned_Offer_Link",
  102: "PDF_Unsigned_Offer_Link",
  103: "HelloSign_Offer_UUID",
  104: "Signed_Offer_Link",
  105: "Onboarding_Email_Sent"
};

// SpreadSheets Names - All of these are constants
  // Just inform Tina and/or Holly to update the content in the spreadsheet itself
var currentDashboardSheetName = "Dashboard";
var updatesSheetName = "Updates";
var projectCodesSheetName = "PaycomProjectCodes";
var checklistSheetName = "Offers Checklist";
var managersSheetName = "PaycomManagersByEmail";
var formConfigSheetName = "Config";

// AGOL constants mostly...
var contractorOfferTemplateFileId = "1Y8G72Toj8Dwn6nmHWJBJ6G8eo1IXT6Pxtnr95DxuH2k";
var standardOfferTemplateFileId = "15DTnmD0Yq6LnOX7ZOgEOg5LGHumwP0piRE-CQ4okd6A";

var salaryRate = 20;
var officialStartDate = "9/1/2022";
var normalEndDate = "4/21/2023";
var seniorEndDate = "5/26/2023";
var currentGraduationYear = 2023;
var raStartDate = "8/27/2022";

var docDraftsFolderId = "16lDd_yyMPLp3VQmaU6eswCockstKnNS2";
var pdfDraftsFolderId = "1aWOoCbzDhFvfNHeqoT0P5xggH2wDP-Tu";
var pdfSignedFolderId = "1r-lrCDmmr-7fJlU2SG_x6B6uVmWvulj5";


function onEdit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  if (sheet.getActiveSheet().getName() == currentDashboardSheetName) {
    // Gets row number of the cell that was modified
    var updatedRow = e.range.getRow(),
        oldValue = e.oldValue;
  
    // Run function to check for the entire sheet row
    var newValues = e.range.getValues()[0],
        firstUpdatedColumn = e.range.getColumn();
    
    Logger.log(newValues, firstUpdatedColumn);

    for (var i = 0; i < newValues.length; i++) {
      var modifiedColumn = i + firstUpdatedColumn,
          newValue = newValues[i];
      
      processUpdate(sheet, updatedRow, modifiedColumn, newValue, oldValue);
    };
  };
};

function processUpdate(sheet, updatedRow, modifiedColumn, newValue, oldValue) {
  // Checks whether the modified cell corresponds to a relevant field.
  // Returns the position as an integer to check whether the modified field matters from our fieldPositions object
  var modifiedField = Object.values(fieldPositions).find(position => {
    return position == modifiedColumn;
  });
  // If a relevant field was modified, proceeds to log the update
  if (!!modifiedField) {   

    handleProjectCodesOnCompanyChange(sheet, updatedRow, modifiedColumn, newValue, oldValue);

    // If there was a change in the value, posts the update in the logs.
    postUpdatedStudentInfoAsUpdate(sheet, updatedRow, modifiedField, newValue);
  }
}


function handleProjectCodesOnCompanyChange(sheet, updatedRow, modifiedColumn, newValue, oldValue) {
  if (modifiedColumn == fieldPositions.company) {
    var dashboardSheet = sheet.getSheetByName(currentDashboardSheetName);
    var impactedProjectCodeCell = dashboardSheet.getRange(updatedRow, fieldPositions.projectCode);

    if (newValue == 'MU') {
      impactedProjectCodeCell.setDataValidation(null);
      impactedProjectCodeCell.setValue("Unassigned");
      postUpdatedStudentInfoAsUpdate(sheet, updatedRow, fieldPositions.projectCode, "Unassigned");
    } else if (newValue == 'MP') {
      var projectCodesSheet = sheet.getSheetByName(projectCodesSheetName);
      var projectCodesList = projectCodesSheet.getRange(2, 2, projectCodesSheet.getLastRow());
      var validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(projectCodesList).build();
      
      impactedProjectCodeCell.setValue("");
      impactedProjectCodeCell.setDataValidation(validationRule);
    } else if (!newValue && !!oldValue) {
      impactedProjectCodeCell.setDataValidation(null);
      impactedProjectCodeCell.setValue("");
    }
  };
};

function postUpdatedStudentInfoAsUpdate(sheet, updatedRow, modifiedField, newValue) {
  Logger.log(newValue);
  if (!!newValue) {
    var updatesSheet = sheet.getSheetByName(updatesSheetName);
    var { studentEmail, studentFullName, paycomID } = getStudentInfoForUpdate(sheet, updatedRow);
    var newRowNum = updatesSheet.getLastRow() + 1
    var actionNum = Object.keys(actionsByColumn).find(code => {
      return modifiedField == code
    });
    var actionCode = actionsByColumn[actionNum];
    
    updatesSheet.getRange(newRowNum, 1).setValue(new Date());
    updatesSheet.getRange(newRowNum, 2).setValue(studentEmail);
    updatesSheet.getRange(newRowNum, 3).setValue(studentFullName);
    updatesSheet.getRange(newRowNum, 4).setValue(paycomID);
    updatesSheet.getRange(newRowNum, 5).setValue(actionCode);
    updatesSheet.getRange(newRowNum, 6).setValue(newValue);
  };
};

function getStudentInfoForUpdate(sheet, updatedRow) {
  var dashboardSheet = sheet.getSheetByName(currentDashboardSheetName);
  // Gets first occurrence since the data is returned as a list of lists with only one elemnt (one sheet row)
  var studentInfo = dashboardSheet.getRange(updatedRow, 1, 1, dashboardSheet.getLastColumn()).getValues()[0];
  Logger.log(studentInfo)
  var studentEmail = studentInfo[2],
      studentFullName = studentInfo[3],
      paycomID = studentInfo[15];

  return { studentEmail, studentFullName, paycomID }
};

function createOfferDocs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = sheet.getSheetByName(checklistSheetName);
  var startRow = 2;
  var dataRange = ss.getRange(startRow, 1, ss.getLastRow(), ss.getLastColumn());
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  
  var today = Utilities.formatDate(new Date(), "GMT-8", "MM/dd/yyyy");
  Logger.log(today)
  
  for (var i = 0; i < data.length; ++i) {
    var studentData = data[i],
        eligibility = studentData[3],
        status = studentData[12],
        offerType = studentData[13],
        docUnsingedOfferLink = studentData[14];

    var hasNoOfferLetterYet = docUnsingedOfferLink == ""

    if (status == "Official" && eligibility == "Yes" && hasNoOfferLetterYet) {      
      if (offerType == 'Standard') {
        generateOfferLetter(sheet, standardOfferTemplateFileId, studentData, i + 4, today);
      } else if (offerType == 'Contractor') {
        generateOfferLetter(sheet, contractorOfferTemplateFileId, studentData, i + 4, today);
      }
    }
  }
}

function generateOfferLetter(sheet, templateID, studentData, rowIndex, today) {
  var newDocID = DriveApp.getFileById(templateID).makeCopy().getId();

  var name = studentData[0],
      expectedGraduationYear = studentData[2].toString(),
      title = studentData[6],
      manager = studentData[7],
      studentId = studentData[9],
      studentsStartDate = title == "RA" ? raStartDate : officialStartDate,
      startDate = Utilities.formatDate(new Date(studentsStartDate), "GMT+3", "MM/dd/yyyy"),
      fileName = name + ' ' + title;
  
  Logger.log(startDate);
  var studentsEndDate = expectedGraduationYear == currentGraduationYear ? seniorEndDate : normalEndDate;
  var endDate = Utilities.formatDate(new Date(studentsEndDate), "GMT+3", "MM/dd/yyyy");
  Logger.log(endDate);

  if(today > startDate) {
    var date = new Date(),
        newStartDate = new Date(date.setTime(date.getTime() + 7*86400000)),
        startDate = Utilities.formatDate(newStartDate, "GMT+3", "MM/dd/yyyy") 
  }
  
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
  postUpdatedStudentInfoAsUpdate(sheet, rowIndex, 101, newDocHyperLink);

  DriveApp.getFileById(newDocID).moveTo(DriveApp.getFolderById(docDraftsFolderId));

  granularConvertPDF(sheet, newDocID, rowIndex);

  SpreadsheetApp.flush();
}

function granularConvertPDF(sheet, newDocID, rowIndex) {          
  var newDoc = DriveApp.getFileById(newDocID);
  var docFolder = DriveApp.getFolderById(pdfDraftsFolderId);
  var docBlob = newDoc.getBlob().getAs('application/pdf');
  var newPDFFile = docFolder.createFile(docBlob);

  newPDFFile.setName(newDoc.getName() + ".pdf");
  
  var newPDFFileID = newPDFFile.getId();
  var newPDFHyperLink = `https://drive.google.com/file/d/${newPDFFileID}/view`;

  // Posts the update in the logs.
  postUpdatedStudentInfoAsUpdate(sheet, rowIndex, 102, newPDFHyperLink);
  SpreadsheetApp.flush();
}


