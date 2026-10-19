
function onFormSubmit(e) {
  const responseSheet = e.range.getSheet();
  const mappingSheet = e.source.getSheetByName("Mail Mapping");

  if (responseSheet.getName() !== "Form Responses") return;

  const row = e.range.getRow();

  const EMPLOYEE_NAME = 3;
  const EMPLOYEE_EMAIL = 4;
  const REQUEST_TYPE = 5;
  const PROBLEM_DESCRIPTION = 6;
  const PRIORITY = 7;

  const REQUEST_ID = 2;
  const CREATED_DATE = 8;
  const CREATED_TIME = 9;
  const STATUS = 10;
  const LAST_UPDATE_DATE = 11;
  const LAST_UPDATE_TIME = 12;
  const NOTES = 13;
  const DEPARTMENT = 14;

  const employeeName = responseSheet.getRange(row, EMPLOYEE_NAME).getValue();

  let employeeEmail = responseSheet
    .getRange(row, EMPLOYEE_EMAIL)
    .getValue()
    .toString()
    .trim()
    .toLowerCase();

  const requestType = responseSheet.getRange(row, REQUEST_TYPE).getValue();
  const problemDescription = responseSheet.getRange(row, PROBLEM_DESCRIPTION).getValue();
  const priority = responseSheet.getRange(row, PRIORITY).getValue();

  if (!employeeEmail.endsWith("@company.com")) {
    responseSheet
      .getRange(row, EMPLOYEE_EMAIL)
      .setNote("Invalid company email. Expected @company.com");
  }

  responseSheet.getRange(row, EMPLOYEE_EMAIL).setValue(employeeEmail);

  // Create Request ID
  const lastRow = responseSheet.getLastRow();

  const requestIds = responseSheet
    .getRange(2, REQUEST_ID, lastRow - 1, 1)
    .getValues()
    .flat()
    .filter(id => id !== "" && !isNaN(id));

  const lastRequestId = requestIds.length > 0
    ? Math.max(...requestIds)
    : 1000;

  const requestId = lastRequestId + 1;
  const now = new Date();

  responseSheet.getRange(row, REQUEST_ID).setValue(requestId);
  responseSheet.getRange(row, CREATED_DATE).setValue(now);
  responseSheet.getRange(row, CREATED_TIME).setValue(now);
  responseSheet.getRange(row, STATUS).setValue("Pending");
  responseSheet.getRange(row, LAST_UPDATE_DATE).setValue(now);
  responseSheet.getRange(row, LAST_UPDATE_TIME).setValue(now);

  // Find department and email
  let department = "";
  let departmentEmail = "";

  if (mappingSheet && mappingSheet.getLastRow() > 1) {
    const mappingData = mappingSheet
      .getRange(2, 1, mappingSheet.getLastRow() - 1, 3)
      .getValues();

    for (const item of mappingData) {
      const mappedRequestType = item[0];

      if (
        String(mappedRequestType).trim().toLowerCase() ===
        String(requestType).trim().toLowerCase()
      ) {
        department = item[1];
        departmentEmail = item[2];
        break;
      }
    }
  }

  responseSheet.getRange(row, DEPARTMENT).setValue(department);

  // Calculate Request Age for the new request
  updateRequestAgeForRow(row);

  if (!departmentEmail) {
    responseSheet.getRange(row, NOTES).setValue(
      "No department email configured for this request type."
    );
  } else {
    const subject = `New IT Support Request - #${requestId}`;

    const body =
      `A new IT support request has been submitted.\n\n` +
      `Request ID: ${requestId}\n` +
      `Employee: ${employeeName}\n` +
      `Employee Email: ${employeeEmail}\n` +
      `Request Type: ${requestType}\n` +
      `Department: ${department}\n` +
      `Priority: ${priority}\n\n` +
      `Problem Description:\n${problemDescription}\n\n` +
      `Status: Pending`;

    MailApp.sendEmail({
      to: departmentEmail,
      subject: subject,
      body: body
    });
  }

  // Keep the original notification behavior
  const subject = `New IT Support Request - #${requestId}`;

  const body =
    `A new IT support request has been submitted.\n\n` +
    `Request ID: ${requestId}\n` +
    `Employee: ${employeeName}\n` +
    `Employee Email: ${employeeEmail}\n` +
    `Request Type: ${requestType}\n` +
    `Priority: ${priority}\n\n` +
    `Problem Description:\n${problemDescription}\n\n` +
    `Status: Pending`;

  if (departmentEmail) {
    MailApp.sendEmail({
      to: departmentEmail,
      subject: subject,
      body: body
    });
  }
}



function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  if (sheet.getName() !== "Form Responses") return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row === 1) return;

  const STATUS = 10;
  const LAST_UPDATE_DATE = 11;
  const LAST_UPDATE_TIME = 12;
  const NOTES = 13;
  const RESOLUTION_TIME = 15;
  const REQUEST_AGE = 17; // Column Q

  // Continue only when Status or Notes changes
  if (col !== STATUS && col !== NOTES) return;

  const now = new Date();

  // Update Last Update Date and Time
  sheet.getRange(row, LAST_UPDATE_DATE).setValue(now);
  sheet.getRange(row, LAST_UPDATE_TIME).setValue(now);

  const status = String(
    sheet.getRange(row, STATUS).getValue()
  ).trim();

  if (col === STATUS) {

    if (status === "Resolved") {
      const createdDate = sheet.getRange(row, 8).getValue();
      const createdTime = sheet.getRange(row, 9).getValue();

      const createdDateTime = combineDateAndTime(
        createdDate,
        createdTime
      );

      const resolutionTime = calculateBusinessHours(
        createdDateTime,
        now
      );

      sheet.getRange(row, RESOLUTION_TIME).setValue(resolutionTime);

      // Clear Request Age immediately when Resolved
      sheet.getRange(row, REQUEST_AGE).clearContent();

    } else {
      // Clear Resolution Time if reopened or changed to another status
      sheet.getRange(row, RESOLUTION_TIME).clearContent();

      // Recalculate Request Age
      updateRequestAgeForRow(row);
    }

  } else if (col === NOTES) {

    // Editing Notes must not change the original Resolution Time
    // Keep updating Request Age only for unresolved requests
    if (status !== "Resolved") {
      updateRequestAgeForRow(row);
    }
  }
}

// Combine separate date and time values into one Date object
function combineDateAndTime(dateValue, timeValue) {
  if (
    !(dateValue instanceof Date) ||
    isNaN(dateValue.getTime()) ||
    !(timeValue instanceof Date) ||
    isNaN(timeValue.getTime())
  ) {
    return null;
  }

  return new Date(
    dateValue.getFullYear(),
    dateValue.getMonth(),
    dateValue.getDate(),
    timeValue.getHours(),
    timeValue.getMinutes(),
    timeValue.getSeconds()
  );
}


// Calculate and save Request Age in column Q
function updateRequestAgeForRow(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Form Responses");

  if (!sheet || row < 2) return;

  const REQUEST_AGE = 17; // Column Q
  const STATUS = 10;
  const CREATED_DATE = 8;
  const CREATED_TIME = 9;

  const status = String(
    sheet.getRange(row, STATUS).getValue()
  ).trim();

  const ageCell = sheet.getRange(row, REQUEST_AGE);

  // Resolved requests should have no Request Age
  if (status === "Resolved") {
    ageCell.clearContent();
    return;
  }

  // Calculate age only for open requests
  if (status !== "Pending" && status !== "In Progress") {
    ageCell.clearContent();
    return;
  }

  const createdDate = sheet.getRange(row, CREATED_DATE).getValue();
  const createdTime = sheet.getRange(row, CREATED_TIME).getValue();

  const createdDateTime = combineDateAndTime(
    createdDate,
    createdTime
  );

  if (!createdDateTime) {
    ageCell.clearContent();
    return;
  }

  const age = calculateBusinessHours(
    createdDateTime,
    new Date()
  );

  ageCell.setValue(age);
}


// Update Request Age for all existing requests
function updateRequestAges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Form Responses");

  if (!sheet) return;

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  for (let row = 2; row <= lastRow; row++) {
    updateRequestAgeForRow(row);
  }
}


// Run this function once to create the hourly trigger
function setupRequestAgeHourlyTrigger() {
  const functionName = "updateRequestAges";

  // Avoid creating duplicate hourly triggers
  const triggers = ScriptApp.getProjectTriggers();

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyHours(1)
    .create();
}


// Calculate business hours, excluding Friday, Saturday, and listed holidays
function calculateBusinessHours(startDate, endDate) {
  if (
    !(startDate instanceof Date) ||
    isNaN(startDate.getTime()) ||
    !(endDate instanceof Date) ||
    isNaN(endDate.getTime())
  ) {
    return 0;
  }

  if (endDate <= startDate) return 0;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const holidaySheet = ss.getSheetByName("Holidays");
  const timezone = ss.getSpreadsheetTimeZone();

  let holidays = [];

  if (holidaySheet && holidaySheet.getLastRow() > 1) {
    holidays = holidaySheet
      .getRange(2, 1, holidaySheet.getLastRow() - 1, 1)
      .getValues()
      .flat()
      .filter(date => date instanceof Date && !isNaN(date.getTime()))
      .map(date =>
        Utilities.formatDate(date, timezone, "yyyy-MM-dd")
      );
  }

  const WORK_START_HOUR = 9;
  const WORK_END_HOUR = 17;

  let totalHours = 0;
  let currentDate = new Date(startDate);

  while (currentDate < endDate) {
    const dayOfWeek = currentDate.getDay();

    const isWeekend =
      dayOfWeek === 5 || dayOfWeek === 6;

    const dateString = Utilities.formatDate(
      currentDate,
      timezone,
      "yyyy-MM-dd"
    );

    const isHoliday = holidays.includes(dateString);

    if (!isWeekend && !isHoliday) {
      const workStart = new Date(currentDate);
      workStart.setHours(WORK_START_HOUR, 0, 0, 0);

      const workEnd = new Date(currentDate);
      workEnd.setHours(WORK_END_HOUR, 0, 0, 0);

      const effectiveStart =
        currentDate > workStart
          ? currentDate
          : workStart;

      const effectiveEnd =
        endDate < workEnd
          ? endDate
          : workEnd;

      if (effectiveEnd > effectiveStart) {
        totalHours +=
          (effectiveEnd - effectiveStart) /
          (1000 * 60 * 60);
      }
    }

    currentDate = new Date(currentDate);
    currentDate.setDate(currentDate.getDate() + 1);
    currentDate.setHours(0, 0, 0, 0);
  }

  return totalHours;
}


function sendDailyReport() {
  const MANAGER_EMAIL = "manager@company.com";
  const SHEET_NAME = "Form Responses";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) return;

  const data = sheet.getDataRange().getValues();

  if (data.length < 1) return;

  // Exclude Timestamp from the Excel report
  const headers = data[0].slice(1);

  const LAST_UPDATE_DATE = 11;
  const today = new Date();
  const timezone = ss.getSpreadsheetTimeZone();

  const todayString = Utilities.formatDate(
    today,
    timezone,
    "yyyy-MM-dd"
  );

  const updatedRows = data.slice(1).filter(row => {
    const lastUpdateDate = row[LAST_UPDATE_DATE - 1];

    if (!(lastUpdateDate instanceof Date)) return false;

    const lastUpdateDateString = Utilities.formatDate(
      lastUpdateDate,
      timezone,
      "yyyy-MM-dd"
    );

    return lastUpdateDateString === todayString;
  });

  if (updatedRows.length === 0) {
    MailApp.sendEmail({
      to: MANAGER_EMAIL,
      subject: "Daily IT Support Report - No Updates",
      body:
        `No IT support requests were updated today.\n\n` +
        `Date: ${todayString}`
    });

    return;
  }

  const tempSpreadsheet = SpreadsheetApp.create(
    `IT Support Daily Report - ${todayString}`
  );

  const tempSheet = tempSpreadsheet.getSheets()[0];

  tempSheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  tempSheet
    .getRange(2, 1, updatedRows.length, headers.length)
    .setValues(updatedRows.map(row => row.slice(1)));

  tempSheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight("bold");

  // Created Date = column 7 in the report
  tempSheet
    .getRange(2, 7, updatedRows.length, 1)
    .setNumberFormat("M/d/yyyy");

  // Created Time = column 8 in the report
  tempSheet
    .getRange(2, 8, updatedRows.length, 1)
    .setNumberFormat("h:mm:ss AM/PM");

  // Last Update Date = column 10 in the report
  tempSheet
    .getRange(2, 10, updatedRows.length, 1)
    .setNumberFormat("M/d/yyyy");

  // Last Update Time = column 11 in the report
  tempSheet
    .getRange(2, 11, updatedRows.length, 1)
    .setNumberFormat("h:mm:ss AM/PM");

  tempSheet.autoResizeColumns(1, headers.length);

  SpreadsheetApp.flush();

  const exportUrl =
    `https://docs.google.com/spreadsheets/d/${tempSpreadsheet.getId()}/export?format=xlsx`;

  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    }
  });

  const excelFile = response
    .getBlob()
    .setName(`IT Support Daily Report - ${todayString}.xlsx`);

  MailApp.sendEmail({
    to: MANAGER_EMAIL,
    subject: `Daily IT Support Report - ${todayString}`,
    body:
      `Hello,\n\n` +
      `Please find attached the IT Support Daily Report for ${todayString}.\n\n` +
      `Updated requests: ${updatedRows.length}\n\n` +
      `Regards`,
    attachments: [excelFile]
  });

  DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
}


function recalculateResolutionTimes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Form Responses");

  if (!sheet) {
    throw new Error('Sheet "Form Responses" not found.');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  // Column numbers
  const CREATED_DATE = 8;
  const CREATED_TIME = 9;
  const STATUS = 10;
  const LAST_UPDATE_DATE = 11;
  const LAST_UPDATE_TIME = 12;
  const RESOLUTION_TIME = 15;

  const data = sheet
    .getRange(2, 1, lastRow - 1, RESOLUTION_TIME)
    .getValues();

  for (let i = 0; i < data.length; i++) {
    const row = i + 2;

    const status = String(data[i][STATUS - 1])
      .trim();

    // Process resolved requests only
    if (status !== "Resolved") continue;

    const createdDate = data[i][CREATED_DATE - 1];
    const createdTime = data[i][CREATED_TIME - 1];
    const lastUpdateDate = data[i][LAST_UPDATE_DATE - 1];
    const lastUpdateTime = data[i][LAST_UPDATE_TIME - 1];

    // Skip rows with missing or invalid dates/times
    if (
      !(createdDate instanceof Date) ||
      isNaN(createdDate.getTime()) ||
      !(createdTime instanceof Date) ||
      isNaN(createdTime.getTime()) ||
      !(lastUpdateDate instanceof Date) ||
      isNaN(lastUpdateDate.getTime()) ||
      !(lastUpdateTime instanceof Date) ||
      isNaN(lastUpdateTime.getTime())
    ) {
      continue;
    }

    const startDateTime = new Date(
      createdDate.getFullYear(),
      createdDate.getMonth(),
      createdDate.getDate(),
      createdTime.getHours(),
      createdTime.getMinutes(),
      createdTime.getSeconds()
    );

    const endDateTime = new Date(
      lastUpdateDate.getFullYear(),
      lastUpdateDate.getMonth(),
      lastUpdateDate.getDate(),
      lastUpdateTime.getHours(),
      lastUpdateTime.getMinutes(),
      lastUpdateTime.getSeconds()
    );

    const resolutionTime = calculateBusinessHours(
      startDateTime,
      endDateTime
    );

    sheet
      .getRange(row, RESOLUTION_TIME)
      .setValue(resolutionTime);
  }
}

