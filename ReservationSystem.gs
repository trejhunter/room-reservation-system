// Room Reservation System
// TJ Hunter, 2020

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();

// Calendars to output appointments to
var calAmphitheater = CalendarApp.getCalendarById('YOUR CALENDAR ID');
var calMainRec = CalendarApp.getCalendarById(
  'YOUR OTHER CALENDAR ID (OPTIONAL)'
);
var calStudyRoom = CalendarApp.getCalendarById(
  'YOUR OTHER CALENDAR ID (OPTIONAL)'
);
var calHours = CalendarApp.getCalendarById('YOUR OTHER CALENDAR ID (OPTIONAL)');
var calHolidays = CalendarApp.getCalendarById('YOUR EMAIL');

// Object from user submission
function Submission(row) {
  //var row = lastRow;
  this.timestamp = sheet.getRange(row, 1).getValue();
  this.firstname = sheet.getRange(row, 2).getValue();
  this.lastname = sheet.getRange(row, 3).getValue();
  this.email = sheet.getRange(row, 4).getValue();
  this.reason = sheet.getRange(row, 5).getValue();
  this.room = sheet.getRange(row, 6).getValue();

  // Split up the selected rooms and convert them to numbers
  var rooms = this.room.split(', ');
  var room1 = rooms[0];

  if (rooms[1] != null) {
    room2 = rooms[1];
  }
  if (rooms[2] != null) {
    room3 = rooms[2];
  }

  for (var i = 0; i < rooms.length; i++) {
    if (rooms[i] == 'Main Rec Center') {
      rooms[i] = 'MainRec';
    } else if (rooms[i] == 'Study Room/Conference Room') {
      rooms[i] = 'StudyRoom';
    }
  }

  this.date = sheet.getRange(row, 7).getValue();
  this.time = sheet.getRange(row, 8).getValue();
  this.duration = sheet.getRange(row, 9).getValue();
  this.organization = sheet.getRange(row, 10).getValue();
  this.idNumber = sheet.getRange(row, 11).getValue();

  //Info not from spreadsheet.
  this.eventTitle1 = '(' + room1 + ') - ' + this.organization;
  if (rooms[1] != null) {
    this.eventTitle2 = '(' + room2 + ') - ' + this.organization;
  }
  if (rooms[2] != null) {
    this.eventTitle3 = '(' + room3 + ') - ' + this.organization;
  }

  this.roomsRequested = 'Room(s) Requested:';
  this.status;
  this.dateString =
    this.date.getMonth() +
    1 +
    '/' +
    this.date.getDate() +
    '/' +
    this.date.getYear();
  this.timeString = this.time.toLocaleTimeString();
  this.date.setHours(this.time.getHours());
  this.date.setMinutes(this.time.getMinutes());
  //this.test = rooms;

  this.calendar1 = eval('cal' + rooms[0]);
  if (rooms[1] != null) {
    this.calendar2 = eval('cal' + rooms[1]);
  }
  if (rooms[2] != null) {
    this.calendar3 = eval('cal' + rooms[2]);
  }

  return this;
}

// Use the duration to create endTime variable
function getEndTime(request) {
  request.endTime = new Date(request.date);
  switch (request.duration) {
    case '1 hour':
      request.endTime.setMinutes(request.date.getMinutes() + 60);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case '1 hour 30 min':
      request.endTime.setMinutes(request.date.getMinutes() + 90);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case '2 hours':
      request.endTime.setMinutes(request.date.getMinutes() + 120);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case '2 hours 30 min':
      request.endTime.setMinutes(request.date.getMinutes() + 150);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case '3 hours':
      request.endTime.setMinutes(request.date.getMinutes() + 180);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case '3 hours 30 min':
      request.endTime.setMinutes(request.date.getMinutes() + 210);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case '4 hours':
      request.endTime.setMinutes(request.date.getMinutes() + 240);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case '4 hours 30 min':
      request.endTime.setMinutes(request.date.getMinutes() + 270);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case '5 hours':
      request.endTime.setMinutes(request.date.getMinutes() + 300);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case '5 hours 30 min':
      request.endTime.setMinutes(request.date.getMinutes() + 330);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
    case '6 hours':
      request.endTime.setMinutes(request.date.getMinutes() + 360);
      request.endTimeString = request.endTime.toLocaleTimeString();
      break;
  }
}

// Check for appointment conflicts
function getConflicts(request) {
  var conflicts1 = request.calendar1.getEvents(request.date, request.endTime);
  var conflictHours = calHours.getEvents(request.date, request.endTime);
  var conflictHolidays = calHolidays.getEvents(request.date, request.endTime);
  if (
    conflicts1.length < 1 &&
    conflictHours.length < 1 &&
    conflictHolidays.length < 1
  ) {
    //request.status = "Approve";
    request.status = 'New';
  } else {
    request.status = 'Conflict';
    sheet.getRange(lastRow, lastColumn - 1).setValue('Reject');
    sheet.getRange(lastRow, lastColumn).setValue('Sent: Conflict');
  }

  if (request.calendar2 != null) {
    var conflicts2 = request.calendar2.getEvents(request.date, request.endTime);
    if (
      conflicts2.length < 1 &&
      conflictHours.length < 1 &&
      conflictHolidays.length < 1
    ) {
      //request.status = "Approve";
      request.status = 'New';
    } else {
      request.status = 'Conflict';
      sheet.getRange(lastRow, lastColumn - 1).setValue('Reject');
      sheet.getRange(lastRow, lastColumn).setValue('Sent: Conflict');
    }
  }

  if (request.calendar3 != null) {
    var conflicts3 = request.calendar3.getEvents(request.date, request.endTime);
    if (
      conflicts3.length < 1 &&
      conflictHours.length < 1 &&
      conflictHolidays.length < 1
    ) {
      //request.status = "Approve";
      request.status = 'New';
    } else {
      request.status = 'Conflict';
      sheet.getRange(lastRow, lastColumn - 1).setValue('Reject');
      sheet.getRange(lastRow, lastColumn).setValue('Sent: Conflict');
    }
  }
}

function getConflicts2(request) {
  var conflicts1 = request.calendar1.getEvents(request.date, request.endTime);
  var noConflicts = false;
  if (conflicts1.length < 1) {
  } else {
    return true;
  }

  if (request.calendar2 != null) {
    var conflicts2 = request.calendar2.getEvents(request.date, request.endTime);
    if (conflicts2.length < 1) {
    } else {
      return true;
    }
  }

  if (request.calendar3 != null) {
    var conflicts3 = request.calendar3.getEvents(request.date, request.endTime);
    if (conflicts3.length < 1) {
    } else {
      return true;
    }
  }
  return noConflicts;
}

// Send an email
function draftEmail(request) {
  request.buttonLink = 'https://forms.gle/JgFf6gbwpB9Kghjh8';
  request.buttonText = 'New Reservation';
  switch (request.status) {
    case 'New':
      var templateid = '16-jte1gQrKwPEICwMaVwI1xMtHRKUHr0XeyUdzt9FCc'; // this is the template file id. You can find this in the URL of the google document template. For example, if your URL Looks like this: https://docs.google.com/document/d/1SDTSW2JCItWMGkA8cDZGwZdAQa13sSpiYhiH-Kla6VA/edit, THEN the ID would be 1SDTSW2JCItWMKkA8cDZGwZdAQa13sSpiYhiH-Kla6VA
      var FOLDER_NAME = 'Output Folder'; // Enter the name of the folder where you want to save your new documents. for example, this could be "Output Folder".

      var newDoc = DocumentApp.create(
        'Contract for  ' +
          request.organization +
          ' reservation on ' +
          request.dateString
      );
      var file = DriveApp.getFileById(newDoc.getId());
      var folder = DriveApp.getFolderById('19_knuCZzEXwkGOBgoW12BHtARmxPo9Wm'); // INSERT the destination folder ID. Once again, this can be found in the URL once you have the folder openned in your browser
      folder.addFile(file);
      var docid = DriveApp.getFileById(templateid).makeCopy().getId();
      var doc = DocumentApp.openById(docid);
      var body = doc.getActiveSection();
      body.replaceText('%PURPOSE%', request.reason);
      body.replaceText('%START TIME%', request.timeString);
      body.replaceText('%END TIME%', request.endTimeString);
      body.replaceText('%ID NUMBER%', request.idNumber); // To add more auto gen fields, add them below along with the column number
      body.replaceText('%ORGANIZATION NAME%', request.organization);
      body.replaceText('%DATE%', request.dateString);

      appendToDoc(doc, newDoc);

      doc.saveAndClose();
      newDoc.saveAndClose();
      var message =
        'Your reservation request has been received. Please sign and return this ' +
        'form in order for your request to be approved and finalized'; // Customize message
      var emailTo = request.email; // replace with your email
      var subject =
        'Request for ' + request.dateString + ' Reservation Received'; // customize subject
      var pdf = DriveApp.getFileById(newDoc.getId())
        .getAs('application/pdf')
        .getBytes();
      var attach = {
        fileName: 'Reservation Contract.pdf',
        content: pdf,
        mimeType: 'application/pdf',
      }; // customize file name: "Autogenerated template"
      MailApp.sendEmail(emailTo, subject, message, { attachments: [attach] });

      DriveApp.getFileById(docid).setTrashed(true);

      request.subject =
        'Request for ' + request.dateString + ' Reservation Received';
      request.header = 'Request Received';
      request.message =
        'Please verify that the times listed below are correct and then sign and return the form attached in the email above.';

      break;
    case 'Review':
      request.email = 'YOUR EMAIL';
      request.subject =
        'New Request for ' + request.room + ' on ' + request.dateString;
      request.header = 'Request Received';
      request.message = 'A new request needs to be reviewed.';
      request.buttonLink =
        'https://docs.google.com/spreadsheets/d/1EnvM-QqHoZyzgFA0U7Je4dpuwtWcNBEaLKlvRJJXQjg/edit?usp=sharing';
      request.buttonText = 'View Request';
      break;
    case 'Approve':
      request.subject =
        'Confirmation: ' +
        request.room +
        ' Reservation for ' +
        request.dateString;
      request.header = 'Confirmation';
      request.message =
        'Your reservation has been scheduled. Someone will be there to let you in at your requested time. If any issues arise please reach out to YOUR EMAIL.';
      break;
    case 'Conflict':
      request.subject =
        'Conflict with ' +
        request.room +
        ' Reservation for ' +
        request.dateString;
      request.header = 'Conflict';
      request.message =
        'There is a scheduling conflict. The Rec Center is either closed at this time or another event has already been scheduled. Please pick another room or time.';
      request.buttonText = 'Reschedule';
      break;
    case 'Reject':
      request.subject =
        'Update on Reservation Request for ' + request.dateString;
      request.header = 'Reschedule';
      request.message =
        'Unfortunately the requested reservation time does not work. Please pick another room or time.';
      request.buttonText = 'Reschedule';
      break;
  }
}

function updateCalendar(request) {
  var event = request.calendar1.createEvent(
    request.eventTitle1,
    request.date,
    request.endTime
  );

  if (request.calendar2 != null) {
    var event2 = request.calendar2.createEvent(
      request.eventTitle2,
      request.date,
      request.endTime
    );
  }

  if (request.calendar3 != null) {
    var event3 = request.calendar3.createEvent(
      request.eventTitle3,
      request.date,
      request.endTime
    );
  }
}

function sendEmail(request) {
  MailApp.sendEmail({
    to: request.email,
    subject: request.subject,
    htmlBody: makeEmail(request),
  });
}

// ~~~~~~~~~~~~~~~~~~ MAIN ~~~~~~~~~~~~~~~~~~~~~

function main() {
  var request = new Submission(lastRow);
  getEndTime(request);
  getConflicts(request);
  draftEmail(request);
  sendEmail(request);

  if (request.status == 'New') {
    request.status = 'Review';
    draftEmail(request);
    sendEmail(request);
  }
}

function approve() {
  var statusChange = new StatusObject();
  while (true) {
    getChangeIndex(statusChange);
    if (statusChange.index == -1) {
      return;
    } else {
      var request = new Submission(statusChange.row);
      getEndTime(request);
      if (statusChange.status) {
        request.status = statusChange.status;
        if (statusChange.status == 'Approve') {
          if (getConflicts2(request)) {
            request.status = 'Conflict';
            sheet.getRange(statusChange.row, lastColumn - 1).setValue('Reject');
            sheet
              .getRange(statusChange.row, lastColumn)
              .setValue('Sent: Conflict');
          } else {
            updateCalendar(request);
          }
        }
        draftEmail(request);
        sendEmail(request);
      }
    }
  }
}

// Convert data of status and notified columns into array
function StatusObject() {
  this.statusArray = sheet.getRange(1, lastColumn - 1, lastRow, 1).getValues();
  this.notifiedArray = sheet.getRange(1, lastColumn, lastRow, 1).getValues();
  this.statusArray = [].concat.apply([], this.statusArray);
  this.notifiedArray = [].concat.apply([], this.notifiedArray);
}

// Get the index of the row that has had a status change
function getChangeIndex(statusChange) {
  statusChange.index = statusChange.notifiedArray.indexOf('');
  statusChange.row = statusChange.index + 1;
  if (statusChange.index == -1) {
    return;
  } else if (statusChange.statusArray[statusChange.index] != '') {
    statusChange.status = statusChange.statusArray[statusChange.index];
    sheet
      .getRange(statusChange.row, lastColumn)
      .setValue('Sent: ' + statusChange.status);
    statusChange.notifiedArray[statusChange.index] = 'update';
  } else {
    statusChange.status = statusChange.statusArray[statusChange.index];
    statusChange.notifiedArray[statusChange.index] = 'no update';
  }
}

function appendToDoc(src, dst) {
  for (var i = 0; i < src.getNumChildren(); i++) {
    appendElementToDoc(dst, src.getChild(i));
  }
}

function appendElementToDoc(doc, object) {
  var type = object.getType();
  var element = object.copy();
  Logger.log('Element type is ' + type);
  if (type == 'PARAGRAPH') {
    doc.appendParagraph(element);
  } else if (type == 'TABLE') {
    doc.appendTable(element);
  }
}
