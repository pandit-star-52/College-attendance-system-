// ==========================================
// ⚙️ API CONFIGURATION SETTINGS
// ==========================================
// Yahan apne Meta Developer Portal se mile details dalein
const WA_TOKEN = 'YOUR_META_WHATSAPP_TOKEN_HERE'; 
const WA_PHONE_ID = 'YOUR_PHONE_NUMBER_ID_HERE';
// ==========================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Advanced School Attendance Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. Teacher Login Verification
function verifyLogin(credentials) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Teachers");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] == credentials || data[i][2] == credentials) { 
      return { success: true, teacherName: data[i][0], assignedClass: data[i][3] };
    }
  }
  return { success: false, message: "Invalid Email or Phone Number" };
}

// 2. Fetch Attendance Data
function getAttendanceData(className, selectedDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentSheet = ss.getSheetByName("Students");
  var attendanceSheet = ss.getSheetByName("Attendance");
  
  var studentData = studentSheet.getDataRange().getValues();
  var students = [];
  for (var i = 1; i < studentData.length; i++) {
    if (studentData[i][2] == className) {
      students.push({ rollNo: studentData[i][0], name: studentData[i][1], phone: studentData[i][3], status: "Present" });
    }
  }

  var attData = attendanceSheet.getDataRange().getValues();
  var existingRecords = {}; 
  for (var j = 1; j < attData.length; j++) {
    if (attData[j][0] == selectedDate && attData[j][1] == className) {
      existingRecords[attData[j][2]] = attData[j][4]; 
    }
  }

  for (var k = 0; k < students.length; k++) {
    if (existingRecords[students[k].rollNo]) {
      students[k].status = existingRecords[students[k].rollNo]; 
    }
  }
  return { students: students, isUpdate: Object.keys(existingRecords).length > 0 };
}

// 3. Save Attendance & TRIGGER WHATSAPP API for Absentees
function saveOrUpdateAttendance(attendanceData, className, selectedDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  var data = sheet.getDataRange().getValues();
  var rowMap = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == selectedDate && data[i][1] == className) {
      rowMap[data[i][2]] = i + 1; 
    }
  }

  attendanceData.forEach(function(student) {
    if (rowMap[student.rollNo]) {
      // Update Mode (Edit)
      sheet.getRange(rowMap[student.rollNo], 5).setValue(student.status); 
    } else {
      // Insert Mode (New)
      sheet.appendRow([selectedDate, className, student.rollNo, student.name, student.status]);
      
      // 👉 WHATSAPP API CALL: Send message only if student is Absent
      if(student.status === "Absent") {
        sendWhatsAppMessage(student.phone, "absent", student.name, selectedDate);
      }
    }
  });
  return "Attendance Saved! Absent alerts sent via WhatsApp.";
}

// 4. Manage Students & Holidays
function addNewStudent(rollNo, name, className, phone) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students").appendRow([rollNo, name, className, phone]);
  return "Student added successfully!";
}

function getHolidays() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Holidays");
  var data = sheet.getDataRange().getValues();
  var holidays = [];
  for (var i = 1; i < data.length; i++) {
    var rawDate = new Date(data[i][0]);
    var d = new Date(rawDate.getTime() - (rawDate.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
    holidays.push({ date: d, name: data[i][1] });
  }
  return holidays;
}

function addHoliday(date, name) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Holidays").appendRow([date, name]);
  return "Holiday Added!";
}

function deleteHoliday(date) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Holidays");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rawDate = new Date(data[i][0]);
    var d = new Date(rawDate.getTime() - (rawDate.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
    if (d === date) { sheet.deleteRow(i + 1); return "Holiday Removed!"; }
  }
}

// ==========================================
// 🚀 WHATSAPP API INTEGRATION MODULE
// ==========================================

function sendWhatsAppMessage(phoneNumber, messageType, param1, param2) {
  if(!phoneNumber) return;

  // Format phone number (Ensure it has country code, e.g., 91 for India)
  var phoneStr = phoneNumber.toString().replace(/\D/g, ''); 
  if(phoneStr.length === 10) phoneStr = '91' + phoneStr;

  var url = 'https://graph.facebook.com/v17.0/' + WA_PHONE_ID + '/messages';
  var payload = {
    "messaging_product": "whatsapp",
    "to": phoneStr,
    "type": "template"
  };

  // Logic for Absent Alert Template
  if (messageType === "absent") {
    payload.template = {
      "name": "absent_alert", // Apne Meta dashboard me is naam ka template banayein
      "language": { "code": "hi" }, // Hindi language template
      "components": [{
        "type": "body",
        "parameters": [
          { "type": "text", "text": param1 }, // Student Name
          { "type": "text", "text": param2 }  // Date
        ]
      }]
    };
  } 
  // Logic for Holiday Alert Template
  else if (messageType === "holiday") {
    payload.template = {
      "name": "holiday_alert", // Apne Meta dashboard me is naam ka template banayein
      "language": { "code": "hi" },
      "components": [{
        "type": "body",
        "parameters": [
          { "type": "text", "text": param1 } // Holiday Name
        ]
      }]
    };
  }

  var options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + WA_TOKEN,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    Logger.log(response.getContentText());
  } catch (e) {
    Logger.log("WhatsApp API Error: " + e.message);
  }
}

// 5. TRIGGER FUNCTION: Send Holiday Alerts 1 Day Before
function sendTomorrowHolidayAlerts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var holidays = ss.getSheetByName("Holidays").getDataRange().getValues();
  
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  var tomorrowStr = new Date(tomorrow.getTime() - (tomorrow.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
  
  var holidayName = "";
  for(var i=1; i<holidays.length; i++){
    var rawDate = new Date(holidays[i][0]);
    var hDate = new Date(rawDate.getTime() - (rawDate.getTimezoneOffset() * 60000)).toISOString().split('T')[0];
    if(hDate === tomorrowStr) { holidayName = holidays[i][1]; break; }
  }
  
  // 👉 WHATSAPP API CALL: If tomorrow is a holiday, loop through students and send messages
  if(holidayName !== "") {
    var students = ss.getSheetByName("Students").getDataRange().getValues();
    for(var j=1; j<students.length; j++){
      var phone = students[j][3]; 
      if(phone) {
         sendWhatsAppMessage(phone, "holiday", holidayName, "");
      }
    }
  }
}

// 6. PDF Auto-Generator API Logic
function generateMonthlyPDF(className, monthName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var newSheetName = className + "_Report_" + new Date().getTime(); 
  var newSheet = ss.insertSheet(newSheetName);
  newSheet.appendRow(["Report for " + className, "Date generated: " + new Date().toLocaleDateString()]);
  var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=pdf&gid=" + newSheet.getSheetId();
  return url; 
}