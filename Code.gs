// ==========================================
// 🎓 COLLEGE ATTENDANCE SYSTEM - Backend
// ==========================================
// Google Apps Script Backend with improved error handling and validation

// ==========================================
// ⚙️ API CONFIGURATION SETTINGS
// ==========================================
const CONFIG = {
  WA_TOKEN: 'YOUR_META_WHATSAPP_TOKEN_HERE',
  WA_PHONE_ID: 'YOUR_PHONE_NUMBER_ID_HERE',
  SHEET_NAMES: {
    TEACHERS: 'Teachers',
    STUDENTS: 'Students',
    ATTENDANCE: 'Attendance',
    HOLIDAYS: 'Holidays'
  },
  TIMEZONE: 'Asia/Kolkata'
};

// ==========================================
// 📝 UTILITY FUNCTIONS
// ==========================================

/**
 * Get active spreadsheet
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Get sheet by name with error handling
 */
function getSheet(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet '${sheetName}' not found. Please create required sheets.`);
  }
  return sheet;
}

/**
 * Validate email format
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Validate phone number (10 digits)
 */
function isValidPhone(phone) {
  const phoneStr = phone.toString().replace(/\D/g, '');
  return phoneStr.length === 10 || phoneStr.length === 12;
}

/**
 * Format phone number with country code
 */
function formatPhoneNumber(phone) {
  let phoneStr = phone.toString().replace(/\D/g, '');
  if (phoneStr.length === 10) {
    phoneStr = '91' + phoneStr;
  }
  return phoneStr;
}

/**
 * Get current date in YYYY-MM-DD format
 */
function getCurrentDate() {
  const now = new Date();
  const tz = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  return tz;
}

/**
 * Serve the HTML frontend
 */
function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setWidth(1000)
      .setHeight(800)
      .setTitle('Smart College Attendance Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput(`<h2>Error Loading App</h2><p>${error.message}</p>`);
  }
}

// ==========================================
// 🔐 AUTHENTICATION
// ==========================================

/**
 * Verify teacher login with email or phone
 */
function verifyLogin(credentials) {
  try {
    if (!credentials || credentials.trim() === '') {
      return { success: false, message: 'Please enter email or phone number' };
    }

    const sheet = getSheet(CONFIG.SHEET_NAMES.TEACHERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const email = data[i][1] ? data[i][1].toString().trim() : '';
      const phone = data[i][2] ? data[i][2].toString().trim() : '';
      const credTrim = credentials.trim().toLowerCase();

      if (email.toLowerCase() === credTrim || phone === credTrim) {
        return {
          success: true,
          teacherName: data[i][0],
          assignedClass: data[i][3],
          teacherId: i
        };
      }
    }

    return { success: false, message: 'Invalid email or phone number' };
  } catch (error) {
    Logger.log('Login Error: ' + error.message);
    return { success: false, message: 'Login failed: ' + error.message };
  }
}

// ==========================================
// 👨‍🎓 STUDENT MANAGEMENT
// ==========================================

/**
 * Get all students for a specific class
 */
function getAttendanceData(className, selectedDate) {
  try {
    if (!className || !selectedDate) {
      return { success: false, message: 'Class and date are required', students: [] };
    }

    const studentSheet = getSheet(CONFIG.SHEET_NAMES.STUDENTS);
    const attendanceSheet = getSheet(CONFIG.SHEET_NAMES.ATTENDANCE);

    const studentData = studentSheet.getDataRange().getValues();
    const students = [];

    // Fetch all students in the selected class
    for (let i = 1; i < studentData.length; i++) {
      if (studentData[i][0] && studentData[i][2] === className) {
        students.push({
          rollNo: studentData[i][0],
          name: studentData[i][1],
          phone: studentData[i][3],
          status: 'Present'
        });
      }
    }

    // Check for existing attendance records for this date
    const attData = attendanceSheet.getDataRange().getValues();
    const existingRecords = {};

    for (let j = 1; j < attData.length; j++) {
      if (attData[j][0] && attData[j][0].toString() === selectedDate && attData[j][1] === className) {
        existingRecords[attData[j][2]] = attData[j][4];
      }
    }

    // Update students with existing attendance status
    for (let k = 0; k < students.length; k++) {
      if (existingRecords[students[k].rollNo]) {
        students[k].status = existingRecords[students[k].rollNo];
      }
    }

    return {
      success: true,
      students: students,
      isUpdate: Object.keys(existingRecords).length > 0,
      totalStudents: students.length
    };
  } catch (error) {
    Logger.log('Attendance Data Error: ' + error.message);
    return { success: false, message: error.message, students: [] };
  }
}

/**
 * Add a new student
 */
function addNewStudent(rollNo, name, className, phone) {
  try {
    // Validation
    if (!rollNo || !name || !className || !phone) {
      return { success: false, message: 'All fields are required' };
    }

    if (!isValidPhone(phone)) {
      return { success: false, message: 'Invalid phone number. Please enter a 10-digit number' };
    }

    const sheet = getSheet(CONFIG.SHEET_NAMES.STUDENTS);
    const data = sheet.getDataRange().getValues();

    // Check for duplicate roll number
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === rollNo && data[i][2] === className) {
        return { success: false, message: `Roll number ${rollNo} already exists in ${className}` };
      }
    }

    sheet.appendRow([rollNo, name, className, phone]);
    return { success: true, message: 'Student added successfully!' };
  } catch (error) {
    Logger.log('Add Student Error: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * Get all students
 */
function getAllStudents() {
  try {
    const sheet = getSheet(CONFIG.SHEET_NAMES.STUDENTS);
    const data = sheet.getDataRange().getValues();
    const students = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        students.push({
          rollNo: data[i][0],
          name: data[i][1],
          className: data[i][2],
          phone: data[i][3]
        });
      }
    }

    return { success: true, students: students };
  } catch (error) {
    return { success: false, message: error.message, students: [] };
  }
}

/**
 * Delete a student
 */
function deleteStudent(rollNo, className) {
  try {
    if (!rollNo || !className) {
      return { success: false, message: 'Roll number and class are required' };
    }

    const sheet = getSheet(CONFIG.SHEET_NAMES.STUDENTS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === rollNo && data[i][2] === className) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Student deleted successfully!' };
      }
    }

    return { success: false, message: 'Student not found' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ==========================================
// 📋 ATTENDANCE MANAGEMENT
// ==========================================

/**
 * Save or update attendance records
 */
function saveOrUpdateAttendance(attendanceData, className, selectedDate) {
  try {
    if (!attendanceData || attendanceData.length === 0) {
      return { success: false, message: 'No attendance data to save' };
    }

    if (!className || !selectedDate) {
      return { success: false, message: 'Class and date are required' };
    }

    const sheet = getSheet(CONFIG.SHEET_NAMES.ATTENDANCE);
    const data = sheet.getDataRange().getValues();
    const rowMap = {};

    // Map existing rows
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === selectedDate && data[i][1] === className) {
        rowMap[data[i][2]] = i + 1;
      }
    }

    const absentees = [];

    attendanceData.forEach(function(student) {
      if (rowMap[student.rollNo]) {
        // Update existing record
        sheet.getRange(rowMap[student.rollNo], 5).setValue(student.status);
      } else {
        // Insert new record
        sheet.appendRow([selectedDate, className, student.rollNo, student.name, student.status]);
      }

      // Collect absentees for WhatsApp notification
      if (student.status === 'Absent' && student.phone) {
        absentees.push({
          name: student.name,
          phone: student.phone,
          date: selectedDate
        });
      }
    });

    // Send WhatsApp notifications asynchronously
    if (absentees.length > 0) {
      absentees.forEach(function(student) {
        try {
          sendWhatsAppMessage(student.phone, 'absent', student.name, student.date);
        } catch (e) {
          Logger.log('WhatsApp Error for ' + student.name + ': ' + e.message);
        }
      });
    }

    return {
      success: true,
      message: 'Attendance saved! ' + absentees.length + ' absence alerts sent via WhatsApp.'
    };
  } catch (error) {
    Logger.log('Save Attendance Error: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * Get attendance report for a date range
 */
function getAttendanceReport(className, startDate, endDate) {
  try {
    if (!className || !startDate || !endDate) {
      return { success: false, message: 'All parameters required', report: [] };
    }

    const sheet = getSheet(CONFIG.SHEET_NAMES.ATTENDANCE);
    const data = sheet.getDataRange().getValues();
    const report = [];

    for (let i = 1; i < data.length; i++) {
      const recordDate = data[i][0] ? data[i][0].toString() : '';
      const recordClass = data[i][1];

      if (recordClass === className && recordDate >= startDate && recordDate <= endDate) {
        report.push({
          date: recordDate,
          rollNo: data[i][2],
          name: data[i][3],
          status: data[i][4]
        });
      }
    }

    return { success: true, report: report };
  } catch (error) {
    return { success: false, message: error.message, report: [] };
  }
}

// ==========================================
// 🎉 HOLIDAY MANAGEMENT
// ==========================================

/**
 * Get all holidays
 */
function getHolidays() {
  try {
    const sheet = getSheet(CONFIG.SHEET_NAMES.HOLIDAYS);
    const data = sheet.getDataRange().getValues();
    const holidays = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        const rawDate = new Date(data[i][0]);
        const d = Utilities.formatDate(rawDate, CONFIG.TIMEZONE, 'yyyy-MM-dd');
        holidays.push({ date: d, name: data[i][1], id: i });
      }
    }

    return holidays.sort((a, b) => new Date(a.date) - new Date(b.date));
  } catch (error) {
    Logger.log('Get Holidays Error: ' + error.message);
    return [];
  }
}

/**
 * Add a new holiday
 */
function addHoliday(date, name) {
  try {
    if (!date || !name) {
      return { success: false, message: 'Date and holiday name are required' };
    }

    // Validate date format
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      return { success: false, message: 'Invalid date format. Use YYYY-MM-DD' };
    }

    const sheet = getSheet(CONFIG.SHEET_NAMES.HOLIDAYS);
    sheet.appendRow([date, name]);
    return { success: true, message: 'Holiday added successfully!' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Delete a holiday
 */
function deleteHoliday(date) {
  try {
    if (!date) {
      return { success: false, message: 'Date is required' };
    }

    const sheet = getSheet(CONFIG.SHEET_NAMES.HOLIDAYS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        const rawDate = new Date(data[i][0]);
        const d = Utilities.formatDate(rawDate, CONFIG.TIMEZONE, 'yyyy-MM-dd');
        if (d === date) {
          sheet.deleteRow(i + 1);
          return { success: true, message: 'Holiday removed successfully!' };
        }
      }
    }

    return { success: false, message: 'Holiday not found' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Check if today is a holiday
 */
function isTodayHoliday() {
  try {
    const today = getCurrentDate();
    const holidays = getHolidays();
    return holidays.some(h => h.date === today);
  } catch (error) {
    return false;
  }
}

// ==========================================
// 💬 WHATSAPP INTEGRATION
// ==========================================

/**
 * Send WhatsApp message via Meta Cloud API
 */
function sendWhatsAppMessage(phoneNumber, messageType, param1, param2) {
  try {
    if (!phoneNumber || !CONFIG.WA_TOKEN || !CONFIG.WA_PHONE_ID) {
      Logger.log('WhatsApp credentials not configured');
      return false;
    }

    // Validate and format phone number
    if (!isValidPhone(phoneNumber)) {
      Logger.log('Invalid phone number: ' + phoneNumber);
      return false;
    }

    const phoneStr = formatPhoneNumber(phoneNumber);
    const url = 'https://graph.facebook.com/v17.0/' + CONFIG.WA_PHONE_ID + '/messages';

    let payload = {
      messaging_product: 'whatsapp',
      to: phoneStr,
      type: 'template'
    };

    // Build template based on message type
    if (messageType === 'absent') {
      payload.template = {
        name: 'absent_alert',
        language: { code: 'hi' },
        components: [
          {
            type: 'body',
            parameters: [
              { type: 'text', text: param1 },
              { type: 'text', text: param2 }
            ]
          }
        ]
      };
    } else if (messageType === 'holiday') {
      payload.template = {
        name: 'holiday_alert',
        language: { code: 'hi' },
        components: [
          {
            type: 'body',
            parameters: [{ type: 'text', text: param1 }]
          }
        ]
      };
    } else {
      Logger.log('Unknown message type: ' + messageType);
      return false;
    }

    const options = {
      method: 'post',
      headers: {
        Authorization: 'Bearer ' + CONFIG.WA_TOKEN,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (response.getResponseCode() === 200) {
      Logger.log('WhatsApp sent to ' + phoneStr + ': ' + result.messages[0].id);
      return true;
    } else {
      Logger.log('WhatsApp Error: ' + JSON.stringify(result));
      return false;
    }
  } catch (error) {
    Logger.log('WhatsApp Exception: ' + error.message);
    return false;
  }
}

/**
 * Send holiday alerts 1 day before
 * (Setup as time-based trigger in Apps Script)
 */
function sendTomorrowHolidayAlerts() {
  try {
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = Utilities.formatDate(tomorrow, CONFIG.TIMEZONE, 'yyyy-MM-dd');

    const holidays = getHolidays();
    const holidayName = holidays.find(h => h.date === tomorrowStr);

    if (holidayName) {
      const students = getSheet(CONFIG.SHEET_NAMES.STUDENTS).getDataRange().getValues();

      for (let j = 1; j < students.length; j++) {
        const phone = students[j][3];
        if (phone) {
          sendWhatsAppMessage(phone, 'holiday', holidayName.name, '');
        }
      }

      Logger.log('Holiday alerts sent for: ' + holidayName.name);
    }
  } catch (error) {
    Logger.log('Send Holiday Alerts Error: ' + error.message);
  }
}

// ==========================================
// 📊 REPORTING
// ==========================================

/**
 * Generate PDF report for a month
 */
function generateMonthlyPDF(className, monthName) {
  try {
    const ss = getSpreadsheet();
    const newSheetName = className + '_Report_' + new Date().getTime();
    const newSheet = ss.insertSheet(newSheetName);

    newSheet.appendRow([
      'Attendance Report - ' + className,
      'Period: ' + monthName,
      'Generated: ' + new Date().toLocaleDateString()
    ]);

    const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + 
                '/export?format=pdf&gid=' + newSheet.getSheetId();

    return { success: true, url: url };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Get attendance statistics
 */
function getAttendanceStats(className, startDate, endDate) {
  try {
    const report = getAttendanceReport(className, startDate, endDate);

    if (!report.success) {
      return { success: false, message: report.message };
    }

    const stats = {};

    report.report.forEach(record => {
      if (!stats[record.rollNo]) {
        stats[record.rollNo] = { name: record.name, present: 0, absent: 0, total: 0 };
      }
      stats[record.rollNo].total++;
      if (record.status === 'Present') {
        stats[record.rollNo].present++;
      } else {
        stats[record.rollNo].absent++;
      }
    });

    // Calculate percentage
    Object.keys(stats).forEach(rollNo => {
      const s = stats[rollNo];
      s.percentage = s.total > 0 ? ((s.present / s.total) * 100).toFixed(2) : 0;
    });

    return { success: true, stats: stats };
  } catch (error) {
    return { success: false, message: error.message };
  }
}
