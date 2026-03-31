// ==================== CONFIG ====================
var CONFIG = {
  SPREADSHEET_ID: '1yRrEV40jMMQCzycruoj8lIM0NPXTRhxXxUfvsaTO2uM',
  DRIVE_FOLDER_ID: '1kch--KBTd15dXHVIKutMAFy7Hj5ARtBo',
  DEFAULT_ADMIN: { username: 'admin', password: 'admin1234', name: 'ผู้ดูแลระบบ', role: 'Admin' },
  SHEETS: {
    Users:             ['ID', 'Username', 'Password', 'Name', 'Role', 'Status'],
    AttendanceLog:     ['LogID', 'Date', 'Name', 'Time_In', 'Time_Out', 'Task_Report', 'Photo_URL', 'Status'],
    WorkCycles:        ['CycleID', 'UserID', 'Name', 'Start_Date', 'End_Date', 'Required_Work_Days', 'Status'],
    WorkPlans:         ['PlanID', 'Submission_ID', 'CycleID', 'UserID', 'Name', 'Plan_Date', 'Plan_Status', 'Notes', 'Completed_LogID', 'Created_At', 'Submitted_At', 'Approved_At'],
    ScheduleRequests:  ['ReqID', 'CycleID', 'UserID', 'Name', 'Original_Date', 'Requested_Date', 'Reason', 'Status', 'Created_At', 'Decision_At']
  }
};

// ==================== WEB APP ENTRY ====================
function doGet(e) {
  ensureSetup_();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบรายงานการปฏิบัติงาน อุทยานแห่งชาติเอราวัณ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==================== SETUP ====================
function ensureSetup_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  for (var name in CONFIG.SHEETS) {
    ensureSheet_(ss, name, CONFIG.SHEETS[name]);
  }
  // สร้าง Admin เริ่มต้นหากยังไม่มี
  var usersSheet = ss.getSheetByName('Users');
  var data = usersSheet.getDataRange().getValues();
  var hasAdmin = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][4] === 'Admin') { hasAdmin = true; break; }
  }
  if (!hasAdmin) {
    var id = 'U' + new Date().getTime();
    usersSheet.appendRow([id, CONFIG.DEFAULT_ADMIN.username, CONFIG.DEFAULT_ADMIN.password, CONFIG.DEFAULT_ADMIN.name, CONFIG.DEFAULT_ADMIN.role, 'Active']);
  }
}

function ensureSheet_(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return sheet;
  }
  // ตรวจสอบคอลัมน์ครบ
  var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (existingHeaders.indexOf(headers[i]) === -1) {
      var nextCol = (sheet.getLastColumn() || 0) + 1;
      sheet.getRange(1, nextCol).setValue(headers[i]).setFontWeight('bold');
      existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
  }
  return sheet;
}

function initializeApp() {
  ensureSetup_();
  return { success: true };
}

// ==================== AUTH ====================
function login(username, password) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Users');
  if (!sheet) { ensureSetup_(); sheet = ss.getSheetByName('Users'); }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === username && data[i][2] === password) {
      var status = data[i][5] || 'Active';
      if (status === 'Pending') {
        return { success: false, message: 'บัญชีของคุณกำลังรอการอนุมัติจากผู้ดูแลระบบ' };
      }
      if (status === 'Unregistered') {
        return { success: false, message: 'กรุณาลงทะเบียนตั้งรหัสผ่านก่อนเข้าใช้งาน' };
      }
      var token = Utilities.getUuid();
      var userProps = PropertiesService.getUserProperties();
      userProps.setProperty('session_' + token, JSON.stringify({
        id: data[i][0], username: data[i][1], name: data[i][3], role: data[i][4]
      }));
      return { success: true, token: token, user: { id: data[i][0], name: data[i][3], role: data[i][4] } };
    }
  }
  return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

function loginAndGetData(username, password) {
  var res = login(username, password);
  if (!res.success) return res;
  if (res.user.role === 'Admin') {
    res.appData = getAdminAppData(res.token);
  } else {
    res.appData = getUserAppData(res.token);
  }
  return res;
}

function restoreSessionAndGetData(token) {
  var user = getSessionUser(token);
  if (!user) return { valid: false };
  var appData;
  if (user.role === 'Admin') {
    appData = getAdminAppData(token);
  } else {
    appData = getUserAppData(token);
  }
  return { valid: true, user: user, appData: appData };
}

function logout(token) {
  if (token) {
    PropertiesService.getUserProperties().deleteProperty('session_' + token);
  }
  return { success: true };
}

function getSessionUser(token) {
  if (!token) return null;
  var json = PropertiesService.getUserProperties().getProperty('session_' + token);
  if (!json) return null;
  return JSON.parse(json);
}

function validateSession_(token) {
  var user = getSessionUser(token);
  if (!user) throw new Error('SESSION_EXPIRED');
  return user;
}

// ==================== DATA FETCHING ====================
function getUserAppData(token) {
  var user = validateSession_(token);
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // รอบงานปัจจุบัน
  var cycles = getSheetData_(ss, 'WorkCycles');
  var myCycles = cycles.filter(function(c) { return c.UserID === user.id; });
  var activeCycle = myCycles.filter(function(c) { return c.Status === 'Active'; })[0] || null;

  // แผนวันทำงาน
  var plans = getSheetData_(ss, 'WorkPlans');
  var myPlans = plans.filter(function(p) { return p.UserID === user.id; });

  // บันทึกเช็คชื่อ
  var logs = getSheetData_(ss, 'AttendanceLog');
  var myLogs = logs.filter(function(l) { return l.Name === user.name; });

  // คำร้องขอสลับวัน
  var requests = getSheetData_(ss, 'ScheduleRequests');
  var myRequests = requests.filter(function(r) { return r.UserID === user.id; });

  // ตรวจสอบเช็คอินวันนี้
  var today = formatDate_(new Date());
  var todayLog = myLogs.filter(function(l) { return l.Date === today; })[0] || null;

  // วันที่อนุมัติแล้ว
  var approvedDates = myPlans.filter(function(p) { return p.Plan_Status === 'Approved'; }).map(function(p) { return p.Plan_Date; });

  // ตรวจสอบว่ามีแผนรออนุมัติหรือไม่
  var hasPendingPlan = myPlans.some(function(p) { return p.Plan_Status === 'Pending'; });

  return {
    user: user,
    activeCycle: activeCycle,
    allCycles: myCycles,
    plans: myPlans,
    approvedDates: approvedDates,
    hasPendingPlan: hasPendingPlan,
    logs: myLogs,
    todayLog: todayLog,
    requests: myRequests,
    today: today,
    currentTime: formatTime_(new Date()),
    canCheckIn: approvedDates.indexOf(today) !== -1 && !todayLog,
    canCheckOut: todayLog && todayLog.Time_In && !todayLog.Time_Out
  };
}

function getAdminAppData(token) {
  var user = validateSession_(token);
  if (user.role !== 'Admin') throw new Error('ACCESS_DENIED');
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  var users = getSheetData_(ss, 'Users').map(function(u) { delete u.Password; return u; });
  var cycles = getSheetData_(ss, 'WorkCycles');
  var plans = getSheetData_(ss, 'WorkPlans');
  var logs = getSheetData_(ss, 'AttendanceLog');
  var requests = getSheetData_(ss, 'ScheduleRequests');

  var today = formatDate_(new Date());
  var todayLogs = logs.filter(function(l) { return l.Date === today; });

  // กลุ่มแผนรออนุมัติ ตาม Submission_ID
  var pendingSubmissions = {};
  plans.filter(function(p) { return p.Plan_Status === 'Pending'; }).forEach(function(p) {
    var sid = p.Submission_ID;
    if (!pendingSubmissions[sid]) {
      pendingSubmissions[sid] = { submissionId: sid, cycleId: p.CycleID, userId: p.UserID, name: p.Name, submittedAt: p.Submitted_At, plans: [] };
    }
    pendingSubmissions[sid].plans.push(p);
  });

  var pendingRequests = requests.filter(function(r) { return r.Status === 'Pending'; });

  // สถิติรายเดือน
  var monthlyStats = {};
  logs.forEach(function(l) {
    if (l.Date) {
      var month = l.Date.substring(0, 7);
      monthlyStats[month] = (monthlyStats[month] || 0) + 1;
    }
  });

  return {
    users: users,
    cycles: cycles,
    plans: plans,
    logs: logs,
    requests: requests,
    today: today,
    todayAttendanceCount: todayLogs.length,
    activeCyclesCount: cycles.filter(function(c) { return c.Status === 'Active'; }).length,
    pendingRequestsCount: pendingRequests.length,
    pendingSubmissions: Object.values(pendingSubmissions),
    pendingRequests: pendingRequests,
    monthlyStats: monthlyStats
  };
}

// ==================== WORK CYCLES ====================
function createWorkCycle(token, userId, userName, startDate, endDateCustom) {
  var user = validateSession_(token);
  if (user.role !== 'Admin') throw new Error('ACCESS_DENIED');
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('WorkCycles');

  var start = new Date(startDate);
  var end;
  if (endDateCustom) {
    end = new Date(endDateCustom);
  } else {
    end = new Date(start);
    end.setDate(end.getDate() + 89);
  }

  var cycleId = 'CYC' + new Date().getTime();
  sheet.appendRow([cycleId, userId, userName, formatDate_(start), formatDate_(end), 30, 'Active']);
  return { success: true, cycleId: cycleId };
}

// ==================== WORK PLANS ====================
function createWorkPlan(token, cycleId, dates) {
  var user = validateSession_(token);
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // ตรวจสอบว่ามีแผนรออนุมัติอยู่หรือไม่
  var plans = getSheetData_(ss, 'WorkPlans');
  var hasPending = plans.some(function(p) {
    return p.UserID === user.id && p.CycleID === cycleId && p.Plan_Status === 'Pending';
  });
  if (hasPending) {
    return { success: false, message: 'คุณมีแผนวันทำงานที่รออนุมัติอยู่แล้ว ไม่สามารถส่งแผนซ้ำได้' };
  }

  var submissionId = 'SUB' + new Date().getTime();
  var now = formatDateTime_(new Date());
  var sheet = ss.getSheetByName('WorkPlans');

  var rows = [];
  dates.forEach(function(date, idx) {
    var planId = 'PLN' + new Date().getTime() + '' + idx;
    rows.push([planId, submissionId, cycleId, user.id, user.name, date, 'Pending', '', '', now, now, '']);
  });
  if (rows.length > 0) {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  return { success: true, submissionId: submissionId, count: dates.length };
}

function updateWorkPlanApprovalStatus(token, submissionId, newStatus) {
  var user = validateSession_(token);
  if (user.role !== 'Admin') throw new Error('ACCESS_DENIED');
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('WorkPlans');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var subIdIdx = headers.indexOf('Submission_ID');
  var statusIdx = headers.indexOf('Plan_Status');
  var approvedAtIdx = headers.indexOf('Approved_At');
  var now = formatDateTime_(new Date());
  var count = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][subIdIdx] === submissionId) {
      sheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
      sheet.getRange(i + 1, approvedAtIdx + 1).setValue(now);
      count++;
    }
  }
  return { success: true, count: count };
}

// ==================== SCHEDULE CHANGE REQUESTS ====================
function createScheduleChangeRequest(token, cycleId, originalDate, requestedDate, reason) {
  var user = validateSession_(token);
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('ScheduleRequests');
  var reqId = 'REQ' + new Date().getTime();
  var now = formatDateTime_(new Date());
  sheet.appendRow([reqId, cycleId, user.id, user.name, originalDate, requestedDate, reason, 'Pending', now, '']);
  return { success: true, reqId: reqId };
}

function updateScheduleRequestStatus(token, reqId, newStatus) {
  var user = validateSession_(token);
  if (user.role !== 'Admin') throw new Error('ACCESS_DENIED');
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // อัพเดทสถานะคำร้อง
  var reqSheet = ss.getSheetByName('ScheduleRequests');
  var reqData = reqSheet.getDataRange().getValues();
  var reqHeaders = reqData[0];
  var reqIdIdx = reqHeaders.indexOf('ReqID');
  var statusIdx = reqHeaders.indexOf('Status');
  var decisionIdx = reqHeaders.indexOf('Decision_At');
  var origDateIdx = reqHeaders.indexOf('Original_Date');
  var newDateIdx = reqHeaders.indexOf('Requested_Date');
  var userIdIdx = reqHeaders.indexOf('UserID');
  var cycleIdIdx = reqHeaders.indexOf('CycleID');
  var now = formatDateTime_(new Date());

  var request = null;
  for (var i = 1; i < reqData.length; i++) {
    if (reqData[i][reqIdIdx] === reqId) {
      reqSheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
      reqSheet.getRange(i + 1, decisionIdx + 1).setValue(now);
      request = { originalDate: reqData[i][origDateIdx], requestedDate: reqData[i][newDateIdx], userId: reqData[i][userIdIdx], cycleId: reqData[i][cycleIdIdx] };
      break;
    }
  }

  // ถ้าอนุมัติ ให้ปรับแผนวันทำงาน
  if (newStatus === 'Approved' && request) {
    var planSheet = ss.getSheetByName('WorkPlans');
    var planData = planSheet.getDataRange().getValues();
    var planHeaders = planData[0];
    var pUserIdIdx = planHeaders.indexOf('UserID');
    var pCycleIdIdx = planHeaders.indexOf('CycleID');
    var pDateIdx = planHeaders.indexOf('Plan_Date');
    var pStatusIdx = planHeaders.indexOf('Plan_Status');

    for (var j = 1; j < planData.length; j++) {
      if (planData[j][pUserIdIdx] === request.userId && planData[j][pCycleIdIdx] === request.cycleId) {
        // เปลี่ยนวันเดิมเป็น Swapped_Out
        if (planData[j][pDateIdx] === request.originalDate && planData[j][pStatusIdx] === 'Approved') {
          planSheet.getRange(j + 1, pStatusIdx + 1).setValue('Swapped_Out');
        }
      }
    }
    // เพิ่มวันใหม่
    var newPlanId = 'PLN' + new Date().getTime();
    planSheet.appendRow([newPlanId, '', request.cycleId, request.userId, '', request.requestedDate, 'Approved', 'สลับจากวันที่ ' + request.originalDate, '', now, '', now]);
  }

  return { success: true };
}

// ==================== CHECK IN / CHECK OUT ====================
function checkIn(token) {
  var user = validateSession_(token);
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var now = new Date();
  var today = formatDate_(now);
  var currentTime = formatTime_(now);
  var currentHour = parseInt(Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH'));
  var currentMinute = parseInt(Utilities.formatDate(now, Session.getScriptTimeZone(), 'mm'));
  var totalMinutes = currentHour * 60 + currentMinute;

  // ตรวจสอบเวลาเปิดให้ลงชื่อ (08:05 น. เป็นต้นไป)
  if (totalMinutes < 485) {
    var waitMin = 485 - totalMinutes;
    return { success: false, message: 'ระบบเปิดให้ลงเวลาเข้างานตั้งแต่ 08:05 น. กรุณารออีก ' + waitMin + ' นาที' };
  }

  // ตรวจสอบว่าเป็นวันที่อยู่ในแผนที่อนุมัติแล้ว
  var plans = getSheetData_(ss, 'WorkPlans');
  var approvedToday = plans.some(function(p) {
    return p.UserID === user.id && p.Plan_Date === today && p.Plan_Status === 'Approved';
  });
  if (!approvedToday) {
    return { success: false, message: 'วันนี้ไม่อยู่ในแผนวันทำงานที่อนุมัติแล้ว ไม่สามารถเช็คอินได้' };
  }

  // ตรวจสอบว่าเช็คอินแล้วหรือยัง
  var logs = getSheetData_(ss, 'AttendanceLog');
  var alreadyIn = logs.some(function(l) { return l.Name === user.name && l.Date === today; });
  if (alreadyIn) {
    return { success: false, message: 'คุณลงเวลาเข้างานวันนี้แล้ว' };
  }

  // ตรวจสอบสถานะสาย (หลัง 08:15 น. = สาย)
  var isLate = totalMinutes > 495;
  var lateStatus = isLate ? 'Late' : 'On_Time';

  var logId = 'LOG' + now.getTime();
  var timeInDisplay = today + ' ' + currentTime;
  var sheet = ss.getSheetByName('AttendanceLog');
  sheet.appendRow([logId, today, user.name, timeInDisplay, '', '', '', lateStatus]);

  // อัพเดท Completed_LogID ใน WorkPlans
  updatePlanCompletedLog_(ss, user.id, today, logId);

  return { success: true, logId: logId, timeIn: timeInDisplay, lateStatus: lateStatus };
}

function checkOut(token, taskReport, photoDataArray) {
  var user = validateSession_(token);
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var now = new Date();
  var today = formatDate_(now);
  var currentHour = parseInt(Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH'));
  var currentMinute = parseInt(Utilities.formatDate(now, Session.getScriptTimeZone(), 'mm'));
  var totalMinutes = currentHour * 60 + currentMinute;

  // ตรวจสอบเวลาเปิดให้รายงานผล (16:00 น. เป็นต้นไป)
  if (totalMinutes < 960) {
    var waitHr = Math.floor((960 - totalMinutes) / 60);
    var waitMn = (960 - totalMinutes) % 60;
    var waitMsg = waitHr > 0 ? waitHr + ' ชั่วโมง ' : '';
    waitMsg += waitMn > 0 ? waitMn + ' นาที' : '';
    return { success: false, message: 'ระบบเปิดให้รายงานผลตั้งแต่ 16:00 น. กรุณารออีก ' + waitMsg };
  }

  // หาบันทึกเช็คอินวันนี้
  var sheet = ss.getSheetByName('AttendanceLog');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var nameIdx = headers.indexOf('Name');
  var dateIdx = headers.indexOf('Date');
  var timeOutIdx = headers.indexOf('Time_Out');
  var taskIdx = headers.indexOf('Task_Report');
  var photoIdx = headers.indexOf('Photo_URL');
  var statusIdx = headers.indexOf('Status');
  var logRow = -1;

  for (var i = 1; i < data.length; i++) {
    var cellDate = data[i][dateIdx];
    var cellDateStr = (cellDate instanceof Date) ? formatDate_(cellDate) : String(cellDate);
    if (data[i][nameIdx] === user.name && cellDateStr === today && !data[i][timeOutIdx]) {
      logRow = i + 1;
      break;
    }
  }

  if (logRow === -1) {
    return { success: false, message: 'ไม่พบบันทึกเวลาเข้างานวันนี้' };
  }

  // อัปโหลดรูปภาพ
  var folderUrl = '';
  if (photoDataArray && photoDataArray.length > 0) {
    folderUrl = uploadPhotos_(user.name, today, photoDataArray);
  }

  var nowOut = new Date();
  var timeOut = formatDate_(nowOut) + ' ' + formatTime_(nowOut);
  var outHour = parseInt(Utilities.formatDate(nowOut, Session.getScriptTimeZone(), 'HH'));
  var outMinute = parseInt(Utilities.formatDate(nowOut, Session.getScriptTimeZone(), 'mm'));
  var outTotalMin = outHour * 60 + outMinute;
  var isLateReport = outTotalMin > 1020;
  var finalStatus = isLateReport ? 'Late_Report' : 'Completed';

  sheet.getRange(logRow, timeOutIdx + 1).setValue(timeOut);
  sheet.getRange(logRow, taskIdx + 1).setValue(taskReport);
  sheet.getRange(logRow, photoIdx + 1).setValue(folderUrl);
  sheet.getRange(logRow, statusIdx + 1).setValue(finalStatus);

  return { success: true, timeOut: timeOut, photoUrl: folderUrl, reportStatus: finalStatus };
}

// ==================== PHOTO MANAGEMENT ====================
function uploadPhotos_(userName, dateStr, photoDataArray) {
  var rootFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);

  // โฟลเดอร์เดือน (เช่น 2026-03)
  var monthStr = dateStr.substring(0, 7);
  var monthFolder = getOrCreateFolder_(rootFolder, monthStr);

  // โฟลเดอร์วันที่ (เช่น 2026-03-30)
  var dateFolder = getOrCreateFolder_(monthFolder, dateStr);

  // โฟลเดอร์รายผู้ใช้
  var userFolder = getOrCreateFolder_(dateFolder, userName);

  photoDataArray.forEach(function(photoData, index) {
    var blob = decodeBase64ToBlob_(photoData.data, photoData.mimeType, photoData.fileName || (userName + '_' + dateStr + '_' + (index + 1) + '.jpg'));
    userFolder.createFile(blob);
  });

  // พยายามแชร์โฟลเดอร์ (ถ้าไม่มีสิทธิ์ก็ข้ามไป ไม่ crash)
  try {
    userFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // ไม่มีสิทธิ์แชร์ ข้ามไป - Admin ยังเข้าถึง Drive ได้โดยตรง
  }

  return userFolder.getUrl();
}

function getOrCreateFolder_(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

function decodeBase64ToBlob_(base64Data, mimeType, fileName) {
  var decoded = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decoded, mimeType || 'image/jpeg', fileName);
  return blob;
}

// ==================== HELPER FUNCTIONS ====================
function getSheetData_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var timeColumns = ['Time_In', 'Time_Out'];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      if (val instanceof Date) {
        if (val.getFullYear() < 1910 || timeColumns.indexOf(headers[j]) !== -1) {
          obj[headers[j]] = formatTime_(val);
        } else {
          obj[headers[j]] = formatDate_(val);
        }
      } else {
        obj[headers[j]] = val;
      }
    }
    result.push(obj);
  }
  return result;
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatTime_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm:ss');
}

function formatDateTime_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function updatePlanCompletedLog_(ss, userId, date, logId) {
  var sheet = ss.getSheetByName('WorkPlans');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var userIdIdx = headers.indexOf('UserID');
  var dateIdx = headers.indexOf('Plan_Date');
  var statusIdx = headers.indexOf('Plan_Status');
  var logIdIdx = headers.indexOf('Completed_LogID');

  for (var i = 1; i < data.length; i++) {
    if (data[i][userIdIdx] === userId && data[i][dateIdx] === date && data[i][statusIdx] === 'Approved') {
      sheet.getRange(i + 1, logIdIdx + 1).setValue(logId);
      break;
    }
  }
}

// ==================== USER MANAGEMENT & REGISTRATION ====================
function getUsers(token) {
  var user = validateSession_(token);
  if (user.role !== 'Admin') throw new Error('ACCESS_DENIED');
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return getSheetData_(ss, 'Users').map(function(u) { delete u.Password; return u; });
}

function getUnregisteredUsers() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Users');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][5] === 'Unregistered') {
      result.push({ id: data[i][0], name: data[i][3] });
    }
  }
  return result;
}

function registerUser(userId, password) {
  if (!userId || !password) return { success: false, message: 'กรุณากรอกข้อมูลให้ครบ' };
  if (password.length < 4) return { success: false, message: 'รหัสผ่านต้องมีอย่างน้อย 4 ตัวอักษร' };

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf('ID');
  var pwIdx = headers.indexOf('Password');
  var statusIdx = headers.indexOf('Status');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(userId) && data[i][statusIdx] === 'Unregistered') {
      sheet.getRange(i + 1, pwIdx + 1).setValue(password);
      sheet.getRange(i + 1, statusIdx + 1).setValue('Pending');
      return { success: true, message: 'ลงทะเบียนสำเร็จ! กรุณารอผู้ดูแลระบบอนุมัติ' };
    }
  }
  return { success: false, message: 'ไม่พบผู้ใช้หรือลงทะเบียนแล้ว' };
}

function addUser(token, username, name) {
  var user = validateSession_(token);
  if (user.role !== 'Admin') throw new Error('ACCESS_DENIED');
  if (!username || !name) return { success: false, message: 'กรุณากรอก Username และชื่อ' };

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === username) return { success: false, message: 'Username "' + username + '" ถูกใช้งานแล้ว' };
  }

  var id = 'U' + new Date().getTime();
  sheet.appendRow([id, username, '', name, 'User', 'Unregistered']);
  return { success: true, userId: id };
}

function approveUser(token, userId) {
  var user = validateSession_(token);
  if (user.role !== 'Admin') throw new Error('ACCESS_DENIED');
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf('ID');
  var statusIdx = headers.indexOf('Status');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(userId) && data[i][statusIdx] === 'Pending') {
      sheet.getRange(i + 1, statusIdx + 1).setValue('Active');
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบผู้ใช้หรือสถานะไม่ถูกต้อง' };
}

function rejectUser(token, userId) {
  var user = validateSession_(token);
  if (user.role !== 'Admin') throw new Error('ACCESS_DENIED');
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf('ID');
  var statusIdx = headers.indexOf('Status');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(userId) && data[i][statusIdx] === 'Pending') {
      sheet.getRange(i + 1, headers.indexOf('Password') + 1).setValue('');
      sheet.getRange(i + 1, statusIdx + 1).setValue('Unregistered');
      return { success: true };
    }
  }
  return { success: false, message: 'ไม่พบผู้ใช้หรือสถานะไม่ถูกต้อง' };
}
