// ==================== CONFIG ====================
var CONFIG = {
  SPREADSHEET_ID: '1yRrEV40jMMQCzycruoj8lIM0NPXTRhxXxUfvsaTO2uM',
  DRIVE_FOLDER_ID: '1kch--KBTd15dXHVIKutMAFy7Hj5ARtBo',
  DEFAULT_ADMIN: { username: 'admin', password: 'admin1234', name: 'ผู้ดูแลระบบ', role: 'Admin' },
  DEV_MODE: true, // ⚠️ เปลี่ยนเป็น false ก่อน deploy จริง — true = ข้ามการตรวจระยะ GPS (ยังบันทึกพิกัดปกติ)
  CHECKIN_LOCATION: { lat: 14.37462, lng: 99.14541, radiusMeters: 1000, name: 'อุทยานแห่งชาติเอราวัณ' },
  SHEETS: {
    Users:             ['ID', 'Username', 'Password', 'Name', 'Role', 'Status'],
    AttendanceLog:     ['LogID', 'Date', 'Name', 'Time_In', 'Time_Out', 'Task_Report', 'Photo_URL', 'Status', 'Latitude', 'Longitude', 'Distance_m', 'Selfie_URL'],
    WorkCycles:        ['CycleID', 'UserID', 'Name', 'Start_Date', 'End_Date', 'Required_Work_Days', 'Status'],
    WorkPlans:         ['PlanID', 'Submission_ID', 'CycleID', 'UserID', 'Name', 'Plan_Date', 'Plan_Status', 'Notes', 'Completed_LogID', 'Created_At', 'Submitted_At', 'Approved_At', 'Day_Type'],
    ScheduleRequests:  ['ReqID', 'CycleID', 'UserID', 'Name', 'Original_Date', 'Requested_Date', 'Reason', 'Status', 'Created_At', 'Decision_At', 'Request_Type']
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

  var now = new Date();
  var today = formatDate_(now);
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

  // สถิติรายเดือน (แยก On_Time / Late)
  var monthlyStats = {};
  var monthlyOnTime = {};
  var monthlyLate = {};
  logs.forEach(function(l) {
    if (l.Date) {
      var month = l.Date.substring(0, 7);
      monthlyStats[month] = (monthlyStats[month] || 0) + 1;
      if (l.Status === 'Late' || l.Status === 'Late_Report') {
        monthlyLate[month] = (monthlyLate[month] || 0) + 1;
      } else {
        monthlyOnTime[month] = (monthlyOnTime[month] || 0) + 1;
      }
    }
  });

  // สถิติวันนี้
  var todayOnTime = todayLogs.filter(function(l) { return l.Status === 'On_Time' || l.Status === 'Completed'; }).length;
  var todayLate = todayLogs.filter(function(l) { return l.Status === 'Late' || l.Status === 'Late_Report'; }).length;

  // สถิติผู้ใช้
  var activeUsers = users.filter(function(u) { return u.Status === 'Active' && u.Role !== 'Admin'; });
  var pendingUsers = users.filter(function(u) { return u.Status === 'Pending'; });
  var unregUsers = users.filter(function(u) { return u.Status === 'Unregistered'; });

  // สถิติรวมทั้งหมด
  var totalOnTime = logs.filter(function(l) { return l.Status === 'On_Time' || l.Status === 'Completed'; }).length;
  var totalLate = logs.filter(function(l) { return l.Status === 'Late' || l.Status === 'Late_Report'; }).length;

  // รายชื่อผู้มาปฏิบัติงานวันนี้
  var todayAttendanceList = todayLogs.map(function(l) {
    return { name: l.Name, timeIn: l.Time_In, timeOut: l.Time_Out || '', status: l.Status, distance: l.Distance_m || '-', selfieUrl: l.Selfie_URL || '' };
  });

  // ===== NEW: จำนวนคนที่ต้องมาวันนี้ (มีแผน Approved วันนี้) =====
  var todayExpectedNames = {};
  plans.forEach(function(p) {
    if (p.Plan_Date === today && p.Plan_Status === 'Approved') {
      todayExpectedNames[p.Name] = true;
    }
  });
  var todayExpectedCount = Object.keys(todayExpectedNames).length;

  // ===== NEW: เวลาเช็คอินเฉลี่ย / เร็วสุด / ช้าสุด วันนี้ =====
  var todayAvgTime = null;
  if (todayLogs.length > 0) {
    var timeMinutes = [];
    var earliest = 9999, latest = 0, earliestName = '', latestName = '';
    todayLogs.forEach(function(l) {
      if (l.Time_In) {
        var timePart = String(l.Time_In).split(' ').pop();
        var tp = timePart.split(':');
        var mins = parseInt(tp[0]) * 60 + parseInt(tp[1]);
        timeMinutes.push(mins);
        if (mins < earliest) { earliest = mins; earliestName = l.Name; }
        if (mins > latest) { latest = mins; latestName = l.Name; }
      }
    });
    if (timeMinutes.length > 0) {
      var avgMins = Math.round(timeMinutes.reduce(function(a, b) { return a + b; }, 0) / timeMinutes.length);
      todayAvgTime = {
        avg: String(Math.floor(avgMins / 60)).padStart(2, '0') + ':' + String(avgMins % 60).padStart(2, '0'),
        earliest: String(Math.floor(earliest / 60)).padStart(2, '0') + ':' + String(earliest % 60).padStart(2, '0'),
        earliestName: earliestName,
        latest: String(Math.floor(latest / 60)).padStart(2, '0') + ':' + String(latest % 60).padStart(2, '0'),
        latestName: latestName
      };
    }
  }

  // ===== NEW: สถิติรายวัน 14 วันล่าสุด =====
  var dailyStats = [];
  for (var di = 13; di >= 0; di--) {
    var d = new Date(now);
    d.setDate(d.getDate() - di);
    var ds = formatDate_(d);
    var dayLogs = logs.filter(function(l) { return l.Date === ds; });
    var dayOnTime = dayLogs.filter(function(l) { return l.Status === 'On_Time' || l.Status === 'Completed'; }).length;
    var dayLate = dayLogs.filter(function(l) { return l.Status === 'Late' || l.Status === 'Late_Report'; }).length;
    dailyStats.push({ date: ds, total: dayLogs.length, onTime: dayOnTime, late: dayLate });
  }

  // ===== NEW: สถิติตามวันในสัปดาห์ (0=อาทิตย์ ... 6=เสาร์) =====
  var weekdayStats = [
    { day: 'อา.', onTime: 0, late: 0, total: 0 },
    { day: 'จ.', onTime: 0, late: 0, total: 0 },
    { day: 'อ.', onTime: 0, late: 0, total: 0 },
    { day: 'พ.', onTime: 0, late: 0, total: 0 },
    { day: 'พฤ.', onTime: 0, late: 0, total: 0 },
    { day: 'ศ.', onTime: 0, late: 0, total: 0 },
    { day: 'ส.', onTime: 0, late: 0, total: 0 }
  ];
  logs.forEach(function(l) {
    if (l.Date) {
      var logDate = new Date(l.Date);
      var dow = logDate.getDay();
      weekdayStats[dow].total++;
      if (l.Status === 'Late' || l.Status === 'Late_Report') {
        weekdayStats[dow].late++;
      } else {
        weekdayStats[dow].onTime++;
      }
    }
  });

  // ===== Engagement Score — สถิติรายบุคคล =====
  // คะแนนต่อวัน (เต็ม 100):
  //   มาทำงาน (มี check-in)       = 30 คะแนน
  //   เช็คอินตรงเวลา (≤ 08:15)    = 30 คะแนน
  //   ส่งรายงาน (มี Time_Out)      = 20 คะแนน
  //   ส่งรายงานตรงเวลา (≤ 17:00)  = 20 คะแนน
  //   ขาดงาน (มีแผนแต่ไม่มา)      =  0 คะแนน
  // engagementScore = (คะแนนรวม / (วันที่มีแผน × 100)) × 100
  var perUserStats = [];
  activeUsers.forEach(function(u) {
    var uLogs = logs.filter(function(l) { return l.Name === u.Name; });

    // วันที่มีแผน Approved ถึงวันนี้ (ไม่นับวันอนาคต, ไม่นับ Swapped_Out)
    var uPlans = plans.filter(function(p) {
      return p.UserID === u.ID && p.Plan_Status === 'Approved' && p.Plan_Date <= today;
    });
    var plannedDays = uPlans.length;

    // วิเคราะห์จาก Time_In / Time_Out จริง (ไม่พึ่ง Status ที่ถูก checkout เขียนทับ)
    var onTimeIn = 0;      // เช้าตรงเวลา (≤ 08:15 = 495 นาที)
    var lateIn = 0;        // เช้าสาย
    var reported = 0;      // ส่งรายงานแล้ว
    var onTimeReport = 0;  // ส่งรายงานตรงเวลา (≤ 17:00 = 1020 นาที)
    var lateReport = 0;    // ส่งรายงานสาย
    var noReport = 0;      // ยังไม่ส่งรายงาน
    var totalPoints = 0;
    var minsArr = [];

    uLogs.forEach(function(l) {
      var dayPts = 30; // มาทำงาน = 30 คะแนน

      // ตรวจเวลาเช็คอินจาก Time_In จริง
      if (l.Time_In) {
        var timePart = String(l.Time_In).split(' ').pop();
        var tp = timePart.split(':');
        var inMins = parseInt(tp[0]) * 60 + parseInt(tp[1]);
        minsArr.push(inMins);
        if (inMins <= 495) {
          onTimeIn++;
          dayPts += 30;
        } else {
          lateIn++;
        }
      }

      // ตรวจการส่งรายงานจาก Time_Out จริง
      if (l.Time_Out && String(l.Time_Out).trim() !== '') {
        reported++;
        dayPts += 20;
        var outPart = String(l.Time_Out).split(' ').pop();
        var op = outPart.split(':');
        var outMins = parseInt(op[0]) * 60 + parseInt(op[1]);
        if (outMins <= 1020) {
          onTimeReport++;
          dayPts += 20;
        } else {
          lateReport++;
        }
      } else {
        noReport++;
      }

      totalPoints += dayPts;
    });

    var attendedDays = uLogs.length;
    var absentDays = Math.max(0, plannedDays - attendedDays);

    // Engagement Score: คิดจากวันที่มีแผนทั้งหมด (วันขาด = 0 คะแนน)
    var maxPoints = plannedDays > 0 ? plannedDays * 100 : 1;
    var engagementScore = plannedDays > 0 ? Math.round((totalPoints / maxPoints) * 100) : 0;

    // เวลาเช็คอินเฉลี่ย
    var avgTimeStr = '-';
    if (minsArr.length > 0) {
      var avgMin = Math.round(minsArr.reduce(function(a, b) { return a + b; }, 0) / minsArr.length);
      avgTimeStr = String(Math.floor(avgMin / 60)).padStart(2, '0') + ':' + String(avgMin % 60).padStart(2, '0');
    }

    perUserStats.push({
      name: u.Name,
      username: u.Username,
      plannedDays: plannedDays,
      attendedDays: attendedDays,
      absentDays: absentDays,
      onTimeIn: onTimeIn,
      lateIn: lateIn,
      reported: reported,
      onTimeReport: onTimeReport,
      lateReport: lateReport,
      noReport: noReport,
      totalPoints: totalPoints,
      engagementScore: engagementScore,
      avgCheckIn: avgTimeStr
    });
  });
  perUserStats.sort(function(a, b) { return b.engagementScore - a.engagementScore || b.attendedDays - a.attendedDays; });

  // ===== NEW: ความคืบหน้ารอบงาน =====
  var activeCycles = cycles.filter(function(c) { return c.Status === 'Active'; });
  var cycleProgress = [];
  activeCycles.forEach(function(c) {
    var cPlansApproved = plans.filter(function(p) { return p.CycleID === c.CycleID && p.Plan_Status === 'Approved'; }).length;
    var cLogsCompleted = logs.filter(function(l) {
      return l.Name === c.Name && l.Date >= c.Start_Date && l.Date <= c.End_Date && l.Time_In;
    }).length;
    var required = parseInt(c.Required_Work_Days) || 30;
    var totalDaysInCycle = Math.max(1, Math.round((new Date(c.End_Date) - new Date(c.Start_Date)) / 86400000));
    var elapsedDays = Math.max(0, Math.round((now - new Date(c.Start_Date)) / 86400000));
    var progressPct = Math.min(100, Math.round((cLogsCompleted / required) * 100));
    var remainDays = Math.max(0, Math.round((new Date(c.End_Date) - now) / 86400000));
    var expectedByNow = required > 0 ? Math.round((elapsedDays / totalDaysInCycle) * required) : 0;
    var behindSchedule = cLogsCompleted < expectedByNow;

    cycleProgress.push({
      name: c.Name,
      cycleId: c.CycleID,
      start: c.Start_Date,
      end: c.End_Date,
      required: required,
      approved: cPlansApproved,
      completed: cLogsCompleted,
      progressPct: progressPct,
      remainDays: remainDays,
      behindSchedule: behindSchedule
    });
  });

  // ===== NEW: เปรียบเทียบสัปดาห์นี้ vs สัปดาห์ก่อน =====
  var thisWeekStart = new Date(now);
  thisWeekStart.setDate(thisWeekStart.getDate() - thisWeekStart.getDay() + 1);
  var lastWeekStart = new Date(thisWeekStart);
  lastWeekStart.setDate(lastWeekStart.getDate() - 7);
  var thisWeekStartStr = formatDate_(thisWeekStart);
  var lastWeekStartStr = formatDate_(lastWeekStart);
  var lastWeekEndStr = formatDate_(new Date(thisWeekStart.getTime() - 86400000));

  var thisWeekLogs = logs.filter(function(l) { return l.Date >= thisWeekStartStr && l.Date <= today; });
  var lastWeekLogs = logs.filter(function(l) { return l.Date >= lastWeekStartStr && l.Date <= lastWeekEndStr; });

  var thisWeekOnTime = thisWeekLogs.filter(function(l) { return l.Status === 'On_Time' || l.Status === 'Completed'; }).length;
  var thisWeekLate = thisWeekLogs.filter(function(l) { return l.Status === 'Late' || l.Status === 'Late_Report'; }).length;
  var lastWeekOnTime = lastWeekLogs.filter(function(l) { return l.Status === 'On_Time' || l.Status === 'Completed'; }).length;
  var lastWeekLate = lastWeekLogs.filter(function(l) { return l.Status === 'Late' || l.Status === 'Late_Report'; }).length;

  var weekComparison = {
    thisWeek: { total: thisWeekLogs.length, onTime: thisWeekOnTime, late: thisWeekLate },
    lastWeek: { total: lastWeekLogs.length, onTime: lastWeekOnTime, late: lastWeekLate },
    totalDiff: thisWeekLogs.length - lastWeekLogs.length,
    onTimeDiff: thisWeekOnTime - lastWeekOnTime,
    lateDiff: thisWeekLate - lastWeekLate
  };

  // ===== NEW: Activity Feed (10 กิจกรรมล่าสุด) =====
  var activities = [];
  logs.forEach(function(l) {
    if (l.Time_In) {
      var timeStr = String(l.Time_In).split(' ').pop();
      activities.push({ type: 'checkin', name: l.Name, date: l.Date, time: timeStr, detail: l.Status === 'Late' ? 'มาสาย' : 'ตรงเวลา' });
    }
    if (l.Time_Out) {
      var timeStr2 = String(l.Time_Out).split(' ').pop();
      activities.push({ type: 'checkout', name: l.Name, date: l.Date, time: timeStr2, detail: 'รายงานผล' });
    }
  });
  requests.forEach(function(r) {
    activities.push({ type: 'request', name: r.Name, date: r.Created_At ? r.Created_At.substring(0, 10) : '', time: r.Created_At ? r.Created_At.substring(11, 16) : '', detail: (r.Request_Type === 'Half_Day' ? 'ขอลาครึ่งวัน' : 'ขอสลับวัน') + ' (' + r.Status + ')' });
  });
  plans.filter(function(p) { return p.Submitted_At; }).forEach(function(p) {
    if (!activities.some(function(a) { return a.type === 'plan' && a.detail === p.Submission_ID; })) {
      activities.push({ type: 'plan', name: p.Name, date: p.Submitted_At ? p.Submitted_At.substring(0, 10) : '', time: p.Submitted_At ? p.Submitted_At.substring(11, 16) : '', detail: p.Submission_ID });
    }
  });
  activities.sort(function(a, b) {
    var da = a.date + ' ' + a.time;
    var db = b.date + ' ' + b.time;
    return da > db ? -1 : da < db ? 1 : 0;
  });
  var recentActivity = activities.slice(0, 10).map(function(a) {
    var icon = a.type === 'checkin' ? 'checkin' : a.type === 'checkout' ? 'checkout' : a.type === 'request' ? 'request' : 'plan';
    return { icon: icon, name: a.name, date: a.date, time: a.time, detail: a.detail };
  });

  return {
    users: users,
    cycles: cycles,
    plans: plans,
    logs: logs,
    requests: requests,
    today: today,
    todayAttendanceCount: todayLogs.length,
    todayExpectedCount: todayExpectedCount,
    todayOnTime: todayOnTime,
    todayLate: todayLate,
    todayAttendanceList: todayAttendanceList,
    todayAvgTime: todayAvgTime,
    activeCyclesCount: activeCycles.length,
    activeUsersCount: activeUsers.length,
    pendingUsersCount: pendingUsers.length,
    unregUsersCount: unregUsers.length,
    pendingRequestsCount: pendingRequests.length,
    pendingPlansCount: Object.keys(pendingSubmissions).length,
    totalOnTime: totalOnTime,
    totalLate: totalLate,
    totalLogs: logs.length,
    pendingSubmissions: Object.values(pendingSubmissions),
    pendingRequests: pendingRequests,
    monthlyStats: monthlyStats,
    monthlyOnTime: monthlyOnTime,
    monthlyLate: monthlyLate,
    dailyStats: dailyStats,
    weekdayStats: weekdayStats,
    perUserStats: perUserStats,
    cycleProgress: cycleProgress,
    weekComparison: weekComparison,
    recentActivity: recentActivity
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
    rows.push([planId, submissionId, cycleId, user.id, user.name, date, 'Pending', '', '', now, now, '', 'Full']);
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
function createScheduleChangeRequest(token, cycleId, originalDate, requestedDate, reason, requestType) {
  var user = validateSession_(token);
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('ScheduleRequests');
  var reqId = 'REQ' + new Date().getTime();
  var now = formatDateTime_(new Date());
  var type = requestType || 'Swap';
  sheet.appendRow([reqId, cycleId, user.id, user.name, originalDate, requestedDate, reason, 'Pending', now, '', type]);
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

  var nameIdx = reqHeaders.indexOf('Name');
  var reqTypeIdx = reqHeaders.indexOf('Request_Type');
  var request = null;
  for (var i = 1; i < reqData.length; i++) {
    if (reqData[i][reqIdIdx] === reqId) {
      reqSheet.getRange(i + 1, statusIdx + 1).setValue(newStatus);
      reqSheet.getRange(i + 1, decisionIdx + 1).setValue(now);
      var reqType = reqTypeIdx >= 0 ? (reqData[i][reqTypeIdx] || 'Swap') : 'Swap';
      request = { originalDate: reqData[i][origDateIdx], requestedDate: reqData[i][newDateIdx], userId: reqData[i][userIdIdx], cycleId: reqData[i][cycleIdIdx], name: reqData[i][nameIdx], type: reqType };
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
    var pDayTypeIdx = planHeaders.indexOf('Day_Type');

    // แปลงวันที่จาก request เป็น string yyyy-MM-dd (แก้ปัญหา Date object === Date object → false)
    var origDateStr = (request.originalDate instanceof Date) ? formatDate_(request.originalDate) : String(request.originalDate);
    var reqDateStr = (request.requestedDate instanceof Date) ? formatDate_(request.requestedDate) : String(request.requestedDate);

    if (request.type === 'Half_Day') {
      // ===== ครึ่งวัน: เปลี่ยน Day_Type ของวันเดิมเป็น Half + เพิ่มวันชดเชยเป็น Half =====
      for (var j = 1; j < planData.length; j++) {
        if (String(planData[j][pUserIdIdx]) === String(request.userId) && String(planData[j][pCycleIdIdx]) === String(request.cycleId)) {
          var planDateStr = (planData[j][pDateIdx] instanceof Date) ? formatDate_(planData[j][pDateIdx]) : String(planData[j][pDateIdx]);
          if (planDateStr === origDateStr && planData[j][pStatusIdx] === 'Approved') {
            if (pDayTypeIdx >= 0) {
              planSheet.getRange(j + 1, pDayTypeIdx + 1).setValue('Half');
            }
          }
        }
      }
      // เพิ่มวันชดเชย (ครึ่งวัน)
      var newPlanId = 'PLN' + new Date().getTime();
      planSheet.appendRow([newPlanId, '', request.cycleId, request.userId, request.name, reqDateStr, 'Approved', 'ชดเชยครึ่งวันจาก ' + origDateStr, '', now, '', now, 'Half']);
    } else {
      // ===== สลับวัน: เปลี่ยนวันเดิมเป็น Swapped_Out + เพิ่มวันใหม่เป็น Full =====
      for (var j = 1; j < planData.length; j++) {
        if (String(planData[j][pUserIdIdx]) === String(request.userId) && String(planData[j][pCycleIdIdx]) === String(request.cycleId)) {
          var planDateStr = (planData[j][pDateIdx] instanceof Date) ? formatDate_(planData[j][pDateIdx]) : String(planData[j][pDateIdx]);
          if (planDateStr === origDateStr && planData[j][pStatusIdx] === 'Approved') {
            planSheet.getRange(j + 1, pStatusIdx + 1).setValue('Swapped_Out');
          }
        }
      }
      // เพิ่มวันใหม่ (เต็มวัน)
      var newPlanId = 'PLN' + new Date().getTime();
      planSheet.appendRow([newPlanId, '', request.cycleId, request.userId, request.name, reqDateStr, 'Approved', 'สลับจากวันที่ ' + origDateStr, '', now, '', now, 'Full']);
    }
  }

  return { success: true };
}

// ==================== CHECK IN / CHECK OUT ====================
function checkIn(token, latitude, longitude, selfieData) {
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

  // ตรวจสอบรูปเซลฟี่
  if (!selfieData || !selfieData.data) {
    return { success: false, message: 'กรุณาถ่ายรูปเซลฟี่เพื่อยืนยันตัวตนก่อนลงเวลา' };
  }

  // ตรวจสอบพิกัด GPS
  if (!latitude || !longitude) {
    return { success: false, message: 'ไม่สามารถระบุตำแหน่งของคุณได้ กรุณาเปิด GPS แล้วลองอีกครั้ง' };
  }
  var loc = CONFIG.CHECKIN_LOCATION;
  var distance = calculateDistance_(latitude, longitude, loc.lat, loc.lng);
  var distanceRounded = Math.round(distance);
  if (distance > loc.radiusMeters && !CONFIG.DEV_MODE) {
    return { success: false, message: 'คุณอยู่ห่างจากจุดลงเวลา (' + loc.name + ') ประมาณ ' + (distanceRounded >= 1000 ? (distanceRounded / 1000).toFixed(1) + ' กม.' : distanceRounded + ' เมตร') + '\nต้องอยู่ในรัศมีไม่เกิน ' + (loc.radiusMeters >= 1000 ? (loc.radiusMeters / 1000) + ' กม.' : loc.radiusMeters + ' เมตร') };
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

  // อัปโหลดรูปเซลฟี่ไป Google Drive
  var selfieUrl = '';
  try {
    selfieUrl = uploadSelfie_(user.name, today, selfieData);
  } catch (e) {
    return { success: false, message: 'ไม่สามารถอัปโหลดรูปเซลฟี่ได้: ' + e.message };
  }

  var logId = 'LOG' + now.getTime();
  var timeInDisplay = today + ' ' + currentTime;
  var sheet = ss.getSheetByName('AttendanceLog');
  sheet.appendRow([logId, today, user.name, timeInDisplay, '', '', '', lateStatus, latitude, longitude, distanceRounded, selfieUrl]);

  // อัพเดท Completed_LogID ใน WorkPlans
  updatePlanCompletedLog_(ss, user.id, today, logId);

  return { success: true, logId: logId, timeIn: timeInDisplay, lateStatus: lateStatus, distance: distanceRounded };
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
function uploadSelfie_(userName, dateStr, selfieData) {
  var rootFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  var monthStr = dateStr.substring(0, 7);
  var monthFolder = getOrCreateFolder_(rootFolder, monthStr);
  var dateFolder = getOrCreateFolder_(monthFolder, dateStr);
  var userFolder = getOrCreateFolder_(dateFolder, userName);

  var fileName = 'selfie_' + userName + '_' + dateStr + '.jpg';
  var blob = decodeBase64ToBlob_(selfieData.data, selfieData.mimeType || 'image/jpeg', fileName);
  var file = userFolder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}

  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400';
}

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

function calculateDistance_(lat1, lng1, lat2, lng2) {
  var R = 6371000;
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLng = (lng2 - lng1) * Math.PI / 180;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
          Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
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
      return { success: true, message: 'ลงทะเบียนสำเร็จ! กรุณารอผู้ดูแลระบบอนุมัติ', username: data[i][headers.indexOf('Username')], name: data[i][headers.indexOf('Name')] };
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

// ==================== ONE-TIME REPAIR ====================
// รันฟังก์ชันนี้ครั้งเดียวเพื่อแก้ข้อมูล WorkPlans ที่สลับวันแล้วแต่วันเดิมยังเป็น Approved
// (เกิดจาก bug Date object comparison ก่อนหน้านี้)
// ⚠️ ลบฟังก์ชันนี้ออกหลังรันเสร็จแล้วได้
function repairSwappedPlans() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // 1. อ่าน ScheduleRequests ที่ Approved แล้ว
  var reqSheet = ss.getSheetByName('ScheduleRequests');
  var reqData = reqSheet.getDataRange().getValues();
  var reqHeaders = reqData[0];
  var rStatusIdx = reqHeaders.indexOf('Status');
  var rOrigIdx = reqHeaders.indexOf('Original_Date');
  var rUserIdIdx = reqHeaders.indexOf('UserID');
  var rCycleIdIdx = reqHeaders.indexOf('CycleID');
  var rTypeIdx = reqHeaders.indexOf('Request_Type');

  var approvedRequests = [];
  for (var i = 1; i < reqData.length; i++) {
    if (reqData[i][rStatusIdx] === 'Approved') {
      var origDate = reqData[i][rOrigIdx];
      approvedRequests.push({
        userId: String(reqData[i][rUserIdIdx]),
        cycleId: String(reqData[i][rCycleIdIdx]),
        originalDate: (origDate instanceof Date) ? formatDate_(origDate) : String(origDate),
        type: rTypeIdx >= 0 ? (reqData[i][rTypeIdx] || 'Swap') : 'Swap'
      });
    }
  }

  if (approvedRequests.length === 0) {
    Logger.log('ไม่มีคำร้องที่ Approved — ไม่ต้องแก้ไข');
    return;
  }

  // 2. อ่าน WorkPlans แล้วหาแถวที่ต้องแก้
  var planSheet = ss.getSheetByName('WorkPlans');
  var planData = planSheet.getDataRange().getValues();
  var planHeaders = planData[0];
  var pUserIdIdx = planHeaders.indexOf('UserID');
  var pCycleIdIdx = planHeaders.indexOf('CycleID');
  var pDateIdx = planHeaders.indexOf('Plan_Date');
  var pStatusIdx = planHeaders.indexOf('Plan_Status');
  var pDayTypeIdx = planHeaders.indexOf('Day_Type');

  var fixCount = 0;

  for (var r = 0; r < approvedRequests.length; r++) {
    var req = approvedRequests[r];
    for (var j = 1; j < planData.length; j++) {
      var planUserId = String(planData[j][pUserIdIdx]);
      var planCycleId = String(planData[j][pCycleIdIdx]);
      var planDate = planData[j][pDateIdx];
      var planDateStr = (planDate instanceof Date) ? formatDate_(planDate) : String(planDate);
      var planStatus = planData[j][pStatusIdx];

      if (planUserId === req.userId && planCycleId === req.cycleId && planDateStr === req.originalDate && planStatus === 'Approved') {
        if (req.type === 'Half_Day') {
          if (pDayTypeIdx >= 0) {
            var currentDayType = planData[j][pDayTypeIdx];
            if (currentDayType !== 'Half') {
              planSheet.getRange(j + 1, pDayTypeIdx + 1).setValue('Half');
              Logger.log('FIXED Half_Day: ' + req.userId + ' | ' + req.originalDate + ' | Day_Type → Half');
              fixCount++;
            }
          }
        } else {
          planSheet.getRange(j + 1, pStatusIdx + 1).setValue('Swapped_Out');
          Logger.log('FIXED Swap: ' + req.userId + ' | ' + req.originalDate + ' | Plan_Status → Swapped_Out');
          fixCount++;
        }
      }
    }
  }

  Logger.log('===== REPAIR COMPLETE: แก้ไข ' + fixCount + ' แถว =====');
}
