// ==================== LINE BOT CONFIG ====================
var BOT_CONFIG = {
  LINE_CHANNEL_ACCESS_TOKEN: 'ใส่_Channel_Access_Token_ตรงนี้',  // ⚠️ ต้องเปลี่ยน
  LINE_GROUP_ID: 'ใส่_Group_ID_ตรงนี้',                          // ⚠️ ต้องเปลี่ยน (ดูวิธีหาในคู่มือ)
  SPREADSHEET_ID: '1yRrEV40jMMQCzycruoj8lIM0NPXTRhxXxUfvsaTO2uM' // ใช้ Sheet เดียวกับระบบหลัก
};

// ==================== WEBHOOK (doPost) ====================
// LINE จะส่ง event มาที่นี่เมื่อมีข้อความในกลุ่ม
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var events = data.events;

    for (var i = 0; i < events.length; i++) {
      var event = events[i];

      // ถ้าเป็นข้อความ "ทดสอบ" → ส่งรายงานทั้ง 2 ช่วงเพื่อทดสอบ
      if (event.type === 'message' && event.message.type === 'text') {
        var text = event.message.text.trim();

        if (text === 'ทดสอบ') {
          var replyToken = event.replyToken;
          var thaiDate = formatThaiDate_(new Date());
          var morningMsg = buildMorningReport_();
          var eveningMsg = buildEveningReport_();
          var testMsg = '🧪 [ทดสอบระบบแจ้งเตือน]\n\n' +
            '━━━━━━━━━━━━━━━━\n' +
            '� รายงานวันที่ ' + thaiDate + ' (ช่วงเช้า)\n' +
            '━━━━━━━━━━━━━━━━\n' +
            morningMsg + '\n\n' +
            '━━━━━━━━━━━━━━━━\n' +
            '� รายงานวันที่ ' + thaiDate + ' (ช่วงเย็น)\n' +
            '━━━━━━━━━━━━━━━━\n' +
            eveningMsg;
          replyMessage_(replyToken, testMsg);
        }
        // ข้อความอื่น → ไม่สนใจ
      }

      // ถ้าบอทถูกเชิญเข้ากลุ่ม → แสดง Group ID เพื่อให้ copy ไปใส่ config
      if (event.type === 'join') {
        var groupId = event.source.groupId;
        replyMessage_(event.replyToken,
          '✅ บอทเข้ากลุ่มเรียบร้อย!\n\n' +
          '📋 Group ID ของกลุ่มนี้คือ:\n' + groupId + '\n\n' +
          '⚠️ กรุณา copy Group ID นี้ไปใส่ใน BOT_CONFIG.LINE_GROUP_ID ในสคริปต์'
        );
      }
    }
  } catch (err) {
    Logger.log('doPost Error: ' + err.message);
  }

  return ContentService.createTextOutput('OK');
}

// ==================== SCHEDULED FUNCTIONS ====================
// ⏰ ตั้ง Trigger ให้รันเวลา 08:30 น.
function sendMorningReport() {
  var msg = '📋 รายงานวันที่ ' + formatThaiDate_(new Date()) + ' (ช่วงเช้า)\n\n' + buildMorningReport_();
  pushMessage_(msg);
}

// ⏰ ตั้ง Trigger ให้รันเวลา 17:30 น.
function sendEveningReport() {
  var msg = '📊 รายงานวันที่ ' + formatThaiDate_(new Date()) + ' (ช่วงเย็น)\n\n' + buildEveningReport_();
  pushMessage_(msg);
}

// ==================== REPORT BUILDERS ====================
function buildMorningReport_() {
  var ss = SpreadsheetApp.openById(BOT_CONFIG.SPREADSHEET_ID);
  var today = formatDate_(new Date());

  // ดึงข้อมูล
  var users = getSheetData_(ss, 'Users');
  var logs = getSheetData_(ss, 'AttendanceLog');
  var plans = getSheetData_(ss, 'WorkPlans');

  // เจ้าหน้าที่ Active (ไม่รวม Admin)
  var activeUsers = users.filter(function(u) { return u.Status === 'Active' && u.Role !== 'Admin'; });

  // คนที่มีแผนวันนี้ (Approved) — รวมแผนที่มาจากการสลับวัน (Name อาจว่าง)
  var todayPlans = plans.filter(function(p) { return p.Plan_Date === today && p.Plan_Status === 'Approved'; });
  var plannedNames = [];
  var halfDayNames = [];
  todayPlans.forEach(function(p) {
    var name = p.Name;
    if (!name || name === '') {
      for (var k = 0; k < users.length; k++) {
        if (users[k].ID === p.UserID) { name = users[k].Name; break; }
      }
    }
    if (name && plannedNames.indexOf(name) === -1) {
      plannedNames.push(name);
      if (p.Day_Type === 'Half') { halfDayNames.push(name); }
    }
  });

  // คนที่เช็คอินวันนี้
  var todayLogs = logs.filter(function(l) { return l.Date === today; });
  var checkedInNames = todayLogs.map(function(l) { return l.Name; });

  // ตัดสินสาย/ตรงเวลาจาก Time_In จริง (หลัง 08:15 = สาย) ไม่พึ่ง Status ที่ถูกเขียนทับตอนเช็คเอาท์
  var lateLogs = [];
  var onTimeLogs = [];
  todayLogs.forEach(function(l) {
    var timeStr = l.Time_In ? l.Time_In.split(' ').pop() : '';
    var parts = timeStr.split(':');
    var h = parseInt(parts[0] || '0');
    var m = parseInt(parts[1] || '0');
    var totalMin = h * 60 + m;
    if (totalMin > 495) { // 495 = 08:15
      lateLogs.push(l);
    } else {
      onTimeLogs.push(l);
    }
  });
  var onTimeCount = onTimeLogs.length;
  var lateCount = lateLogs.length;

  // คนที่มีแผนแต่ยังไม่เช็คอิน
  var notCheckedIn = plannedNames.filter(function(name) {
    return checkedInNames.indexOf(name) === -1;
  });

  // คนที่หยุดวันนี้ (Active แต่ไม่มีแผนวันนี้)
  var offToday = activeUsers.filter(function(u) {
    return plannedNames.indexOf(u.Name) === -1;
  });

  // สร้างข้อความ
  var msg = '';
  msg += '👥 เจ้าหน้าที่ทั้งหมด: ' + activeUsers.length + ' คน\n';
  msg += '📅 มีแผนปฏิบัติงานวันนี้: ' + plannedNames.length + ' คน\n';
  msg += '✅ เช็คชื่อแล้ว: ' + todayLogs.length + ' คน\n';
  msg += '   • ตรงเวลา: ' + onTimeCount + ' คน\n';
  msg += '   • มาสาย: ' + lateCount + ' คน\n';

  if (lateCount > 0) {
    msg += '\n😅 รายชื่อผู้มาสาย:\n';
    lateLogs.forEach(function(l, i) {
      msg += '   ' + (i + 1) + '. ' + l.Name + ' (เข้า ' + (l.Time_In ? l.Time_In.split(' ').pop() : '-') + ')\n';
    });
  }

  if (halfDayNames.length > 0) {
    msg += '\n⏰ ปฏิบัติงานครึ่งวัน (' + halfDayNames.length + ' คน):\n';
    halfDayNames.forEach(function(name, i) {
      msg += '   ' + (i + 1) + '. ' + name + '\n';
    });
  }

  if (notCheckedIn.length > 0) {
    msg += '\n⚠️ มีแผนแต่ยังไม่เช็คชื่อ (' + notCheckedIn.length + ' คน):\n';
    notCheckedIn.forEach(function(name, i) {
      msg += '   ' + (i + 1) + '. ' + name + '\n';
    });
  }

  if (offToday.length > 0) {
    msg += '\n🏖️ ผู้หยุดวันนี้ (' + offToday.length + ' คน):\n';
    offToday.forEach(function(u, i) {
      msg += '   ' + (i + 1) + '. ' + u.Name + '\n';
    });
  }

  // คำขอเปลี่ยนวันที่รอดำเนินการ
  var schedReqs = getSheetData_(ss, 'ScheduleRequests');
  var pendingReqs = schedReqs.filter(function(r) { return r.Status === 'Pending'; });
  if (pendingReqs.length > 0) {
    msg += '\n📨 คำขอเปลี่ยนแผนรออนุมัติ (' + pendingReqs.length + ' รายการ):\n';
    pendingReqs.forEach(function(r, i) {
      var typeLabel = r.Request_Type === 'Half_Day' ? 'ครึ่งวัน' : 'สลับวัน';
      msg += '   ' + (i + 1) + '. ' + r.Name + ' [' + typeLabel + '] (' + r.Original_Date + ' → ' + r.Requested_Date + ')\n';
    });
  } else {
    msg += '\n✅ ไม่มีคำขอเปลี่ยนแผนรออนุมัติ';
  }

  return msg.trim();
}

function buildEveningReport_() {
  var ss = SpreadsheetApp.openById(BOT_CONFIG.SPREADSHEET_ID);
  var today = formatDate_(new Date());

  // ดึงข้อมูล
  var logs = getSheetData_(ss, 'AttendanceLog');

  // คนที่เช็คอินวันนี้
  var todayLogs = logs.filter(function(l) { return l.Date === today; });
  var totalCheckedIn = todayLogs.length;

  // คนที่ส่งรายงานแล้ว (มี Time_Out)
  var reported = todayLogs.filter(function(l) { return l.Time_Out && l.Time_Out !== ''; });
  var reportedCount = reported.length;

  // คนที่ยังไม่ส่งรายงาน
  var notReported = todayLogs.filter(function(l) { return !l.Time_Out || l.Time_Out === ''; });

  // คนที่ส่งรายงานล่าช้า
  var lateReports = reported.filter(function(l) { return l.Status === 'Late_Report'; });
  var lateReportCount = lateReports.length;

  // คนที่ส่งตรงเวลา
  var onTimeReportCount = reported.filter(function(l) { return l.Status === 'Completed'; }).length;

  // สร้างข้อความ
  var msg = '';
  msg += '📋 ผู้มาปฏิบัติงานวันนี้: ' + totalCheckedIn + ' คน\n';
  msg += '✅ ส่งรายงานแล้ว: ' + reportedCount + '/' + totalCheckedIn + ' คน';
  msg += (reportedCount === totalCheckedIn && totalCheckedIn > 0) ? ' ✨ ครบ!\n' : '\n';
  msg += '   • ตรงเวลา: ' + onTimeReportCount + ' คน\n';
  msg += '   • ส่งล่าช้า: ' + lateReportCount + ' คน\n';

  if (lateReportCount > 0) {
    msg += '\n⏰ รายชื่อผู้ส่งรายงานล่าช้า:\n';
    lateReports.forEach(function(l, i) {
      msg += '   ' + (i + 1) + '. ' + l.Name + '\n';
    });
  }

  if (notReported.length > 0) {
    msg += '\n❌ ยังไม่ส่งรายงาน (' + notReported.length + ' คน):\n';
    notReported.forEach(function(l, i) {
      msg += '   ' + (i + 1) + '. ' + l.Name + '\n';
    });
  } else if (totalCheckedIn > 0) {
    msg += '\n🎉 ทุกคนส่งรายงานครบถ้วนแล้ว!';
  } else {
    msg += '\n📭 ไม่มีผู้ปฏิบัติงานวันนี้';
  }

  // คำขอเปลี่ยนวันที่รอดำเนินการ
  var schedReqs = getSheetData_(ss, 'ScheduleRequests');
  var pendingReqs = schedReqs.filter(function(r) { return r.Status === 'Pending'; });
  if (pendingReqs.length > 0) {
    msg += '\n\n📨 คำขอเปลี่ยนแผนรออนุมัติ (' + pendingReqs.length + ' รายการ):\n';
    pendingReqs.forEach(function(r, i) {
      var typeLabel = r.Request_Type === 'Half_Day' ? 'ครึ่งวัน' : 'สลับวัน';
      msg += '   ' + (i + 1) + '. ' + r.Name + ' [' + typeLabel + '] (' + r.Original_Date + ' → ' + r.Requested_Date + ')\n';
    });
  } else {
    msg += '\n\n✅ ไม่มีคำขอเปลี่ยนแผนรออนุมัติ';
  }

  return msg.trim();
}

// ==================== LINE API ====================
// ส่งข้อความตอบกลับ (reply)
function replyMessage_(replyToken, text) {
  var url = 'https://api.line.me/v2/bot/message/reply';
  var payload = {
    replyToken: replyToken,
    messages: [{ type: 'text', text: text }]
  };
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + BOT_CONFIG.LINE_CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

// ส่งข้อความไปยังกลุ่ม (push)
function pushMessage_(text) {
  var url = 'https://api.line.me/v2/bot/message/push';
  var payload = {
    to: BOT_CONFIG.LINE_GROUP_ID,
    messages: [{ type: 'text', text: text }]
  };
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + BOT_CONFIG.LINE_CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

// ==================== HELPER FUNCTIONS ====================
function getSheetData_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      // Date object → ตรวจว่ามีเวลาหรือไม่
      if (val instanceof Date) {
        var h = val.getHours(), m = val.getMinutes(), s = val.getSeconds();
        if (h === 0 && m === 0 && s === 0) {
          obj[headers[j]] = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          obj[headers[j]] = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        }
      } else {
        obj[headers[j]] = val !== undefined ? String(val) : '';
      }
    }
    results.push(obj);
  }
  return results;
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatThaiDate_(date) {
  var thaiDays = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];
  var thaiMonths = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  var dow = date.getDay();
  var d = parseInt(Utilities.formatDate(date, Session.getScriptTimeZone(), 'd'));
  var m = parseInt(Utilities.formatDate(date, Session.getScriptTimeZone(), 'M')) - 1;
  var y = parseInt(Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy')) + 543;
  return 'วัน' + thaiDays[dow] + 'ที่ ' + d + ' ' + thaiMonths[m] + ' ' + y;
}

function formatThaiDateTime_(date) {
  var thaiMonths = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  var d = parseInt(Utilities.formatDate(date, Session.getScriptTimeZone(), 'd'));
  var m = parseInt(Utilities.formatDate(date, Session.getScriptTimeZone(), 'M')) - 1;
  var y = parseInt(Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy')) + 543;
  var time = Utilities.formatDate(date, Session.getScriptTimeZone(), 'HH:mm');
  return d + ' ' + thaiMonths[m] + ' ' + y + ' เวลา ' + time + ' น.';
}

// ==================== TRIGGER SETUP ====================
// รันฟังก์ชันนี้ครั้งเดียวเพื่อตั้ง Trigger อัตโนมัติ
function setupTriggers() {
  // ลบ trigger เดิมทั้งหมดของโปรเจคนี้
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });

  // ตั้ง Trigger ช่วงเช้า 08:30
  ScriptApp.newTrigger('sendMorningReport')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .nearMinute(30)
    .create();

  // ตั้ง Trigger ช่วงเย็น 17:30
  ScriptApp.newTrigger('sendEveningReport')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .nearMinute(30)
    .create();

  Logger.log('✅ ตั้ง Trigger สำเร็จ: เช้า 08:30 + เย็น 17:30');
}
