/**
 * EduAlimun Exam Portal — Google Apps Script backend (v4)
 *
 * শিক্ষার্থী নিজের মোবাইল নম্বর দিয়ে ঢোকে (কোড লাগে না)।
 * - এডমিন "নিবন্ধিত শিক্ষার্থী" তালিকা দিলে: শুধু ওই নম্বরগুলো ঢুকতে পারবে।
 *   (তালিকা খালি থাকলে: যেকোনো নম্বর চলবে — শুধু একই নম্বর/ডিভাইসে দ্বিতীয়বার আটকাবে)
 * - একই নম্বর বা একই ডিভাইস থেকে একটি পরীক্ষা একবারই।
 *
 * শিট (অটো তৈরি হবে): Exam, Leaderboard, ExamAttempts, Students
 */

// ⚠️ নিজের পাসওয়ার্ড বসান (এটা শুধু সার্ভারে থাকে, পেজের কোডে নয়)
var ADMIN_PASSWORD = 'EduAlimun123@#';

// একই ডিভাইস থেকে দ্বিতীয়জনকে আটকাবে কিনা (ভাইবোন একই ফোন ব্যবহার করলে false করুন)
var BLOCK_SAME_DEVICE = true;
// একই নামে দ্বিতীয়জনকে আটকাবে কিনা (একই নামের দুজন শিক্ষার্থী থাকলে false রাখুন)
var BLOCK_SAME_NAME = false;

var EXAM_HEADERS = ['examId', 'subject', 'duration', 'examDate', 'questions'];
var LB_HEADERS = ['name', 'score', 'wrong', 'phone', 'time'];
var STU_HEADERS = ['phone', 'name', 'createdAt'];
var ATT_HEADERS = ['examId', 'phone', 'deviceId', 'name', 'startedAt', 'submitted'];

/* ---------- helpers ---------- */
function sheet_(name, headers, plainText) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    // ফোন নম্বরের শুরুর 0 যেন হারিয়ে না যায়
    if (plainText) sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');
    sh.appendRow(headers);
  }
  return sh;
}
function examSheet_() { return sheet_('Exam', EXAM_HEADERS, true); }
function lbSheet_() { return sheet_('Leaderboard', LB_HEADERS, true); }
function stuSheet_() { return sheet_('Students', STU_HEADERS, true); }
function attSheet_() { return sheet_('ExamAttempts', ATT_HEADERS, true); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function isAdmin_(pw) { return String(pw || '') === ADMIN_PASSWORD; }

function readRows_(sh, cols) {
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues();
}
function clearRows_(sh) {
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// বাংলা সংখ্যা → ইংরেজি, +880 / 880 বাদ, শুধু 01XXXXXXXXX গ্রহণযোগ্য
function normPhone_(raw) {
  var bn = '০১২৩৪৫৬৭৮৯', s = '';
  String(raw || '').split('').forEach(function (ch) {
    var i = bn.indexOf(ch);
    s += i >= 0 ? String(i) : ch;
  });
  s = s.replace(/\D/g, '');
  if (s.indexOf('880') === 0 && s.length === 13) s = s.substring(2);
  if (s.length === 10 && s.charAt(0) === '1') s = '0' + s;
  return /^01[3-9]\d{8}$/.test(s) ? s : '';
}
function normName_(n) { return String(n || '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function cleanName_(n) { return String(n || '').replace(/[<>&"'`]/g, '').replace(/\s+/g, ' ').trim().substring(0, 40); }

function getExam_() {
  var rows = readRows_(examSheet_(), 5);
  if (!rows.length) return { examId: '', subject: '', duration: 0, examDate: '', questions: [] };
  var r = rows[0], qs = [];
  try { qs = JSON.parse(r[4] || '[]'); } catch (err) { qs = []; }
  return {
    examId: String(r[0]), subject: String(r[1]), duration: Number(r[2]),
    examDate: String(r[3]), questions: qs
  };
}
function getLeaderboard_() {
  // ফোন নম্বর পাবলিক ডাটায় যায় না
  return readRows_(lbSheet_(), 3).map(function (r) {
    return { name: String(r[0]), score: Number(r[1]), wrong: Number(r[2]) };
  });
}

/* ---------- GET ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || 'getData';

  if (action === 'getData') return json_({ exam: getExam_(), leaderboard: getLeaderboard_() });
  if (action === 'start') return json_(startAttempt_(p));

  if (action === 'verify') return json_({ ok: isAdmin_(p.password) });
  if (action === 'listAttempts') {
    if (!isAdmin_(p.password)) return json_({ ok: false, reason: 'unauthorized' });
    return json_(overview_());
  }
  return json_({ ok: false, reason: 'unknown_action' });
}

/* পরীক্ষা শুরুর সময় — যাচাই + রেজিস্টার (একসাথে দুজন চাপলেও লক থাকায় নিরাপদ) */
function startAttempt_(p) {
  var phone = normPhone_(p.phone);
  var name = cleanName_(p.name);
  var deviceId = String(p.deviceId || '').trim();
  if (!name) return { ok: false, reason: 'bad_name' };
  if (!phone) return { ok: false, reason: 'bad_phone' };
  if (!deviceId) return { ok: false, reason: 'bad_request' };

  return withLock_(function () {
    var exam = getExam_();
    if (!exam.examId || !exam.questions.length) return { ok: false, reason: 'noexam' };
    if (p.examId && String(p.examId) !== exam.examId) return { ok: false, reason: 'stale' };

    // নিবন্ধিত তালিকা থাকলে শুধু সেই নম্বরই চলবে, আর নাম আসবে তালিকা থেকে
    var reg = readRows_(stuSheet_(), 3);
    if (reg.length) {
      var hit = null;
      for (var k = 0; k < reg.length; k++) { if (String(reg[k][0]) === phone) { hit = reg[k]; break; } }
      if (!hit) return { ok: false, reason: 'not_registered' };
      name = cleanName_(hit[1]) || name;
    }

    var nameN = normName_(name);
    var dup = readRows_(attSheet_(), 6).some(function (r) {
      if (String(r[0]) !== exam.examId) return false;
      if (String(r[1]) === phone) return true;
      if (BLOCK_SAME_DEVICE && String(r[2]) === deviceId) return true;
      if (BLOCK_SAME_NAME && normName_(r[3]) === nameN) return true;
      return false;
    });
    if (dup) return { ok: false, reason: 'already' };

    attSheet_().appendRow([exam.examId, phone, deviceId, name, new Date().toISOString(), '']);
    return { ok: true, name: name };
  });
}

/* এডমিন প্যানেলের তালিকা: নিবন্ধন চালু থাকলে নিবন্ধিতদের, নইলে যারা শুরু করেছে তাদের */
function overview_() {
  var exam = getExam_();
  var att = readRows_(attSheet_(), 6).filter(function (r) { return String(r[0]) === exam.examId; });
  function statusOf(phone) {
    var st = '';
    att.forEach(function (r) { if (String(r[1]) === phone) st = String(r[5]) === 'yes' ? 'submitted' : 'started'; });
    return st;
  }
  var reg = readRows_(stuSheet_(), 3);
  if (reg.length) {
    return { ok: true, registryOn: true, rows: reg.map(function (r) {
      return { name: String(r[1]), phone: String(r[0]), status: statusOf(String(r[0])), registered: true };
    }) };
  }
  return { ok: true, registryOn: false, rows: att.map(function (r) {
    return { name: String(r[3]), phone: String(r[1]), status: String(r[5]) === 'yes' ? 'submitted' : 'started', registered: false };
  }) };
}

/* ---------- POST ---------- */
function doPost(e) {
  var d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, reason: 'bad_json' }); }

  if (d.action === 'submit') return json_(submit_(d));

  if (!isAdmin_(d.password)) return json_({ ok: false, reason: 'unauthorized' });
  if (d.action === 'publish') return json_(publish_(d));
  if (d.action === 'reset') return json_(reset_());
  if (d.action === 'resetAttempt') return json_(resetAttempt_(d));
  if (d.action === 'addStudents') return json_(addStudents_(d));
  if (d.action === 'deleteStudent') return json_(deleteStudent_(d));
  if (d.action === 'clearStudents') return json_(withLock_(function () { clearRows_(stuSheet_()); return { ok: true }; }));
  return json_({ ok: false, reason: 'unknown_action' });
}

function publish_(d) {
  return withLock_(function () {
    var sh = examSheet_();
    clearRows_(sh);
    sh.appendRow([
      String(Date.now()),          // নতুন examId — নতুন পরীক্ষায় সবাই আবার ১ বার সুযোগ পাবে
      d.subject || '', String(Number(d.duration) || 0), d.examDate || '',
      JSON.stringify(d.questions || [])
    ]);
    return { ok: true };
  });
}

function reset_() {
  return withLock_(function () {
    clearRows_(examSheet_());
    clearRows_(lbSheet_());
    clearRows_(attSheet_());
    return { ok: true };
  });
}

function addStudents_(d) {
  return withLock_(function () {
    var sh = stuSheet_(), have = {};
    readRows_(sh, 3).forEach(function (r) { have[String(r[0])] = true; });
    var now = new Date().toISOString(), rows = [];
    (d.students || []).forEach(function (st) {
      var phone = normPhone_(st.phone);
      if (!phone || have[phone]) return;           // ভুল নম্বর / ডুপ্লিকেট বাদ
      have[phone] = true;
      rows.push([phone, cleanName_(st.name) || phone, now]);
    });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
    return { ok: true, added: rows.length };
  });
}

function deleteStudent_(d) {
  return withLock_(function () {
    var phone = normPhone_(d.phone), sh = stuSheet_(), rows = readRows_(sh, 3);
    for (var i = rows.length - 1; i >= 0; i--) { if (String(rows[i][0]) === phone) sh.deleteRow(i + 2); }
    return { ok: true };
  });
}

/* কারো নেট/ফোনের সমস্যা হলে এডমিন তাকে চলতি পরীক্ষায় আবার সুযোগ দিতে পারবে */
function resetAttempt_(d) {
  return withLock_(function () {
    var exam = getExam_(), phone = normPhone_(d.phone);
    if (!phone) return { ok: false, reason: 'bad_phone' };

    var att = attSheet_(), rows = readRows_(att, 6);
    for (var i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][0]) === exam.examId && String(rows[i][1]) === phone) att.deleteRow(i + 2);
    }
    var lb = lbSheet_(), lrows = readRows_(lb, 4);
    for (var j = lrows.length - 1; j >= 0; j--) {
      if (String(lrows[j][3]) === phone) lb.deleteRow(j + 2);
    }
    return { ok: true };
  });
}

function submit_(d) {
  var phone = normPhone_(d.phone);
  var deviceId = String(d.deviceId || '').trim();
  if (!phone || !deviceId) return { ok: false, reason: 'bad_request' };

  return withLock_(function () {
    var exam = getExam_();
    if (!exam.examId) return { ok: false, reason: 'noexam' };
    if (d.examId && String(d.examId) !== exam.examId) return { ok: false, reason: 'stale' };

    var att = attSheet_(), rows = readRows_(att, 6), idx = -1;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === exam.examId && String(rows[i][1]) === phone && String(rows[i][2]) === deviceId) { idx = i; break; }
    }
    if (idx === -1) return { ok: false, reason: 'no_attempt' };            // শুরুই করেনি
    if (String(rows[idx][5]) === 'yes') return { ok: false, reason: 'already_submitted' };

    att.getRange(idx + 2, 6).setValue('yes');
    // নাম আসে শুরুর সময় রেজিস্টার করা নাম থেকে
    lbSheet_().appendRow([rows[idx][3], Number(d.score) || 0, Number(d.wrong) || 0, phone, new Date().toISOString()]);
    return { ok: true };
  });
}
