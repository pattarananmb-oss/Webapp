/* Kru Ploy Tech Classroom - Code.gs */
// ===================== CONFIG =====================
var CFG = {
  DB_KEY: 'KPTC_DB_ID',
  DRIVE_KEY: 'KPTC_DRIVE_ID',
  DB_NAME: 'Kru Ploy Tech Classroom - Database',
  TTL: 21600,
  DEFAULT_SESSION_MIN: 360,
  ADMIN_USER: 'P@ttAdmin',
  ADMIN_PASS: 'Kruploy@2025'
};

// ===================== SCHEMA =====================
var SCHEMA = {
  Users: ['userId','username','password','fullName','email','role','studentId','className','phone','avatar','active','mustChange','createdAt','updatedAt'],
  Roles: ['role','description'],
  Permissions: ['role','permission'],
  Topics: ['topicId','name','description','active','createdAt'],
  Lessons: ['lessonId','topicId','title','content','sortOrder','active','createdAt'],
  Questions: ['questionId','topicId','lessonId','question','choiceA','choiceB','choiceC','choiceD','answer','explanation','difficulty','active','createdAt'],
  Exams: ['examId','title','topicId','durationMin','questionCount','passPercent','active','createdBy','createdAt'],
  ExamQuestions: ['examId','questionId'],
  ExamSessions: ['sessionId','examId','studentId','startAt','endAt','status','createdAt'],
  ExamAnswers: ['sessionId','questionId','choice','savedAt'],
  Attempts: ['attemptId','sessionId','examId','studentId','score','total','correct','wrong','percent','passed','submittedAt'],
  AttemptAnswers: ['attemptId','questionId','choice','correctAnswer','isCorrect'],
  Documents: ['documentId','title','description','topicId','fileId','fileUrl','uploadedBy','active','createdAt'],
  AuditLogs: ['timestamp','userId','action','entity','entityId','details','userAgent'],
  SystemSettings: ['key','value']
};

var PERMS = {
  admin: ['VIEW_CONTENT','MANAGE_CONTENT','MANAGE_QUESTIONS','IMPORT_QUESTIONS','MANAGE_DOCUMENTS','VIEW_RESULTS','VIEW_ANALYTICS','MANAGE_USERS','MANAGE_SETTINGS','VIEW_AUDIT','CHANGE_OWN_PASSWORD'],
  teacher: ['VIEW_CONTENT','MANAGE_CONTENT','MANAGE_QUESTIONS','IMPORT_QUESTIONS','MANAGE_DOCUMENTS','VIEW_RESULTS','VIEW_ANALYTICS','CHANGE_OWN_PASSWORD'],
  student: ['VIEW_CONTENT','START_EXAM','VIEW_OWN_RESULT','CHANGE_OWN_PASSWORD']
};

var ENT = {
  topic:    {sh:'Topics',    id:'topicId',    pre:'T', perm:'MANAGE_CONTENT',  children:[['Lessons','topicId'],['Questions','topicId'],['Exams','topicId'],['Documents','topicId']]},
  lesson:   {sh:'Lessons',   id:'lessonId',   pre:'L', perm:'MANAGE_CONTENT',  children:[['Questions','lessonId']], parents:[['Topics','topicId']]},
  question: {sh:'Questions', id:'questionId', pre:'Q', perm:'MANAGE_QUESTIONS', children:[['ExamQuestions','questionId']], parents:[['Topics','topicId']]},
  exam:     {sh:'Exams',     id:'examId',     pre:'E', perm:'MANAGE_CONTENT',  children:[['ExamSessions','examId']], parents:[['Topics','topicId']]},
  user:     {sh:'Users',     id:'userId',     pre:'U', perm:'MANAGE_USERS',    children:[]}
};

function doGet() {

  safeRun_('doGet', function () { ensureInstalled_(); });

  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Kru Ploy Tech Classroom')

    .addMetaTag('viewport', 'width=device-width, initial-scale=1');

}



// ===================== UTILITIES / ERRORS =====================

function now_() { return new Date().toISOString(); }

function newId_(p) { return p + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase(); }

function s_(v) { return v === null || v === undefined ? '' : String(v).trim(); }

function isActive_(v) {

  var x = s_(v).toLowerCase();

  return ['false','0','no','n','disabled','inactive','draft','ปิด','ปิดใช้งาน'].indexOf(x) < 0;

}

function logError(where, e) {

  try { audit_('', 'ERROR', where, '', String(e && e.stack || e)); } catch (x) {}

  console.error(where + ': ' + (e && e.stack || e));

}

function handleError(where, e) { logError(where, e); return { ok: false, error: String(e && e.message || e) }; }

function safeRun_(where, fn) { try { return { ok: true, data: fn() }; } catch (e) { return handleError(where, e); } }

function fail_(m) { throw new Error(m); }



// ===================== DATABASE =====================

function props_() { return PropertiesService.getScriptProperties(); }

function db_() {

  var id = props_().getProperty(CFG.DB_KEY);

  if (!id) fail_('ยังไม่ได้ติดตั้งระบบ: รัน setupSystem() ก่อน');

  return SpreadsheetApp.openById(id);

}

function sheet_(n) {

  var sh = db_().getSheetByName(n);

  if (!sh) fail_('ไม่พบ Sheet: ' + n);

  return sh;

}

function rows_(n) {

  var sh = sheet_(n), v = sh.getDataRange().getValues();

  if (v.length < 2) return [];

  var h = v[0], out = [];

  for (var i = 1; i < v.length; i++) {

    var o = {}, empty = true;

    for (var j = 0; j < h.length; j++) {

      var c = v[i][j]; if (c instanceof Date) c = c.toISOString();

      o[h[j]] = c; if (c !== '') empty = false;

    }

    if (!empty) out.push(o);

  }

  return out;

}

function add_(n, o) {

  var h = SCHEMA[n];

  sheet_(n).appendRow(h.map(function (k) { return o[k] === undefined ? '' : o[k]; }));

}

function addMany_(n, list) {

  if (!list.length) return;

  var h = SCHEMA[n], sh = sheet_(n);

  var data = list.map(function (o) { return h.map(function (k) { return o[k] === undefined ? '' : o[k]; }); });

  sh.getRange(sh.getLastRow() + 1, 1, data.length, h.length).setValues(data);

}

function findRow_(n, field, id) {

  var sh = sheet_(n), h = SCHEMA[n], col = h.indexOf(field);

  var last = sh.getLastRow(); if (last < 2) return 0;

  var vals = sh.getRange(2, col + 1, last - 1, 1).getValues();

  for (var i = 0; i < vals.length; i++) if (s_(vals[i][0]) === s_(id)) return i + 2;

  return 0;

}

function update_(n, field, id, patch) {

  var r = findRow_(n, field, id); if (!r) fail_('ไม่พบข้อมูล ' + id);

  var sh = sheet_(n), h = SCHEMA[n], cur = sh.getRange(r, 1, 1, h.length).getValues()[0];

  h.forEach(function (k, i) { if (patch[k] !== undefined) cur[i] = patch[k]; });

  sh.getRange(r, 1, 1, h.length).setValues([cur]);

}

function delete_(n, field, id) { var r = findRow_(n, field, id); if (r) sheet_(n).deleteRow(r); }

function count_(n) { return Math.max(0, sheet_(n).getLastRow() - 1); }



// ===================== INSTALL =====================

function ensureInstalled_() {

  var p = props_();

  if (p.getProperty(CFG.DB_KEY) && p.getProperty('KPTC_INSTALLED') === '1') return false;

  setupSystem(); return true;

}

function setupSystem() {

  var lock = LockService.getScriptLock(); lock.waitLock(30000);

  try {

    var p = props_(), id = p.getProperty(CFG.DB_KEY), fresh = false, ss;

    if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { fail_('เปิด DB เดิมไม่ได้ (' + id + ') ไม่สร้างใหม่เพื่อกันข้อมูลหาย: ' + e.message); } }

    else { ss = SpreadsheetApp.create(CFG.DB_NAME); p.setProperty(CFG.DB_KEY, ss.getId()); fresh = true; }

    // sheets + headers (never clears existing data)

    Object.keys(SCHEMA).forEach(function (n) {

      var sh = ss.getSheetByName(n);

      if (!sh) {

        sh = ss.insertSheet(n);

        sh.getRange(1, 1, sh.getMaxRows(), SCHEMA[n].length).setNumberFormat('@');

      }

      if (sh.getLastRow() < 1) {

        sh.getRange(1, 1, 1, SCHEMA[n].length).setValues([SCHEMA[n]]).setFontWeight('bold').setBackground('#e0f2fe');

        sh.setFrozenRows(1);

      }

    });

    var def = ss.getSheetByName('Sheet1'); if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

    // Drive folder

    if (!p.getProperty(CFG.DRIVE_KEY)) p.setProperty(CFG.DRIVE_KEY, DriveApp.createFolder('Kru Ploy Tech Classroom - Files').getId());

    // seed once

    if (!getSetting_('seeded')) {

      seed_();

      setSetting_('seeded', now_());

    }

    p.setProperty('KPTC_INSTALLED', '1');

    return systemHealth_();

  } finally { lock.releaseLock(); }

}

function getSetting_(k) { var r = rows_('SystemSettings').filter(function (x) { return x.key === k; })[0]; return r ? r.value : ''; }

function setSetting_(k, v) { var r = findRow_('SystemSettings', 'key', k); if (r) update_('SystemSettings', 'key', k, { value: v }); else add_('SystemSettings', { key: k, value: v }); }



// ===================== SEED (first install only) =====================

function seed_() {

  var t = now_();

  addMany_('Roles', [{ role: 'admin', description: 'ผู้ดูแลระบบ' }, { role: 'teacher', description: 'ครูผู้สอน' }, { role: 'student', description: 'นักเรียน' }]);

  var pr = []; Object.keys(PERMS).forEach(function (r) { PERMS[r].forEach(function (x) { pr.push({ role: r, permission: x }); }); });

  addMany_('Permissions', pr);

  var topics = ['Hardware','Operating System','Microsoft Word','Microsoft Excel','Network','Cybersecurity','AI & Generative AI','Internet & Cloud','Database & SQL','Programming & Algorithm'];

  var desc = ['อุปกรณ์คอมพิวเตอร์','ระบบปฏิบัติการ','โปรแกรมประมวลผลคำ','โปรแกรมตารางคำนวณ','เครือข่ายคอมพิวเตอร์','ความปลอดภัยไซเบอร์','ปัญญาประดิษฐ์และ Generative AI','อินเทอร์เน็ตและคลาวด์','ฐานข้อมูลและ SQL','การเขียนโปรแกรมและอัลกอริทึม'];

  var TL = [], LL = [];

  topics.forEach(function (n, i) {

    var id = 'T' + ('0' + (i + 1)).slice(-2);

    TL.push({ topicId: id, name: n, description: desc[i], active: 'TRUE', createdAt: t });

    LL.push({ lessonId: 'L' + ('0' + (i + 1)).slice(-2), topicId: id, title: 'บทนำ: ' + n, content: 'เนื้อหาเบื้องต้นของหัวข้อ ' + n + ' (ครูแก้ไขเนื้อหาได้ในเมนูบทเรียน)', sortOrder: i + 1, active: 'TRUE', createdAt: t });

  });

  addMany_('Topics', TL); addMany_('Lessons', LL);

  // topic|question|A|B|C|D|answer

  var Q = [

   'T01|หน่วยประมวลผลกลางคือข้อใด|CPU|RAM|SSD|GPU|A','T01|หน่วยความจำชั่วคราวคือ|ROM|RAM|HDD|DVD|B','T01|อุปกรณ์ใดเป็นอุปกรณ์นำเข้า|จอภาพ|ลำโพง|คีย์บอร์ด|เครื่องพิมพ์|C',

   'T02|ข้อใดคือระบบปฏิบัติการ|Windows|Excel|Chrome|Zoom|A','T02|นามสกุลไฟล์ที่เรียกใช้โปรแกรมใน Windows คือ|.txt|.exe|.jpg|.pdf|B','T02|ปุ่มลัดคัดลอกคือ|Ctrl+V|Ctrl+X|Ctrl+C|Ctrl+Z|C',

   'T03|ปุ่มลัดบันทึกไฟล์ใน Word คือ|Ctrl+S|Ctrl+P|Ctrl+N|Ctrl+O|A','T03|ตัวหนาใน Word ใช้ปุ่ม|Ctrl+I|Ctrl+B|Ctrl+U|Ctrl+K|B','T03|การจัดข้อความกึ่งกลางใช้ปุ่ม|Ctrl+L|Ctrl+R|Ctrl+E|Ctrl+J|C',

   'T04|สูตรหาผลรวมใน Excel คือ|=SUM()|=AVG()|=MAX()|=MIN()|A','T04|สูตรหาค่าเฉลี่ยคือ|=SUM()|=COUNT()|=AVERAGE()|=IF()|C','T04|ข้อมูลใน 1 ช่องเรียกว่า|Row|Cell|Sheet|Book|B',

   'T05|อุปกรณ์เชื่อมต่อเครือข่ายไร้สายคือ|Switch|Router|Hub|Modem เท่านั้น|B','T05|LAN ย่อมาจาก|Local Area Network|Long Area Net|Large Area Network|Link Area Node|A','T05|IP Address ใช้ทำอะไร|ระบุที่อยู่อุปกรณ์|เก็บไฟล์|ป้องกันไวรัส|ตกแต่งเว็บ|A',

   'T06|รหัสผ่านที่ดีควร|ใช้วันเกิด|สั้นและจำง่าย|ยาวและผสมอักขระ|ใช้ซ้ำทุกเว็บ|C','T06|Phishing คือ|การหลอกขโมยข้อมูล|การสำรองข้อมูล|การเข้ารหัส|การอัปเดต|A','T06|2FA คือ|ยืนยันตัวตน 2 ชั้น|เครื่องสแกนไวรัส|ไฟร์วอลล์|เว็บเบราว์เซอร์|A',

   'T07|AI ย่อมาจาก|Artificial Intelligence|Auto Internet|Advanced Input|Active Interface|A','T07|Generative AI ทำอะไรได้|สร้างเนื้อหาใหม่|ซ่อมฮาร์ดแวร์|ชาร์จแบตเตอรี่|ลดไฟฟ้า|A','T07|Prompt คืออะไร|คำสั่งที่ให้ AI|ชื่อไฟล์|รหัสผ่าน|ชนิดจอภาพ|A',

   'T08|URL คือ|ที่อยู่เว็บไซต์|ชื่อผู้ใช้|ไวรัส|ภาษาโปรแกรม|A','T08|Cloud Storage ตัวอย่างคือ|Google Drive|Notepad|Paint|Calculator|A','T08|เบราว์เซอร์คือ|Chrome|Word|Excel|Scratch|A',

   'T09|SQL ใช้สำหรับ|จัดการฐานข้อมูล|ตัดต่อวิดีโอ|วาดภาพ|เล่นเกม|A','T09|คำสั่งดึงข้อมูลคือ|INSERT|SELECT|DELETE|DROP|B','T09|Primary Key คือ|ค่าที่ไม่ซ้ำในตาราง|รหัสผ่าน|ชื่อตาราง|ชนิดไฟล์|A',

   'T10|ลำดับขั้นตอนแก้ปัญหาเรียกว่า|อัลกอริทึม|ฮาร์ดแวร์|เครือข่าย|ไวรัส|A','T10|โครงสร้างวนซ้ำคือ|Loop|If|Print|Input|A','T10|ตัวแปรใช้สำหรับ|เก็บค่าข้อมูล|พิมพ์เอกสาร|ต่ออินเทอร์เน็ต|ปิดเครื่อง|A'

  ];

  var QL = Q.map(function (r, i) { var a = r.split('|');

    return { questionId: 'Q' + ('00' + (i + 1)).slice(-3), topicId: a[0], lessonId: 'L' + a[0].substring(1), question: a[1], choiceA: a[2], choiceB: a[3], choiceC: a[4], choiceD: a[5], answer: a[6], explanation: '', difficulty: 'easy', active: 'TRUE', createdAt: t }; });

  addMany_('Questions', QL);

  var ex = [], eq = [];

  topics.forEach(function (n, i) {

    var id = 'E' + ('0' + (i + 1)).slice(-2), tid = 'T' + ('0' + (i + 1)).slice(-2);

    ex.push({ examId: id, title: 'แบบทดสอบ ' + n, topicId: tid, durationMin: 10, questionCount: 3, passPercent: 60, active: 'TRUE', createdBy: 'seed', createdAt: t });

    QL.filter(function (q) { return q.topicId === tid; }).forEach(function (q) { eq.push({ examId: id, questionId: q.questionId }); });

  });

  addMany_('Exams', ex); addMany_('ExamQuestions', eq);

  add_('Users', { userId: 'U_ADMIN', username: CFG.ADMIN_USER, password: hash_(CFG.ADMIN_PASS), fullName: 'Master Admin', email: 'admin@kptc.local', role: 'admin', active: 'TRUE', mustChange: 'TRUE', createdAt: t, updatedAt: t });

  setSetting_('APP_TITLE','Kru Ploy Tech Classroom');

  setSetting_('APP_SUBTITLE','เข้าสู่ระบบห้องเรียนออนไลน์');

  setSetting_('FOOTER_TEXT','© 2026 Kru Ploy Tech Classroom. All rights reserved.');

  setSetting_('THEME','light');

  setSetting_('LANG','th');

  setSetting_('COPY_NOTICE','เนื้อหาบทเรียนและข้อสอบเป็นทรัพย์สินทางปัญญา ห้ามคัดลอก ทำซ้ำ บันทึกหน้าจอ หรือนำไปเผยแพร่โดยไม่ได้รับอนุญาต');

}



// ===================== AUTH =====================

function hash_(pw) {

  var salt = Utilities.getUuid().substring(0, 12);

  return salt + '$' + digest_(salt + pw);

}

function digest_(x) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, x, Utilities.Charset.UTF_8)); }

function verify_(pw, stored) { var p = s_(stored).split('$'); return p.length === 2 && digest_(p[0] + pw) === p[1]; }

function pub_(u) { var c = {}; for (var k in u) if (k !== 'password') c[k] = u[k]; return c; }

function login(username, password) {

  var key=s_(username).toLowerCase();

  var u = rows_('Users').filter(function (x) { return s_(x.username).toLowerCase() === key || s_(x.email).toLowerCase() === key; })[0];

  if (!u || !isActive_(u.active) || !verify_(password, u.password)) { audit_('', 'LOGIN_FAIL', 'User', s_(username), ''); fail_('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'); }

  var token = Utilities.getUuid() + Utilities.getUuid();

  CacheService.getScriptCache().put('S_' + token, u.userId, sessionTtlSeconds_());

  audit_(u.userId, 'LOGIN', 'User', u.userId, '');

  return { token: token, user: pub_(u), permissions: PERMS[u.role] || [] };

}

function sessionTtlMinutes_() {

  var n = Number(getSetting_('SESSION_TIMEOUT_MIN') || CFG.DEFAULT_SESSION_MIN / 60);

  if (!isFinite(n)) n = CFG.DEFAULT_SESSION_MIN / 60;

  return Math.max(5, Math.min(1440, Math.floor(n)));

}

function sessionTtlSeconds_() { return sessionTtlMinutes_() * 60; }



function session_(token) {

  var id = token && CacheService.getScriptCache().get('S_' + token);

  if (!id) fail_('SESSION_EXPIRED');

  var u = rows_('Users').filter(function (x) { return s_(x.userId) === id; })[0];

  if (!u || !isActive_(u.active)) fail_('SESSION_EXPIRED');

  return u;

}

function logout_(token, u) { CacheService.getScriptCache().remove('S_' + token); audit_(u.userId, 'LOGOUT', 'User', u.userId, ''); return true; }

function getSession_(u) { return { user: pub_(u), permissions: PERMS[u.role] || [] }; }



// ===================== PERMISSION =====================

function can_(u, perm) { return (PERMS[u.role] || []).indexOf(perm) >= 0; }

function need_(u, perm) { if (!can_(u, perm)) fail_('ไม่มีสิทธิ์: ' + perm); }



// ===================== API GATEWAY =====================

// Client calls only api(token, name, payload). Every call re-checks session + permission on the server.

function api(token, name, p) {

  return safeRun_('api:' + name, function () {

    p = p || {};

    if (name === 'login') return login(p.username, p.password);

    if (name === 'googleLogin') return googleLogin_(p.credential);

    if (name === 'getPublicSettings') return getPublicSettings_();

    var u = session_(token);

    switch (name) {

      case 'getSession': return getSession_(u);

      case 'logout': return logout_(token, u);

      case 'changePassword': return changePassword_(u, p);

      case 'getData': return getData_(u);

      case 'save': return saveEntity_(u, p);

      case 'delete': return deleteEntity_(u, p);

      case 'importPreview': need_(u, 'IMPORT_QUESTIONS'); return importQuestions_(u, p.rows, p.mode, true);

      case 'importQuestions': need_(u, 'IMPORT_QUESTIONS'); return importQuestions_(u, p.rows, p.mode, false);

      case 'startExam': need_(u, 'START_EXAM'); return startExam_(u, p.examId);

      case 'startMixedExam': need_(u, 'START_EXAM'); return startMixedExam_(u, p);

      case 'submitExam': need_(u, 'START_EXAM'); return submitExam_(u, p.sessionId, p.answers || {});

      case 'uploadDocument': need_(u, 'MANAGE_DOCUMENTS'); return uploadDocument_(u, p);

      case 'deleteDocument': need_(u, 'MANAGE_DOCUMENTS'); return deleteDocument_(u, p.documentId);

      case 'systemHealth': need_(u, 'MANAGE_SETTINGS'); return systemHealth_();

      case 'getSettings': need_(u, 'MANAGE_SETTINGS'); return getSettings_(u);

      case 'saveSetting': need_(u, 'MANAGE_SETTINGS'); return saveSetting_(u, p);

      case 'previewExam': need_(u, 'MANAGE_CONTENT'); return previewExam_(u, p.examId);

    }

    fail_('ไม่รู้จักคำสั่ง ' + name);

  });

}



// ===================== DATA LAYER =====================

function act_(list) { return list.filter(function (x) { return isActive_(x.active); }); }

function getData_(u) {

  var d = { role: u.role, topics: [], lessons: [], documents: [], exams: [], questions: [], results: [] };

  var staff = can_(u, 'MANAGE_CONTENT');

  var topics = rows_('Topics'), lessons = rows_('Lessons'), docs = rows_('Documents'), exams = rows_('Exams');

  d.topics = staff ? topics : act_(topics);

  d.lessons = staff ? lessons : act_(lessons);

  d.documents = staff ? docs : act_(docs);

  d.exams = staff ? exams : act_(exams);

  var qs = rows_('Questions'), qcount = {};

  qs.forEach(function (q) { if (isActive_(q.active)) qcount[q.topicId] = (qcount[q.topicId] || 0) + 1; });

  d.questionCounts = qcount;

  if (can_(u, 'MANAGE_QUESTIONS')) d.questions = qs;

  d.results = getResults_(u);

  if (u.role === 'student') d.development = getStudentDevelopment_(u);

  if (can_(u, 'VIEW_RESULTS')) {

    d.students = rows_('Users').filter(function (x) { return x.role === 'student'; }).map(pub_);

    d.analytics = getAnalytics_(u);

  }

  if (can_(u, 'MANAGE_USERS')) d.users = rows_('Users').map(pub_);

  if (can_(u, 'VIEW_AUDIT')) d.audit = rows_('AuditLogs').slice(-300).reverse();

  if (can_(u, 'MANAGE_SETTINGS')) d.health = systemHealth_();

  return d;

}

function getResults_(u) {

  var at = rows_('Attempts'), ex = {}, us = {};

  rows_('Exams').forEach(function (e) { ex[e.examId] = e.title; });

  rows_('Users').forEach(function (x) { us[x.userId] = x.fullName || x.username; });

  if (u.role === 'student') at = at.filter(function (a) { return s_(a.studentId) === s_(u.userId); });

  else need_(u, 'VIEW_RESULTS');

  return at.map(function (a) { a.examTitle = ex[a.examId] || a.examId; a.studentName = us[a.studentId] || a.studentId; return a; }).reverse();

}

function getStudentDevelopment_(u) {

  var attempts = rows_('Attempts').filter(function(a){ return s_(a.studentId)===s_(u.userId); });

  var ids = {}; attempts.forEach(function(a){ ids[s_(a.attemptId)] = 1; });

  var answers = rows_('AttemptAnswers').filter(function(a){ return ids[s_(a.attemptId)] && String(a.isCorrect).toUpperCase()!=='TRUE'; });

  var qs = {}; rows_('Questions').forEach(function(q){ qs[s_(q.questionId)] = q; });

  var topics = {}, lessons = {};

  answers.forEach(function(a){ var q=qs[s_(a.questionId)]; if(!q) return; var tk=s_(q.topicId); topics[tk]=(topics[tk]||0)+1; var lk=s_(q.lessonId); if(lk) lessons[lk]=(lessons[lk]||0)+1; });

  var tnames={}, lnames={}; rows_('Topics').forEach(function(x){tnames[s_(x.topicId)]=s_(x.name)}); rows_('Lessons').forEach(function(x){lnames[s_(x.lessonId)]={title:s_(x.title),topicId:s_(x.topicId)}});

  var totalWrong=answers.length;

  return {totalWrong:totalWrong, byTopic:Object.keys(topics).map(function(k){return {topicId:k,name:tnames[k]||k,wrong:topics[k],percent:totalWrong?Math.round(topics[k]/totalWrong*1000)/10:0};}).sort(function(a,b){return b.wrong-a.wrong;}), byLesson:Object.keys(lessons).map(function(k){var x=lnames[k]||{};return {lessonId:k,title:x.title||k,topicId:x.topicId||'',wrong:lessons[k]};}).sort(function(a,b){return b.wrong-a.wrong;}).slice(0,10)};

}



function getAnalytics_(u) {

  need_(u, 'VIEW_ANALYTICS');

  var at = rows_('Attempts'), ex = {};

  rows_('Exams').forEach(function (e) { ex[e.examId] = e; });

  var pct = at.map(function (a) { return Number(a.percent) || 0; });

  var passed = at.filter(function (a) { return s_(a.passed).toUpperCase() === 'TRUE'; }).length;

  var byTopic = {}, byStudent = {};

  at.forEach(function (a) {

    var tid = ex[a.examId] ? ex[a.examId].topicId : '?';

    (byTopic[tid] = byTopic[tid] || []).push(Number(a.percent) || 0);

    (byStudent[a.studentId] = byStudent[a.studentId] || []).push(Number(a.percent) || 0);

  });

  var avg = function (l) { return l.length ? Math.round(l.reduce(function (a, b) { return a + b; }, 0) / l.length * 10) / 10 : 0; };

  var names = {}; rows_('Users').forEach(function (x) { names[x.userId] = x.fullName || x.username; });

  return { attempts: at.length, passed: passed, passRate: at.length ? Math.round(passed / at.length * 1000) / 10 : 0,

    avg: avg(pct), max: pct.length ? Math.max.apply(null, pct) : 0, min: pct.length ? Math.min.apply(null, pct) : 0,

    byTopic: Object.keys(byTopic).map(function (k) { return { topicId: k, count: byTopic[k].length, avg: avg(byTopic[k]) }; }),

    byStudent: Object.keys(byStudent).map(function (k) { return { studentId: k, name: names[k] || k, count: byStudent[k].length, avg: avg(byStudent[k]) }; }) };

}



// ===================== GENERIC CRUD (Topics/Lessons/Questions/Exams/Users) =====================

function saveEntity_(u, p) {

  var E = ENT[p.entity]; if (!E) fail_('entity ไม่ถูกต้อง');

  need_(u, E.perm);

  var o = p.obj || {}, t = now_(), id = s_(o[E.id]);

  if (p.entity === 'user') return saveUser_(u, o, p.isNew);

  if (!id && p.isNew) id = newId_(E.pre);

  if (!id) fail_('ไม่มี ID');

  (E.parents || []).forEach(function (pa) { if (s_(o[pa[1]]) && !findRow_(pa[0], pa[1], o[pa[1]])) fail_('ไม่พบ ' + pa[1] + ' = ' + o[pa[1]]); });

  if (p.entity === 'question') {

    ['question','choiceA','choiceB','choiceC','choiceD'].forEach(function (k) { if (!s_(o[k])) fail_('กรุณากรอก ' + k); });

    if (['A','B','C','D'].indexOf(s_(o.answer).toUpperCase()) < 0) fail_('answer ต้องเป็น A-D');

    o.answer = s_(o.answer).toUpperCase();

  }

  o[E.id] = id;

  if (p.isNew) {

    if (findRow_(E.sh, E.id, id)) fail_('ID ซ้ำ: ' + id);

    o.createdAt = t; if (p.entity === 'exam') o.createdBy = u.userId;

    if (o.active === undefined || o.active === '') o.active = 'TRUE';

    add_(E.sh, o);

    if (p.entity === 'exam') buildExamQuestions_(o);

  } else {

    update_(E.sh, E.id, id, o);

    if (p.entity === 'exam' && p.rebuild) buildExamQuestions_(o);

  }

  audit_(u.userId, (p.isNew ? 'CREATE_' : 'UPDATE_') + p.entity.toUpperCase(), E.sh, id, '');

  return { id: id };

}

function buildExamQuestions_(e) {

  var old = sheet_('ExamQuestions'), v = old.getDataRange().getValues();

  for (var i = v.length - 1; i >= 1; i--) if (s_(v[i][0]) === s_(e.examId)) old.deleteRow(i + 1);

  var pool = act_(rows_('Questions')).filter(function (q) { return s_(q.topicId) === s_(e.topicId); });

  pool.sort(function () { return Math.random() - 0.5; });

  var n = Number(e.questionCount) || pool.length;

  addMany_('ExamQuestions', pool.slice(0, n).map(function (q) { return { examId: e.examId, questionId: q.questionId }; }));

}

function deleteEntity_(u, p) {

  var E = ENT[p.entity]; if (!E) fail_('entity ไม่ถูกต้อง');

  need_(u, E.perm);

  if (p.entity === 'user' && s_(p.id) === s_(u.userId)) fail_('ลบบัญชีตัวเองไม่ได้');

  for (var i = 0; i < E.children.length; i++) {

    var c = E.children[i];

    if (p.entity === 'exam' && c[0] === 'ExamSessions') continue;

    if (rows_(c[0]).some(function (r) { return s_(r[c[1]]) === s_(p.id); })) fail_('ลบไม่ได้: ยังมีข้อมูลใน ' + c[0] + ' ที่อ้างถึง ' + p.id + ' (ลบข้อมูลลูกก่อน หรือปิดใช้งานแทน)');

  }

  if (p.entity === 'exam') { var eq = sheet_('ExamQuestions'), v = eq.getDataRange().getValues(); for (var j = v.length - 1; j >= 1; j--) if (s_(v[j][0]) === s_(p.id)) eq.deleteRow(j + 1); }

  delete_(E.sh, E.id, p.id);

  audit_(u.userId, 'DELETE_' + p.entity.toUpperCase(), E.sh, p.id, '');

  return true;

}

function saveUser_(u, o, isNew) {
  var t = now_(), un = s_(o.username), email = s_(o.email).toLowerCase();
  if (!/^[A-Za-z0-9._@\-]{3,40}$/.test(un)) fail_('username ต้องยาว 3-40 ตัว และใช้ได้เฉพาะ A-Z a-z 0-9 . _ - @');
  if (['admin','teacher','student'].indexOf(s_(o.role).toLowerCase()) < 0) fail_('role ไม่ถูกต้อง');
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) fail_('รูปแบบ Email ไม่ถูกต้อง');
  o.username = un;
  o.email = email;

  var all = rows_('Users');
  var dupU = all.some(function (x) { return s_(x.username).toLowerCase() === un.toLowerCase() && s_(x.userId) !== s_(o.userId); });
  var dupE = email && all.some(function (x) { return s_(x.email).toLowerCase() === email && s_(x.userId) !== s_(o.userId); });
  if (dupU) fail_('username ซ้ำ');
  if (dupE) fail_('email ซ้ำ');

  var existing = !isNew ? all.filter(function (x) { return s_(x.userId) === s_(o.userId); })[0] : null;
  if (!isNew && !existing) fail_('ไม่พบผู้ใช้');
  if (!isNew && s_(o.userId) === s_(u.userId)) {
    if (s_(o.role).toLowerCase() !== 'admin') fail_('ไม่สามารถลดสิทธิ์บัญชี Admin ของตนเองได้');
    if (String(o.active).toUpperCase() !== 'TRUE') fail_('ไม่สามารถปิดใช้งานบัญชี Admin ของตนเองได้');
  }

  var plain = '';
  if (isNew) {
    if (!o.password || s_(o.password).length < 8) fail_('รหัสผ่านต้องยาวอย่างน้อย 8 ตัวอักษร');
    plain = o.password;
    o.userId = newId_('U');
    o.password = hash_(o.password);
    o.mustChange = 'TRUE';
    o.createdAt = t;
    o.updatedAt = t;
    if (o.active === undefined || o.active === '') o.active = 'TRUE';
    add_('Users', o);
  } else {
    if (o.password) {
      if (s_(o.password).length < 8) fail_('รหัสผ่านต้องยาวอย่างน้อย 8 ตัวอักษร');
      plain = o.password;
      o.password = hash_(o.password);
      o.mustChange = 'TRUE';
    } else delete o.password;
    o.updatedAt = t;
    update_('Users', 'userId', o.userId, o);
  }

  audit_(u.userId, isNew ? 'CREATE_USER' : 'UPDATE_USER', 'Users', o.userId, o.role);
  var mail = { sent: false, error: '' };
  if (plain) {
    var to = email || (rows_('Users').filter(function (x) { return s_(x.userId) === s_(o.userId); })[0] || {}).email;
    if (!to) mail.error = 'ไม่มีอีเมลผู้ใช้';
    else mail = sendCredentials_(to, o.fullName, un, plain, isNew);
    audit_(u.userId, mail.sent ? 'EMAIL_SENT' : 'EMAIL_FAIL', 'Users', o.userId, mail.error);
  }
  return { id: o.userId, mailSent: mail.sent, mailError: mail.error };
}

function sendCredentials_(to, name, username, pass, isNew) {

  try {

    var url = ScriptApp.getService().getUrl();

    MailApp.sendEmail({ to: to, name: 'Kru Ploy Tech Classroom',

      subject: (isNew ? 'บัญชีใหม่' : 'รีเซ็ตรหัสผ่าน') + ' - Kru Ploy Tech Classroom',

      htmlBody: '<p>สวัสดี ' + (name || username) + '</p><p>Username: <b>' + username + '</b><br>รหัสผ่านชั่วคราว: <b>' + pass + '</b></p><p>เข้าสู่ระบบที่ <a href="' + url + '">' + url + '</a><br>ระบบจะให้เปลี่ยนรหัสผ่านทันทีหลังเข้าสู่ระบบครั้งแรก</p>' });

    return { sent: true, error: '' };

  } catch (e) { logError('sendCredentials_', e); return { sent: false, error: String(e.message || e) }; }

}

function changePassword_(u, p) {

  need_(u, 'CHANGE_OWN_PASSWORD');

  if (!verify_(p.oldPassword, u.password)) fail_('รหัสผ่านเดิมไม่ถูกต้อง');

  if (s_(p.newPassword).length < 8) fail_('รหัสผ่านใหม่ต้องยาวอย่างน้อย 8 ตัวอักษร');

  if (p.newPassword === p.oldPassword) fail_('รหัสผ่านใหม่ต้องต่างจากเดิม');

  update_('Users', 'userId', u.userId, { password: hash_(p.newPassword), mustChange: 'FALSE', updatedAt: now_() });

  audit_(u.userId, 'CHANGE_PASSWORD', 'Users', u.userId, '');

  return true;

}





// ===================== GOOGLE LOGIN =====================

function getPublicSettings_(){return {googleClientId:getSetting_('GOOGLE_CLIENT_ID'),theme:getSetting_('THEME')||'light',lang:getSetting_('LANG')||'th',title:getSetting_('APP_TITLE')||'Kru Ploy Tech Classroom',subtitle:getSetting_('APP_SUBTITLE')||'เข้าสู่ระบบห้องเรียนออนไลน์',footer:getSetting_('FOOTER_TEXT')||'© 2026 Kru Ploy Tech Classroom. All rights reserved.',copyNotice:getSetting_('COPY_NOTICE')||'',sessionTimeoutMin:sessionTtlMinutes_()};}

function googleLogin_(credential) {
  credential = s_(credential);
  if (!credential) fail_('ไม่พบ Google Credential');
  var cid = getSetting_('GOOGLE_CLIENT_ID');
  if (!cid) fail_('ผู้ดูแลระบบยังไม่ได้ตั้งค่า Google Client ID');
  var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(credential), {muteHttpExceptions:true});
  var data = JSON.parse(res.getContentText() || '{}');
  var issuer = s_(data.iss);
  if (res.getResponseCode() !== 200 || issuer !== 'https://accounts.google.com' || !data.email || String(data.email_verified).toLowerCase() !== 'true' || s_(data.aud) !== cid) fail_('ไม่สามารถยืนยันบัญชี Google ได้');
  var email = s_(data.email).toLowerCase();
  var u = rows_('Users').filter(function (x) { return s_(x.email).toLowerCase() === email; })[0];
  if (!u || !isActive_(u.active)) { audit_('', 'GOOGLE_LOGIN_REJECT', 'User', email, 'unknown_or_inactive'); fail_('อีเมล Google นี้ยังไม่ได้ลงทะเบียนในระบบ'); }
  var token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('S_' + token, u.userId, sessionTtlSeconds_());
  audit_(u.userId, 'GOOGLE_LOGIN', 'User', u.userId, email);
  return {token:token, user:pub_(u), permissions:PERMS[u.role] || []};
}

function getSettings_(u){var allowed=['APP_TITLE','APP_SUBTITLE','FOOTER_TEXT','THEME','LANG','COPY_NOTICE','GOOGLE_CLIENT_ID','SESSION_TIMEOUT_MIN'],out={};rows_('SystemSettings').forEach(function(x){if(allowed.indexOf(s_(x.key))>=0)out[x.key]=x.value;});return out;}

function saveSetting_(u,p){var allowed=['APP_TITLE','APP_SUBTITLE','FOOTER_TEXT','THEME','LANG','COPY_NOTICE','GOOGLE_CLIENT_ID','SESSION_TIMEOUT_MIN'],k=s_(p.key),v=s_(p.value);if(allowed.indexOf(k)<0)fail_('ไม่อนุญาตให้แก้ไข setting นี้');if(k==='THEME'&&['light','dark'].indexOf(v)<0)fail_('ธีมไม่ถูกต้อง');if(k==='LANG'&&['th','en'].indexOf(v)<0)fail_('ภาษาไม่ถูกต้อง');if(k==='SESSION_TIMEOUT_MIN'&&(Number(v)<5||Number(v)>1440))fail_('Session Timeout ต้องอยู่ระหว่าง 5-1440 นาที');setSetting_(k,v);audit_(u.userId,'UPDATE_SETTING','SystemSettings',k,'');return {key:k,value:v};}

function previewExam_(u,examId){var e=rows_('Exams').filter(function(x){return s_(x.examId)===s_(examId);})[0];if(!e)fail_('ไม่พบชุดสอบ');var ids=rows_('ExamQuestions').filter(function(x){return s_(x.examId)===s_(examId);}).map(function(x){return s_(x.questionId);});var qs=rows_('Questions').filter(function(q){return ids.indexOf(s_(q.questionId))>=0&&isActive_(q.active);});return {exam:e,questions:qs.map(function(q){return {questionId:q.questionId,question:q.question,choiceA:q.choiceA,choiceB:q.choiceB,choiceC:q.choiceC,choiceD:q.choiceD};})};}



// ===================== QUESTION IMPORT =====================

// rows come from the browser (CSV parsed client side; XLSX -> save as CSV). Validation/duplicates are server side.

function importQuestions_(u, rows, mode, dry) {

  rows = rows || []; mode = mode || 'skip';

  var topics = {}, lessons = {};

  rows_('Topics').forEach(function (x) { topics[s_(x.topicId)] = 1; });

  rows_('Lessons').forEach(function (x) { lessons[s_(x.lessonId)] = s_(x.topicId); });

  var exist = rows_('Questions'), byId = {}, byText = {};

  exist.forEach(function (q) { byId[s_(q.questionId)] = 1; byText[s_(q.topicId) + '|' + s_(q.question).toLowerCase()] = 1; });

  var seenId = {}, seenText = {}, res = { ok: [], errors: [], dups: [], total: rows.length }, toAdd = [], toUpd = [];

  rows.forEach(function (r, i) {

    var n = i + 2, er = [], o = {};

    ['questionId','topicId','lessonId','question','choiceA','choiceB','choiceC','choiceD','answer','explanation','difficulty','active'].forEach(function (k) { o[k] = s_(r[k]); });

    if (!topics[o.topicId]) er.push('topicId ไม่มีในระบบ');

    if (!o.lessonId) er.push('กรุณาระบุ lessonId ให้สัมพันธ์กับ topicId');

    else if (!lessons[o.lessonId] || lessons[o.lessonId] !== o.topicId) er.push('lessonId ไม่มี/ไม่ตรง topic');

    ['question','choiceA','choiceB','choiceC','choiceD'].forEach(function (k) { if (!o[k]) er.push('ขาด ' + k); });

    o.answer = o.answer.toUpperCase();

    if (['A','B','C','D'].indexOf(o.answer) < 0) er.push('answer ต้องเป็น A-D');

    o.difficulty = (o.difficulty || 'medium').toLowerCase();

    if (['easy','medium','hard'].indexOf(o.difficulty) < 0) er.push('difficulty ต้องเป็น easy/medium/hard');

    if (o.active === '') o.active = 'TRUE';

    if (['TRUE','FALSE'].indexOf(o.active.toUpperCase()) < 0) er.push('active ต้องเป็น TRUE/FALSE'); else o.active=o.active.toUpperCase();

    if (er.length) { res.errors.push({ row: n, message: er.join(', ') }); return; }

    var tk = o.topicId + '|' + o.question.toLowerCase();

    var isDup = (o.questionId && (byId[o.questionId] || seenId[o.questionId])) || byText[tk] || seenText[tk];

    if (o.questionId) seenId[o.questionId] = 1; seenText[tk] = 1;

    if (isDup) {

      res.dups.push({ row: n, questionId: o.questionId, question: o.question });

      if (mode === 'update' && o.questionId && byId[o.questionId]) toUpd.push(o);

      return;

    }

    if (!o.questionId) o.questionId = newId_('Q');

    o.createdAt = now_(); toAdd.push(o); res.ok.push({ row: n, questionId: o.questionId, question: o.question });

  });

  if (!dry) {

    addMany_('Questions', toAdd);

    toUpd.forEach(function (o) { update_('Questions', 'questionId', o.questionId, o); });

    res.imported = toAdd.length; res.updated = toUpd.length;

    audit_(u.userId, 'IMPORT_QUESTIONS', 'Questions', '', 'added=' + toAdd.length + ' updated=' + toUpd.length + ' errors=' + res.errors.length);

  }

  return res;

}



// ===================== EXAMS =====================

function startMixedExam_(u, p) {

  var topics=(p.topicIds||[]).map(function(x){return s_(x);}).filter(function(x){return x;});

  if(!topics.length) fail_('กรุณาเลือกอย่างน้อย 1 หมวด');

  var count=Math.max(1,Math.min(200,Number(p.questionCount)||20));

  var duration=Math.max(1,Math.min(300,Number(p.durationMin)||30));

  var pass=Math.max(0,Math.min(100,Number(p.passPercent)||50));

  var qs=rows_('Questions').filter(function(q){return topics.indexOf(s_(q.topicId))>=0&&isActive_(q.active);});

  if(qs.length<count) fail_('ข้อสอบไม่เพียงพอตามเงื่อนไขที่เลือก (มี '+qs.length+' ข้อ)');

  qs.sort(function(){return Math.random()-.5;}); qs=qs.slice(0,count);

  var ids=qs.map(function(q){return s_(q.questionId);});

  var sid=newId_('S'),start=new Date(),end=new Date(start.getTime()+duration*60000);

  var encoded='MIXED|'+duration+'|'+pass+'|'+ids.join(',');

  add_('ExamSessions',{sessionId:sid,examId:encoded,studentId:u.userId,startAt:start.toISOString(),endAt:end.toISOString(),status:'open',createdAt:now_()});

  audit_(u.userId,'START_MIXED_EXAM','ExamSessions',sid,topics.join(','));

  return {sessionId:sid,exam:{examId:encoded,title:'Mixed Topics Exam',topicId:'MIXED',durationMin:duration,questionCount:count,passPercent:pass},endAt:end.toISOString(),questions:qs.map(function(q){return {questionId:q.questionId,question:q.question,choiceA:q.choiceA,choiceB:q.choiceB,choiceC:q.choiceC,choiceD:q.choiceD};})};

}

function startExam_(u, examId) {

  var e = rows_('Exams').filter(function (x) { return s_(x.examId) === s_(examId); })[0];

  if (!e || !isActive_(e.active)) fail_('ไม่พบข้อสอบหรือปิดอยู่');

  var ids = rows_('ExamQuestions').filter(function (x) { return s_(x.examId) === s_(examId); }).map(function (x) { return s_(x.questionId); });

  var qs = rows_('Questions').filter(function (q) { return ids.indexOf(s_(q.questionId)) >= 0 && isActive_(q.active); });

  if (!qs.length) fail_('ข้อสอบนี้ยังไม่มีคำถาม');

  var dur = Number(e.durationMin) || 10, start = new Date(), end = new Date(start.getTime() + dur * 60000);

  var sid = newId_('S');

  add_('ExamSessions', { sessionId: sid, examId: examId, studentId: u.userId, startAt: start.toISOString(), endAt: end.toISOString(), status: 'open', createdAt: now_() });

  audit_(u.userId, 'START_EXAM', 'ExamSessions', sid, examId);

  return { sessionId: sid, exam: e, endAt: end.toISOString(), questions: qs.map(function (q) { return { questionId: q.questionId, question: q.question, choiceA: q.choiceA, choiceB: q.choiceB, choiceC: q.choiceC, choiceD: q.choiceD }; }) };

}

function submitExam_(u, sessionId, answers) {

  var lock = LockService.getScriptLock(); lock.waitLock(20000);

  try {

    var ses = rows_('ExamSessions').filter(function (x) { return s_(x.sessionId) === s_(sessionId); })[0];

    if (!ses || s_(ses.studentId) !== s_(u.userId)) fail_('ไม่พบรอบสอบ');

    if (ses.status !== 'open') fail_('รอบสอบนี้ส่งแล้ว');

    if (new Date() > new Date(new Date(ses.endAt).getTime() + 30000)) { update_('ExamSessions', 'sessionId', sessionId, { status: 'expired' }); fail_('หมดเวลาสอบ'); }

    var mixed = s_(ses.examId).indexOf('MIXED|') === 0;

    var e, ids=[];

    if (mixed) {

      var mp=s_(ses.examId).split('|');

      e={examId:ses.examId,title:'Mixed Topics Exam',topicId:'MIXED',durationMin:Number(mp[1])||30,passPercent:Number(mp[2])||50};

      ids=(mp[3]||'').split(',').filter(function(x){return x;});

    } else {

      e = rows_('Exams').filter(function (x) { return s_(x.examId) === s_(ses.examId); })[0];

      ids = rows_('ExamQuestions').filter(function (x) { return s_(x.examId) === s_(ses.examId); }).map(function (x) { return s_(x.questionId); });

    }

    var qs = rows_('Questions').filter(function (q) { return ids.indexOf(s_(q.questionId)) >= 0; });

    var aid = newId_('A'), correct = 0, ea = [], aa = [], review = [];

    qs.forEach(function (q) {

      var ch = s_(answers[q.questionId]).toUpperCase(), ok = ch === s_(q.answer).toUpperCase();

      if (ok) correct++;

      ea.push({ sessionId: sessionId, questionId: q.questionId, choice: ch, savedAt: now_() });

      aa.push({ attemptId: aid, questionId: q.questionId, choice: ch, correctAnswer: q.answer, isCorrect: ok ? 'TRUE' : 'FALSE' });

      review.push({ questionId: q.questionId, question: q.question, choice: ch, correctAnswer: q.answer, isCorrect: ok, explanation: q.explanation });

    });

    var total = qs.length, pct = total ? Math.round(correct / total * 1000) / 10 : 0, passed = pct >= (Number(e.passPercent) || 60);

    addMany_('ExamAnswers', ea); addMany_('AttemptAnswers', aa);

    var at = { attemptId: aid, sessionId: sessionId, examId: ses.examId, studentId: u.userId, score: correct, total: total, correct: correct, wrong: total - correct, percent: pct, passed: passed ? 'TRUE' : 'FALSE', submittedAt: now_() };

    add_('Attempts', at);

    update_('ExamSessions', 'sessionId', sessionId, { status: 'submitted' });

    audit_(u.userId, 'SUBMIT_EXAM', 'Attempts', aid, 'score=' + correct + '/' + total);

    return { attempt: at, review: review };

  } finally { lock.releaseLock(); }

}



// ===================== DOCUMENTS =====================

function uploadDocument_(u, p) {

  if (!p.base64 || !/\.pdf$/i.test(s_(p.fileName))) fail_('รองรับเฉพาะไฟล์ PDF');

  if (p.base64.length > 9000000) fail_('ไฟล์ใหญ่เกินไป (จำกัดประมาณ 6 MB)');

  var blob = Utilities.newBlob(Utilities.base64Decode(p.base64), 'application/pdf', p.fileName);

  var f = DriveApp.getFolderById(props_().getProperty(CFG.DRIVE_KEY)).createFile(blob);

  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { logError('share', e); }

  var id = newId_('D');

  add_('Documents', { documentId: id, title: s_(p.title) || p.fileName, description: s_(p.description), topicId: s_(p.topicId), fileId: f.getId(), fileUrl: f.getUrl(), uploadedBy: u.userId, active: 'TRUE', createdAt: now_() });

  audit_(u.userId, 'UPLOAD_DOCUMENT', 'Documents', id, p.fileName);

  return { id: id };

}

function deleteDocument_(u, id) {

  var d = rows_('Documents').filter(function (x) { return s_(x.documentId) === s_(id); })[0];

  if (!d) fail_('ไม่พบเอกสาร');

  try { DriveApp.getFileById(d.fileId).setTrashed(true); } catch (e) { logError('trash', e); }

  delete_('Documents', 'documentId', id);

  audit_(u.userId, 'DELETE_DOCUMENT', 'Documents', id, '');

  return true;

}



// ===================== AUDIT =====================

function audit_(userId, action, entity, entityId, details) {

  try { add_('AuditLogs', { timestamp: now_(), userId: userId, action: action, entity: entity, entityId: entityId, details: String(details).substring(0, 500), userAgent: '' }); } catch (e) { console.error(e); }

}



// ===================== AUTHORIZATION HELPER =====================
// Run once manually from the Apps Script editor to grant required service scopes.
function authorizeAll() {
  var p = props_(), dbId = p.getProperty(CFG.DB_KEY), driveId = p.getProperty(CFG.DRIVE_KEY);
  if (dbId) SpreadsheetApp.openById(dbId).getName();
  if (driveId) DriveApp.getFolderById(driveId).getName();
  MailApp.getRemainingDailyQuota();
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=invalid', {muteHttpExceptions:true}).getResponseCode();
  return 'Authorization check completed';
}

// ===================== SYSTEM HEALTH =====================

function systemHealth() { return systemHealth_(); }

function systemHealth_() {

  var p = props_(), id = p.getProperty(CFG.DB_KEY), h = { dbId: id, driveId: p.getProperty(CFG.DRIVE_KEY), installed: p.getProperty('KPTC_INSTALLED') === '1' };

  if (!id) return h;

  var ss = SpreadsheetApp.openById(id);

  h.name = ss.getName(); h.url = ss.getUrl();

  h.counts = {}; ['Users','Topics','Lessons','Questions','Exams','Documents','Attempts'].forEach(function (n) { h.counts[n] = count_(n); });

  h.sameDbForAllRoles = true; // every role goes through db_() -> same KPTC_DB_ID

  try { DriveApp.getFolderById(h.driveId); h.drivePermission = 'OK'; } catch (e) { h.drivePermission = 'ERROR: ' + e.message; }

  h.sheetsMissing = Object.keys(SCHEMA).filter(function (n) { return !ss.getSheetByName(n); });

  return h;

}
