/**
 * Class Scoreboard — Google Sheets Backend (FULL)
 * 修正版：欄位 ↔ 功能 連結
 * - Announcements：新增 comments(JSON)；讀寫自動 JSON 處理
 * - studentMessages：走 Settings(JSON) 專用通道
 * - Quizzes：新增 startDate；讀取時回填 choices{} / answer
 * - 其它：增量寫入、整表覆寫、append 歷程、容量裁切、60s 快取、Lock、錯誤 Logs、真刪
 */

////////////////////////////
// Basic
////////////////////////////

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate().setTitle('班級管理系統');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function _nowISO() { return new Date().toISOString(); }

function _logError(fn, msg) {
  try {
    var sh = _ensureExactSheet('Logs', ['time','function','message']);
    sh.appendRow([_nowISO(), fn, String(msg)]);
  } catch (e) {
    // 最後防線不拋出
  }
}

function _getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

////////////////////////////
// Schema (表頭)
////////////////////////////

var H = {
  Students:        ['id','name','groupId','groupName','score'],
  Groups:          ['id','name','score'],
  Rewards:         ['id','name','points','quantity','image','description'],
  // ExchangeHistory: expanded to include full details for display and processing
  ExchangeHistory: ['id','studentId','studentName','groupName','rewardId','rewardName','points','status','requestDate','approveDate','rejectDate'],
  // ScoreHistory: JSON field affectedStudents holds list of changes
  ScoreHistory:    ['id','type','targetId','targetName','groupName','scoreChange','reason','date','affectedStudents'], // JSON
  // Announcements: comments is stored as JSON
  Announcements:   ['id','title','content','link','date','comments'], // comments: JSON
  // Quizzes: include type, scores, createDate, endDate. winners is numeric (top N winners)
  Quizzes:         ['id','winners','title','question','choiceA','choiceB','choiceC','choiceD','correct','type','scores','startType','startDate','startTime','status','createDate','endDate'],
  QuizAnswers:     ['id','quizId','studentId','studentName','rank','scoreAwarded','answer','isCorrect','submitTime'],
  // StudentMessages: each message row; replies stored as JSON
  StudentMessages: ['id','studentId','studentName','groupName','content','visibility','date','replies'],
  Settings:        ['key','value','updatedAt']
};

// 需要 JSON.parse 的欄位
var JSON_FIELDS = {
  ScoreHistory: ['affectedStudents'],
  Announcements: ['comments'],
  // Quizzes: scores is stored as JSON string; winners is numeric
  Quizzes: ['scores'],
  // StudentMessages: replies is stored as JSON string
  StudentMessages: ['replies']
};

////////////////////////////
// Sheet helpers
////////////////////////////

function _ensureExactSheet(name, headers) {
  var ss = _getSS();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }
  // 確保表頭一致（以你提供的 schema 為準）
  var lastCol = sh.getLastColumn();
  if (lastCol === 0) lastCol = headers.length; // 空表保護
  var cur = sh.getRange(1, 1, 1, Math.max(lastCol, headers.length)).getValues()[0];
  var changed = false;
  for (var i = 0; i < headers.length; i++) {
    if (cur[i] !== headers[i]) { cur[i] = headers[i]; changed = true; }
  }
  if (changed) {
    sh.getRange(1, 1, 1, headers.length).setValues([cur.slice(0, headers.length)]);
    if (sh.getLastColumn() > headers.length) sh.deleteColumns(headers.length + 1, sh.getLastColumn() - headers.length);
  }
  return sh;
}

function _valForCell(key, val, jsonFields) {
  if (jsonFields && jsonFields.indexOf(key) >= 0) {
    try { return (val === undefined || val === null) ? '' : JSON.stringify(val); }
    catch(e){ return (val === undefined || val === null) ? '' : String(val); }
  }
  return (val === undefined || val === null) ? '' : val;
}

// 全取（支援快取與 JSON.parse）
function _getAllObjects(name, headers, jsonFieldsOpt) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('ALL__' + name);
  if (hit) {
    try { return JSON.parse(hit); } catch(e){}
  }
  var sh = _ensureExactSheet(name, headers);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, headers.length).getValues();
  var jsonFields = jsonFieldsOpt || (JSON_FIELDS[name] || []);
  var arr = values.map(function(row){
    var o = {};
    for (var i = 0; i < headers.length; i++) {
      var k = headers[i], v = row[i];
      if (jsonFields.indexOf(k) >= 0 && typeof v === 'string' && v) {
        try { o[k] = JSON.parse(v); } catch(e) { o[k] = v; }
      } else {
        o[k] = v;
      }
    }
    return o;
  });

  // 相容處理：回填 quizzes 舊欄位 (choices/answer)
  if (name === 'Quizzes') {
    arr.forEach(function(q){
      if (!q.choices) {
        q.choices = {
          A: q.choiceA || '',
          B: q.choiceB || '',
          C: q.choiceC || '',
          D: q.choiceD || ''
        };
      }
      if (q.answer === undefined && q.correct !== undefined) q.answer = q.correct;
    });
  }

  try { cache.put('ALL__' + name, JSON.stringify(arr), 60); } catch (e) {}
  return arr;
}

// 整表覆寫（保留表頭）：for rewards 等需要支援刪除的集合
function _replaceAllRows(name, headers, items, jsonFieldsOpt) {
  var sh = _ensureExactSheet(name, headers);
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ _logError('_replaceAllRows-lock', e); }

  try {
    // 清除舊資料（保留第一列表頭）
    var last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);
    // 寫入新資料
    var jsonFields = jsonFieldsOpt || (JSON_FIELDS[name] || []);
    var rows = (items || []).map(function(it){
      return headers.map(function(k){ return _valForCell(k, it[k], jsonFields); });
    });
    if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    // 清快取
    CacheService.getScriptCache().remove('ALL__' + name);
    return { status: 'replaced', count: rows.length };
  } catch (e) {
    _logError('_replaceAllRows', e.toString());
    return { status: 'error', error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(e){}
  }
}

// 增量 upsert：依 id 覆寫/新增
function _upsertRows(name, headers, items, idField, jsonFieldsOpt) {
  var sh = _ensureExactSheet(name, headers);
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ _logError('_upsertRows-lock', e); }

  try {
    var existing = _getAllObjects(name, headers, jsonFieldsOpt);
    var index = {};
    for (var i = 0; i < existing.length; i++) index[String(existing[i][idField])] = i + 2;

    var toAppend = [];
    var jsonFields = jsonFieldsOpt || (JSON_FIELDS[name] || []);

    for (var j = 0; j < (items || []).length; j++) {
      var it = items[j] || {};
      var id = String(it[idField] || '');
      if (!id) { id = String(new Date().getTime()) + '-' + j; it[idField] = id; }
      if (index[id]) {
        var row = index[id];
        var vals = headers.map(function(key){ return _valForCell(key, (it[key] !== undefined ? it[key] : ''), jsonFields); });
        sh.getRange(row, 1, 1, headers.length).setValues([vals]);
      } else {
        var valsNew = headers.map(function(key){ return _valForCell(key, (it[key] !== undefined ? it[key] : ''), jsonFields); });
        toAppend.push(valsNew);
      }
    }
    if (toAppend.length) {
      var BATCH = 200, start = 0;
      while (start < toAppend.length) {
        var part = toAppend.slice(start, start + BATCH);
        sh.getRange(sh.getLastRow() + 1, 1, part.length, headers.length).setValues(part);
        start += BATCH;
      }
    }
    CacheService.getScriptCache().remove('ALL__' + name);
    return { status: 'upserted', appended: toAppend.length };
  } catch (e) {
    _logError('_upsertRows', e.toString());
    return { status: 'error', error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(e){}
  }
}

// 只新增（避免覆寫歷史）：for ScoreHistory / QuizAnswers
function _appendNewById(name, headers, items, jsonFieldsOpt) {
  var sh = _ensureExactSheet(name, headers);
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(e){ _logError('_appendNewById-lock', e); }

  try {
    var existing = _getAllObjects(name, headers, jsonFieldsOpt);
    var existingIds = {};
    for (var i = 0; i < existing.length; i++) existingIds[String(existing[i].id)] = true;

    var toAppend = [];
    var jsonFields = jsonFieldsOpt || (JSON_FIELDS[name] || []);
    (items || []).forEach(function(it, j){
      var rowObj = Object.assign({}, it);
      if (!rowObj.id) rowObj.id = String(new Date().getTime()) + '-' + j;
      if (!existingIds[String(rowObj.id)]) {
        var vals = headers.map(function(k){ return _valForCell(k, (rowObj[k] !== undefined ? rowObj[k] : ''), jsonFields); });
        toAppend.push(vals);
      }
    });
    if (toAppend.length) {
      sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, headers.length).setValues(toAppend);
    }
    CacheService.getScriptCache().remove('ALL__' + name);
    return { status: 'appended', count: toAppend.length };
  } catch (e) {
    _logError('_appendNewById', e.toString());
    return { status: 'error', error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(e){}
  }
}

// 容量裁切（保留最新 N 筆）
function _trimMaxRows(name, headers, maxRows) {
  var sh = _ensureExactSheet(name, headers);
  var last = sh.getLastRow();
  var dataRows = last - 1;
  if (dataRows > maxRows) {
    var toDelete = dataRows - maxRows;
    // 刪最舊的（自第 2 列起）
    sh.deleteRows(2, toDelete);
    CacheService.getScriptCache().remove('ALL__' + name);
    return { trimmed: toDelete };
  }
  return { trimmed: 0 };
}

////////////////////////////
// Settings helpers
////////////////////////////

function _getSetting(key) {
  var sh = _ensureExactSheet('Settings', H.Settings);
  var last = sh.getLastRow();
  if (last < 2) return '';
  var values = sh.getRange(2, 1, last - 1, 2).getValues(); // key,value
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(key)) return values[i][1];
  }
  return '';
}

function _setSetting(key, value) {
  var sh = _ensureExactSheet('Settings', H.Settings);
  var last = sh.getLastRow();
  for (var r = 2; r <= last; r++) {
    if (String(sh.getRange(r, 1).getValue()) === String(key)) {
      sh.getRange(r, 2, 1, 2).setValues([[value, _nowISO()]]);
      return { status: 'updated' };
    }
  }
  sh.appendRow([key, value, _nowISO()]);
  return { status: 'created' };
}

function _getJSONSetting(key, defaultVal) {
  var txt = _getSetting(key);
  if (!txt) return (defaultVal === undefined ? null : defaultVal);
  try { return JSON.parse(txt); } catch(e){ return (defaultVal === undefined ? null : defaultVal); }
}

function _setJSONSetting(key, obj) {
  try {
    var txt = JSON.stringify(obj || null);
    return _setSetting(key, txt);
  } catch (e) {
    return _setSetting(key, String(obj));
  }
}

////////////////////////////
// Public API for Frontend
////////////////////////////

function getStorage(key) {
  try {
    switch (String(key)) {
      case 'students':         return _getAllObjects('Students', H.Students);
      case 'groups':           return _getAllObjects('Groups', H.Groups);
      case 'rewards':          return _getAllObjects('Rewards', H.Rewards);
      case 'exchangeRequests': return _getAllObjects('ExchangeHistory', H.ExchangeHistory);
      case 'scoreHistory':     return _getAllObjects('ScoreHistory', H.ScoreHistory, JSON_FIELDS.ScoreHistory);
      case 'announcements':    return _getAllObjects('Announcements', H.Announcements, JSON_FIELDS.Announcements);
      case 'quizzes':          return _getAllObjects('Quizzes', H.Quizzes, JSON_FIELDS.Quizzes); // 讀取時會自動補 choices/answer
      case 'quizAnswers':      return _getAllObjects('QuizAnswers', H.QuizAnswers);
      case 'studentMessages':  {
        // Load messages from dedicated sheet; parse replies JSON field
        var msgs = _getAllObjects('StudentMessages', H.StudentMessages, JSON_FIELDS.StudentMessages);
        // Always return array; do not fallback to Settings JSON storage
        return msgs || [];
      }
      case 'classTitle':       return _getSetting('classTitle') || '';
      case 'loginAttempts':    return _getSetting('loginAttempts') || 0;
      case 'lockoutTime':      return _getSetting('lockoutTime') || '';
      default:                 return _getSetting(String(key)); // 其餘單值設定
    }
  } catch (e) {
    _logError('getStorage:' + key, e.toString());
    throw e;
  }
}

function setStorage(key, value) {
  try {
    switch (String(key)) {

      case 'students':         return _upsertRows('Students', H.Students, value || [], 'id');

      case 'groups':           return _upsertRows('Groups', H.Groups, value || [], 'id');

      // 獎品：整表覆寫（含刪除）
      case 'rewards': {
        var items = (value || []).map(function(r){
          return {
            id: r.id, name: r.name, points: r.points, quantity: r.quantity,
            image: r.image || '', description: r.description || ''
          };
        });
        return _replaceAllRows('Rewards', H.Rewards, items);
      }

      // 公告（comments 將以 JSON 字串儲存）
      case 'announcements': {
        var itemsA = (value || []).map(function(a){
          return {
            id: a.id, title: a.title, content: a.content, link: a.link || '',
            date: a.date || _nowISO(), comments: (a.comments || [])
          };
        });
        return _upsertRows('Announcements', H.Announcements, itemsA, 'id', JSON_FIELDS.Announcements);
      }

      // quizzes：支援 choices / answer / scores / type / createDate / endDate；補 startDate 與 startTime
      case 'quizzes': {
        var itemsQ = (value || []).map(function(q, idx){
          var out = {};
          // assign id if missing
          out.id = q.id || (String(new Date().getTime()) + '-' + idx);
          out.title = q.title;
          out.question = q.question;
          // winners (top N) should be numeric
          if (q.winners !== undefined && q.winners !== null && q.winners !== '') {
            out.winners = parseInt(q.winners, 10);
            if (isNaN(out.winners)) out.winners = 0;
          } else {
            out.winners = 0;
          }
          // choices mapping
          if (q.choices) {
            out.choiceA = q.choices.A || '';
            out.choiceB = q.choices.B || '';
            out.choiceC = q.choices.C || '';
            out.choiceD = q.choices.D || '';
          } else {
            out.choiceA = q.choiceA || '';
            out.choiceB = q.choiceB || '';
            out.choiceC = q.choiceC || '';
            out.choiceD = q.choiceD || '';
          }
          // correct answer: from correct or answer field
          if (q.correct !== undefined && q.correct !== null) {
            out.correct = q.correct;
          } else if (q.answer !== undefined && q.answer !== null) {
            out.correct = q.answer;
          } else {
            out.correct = '';
          }
          // type of quiz ('choice' or 'text')
          out.type = q.type || '';
          // scores mapping (per rank) stored as JSON
          out.scores = q.scores || {};
          // startType: immediate or scheduled
          // default to 'scheduled' instead of 'schedule' to match frontend values
          out.startType = q.startType || 'scheduled';
          if (out.startType === 'immediate') {
            var nowT = new Date();
            out.startDate = Utilities.formatDate(nowT, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            out.startTime = Utilities.formatDate(nowT, Session.getScriptTimeZone(), 'HH:mm');
            out.status = q.status || 'active';
          } else {
            out.startDate = q.startDate || '';
            out.startTime = q.startTime || '';
            out.status = q.status || 'scheduled';
          }
          // createDate and endDate (optional)
          out.createDate = q.createDate || '';
          out.endDate = q.endDate || '';
          return out;
        });
        return _upsertRows('Quizzes', H.Quizzes, itemsQ, 'id', JSON_FIELDS.Quizzes);
      }

      // 歷程與作答：只新增
      case 'scoreHistory': {
        var itemsH = (value || []).map(function(h){
          return {
            id: h.id, type: h.type, targetId: h.targetId, targetName: h.targetName,
            groupName: h.groupName || '', scoreChange: h.scoreChange, reason: h.reason || '',
            date: h.date || _nowISO(), affectedStudents: h.affectedStudents || []
          };
        });
        var resH = _appendNewById('ScoreHistory', H.ScoreHistory, itemsH, JSON_FIELDS.ScoreHistory);
        // 容量上限：1 萬筆
        _trimMaxRows('ScoreHistory', H.ScoreHistory, 10000);
        return resH;
      }

      case 'quizAnswers': {
        var itemsQA = (value || []).map(function(a){
          return {
            id: a.id, quizId: a.quizId, studentId: a.studentId, studentName: a.studentName,
            rank: a.rank || '', scoreAwarded: a.scoreAwarded || 0, answer: a.answer || '',
            isCorrect: a.isCorrect ? true : false, submitTime: a.submitTime || _nowISO()
          };
        });
        var resQA = _appendNewById('QuizAnswers', H.QuizAnswers, itemsQA);
        // 容量上限：5000 筆
        _trimMaxRows('QuizAnswers', H.QuizAnswers, 5000);
        return resQA;
      }

      case 'exchangeRequests': {
        // Upsert full exchange request records, preserving all needed fields
        var itemsE = (value || []).map(function(r, idx){
          var obj = {};
          // assign id if missing
          obj.id = r.id || (String(new Date().getTime()) + '-' + idx);
          obj.studentId = r.studentId;
          obj.studentName = r.studentName || '';
          obj.groupName = r.groupName || '';
          obj.rewardId = r.rewardId;
          obj.rewardName = r.rewardName || '';
          obj.points = r.points || 0;
          obj.status = r.status || 'pending';
          // requestDate property used by frontend; fallback to original date or now
          obj.requestDate = r.requestDate || r.date || _nowISO();
          obj.approveDate = r.approveDate || '';
          obj.rejectDate = r.rejectDate || '';
          return obj;
        });
        return _upsertRows('ExchangeHistory', H.ExchangeHistory, itemsE, 'id');
      }

      // studentMessages：儲存為 StudentMessages 工作表，replies 序列化
      case 'studentMessages': {
        var itemsM = (value || []).map(function(m, idx){
          var obj = {};
          obj.id = m.id || (String(new Date().getTime()) + '-' + idx);
          obj.studentId = m.studentId;
          obj.studentName = m.studentName || '';
          obj.groupName = m.groupName || '';
          obj.content = m.content || '';
          obj.visibility = m.visibility || '';
          obj.date = m.date || _nowISO();
          obj.replies = m.replies || [];
          return obj;
        });
        return _replaceAllRows('StudentMessages', H.StudentMessages, itemsM, JSON_FIELDS.StudentMessages);
      }

      // 單值設定
      case 'classTitle':
      case 'loginAttempts':
      case 'lockoutTime': {
        return _setSetting(String(key), (value === undefined ? '' : value));
      }

      default: {
        // 其餘未知 key 一律當作 Settings 單值
        return _setSetting(String(key), (value === undefined ? '' : value));
      }
    }
  } catch (e) {
    _logError('setStorage:' + key, e.toString());
    throw e;
  }
}

// 一次批次儲存（前端若把所有集合傳進來）
function saveAllData(data) {
  try {
    if (data && data.students          !== undefined) setStorage('students', data.students);
    if (data && data.groups            !== undefined) setStorage('groups', data.groups);
    if (data && data.rewards           !== undefined) setStorage('rewards', data.rewards);
    if (data && data.history           !== undefined) setStorage('scoreHistory', data.history);
    if (data && data.exchangeRequests  !== undefined) setStorage('exchangeRequests', data.exchangeRequests);
    if (data && data.announcements     !== undefined) setStorage('announcements', data.announcements);
    if (data && data.quizzes           !== undefined) setStorage('quizzes', data.quizzes);
    if (data && data.quizAnswers       !== undefined) setStorage('quizAnswers', data.quizAnswers);
    if (data && data.studentMessages   !== undefined) setStorage('studentMessages', data.studentMessages);
    if (data && data.classTitle        !== undefined) setStorage('classTitle', data.classTitle);
    return { success: true };
  } catch (e) {
    _logError('saveAllData', e.toString());
    return { success: false, error: e.toString() };
  }
}

// 真刪：依 key / id 刪除對應資料列
function deleteStorage(key, id) {
  try {
    var sh, headers;
    switch (String(key)) {
      case 'students':         sh = _ensureExactSheet('Students', H.Students); headers = H.Students; break;
      case 'groups':           sh = _ensureExactSheet('Groups', H.Groups); headers = H.Groups; break;
      case 'rewards':          sh = _ensureExactSheet('Rewards', H.Rewards); headers = H.Rewards; break;
      case 'announcements':    sh = _ensureExactSheet('Announcements', H.Announcements); headers = H.Announcements; break;
      case 'quizzes':          sh = _ensureExactSheet('Quizzes', H.Quizzes); headers = H.Quizzes; break;
      case 'exchangeRequests': sh = _ensureExactSheet('ExchangeHistory', H.ExchangeHistory); headers = H.ExchangeHistory; break;
      case 'scoreHistory':     sh = _ensureExactSheet('ScoreHistory', H.ScoreHistory); headers = H.ScoreHistory; break;
      case 'quizAnswers':      sh = _ensureExactSheet('QuizAnswers', H.QuizAnswers); headers = H.QuizAnswers; break;
      default:                 return { status: 'error', message: 'Unsupported key: ' + key };
    }
    var last = sh.getLastRow();
    if (last < 2) return { status: 'notfound' };
    var values = sh.getRange(2, 1, last - 1, headers.length).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]) === String(id)) {
        sh.deleteRow(i + 2);
        try { CacheService.getScriptCache().remove('ALL__' + key); } catch(e){}
        return { status: 'deleted' };
      }
    }
    return { status: 'notfound' };
  } catch (e) {
    _logError('deleteStorage:' + key, e.toString());
    return { status: 'error', error: e.toString() };
  }
}

////////////////////////////
// Atomic Operations (concurrent-safe)
////////////////////////////

// Read directly from sheet, bypass cache
function _readFresh(name, headers, jsonFieldsOpt) {
  var sh = _ensureExactSheet(name, headers);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, headers.length).getValues();
  var jsonFields = jsonFieldsOpt || (JSON_FIELDS[name] || []);
  return values.map(function(row) {
    var o = {};
    for (var i = 0; i < headers.length; i++) {
      var k = headers[i], v = row[i];
      if (jsonFields.indexOf(k) >= 0 && typeof v === 'string' && v) {
        try { o[k] = JSON.parse(v); } catch(e) { o[k] = v; }
      } else { o[k] = v; }
    }
    return o;
  });
}

// Upsert without acquiring lock (caller must hold lock).
// existing: full array read from sheet (used for row-number lookup).
// items: only the records to write.
function _upsertNoLock(sh, name, headers, existing, items, idField, jsonFieldsOpt) {
  var index = {};
  existing.forEach(function(row, i) { index[String(row[idField])] = i + 2; });
  var jsonFields = jsonFieldsOpt || (JSON_FIELDS[name] || []);
  var toAppend = [];
  (items || []).forEach(function(it) {
    var id = String(it[idField] || '');
    var vals = headers.map(function(k) { return _valForCell(k, it[k] !== undefined ? it[k] : '', jsonFields); });
    if (index[id]) {
      sh.getRange(index[id], 1, 1, headers.length).setValues([vals]);
    } else {
      toAppend.push(vals);
    }
  });
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, headers.length).setValues(toAppend);
  }
  CacheService.getScriptCache().remove('ALL__' + name);
}

// Append rows directly without lock or duplicate check (caller must hold lock and ensure unique IDs).
function _appendDirectNoLock(sh, name, headers, items, jsonFieldsOpt) {
  var jsonFields = jsonFieldsOpt || (JSON_FIELDS[name] || []);
  var rows = (items || []).map(function(it) {
    return headers.map(function(k) { return _valForCell(k, it[k] !== undefined ? it[k] : '', jsonFields); });
  });
  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }
  CacheService.getScriptCache().remove('ALL__' + name);
}

function _makeHistoryId() {
  return String(Date.now()) + '-' + Math.floor(Math.random() * 100000);
}

function _nowLocale() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
}

