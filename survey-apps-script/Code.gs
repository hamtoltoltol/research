/**
 * 교내 연구진흥사업 만족도 설문조사 - Google Apps Script 웹 앱
 *
 * 문항과 안내문은 이 스프레드시트의 탭(안내문구, 공통질문, 사업목록, 사업안내, 사업별질문, 마무리질문)에 적혀 있고,
 * 설문 화면은 그 내용을 읽어서 만들어집니다. 문구를 고치려면 코드가 아니라 시트의 칸을 수정하세요.
 *
 * 응답 저장 위치(같은 스프레드시트)
 *  - "응답" 탭: 담당자가 보는 결과. 응답자 1명 = 1행, 문항 = 열. 이메일은 저장하지 않습니다.
 *  - "원본" 탭: 중복 응답 방지와 이어 쓰기용(숨김). 응답자 식별값(이메일의 암호화 해시)이 들어 있으므로 공유하지 마세요.
 */

// 응답을 허용할 이메일 도메인. 비우면 로그인된 모든 계정을 허용합니다.
const ALLOWED_DOMAINS = ['yonsei.ac.kr'];

const TAB = { notice: '안내문구', common: '공통질문', closing: '마무리질문', biz: '사업목록', info: '사업안내', bizq: '사업별질문' };
const RESULT_SHEET = '응답';
const RAW_SHEET = '원본';
const MAX_PAYLOAD_CHARS = 120000;
const CACHE_SECONDS = 300;
const MAX_OPTIONS = 15;
const TYPE_MAP = { '단일선택': 'single', '복수선택': 'multi', '서술형': 'text' };

/* ====================================================================
 * 웹 앱
 * ==================================================================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('교내 연구진흥사업 만족도 설문조사')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 화면을 만드는 데 필요한 문항 설정과, 접속한 사람의 이전 응답을 한 번에 돌려줍니다. */
function boot() {
  const who = currentUser_();
  const webUrl = ScriptApp.getService().getUrl() || '';
  if (!who.ok) return { state: { ok: false, message: who.message, email: who.email || '', domain: ALLOWED_DOMAINS.join(', ') }, webUrl: webUrl };
  const config = getConfig_();
  const sheet = rawSheet_();
  const row = findRow_(sheet, who.key);
  let data = null;
  if (row) {
    const v = sheet.getRange(row, 1, 1, 5).getValues()[0];
    try {
      data = { answers: JSON.parse(v[3]), completed: v[2] === true || v[2] === 'TRUE' };
    } catch (e) {
      data = null;
    }
  }
  return { config: config, state: { ok: true, data: data }, webUrl: webUrl };
}

/** 응답을 저장합니다. 같은 계정이 다시 저장하면 기존 응답이 새 내용으로 바뀝니다. */
function saveResponse(payload) {
  const who = currentUser_();
  if (!who.ok) return who;
  if (!payload || typeof payload !== 'object' || !payload.answers || typeof payload.answers !== 'object') {
    return { ok: false, message: '응답 형식이 올바르지 않습니다.' };
  }
  const cfg = getConfig_();
  const cols = columns_(cfg);
  const known = {};
  cols.forEach(function (c) { known[c[1]] = true; });
  const answers = {};
  Object.keys(payload.answers).forEach(function (k) { if (known[k]) answers[k] = payload.answers[k]; });
  const answersJson = JSON.stringify(answers);
  if (answersJson.length > MAX_PAYLOAD_CHARS) return { ok: false, message: '응답 크기가 너무 큽니다.' };

  const completed = payload.completed === true;
  const now = nowText_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const raw = rawSheet_();
    const res = resultSheet_();
    let row = findRow_(raw, who.key);
    const isNew = !row;
    if (isNew) row = Math.max(raw.getLastRow(), 1) + 1;
    let submittedAt = isNew ? '' : String(raw.getRange(row, 5).getValue() || '');
    if (completed && !submittedAt) submittedAt = now;
    raw.getRange(row, 1, 1, 5).setNumberFormat('@').setValues([[who.key, now, completed, answersJson, submittedAt]]);

    const headers = resultHeaders_(cols);
    if (!sameHeaders_(res, headers)) {
      rebuildResults_(cfg);
    } else {
      const line = resultLine_(completed, submittedAt, now, cols, answers);
      res.getRange(row, 1, 1, line.length).setNumberFormat('@').setValues([line]);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ====================================================================
 * 시트 메뉴
 * ==================================================================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('설문 관리')
    .addItem('① 샘플 문항 불러오기 (처음 한 번)', 'setupSheets')
    .addItem('② 변경 내용 점검', 'checkConfig')
    .addItem('③ 변경 내용 바로 반영하기', 'applyChanges')
    .addItem('④ 설문 링크 보기', 'showLink')
    .addToUi();
}

function setupSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = Object.keys(TAB).some(function (k) {
    const sh = ss.getSheetByName(TAB[k]);
    return sh && sh.getLastRow() > 1;
  });
  if (exists) {
    const r = ui.alert('이미 문항이 있습니다', '샘플 문항으로 다시 채우면 지금 적혀 있는 문항과 안내문이 모두 지워집니다. 계속할까요?', ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) return;
  }
  writeTab_(TAB.notice, ['항목', '내용'], SEED.notice, [140, 560], null);
  writeTab_(TAB.common, SEED.heads.question, SEED.common, questionWidths_(), 'question');
  writeTab_(TAB.biz, ['번호', '사업명', '홈페이지 링크', '사용', 'ID'], SEED.bizList, [70, 320, 360, 70, 90], 'biz');
  writeTab_(TAB.info, ['사업번호', '소제목', '내용', '표시'], SEED.info, [80, 200, 620, 100], 'info');
  writeTab_(TAB.bizq, SEED.heads.bizq, SEED.bizq, [90].concat(questionWidths_()), 'bizq');
  writeTab_(TAB.closing, SEED.heads.question, SEED.closing, questionWidths_(), 'question');
  const order = [TAB.notice, TAB.common, TAB.biz, TAB.info, TAB.bizq, TAB.closing];
  order.forEach(function (n, i) { ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 1); });
  getConfigFresh_();
  ui.alert('샘플 문항을 불러왔습니다', '각 탭의 안내 행(맨 위 메모)을 읽고 문구를 수정하세요. 수정 후에는 5분 안에 설문에 반영되고, 바로 반영하려면 메뉴 ③을 누르세요.', ui.ButtonSet.OK);
}

function checkConfig() {
  const problems = [];
  const cfg = buildConfig_(problems);
  const ui = SpreadsheetApp.getUi();
  if (problems.length) {
    ui.alert('확인이 필요한 부분 ' + problems.length + '개', problems.slice(0, 25).join('\n') + (problems.length > 25 ? '\n… 외 ' + (problems.length - 25) + '개' : ''), ui.ButtonSet.OK);
  } else {
    ui.alert('문제 없음', '공통질문 ' + cfg.common.length + '개, 사업 ' + cfg.businesses.length + '개, 마무리 질문 ' + cfg.closing.length + '개를 정상적으로 읽었습니다.', ui.ButtonSet.OK);
  }
}

function applyChanges() {
  const cfg = getConfigFresh_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { rebuildResults_(cfg); } finally { lock.releaseLock(); }
  SpreadsheetApp.getUi().alert('반영했습니다', '설문 화면에 바로 반영되었고, "응답" 탭의 열 구성도 현재 문항에 맞게 다시 만들었습니다.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function showLink() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('설문 링크', url ? url : '아직 웹 앱으로 배포되지 않았습니다. 배포 > 새 배포 > 웹 앱으로 배포해 주세요.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ====================================================================
 * 설정 읽기
 * ==================================================================== */

function getConfig_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('cfg');
  if (hit) {
    try {
      return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(hit), 'application/x-gzip')).getDataAsString());
    } catch (e) { /* 캐시가 깨졌으면 다시 읽는다 */ }
  }
  return getConfigFresh_();
}

function getConfigFresh_() {
  ensureIds_();
  const cfg = buildConfig_([]);
  try {
    const packed = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(cfg))).getBytes());
    if (packed.length < 90000) CacheService.getScriptCache().put('cfg', packed, CACHE_SECONDS);
  } catch (e) { /* 캐시를 못 써도 동작에는 문제 없음 */ }
  return cfg;
}

function buildConfig_(problems) {
  const meta = readNotice_();
  const hidden = {};
  const businesses = readBusinesses_(problems, hidden);
  const byNo = {};
  businesses.forEach(function (b) { byNo[b.no] = b; });

  readInfo_(businesses, byNo, hidden, problems);

  const names = businesses.map(function (b) { return b.name; });
  const common = readQuestions_(TAB.common, 0, names, problems);
  const closing = readQuestions_(TAB.closing, 0, names, problems);

  // 사업별 질문: 사업번호 칸이 '전체'이면 모든 사업에, 번호(쉼표로 여러 개)이면 해당 사업에만 붙습니다.
  const bizqSheet = getSheet_(TAB.bizq);
  const bizRows = bizqSheet ? bizqSheet.getDataRange().getDisplayValues().slice(1) : [];
  businesses.forEach(function (b) { b.questions = []; });
  const parsed = readQuestionRows_(bizRows, 1, names, TAB.bizq, problems);
  parsed.forEach(function (item) {
    const target = String(item.row[0] || '').trim();
    const nos = target === '' || target === '전체' ? null : target.split(',').map(function (x) { return x.trim(); });
    if (nos) nos.forEach(function (n) { if (!byNo[n] && !hidden[n]) problems.push('[' + TAB.bizq + '] ' + item.line + '행: 사업번호 "' + n + '"가 사업목록에 없습니다.'); });
    businesses.forEach(function (b) {
      if (!nos || nos.indexOf(b.no) !== -1) {
        const q = JSON.parse(JSON.stringify(item.q));
        q.id = b.id + '_' + item.q.id;
        b.questions.push(q);
      }
    });
  });
  businesses.forEach(function (b) { numberQuestions_(b.questions); });

  if (!businesses.length) problems.push('[' + TAB.biz + '] 사용 중인 사업이 하나도 없습니다.');
  return { meta: meta, common: common, businesses: businesses, closing: closing };
}

function readNotice_() {
  const sh = getSheet_(TAB.notice);
  const rows = sh ? sh.getDataRange().getDisplayValues().slice(1) : [];
  const m = {};
  rows.forEach(function (r) {
    const k = String(r[0] || '').trim(), v = String(r[1] || '');
    if (!k || !v.trim()) return;
    (m[k] = m[k] || []).push(v);
  });
  const one = function (k, d) { return m[k] ? m[k][0] : d; };
  return {
    title: one('설문 제목', '설문조사'),
    intro: m['소개 문단'] || [],
    period: one('조사기간', ''), target: one('조사대상', ''),
    commonTitle: one('공통질문 제목', '공통질문'), commonLead: one('공통질문 안내', ''),
    selectTitle: one('사업 선택 제목', '사업 선택'), selectLead: one('사업 선택 안내', ''),
    selectQ: one('사업 선택 질문', '참여한 경험이 있는 사업을 모두 선택하여 주십시오. (복수응답 가능)'),
    selectHint: one('사업 선택 보조설명', ''), noneLabel: one('해당없음 보기', ''),
    bizNotice: m['사업별 질문 시작 안내'] || [],
    closingNotice: one('마무리 질문 시작 안내', ''), closingTitle: one('마무리 제목', '마무리 질문'),
    signoff: one('맺음말', ''), signer: one('서명', '')
  };
}

function readBusinesses_(problems, hidden) {
  const sh = getSheet_(TAB.biz);
  const rows = sh ? sh.getDataRange().getDisplayValues() : [];
  const out = [], seen = {};
  rows.slice(1).forEach(function (r, i) {
    const no = String(r[0] || '').trim(), name = String(r[1] || '').trim();
    if (!no && !name) return;
    if (String(r[3] || '').trim() === '아니오') { hidden[no] = true; return; }
    if (!no || !name) { problems.push('[' + TAB.biz + '] ' + (i + 2) + '행: 번호와 사업명을 모두 적어 주세요.'); return; }
    if (seen[no]) { problems.push('[' + TAB.biz + '] ' + (i + 2) + '행: 번호 "' + no + '"가 중복되었습니다.'); return; }
    seen[no] = true;
    const link = String(r[2] || '').trim();
    if (link && link.indexOf('https://') !== 0) problems.push('[' + TAB.biz + '] ' + (i + 2) + '행: 홈페이지 링크는 https:// 로 시작해야 합니다. (링크가 화면에 표시되지 않습니다)');
    out.push({ id: String(r[4] || '').trim(), no: no, name: name, link: link.indexOf('https://') === 0 ? link : '', sections: [], questions: [] });
  });
  return out;
}

function readInfo_(businesses, byNo, hidden, problems) {
  const sh = getSheet_(TAB.info);
  const rows = sh ? sh.getDataRange().getDisplayValues().slice(1) : [];
  const KIND = { '일반': 'ln', '들여쓰기': 'sub', '강조': 'warn' };
  rows.forEach(function (r, i) {
    const no = String(r[0] || '').trim(), head = String(r[1] || '').trim(), text = String(r[2] || '').trim(), kind = String(r[3] || '일반').trim();
    if (!no && !head && !text) return;
    const b = byNo[no];
    if (!b) { if (no && !hidden[no]) problems.push('[' + TAB.info + '] ' + (i + 2) + '행: 사업번호 "' + no + '"가 사업목록에 없거나 사용 중이 아닙니다.'); return; }
    let sec = b.sections[b.sections.length - 1];
    if (!sec || sec.h !== head) { sec = { h: head, items: [] }; b.sections.push(sec); }
    if (!text) return;
    if (kind === '표머리' || kind === '표행') {
      const cells = text.split('|').map(function (c) { return c.trim(); });
      let tb = sec.items[sec.items.length - 1];
      if (!tb || tb.type !== 'table') { tb = { type: 'table', head: null, rows: [] }; sec.items.push(tb); }
      if (kind === '표머리') tb.head = cells; else tb.rows.push(cells);
    } else {
      if (!KIND[kind]) problems.push('[' + TAB.info + '] ' + (i + 2) + '행: 표시 "' + kind + '"는 일반, 들여쓰기, 강조, 표머리, 표행 중 하나여야 합니다.');
      sec.items.push({ type: 'line', t: text, cls: KIND[kind] || 'ln' });
    }
  });
  businesses.forEach(function (b) {
    b.sections.forEach(function (s) {
      s.items.forEach(function (it) { if (it.type === 'table' && !it.head) it.head = it.rows.length ? it.rows[0].map(function () { return ''; }) : []; });
    });
  });
}

function readQuestions_(tabName, offset, bizNames, problems) {
  const sh = getSheet_(tabName);
  const rows = sh ? sh.getDataRange().getDisplayValues().slice(1) : [];
  const list = readQuestionRows_(rows, offset, bizNames, tabName, problems).map(function (x) { return x.q; });
  numberQuestions_(list);
  return list;
}

/** 질문 행을 읽어 문항 객체로 바꿉니다. 열 순서: 번호, 질문, 유형, 필수, 보기1~15, 기타입력 보기, 최대 선택 수, 정확히 선택 수, 안내문, ID */
function readQuestionRows_(rows, off, bizNames, tabName, problems) {
  const out = [];
  rows.forEach(function (r, i) {
    const line = i + 2;
    const text = String(r[off + 1] || '').trim();
    if (!text) return;
    const typeKo = String(r[off + 2] || '').trim();
    const type = TYPE_MAP[typeKo];
    if (!type) { problems.push('[' + tabName + '] ' + line + '행: 유형은 단일선택, 복수선택, 서술형 중 하나여야 합니다. (이 문항은 화면에 나오지 않습니다)'); return; }
    let opts = [];
    for (let k = 0; k < MAX_OPTIONS; k++) {
      const o = String(r[off + 4 + k] || '').trim();
      if (!o) continue;
      if (o === '[사업목록]') opts = opts.concat(bizNames); else opts.push(o);
    }
    if (type !== 'text' && opts.length < 2) { problems.push('[' + tabName + '] ' + line + '행: 선택형 문항은 보기가 2개 이상 필요합니다. (이 문항은 화면에 나오지 않습니다)'); return; }
    const other = String(r[off + 19] || '').trim();
    if (other && opts.indexOf(other) === -1) problems.push('[' + tabName + '] ' + line + '행: "기타입력 보기"는 보기 중 하나와 똑같이 적어야 합니다.');
    const max = parseInt(r[off + 20], 10) || 0;
    const exact = parseInt(r[off + 21], 10) || 0;
    if (exact && exact > opts.length) problems.push('[' + tabName + '] ' + line + '행: "정확히 선택 수"가 보기 수보다 큽니다.');
    if (type !== 'multi' && (max || exact)) problems.push('[' + tabName + '] ' + line + '행: 최대/정확히 선택 수는 복수선택 문항에만 쓸 수 있습니다.');
    const id = String(r[off + 23] || '').trim();
    const q = {
      id: id || ('row' + line), no: String(r[off] || '').trim(), text: text, type: type, opts: type === 'text' ? [] : opts,
      opt: String(r[off + 3] || '').trim() === '아니오',
      other: type === 'text' || opts.indexOf(other) === -1 ? '' : other,
      max: type === 'multi' ? max : 0, exact: type === 'multi' ? exact : 0,
      hint: String(r[off + 22] || '').replace(/\[사업목록\]/g, bizNames.join(', '))
    };
    out.push({ q: q, row: r, line: line });
  });
  return out;
}

/** 번호 칸이 비어 있으면 앞 번호 + 1 (첫 문항은 1) 로 자동 번호를 붙입니다. */
function numberQuestions_(list) {
  let last = 0;
  list.forEach(function (q) {
    if (q.no) { const n = parseInt(q.no, 10); if (!isNaN(n)) last = n; return; }
    last += 1; q.no = String(last);
  });
}

/* ====================================================================
 * 응답 시트
 * ==================================================================== */

/** 응답 탭의 열 목록: [머리글, 응답 키] */
function columns_(cfg) {
  const cols = [];
  const add = function (prefix, q) {
    cols.push([prefix + ' Q' + q.no + ' ' + q.text, q.id]);
    if (q.other) cols.push([prefix + ' Q' + q.no + ' ' + q.text + ' (기타 내용)', q.id + '_etc']);
  };
  cfg.common.forEach(function (q) { add('[공통]', q); });
  cols.push(['[사업 선택] Q1 ' + cfg.meta.selectQ, 'sel']);
  cfg.businesses.forEach(function (b) { b.questions.forEach(function (q) { add('[' + b.no + '. ' + b.name + ']', q); }); });
  cfg.closing.forEach(function (q) { add('[마무리]', q); });
  return cols;
}

function resultHeaders_(cols) {
  return ['상태', '제출일시', '수정일시'].concat(cols.map(function (c) { return cell_(c[0]); }));
}

function resultLine_(completed, submittedAt, updatedAt, cols, answers) {
  return [completed ? '제출 완료' : '작성 중', submittedAt, updatedAt].concat(cols.map(function (c) {
    const v = answers[c[1]];
    return cell_(Array.isArray(v) ? v.join(' | ') : v);
  }));
}

function sameHeaders_(res, headers) {
  if (res.getLastColumn() !== headers.length) return false;
  const cur = res.getRange(1, 1, 1, headers.length).getValues()[0];
  for (let i = 0; i < headers.length; i++) if (cur[i] !== headers[i]) return false;
  return true;
}

/** 원본 탭의 모든 응답으로 응답 탭을 새로 만듭니다. (문항이 바뀐 뒤 열이 어긋나지 않게 합니다) */
function rebuildResults_(cfg) {
  const raw = rawSheet_();
  const res = resultSheet_();
  const cols = columns_(cfg);
  const headers = resultHeaders_(cols);
  const n = raw.getLastRow() - 1;
  const rows = [];
  if (n > 0) {
    raw.getRange(2, 1, n, 5).getValues().forEach(function (v) {
      let answers = {};
      try { answers = JSON.parse(v[3] || '{}'); } catch (e) { answers = {}; }
      rows.push(resultLine_(v[2] === true || v[2] === 'TRUE', String(v[4] || ''), String(v[1] || ''), cols, answers));
    });
  }
  res.clear();
  const all = [headers].concat(rows);
  res.getRange(1, 1, all.length, headers.length).setNumberFormat('@').setValues(all);
  res.setFrozenRows(1);
  res.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e4f1f5');
}

/* ====================================================================
 * 내부 도구
 * ==================================================================== */

function currentUser_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) {
    // 학교 도메인 밖의 계정(개인 Gmail 등)은 구글이 이메일을 알려 주지 않아 이 경우로 들어옵니다.
    return { ok: false, message: '지금 로그인된 구글 계정이 학교 계정이 아니거나 확인할 수 없습니다.' };
  }
  if (ALLOWED_DOMAINS.length && ALLOWED_DOMAINS.indexOf(email.split('@')[1]) === -1) {
    return { ok: false, email: email, message: '이 설문은 학교 구글 계정으로만 응답할 수 있습니다.' };
  }
  return { ok: true, key: hash_(email) };
}

function hash_(email) {
  const props = PropertiesService.getScriptProperties();
  let salt = props.getProperty('SALT');
  if (!salt) {
    salt = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('SALT', salt);
  }
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + email, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function nowText_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'); }

function cell_(v) { return String(v === undefined || v === null ? '' : v).slice(0, 5000); }

function getSheet_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }

function findRow_(sheet, key) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const keys = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) if (keys[i][0] === key) return i + 2;
  return 0;
}

function rawSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(RAW_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RAW_SHEET);
    sh.getRange(1, 1, 1, 5).setValues([['응답자 식별값(해시)', '수정일시', '제출 완료', '응답 원문(JSON)', '제출일시']]);
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function resultSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(RESULT_SHEET) || ss.insertSheet(RESULT_SHEET);
}

/** 새로 추가한 문항·사업 행에 ID가 비어 있으면 자동으로 채웁니다. (ID는 응답을 문항과 연결하는 값이라 고치지 마세요) */
function ensureIds_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const props = PropertiesService.getScriptProperties();
    let seq = parseInt(props.getProperty('IDSEQ') || '0', 10);
    const fill = function (tab, textCol, idCol, prefix) {
      const sh = getSheet_(tab);
      if (!sh || sh.getLastRow() < 2) return;
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, idCol).getDisplayValues();
      const used = {};
      rows.forEach(function (r) { const id = String(r[idCol - 1] || '').trim(); if (id) used[id] = true; });
      const updates = [];
      let changed = false;
      rows.forEach(function (r) {
        let id = String(r[idCol - 1] || '').trim();
        if (!id && String(r[textCol - 1] || '').trim()) {
          do { seq += 1; id = prefix + seq; } while (used[id]);
          used[id] = true; changed = true;
        }
        updates.push([id]);
      });
      if (changed) sh.getRange(2, idCol, updates.length, 1).setNumberFormat('@').setValues(updates);
    };
    fill(TAB.common, 2, 24, 'q');
    fill(TAB.closing, 2, 24, 'q');
    fill(TAB.bizq, 3, 25, 'q');
    fill(TAB.biz, 2, 5, 'x');
    props.setProperty('IDSEQ', String(seq));
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 시트 서식 ---------- */

function questionWidths_() {
  const w = [60, 380, 90, 60];
  for (let i = 0; i < MAX_OPTIONS; i++) w.push(160);
  return w.concat([160, 80, 80, 240, 90]);
}

function writeTab_(name, head, rows, widths, kind) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  const cols = head.length;
  const all = [head].concat(rows);
  sh.getRange(1, 1, all.length, cols).setNumberFormat('@').setValues(all);
  sh.getRange(1, 1, 1, cols).setFontWeight('bold').setBackground('#0b6b8a').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, Math.max(rows.length, 1), cols).setWrap(true).setVerticalAlignment('top');
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  // 입력 실수를 줄이는 드롭다운
  const list = function (col, values) {
    sh.getRange(2, col, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build());
  };
  const off = kind === 'bizq' ? 1 : 0;
  if (kind === 'question' || kind === 'bizq') {
    list(off + 3, ['단일선택', '복수선택', '서술형']);
    list(off + 4, ['예', '아니오']);
    sh.getRange(1, off + 24, 1, 1).setBackground('#8a9ba5');
    sh.getRange(2, off + 24, 500, 1).setBackground('#eef2f4');
  }
  if (kind === 'biz') { list(4, ['예', '아니오']); sh.getRange(1, 5, 1, 1).setBackground('#8a9ba5'); sh.getRange(2, 5, 500, 1).setBackground('#eef2f4'); }
  if (kind === 'info') list(4, ['일반', '들여쓰기', '강조', '표머리', '표행']);
  const notes = {
    '안내문구': '왼쪽 항목 이름은 바꾸지 말고 오른쪽 내용만 고치세요. "소개 문단", "사업별 질문 시작 안내"는 행을 추가해 문단을 늘릴 수 있습니다. 칸 안에서 줄바꿈은 Alt+Enter(맥: Option+Enter)입니다.',
    '공통질문': '한 줄이 한 문항입니다. 보기는 보기1, 보기2… 칸에 하나씩 적고, 쓰지 않는 칸은 비워 두세요. 유형과 필수는 목록에서 고릅니다. 새 문항은 마지막 줄 아래에 적으면 되고 ID는 자동으로 채워집니다. 번호 칸이 비어 있으면 앞 번호에 이어 자동 번호가 붙습니다.',
    '사업목록': '화면에 나오는 사업 목록입니다. 사용 칸을 "아니오"로 바꾸면 이 사업은 설문에서 빠집니다. 번호는 사업안내·사업별질문 탭과 연결됩니다. 새 사업은 마지막 줄 아래에 적으세요.',
    '사업안내': '사업 선택 후 각 사업 화면 위쪽에 나오는 안내문입니다. 한 줄이 한 문장입니다. 같은 사업번호와 소제목이 이어지면 한 묶음으로 보입니다. 표시: 일반/들여쓰기/강조(빨간 글씨), 표는 표머리 줄 1개 + 표행 줄 여러 개로 적고 칸은 | 로 나눕니다. 표 칸 앞에 !를 붙이면 빨간 글씨입니다.',
    '사업별질문': '선택한 사업마다 나오는 질문입니다. 사업번호가 "전체"이면 모든 사업에 나오고, 3-1 또는 3-1,3-2 처럼 적으면 그 사업에만 나옵니다. 나머지 칸 규칙은 공통질문 탭과 같습니다.',
    '마무리질문': '마지막에 나오는 공통문항입니다. 칸 규칙은 공통질문 탭과 같습니다. 보기1에 [사업목록]이라고 적으면 사업 이름이 보기로 자동 채워집니다.'
  };
  sh.getRange(1, 1).setNote(notes[name] || '');
}
