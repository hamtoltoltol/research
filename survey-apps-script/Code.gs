/**
 * 교내 연구진흥사업 만족도 설문조사 - Google Apps Script 웹 앱
 *
 * 구글 시트에서 [확장 프로그램 > Apps Script]로 열어 이 파일을 Code.gs 에 붙여넣고,
 * 같은 프로젝트에 'Index' 이름의 HTML 파일을 만들어 Index.html 내용을 붙여넣습니다.
 * 배포: 실행 계정 = 나, 액세스 권한 = 학교 도메인 내 모든 사용자.
 *
 * 응답 저장 위치(같은 스프레드시트)
 *  - "응답" 탭: 담당자가 보는 결과. 응답자 1명 = 1행, 문항 = 열. 이메일은 저장하지 않습니다.
 *  - "원본" 탭: 중복 응답 방지와 이어 쓰기용. 응답자 식별값(이메일의 암호화 해시)과 응답 원문(JSON)이 들어 있으므로
 *               편집자에게 공유하지 말고 숨김 처리해 두세요.
 */

// 응답을 허용할 이메일 도메인. 비우면 로그인된 모든 계정을 허용합니다.
const ALLOWED_DOMAINS = ['yonsei.ac.kr'];

const RESULT_SHEET = '응답';
const RAW_SHEET = '원본';
const MAX_PAYLOAD_CHARS = 120000;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('교내 연구진흥사업 만족도 설문조사')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 접속한 사람의 이전 응답을 돌려줍니다. 로그인 계정을 확인할 수 없으면 ok:false 입니다. */
function getMyState() {
  const who = currentUser_();
  if (!who.ok) return who;
  const sheet = rawSheet_();
  const row = findRow_(sheet, who.key);
  if (!row) return { ok: true, data: null };
  const values = sheet.getRange(row, 1, 1, 4).getValues()[0];
  let data = null;
  try {
    data = { answers: JSON.parse(values[3]), completed: values[2] === true || values[2] === 'TRUE' };
  } catch (e) {
    data = null;
  }
  return { ok: true, data: data };
}

/** 응답을 저장합니다. 같은 계정이 다시 저장하면 기존 응답이 새 내용으로 바뀝니다. */
function saveResponse(payload) {
  const who = currentUser_();
  if (!who.ok) return who;
  if (!payload || typeof payload !== 'object' || !payload.answers || !Array.isArray(payload.cells) || !Array.isArray(payload.headers)) {
    return { ok: false, message: '응답 형식이 올바르지 않습니다.' };
  }
  const answersJson = JSON.stringify(payload.answers);
  if (answersJson.length > MAX_PAYLOAD_CHARS || payload.cells.length > 400 || payload.headers.length !== payload.cells.length) {
    return { ok: false, message: '응답 크기가 너무 큽니다.' };
  }
  const completed = payload.completed === true;
  const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const raw = rawSheet_();
    const res = resultSheet_();
    let row = findRow_(raw, who.key);
    const isNew = !row;
    if (isNew) row = Math.max(raw.getLastRow(), 1) + 1;

    // 제출일시는 처음 제출한 시각을 유지합니다.
    let submittedAt = '';
    if (!isNew) submittedAt = String(res.getRange(row, 2).getValue() || '');
    if (completed && !submittedAt) submittedAt = now;

    raw.getRange(row, 1, 1, 4).setValues([[who.key, now, completed, answersJson]]);

    const headers = ['상태', '제출일시', '수정일시'].concat(payload.headers.map(cell_));
    if (res.getLastColumn() !== headers.length || res.getRange(1, 1).getValue() !== headers[0]) {
      res.getRange(1, 1, 1, headers.length).setNumberFormat('@').setValues([headers]);
      res.setFrozenRows(1);
    }
    const line = [completed ? '제출 완료' : '작성 중', submittedAt, now].concat(payload.cells.map(cell_));
    // 텍스트 서식으로 저장해 =, + 로 시작하는 응답이 수식으로 실행되지 않게 합니다.
    res.getRange(row, 1, 1, line.length).setNumberFormat('@').setValues([line]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 내부 함수 ---------- */

function currentUser_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) {
    return { ok: false, message: '로그인한 Google 계정을 확인할 수 없습니다. 학교 Google 계정으로 로그인한 뒤 다시 열어 주세요.' };
  }
  if (ALLOWED_DOMAINS.length && ALLOWED_DOMAINS.indexOf(email.split('@')[1]) === -1) {
    return { ok: false, message: '허용된 학교 Google 계정으로 로그인한 경우에만 응답할 수 있습니다.' };
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

function cell_(v) {
  return String(v === undefined || v === null ? '' : v).slice(0, 5000);
}

function findRow_(sheet, key) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const keys = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (keys[i][0] === key) return i + 2;
  }
  return 0;
}

function rawSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(RAW_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RAW_SHEET);
    sh.getRange(1, 1, 1, 4).setValues([['응답자 식별값(해시)', '수정일시', '제출 완료', '응답 원문(JSON)']]);
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function resultSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(RESULT_SHEET) || ss.insertSheet(RESULT_SHEET);
}
