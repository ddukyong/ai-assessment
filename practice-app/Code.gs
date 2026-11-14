/**
 * 서·논술형 AI 설계 실습 서버 (Google Apps Script)
 * - 연수 참여자가 만든 설계안을 스프레드시트 '설계안' 시트에 한 사람당 한 줄로 저장
 * - 같은 연수 코드 + 이름 + 비밀번호(숫자 4자리)로 다른 기기에서 이어서 할 수 있음
 * - 스프레드시트의 [확장 프로그램 > Apps Script]에 이 파일 전체를 붙여 넣고 웹 앱으로 배포
 */

/* ===== 연수 담당자가 바꾸는 곳 ===== */
const COURSE_CODES = ["ai2026"];   // 참여자에게 알려 줄 연수 코드 (여러 개 가능)
const SHEET_ID = "";               // 스프레드시트에 붙은 스크립트면 비워 둠
/* ================================== */

const SHEET_NAME = "설계안";
const HEAD = ["저장 시각", "연수 코드", "이름", "학교급·학년·과목", "진행 단계", "점검 완료", "평가 요소", "발문", "자료", "조건",
  "채점기준", "예시답안", "채점기준 수정", "AI 활용 안내문", "피드백(교사 확정)", "확인값", "데이터1", "데이터2", "데이터3", "데이터4"];
const KEY_COL = 15, DATA_COL = 16, CHUNK = 45000, CELL_MAX = 45000;

function doGet(){ return out_({ ok:true, msg:"서·논술형 AI 설계 실습 서버가 작동 중입니다." }); }

function doPost(e){
  let r;
  try { r = handle_(JSON.parse(e.postData.contents)); }
  catch(err){ r = { ok:false, error:String(err && err.message || err) }; }
  return out_(r);
}

function handle_(p){
  const code = String(p.code || "").trim();
  const name = String(p.name || "").trim().slice(0, 20);
  const pin = String(p.pin || "").trim();
  if (COURSE_CODES.indexOf(code) < 0) throw new Error("연수 코드를 다시 확인해 주세요.");
  if (!name) throw new Error("이름을 써 주세요.");
  if (!/^\d{4}$/.test(pin)) throw new Error("비밀번호는 숫자 4자리로 써 주세요.");

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_();
    const rows = sh.getDataRange().getValues();
    let idx = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === code && String(rows[i][2]) === name) { idx = i; break; }
    }
    const key = hash_(code + "|" + name + "|" + pin);
    if (idx >= 0 && String(rows[idx][KEY_COL]) !== key)
      throw new Error("이미 있는 이름인데 비밀번호가 달라요. 내 이름이 맞다면 비밀번호를 확인하고, 다른 분과 이름이 같다면 이름 뒤에 학교를 붙여 주세요.");

    if (p.action === "load") {
      if (idx < 0) return { ok:true, data:null };
      const json = rows[idx].slice(DATA_COL, DATA_COL + 4).join("");
      return { ok:true, data: json ? JSON.parse(json) : null };
    }
    if (p.action === "save") {
      const s = p.summary || {};
      const json = JSON.stringify(p.data || {});
      if (json.length > CHUNK * 4) throw new Error("붙여 넣은 AI 결과가 너무 길어요. AI 결과 칸의 내용을 줄인 뒤 다시 저장해 주세요.");
      const cut = t => String(t || "").slice(0, CELL_MAX);
      const row = [new Date(), code, name, cut(s.who), cut(s.step), cut(s.checks), cut(s.elems), cut(s.q), cut(s.d), cut(s.c),
        cut(s.rubric), cut(s.model), cut(s.fix), cut(s.notice), cut(s.feedback), key];
      for (let i = 0; i < 4; i++) row.push(json.slice(i * CHUNK, (i + 1) * CHUNK));
      if (idx >= 0) sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);
      else sh.appendRow(row);
      return { ok:true };
    }
    throw new Error("알 수 없는 요청이에요.");
  } finally { lock.releaseLock(); }
}

function sheet_(){
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.setFrozenColumns(3);
    sh.hideColumns(KEY_COL + 1, 5);   // 확인값과 데이터 칸은 숨김
  }
  return sh;
}

function hash_(t){
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t, Utilities.Charset.UTF_8));
}

function out_(o){
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
