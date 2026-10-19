/**
 * 수능 단어장 → 구글 시트 자동 기록
 *
 * 앱이 학습을 끝낼 때마다 보내는 내용:
 *   { rows: [{ id, row: [날짜, 시간, 학생, Day, 항목, 범위, 푼 문제, 맞힘, 정답률(%), 틀린 것, 중간 종료] }],
 *     status: { modes, days, done, dayDone } }
 *
 * 탭 1 "학습 기록": 한 번 학습할 때마다 한 줄. 마지막 열(기록ID)로 중복을 거릅니다.
 * 탭 2 "Day별 과제": 매번 새로 그리는 Day별 과제 완료 현황표.
 *
 * 설치 방법은 저장소의 SETUP-SHEET.md를 보세요.
 */

const LOG_SHEET = "학습 기록";
const DAY_SHEET = "Day별 과제";
const HEAD = ["날짜", "시간", "학생", "Day", "항목", "범위", "푼 문제", "맞힘", "정답률(%)", "틀린 것", "중간 종료", "기록ID"];
const NUM_COLS = [3, 6, 7, 8]; // Day, 푼 문제, 맞힘, 정답률 (0부터 센 위치)

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    const added = appendRows_(body.rows || []);
    if (body.status) writeDays_(body.status);
    return json_({ ok: true, added: added });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 브라우저에서 웹 앱 주소를 열었을 때 연결 확인용
function doGet() {
  return ContentService.createTextOutput("수능 단어장 기록 연결이 준비됐어요.");
}

function appendRows_(rows) {
  const sh = getSheet_(LOG_SHEET, HEAD);
  const last = sh.getLastRow();
  const ids = new Set(last > 1 ? sh.getRange(2, HEAD.length, last - 1, 1).getValues().map(r => String(r[0])) : []);
  const fresh = [];
  rows.forEach(r => {
    if (!r || !r.id || ids.has(r.id)) return;
    ids.add(r.id);
    const row = HEAD.slice(0, -1).map((_, i) => {
      const v = r.row[i] === undefined ? "" : r.row[i];
      return NUM_COLS.indexOf(i) >= 0 && v !== "" && !isNaN(v) ? Number(v) : String(v);
    });
    fresh.push(row.concat(r.id));
  });
  if (fresh.length) sh.getRange(sh.getLastRow() + 1, 1, fresh.length, HEAD.length).setValues(fresh);
  return fresh.length;
}

function writeDays_(st) {
  const modes = st.modes || [], days = st.days || [], done = st.done || {}, dayDone = st.dayDone || {};
  const head = ["Day", "단원", "범위"].concat(modes.map(m => m.t), ["완료한 과제", "Day 완료일"]);
  const sh = getSheet_(DAY_SHEET, null);
  sh.clear();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold").setBackground("#EDEFF4");
  sh.setFrozenRows(1);
  if (!days.length) return;

  const values = [], colors = [], notes = [];
  days.forEach(d => {
    const dd = done[d.day] || {};
    const row = [d.day, d.ch || "", d.range || ""], bg = [null, null, null], nt = ["", "", ""];
    let n = 0;
    modes.forEach(m => {
      const x = dd[m.id];
      if (x) {
        n++;
        row.push(x.acc / 100);
        bg.push(x.acc >= 80 ? "#E2F4EB" : x.acc >= 60 ? "#FFF3B0" : "#FCEBED");
        nt.push(x.t + " 완료 · " + x.n + "문제");
      } else { row.push(""); bg.push(null); nt.push(""); }
    });
    row.push(n + " / " + modes.length, dayDone[d.day] || "");
    bg.push(null, dayDone[d.day] ? "#E2F4EB" : null);
    nt.push("", "");
    values.push(row); colors.push(bg); notes.push(nt);
  });
  const body = sh.getRange(2, 1, values.length, head.length);
  body.setValues(values).setBackgrounds(colors).setNotes(notes).setHorizontalAlignment("center");
  if (modes.length) sh.getRange(2, 4, values.length, modes.length).setNumberFormat("0%");
  sh.getRange(2, 2, values.length, 1).setHorizontalAlignment("left");
  sh.getRange(values.length + 3, 1).setValue("숫자는 과제를 끝냈을 때의 첫 시도 정답률이에요. 칸에 마우스를 올리면 완료한 날짜가 보여요. 마지막 갱신: " +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")).setFontColor("#5E6782");
  sh.autoResizeColumns(1, head.length);
}

function getSheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name, name === LOG_SHEET ? 0 : 1);
    if (head) {
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold").setBackground("#EDEFF4");
      sh.setFrozenRows(1);
      sh.getRange("A:B").setNumberFormat("@"); // 날짜·시간을 앱에 적힌 그대로
      sh.hideColumns(head.length);             // 기록ID는 숨김 (중복 확인용)
    }
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 설치 확인용: 편집기에서 이 함수를 한 번 실행하면 권한 승인 창이 뜨고, 시트에 탭 두 개가 생깁니다.
function setup() {
  getSheet_(LOG_SHEET, HEAD);
  writeDays_({ modes: [], days: [], done: {}, dayDone: {} });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ["시트1", "Sheet1"].forEach(n => { const s = ss.getSheetByName(n); if (s && s.getLastRow() === 0) ss.deleteSheet(s); });
}
