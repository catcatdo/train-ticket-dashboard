/**
 * 열차표 기록 자동 정리 스크립트
 * - 도착시간이 현재시각보다 과거인 행 삭제
 * - 메모/LED문구는 남은 행에서 수정하지 않음
 *
 * 사용 전:
 * 1) 스프레드시트 > 확장 프로그램 > Apps Script 열기
 * 2) 이 파일 내용 붙여넣기
 * 3) setupCleanupTrigger() 1회 실행(권한 허용)
 */

const SHEET_NAME = '시트1';
const TIMEZONE = 'Asia/Seoul';

function cleanupCompletedTrainRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`시트를 찾을 수 없음: ${SHEET_NAME}`);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return; // 데이터 없음

  // 헤더 제외 전체 데이터
  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const values = range.getValues();

  // 컬럼 인덱스 (1-based):
  // A 날짜, B 출발역, C 도착역, D 출발시간, E 도착시간, F 열차번호, G 좌석, H 메모, I LED문구
  const COL_DATE = 1; // A
  const COL_ARRIVAL = 5; // E

  const now = new Date();
  const rowsToDelete = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const dateCell = row[COL_DATE - 1];
    const arrivalCell = row[COL_ARRIVAL - 1];

    const arrivalDateTime = parseArrivalDateTime(dateCell, arrivalCell, TIMEZONE);
    if (!arrivalDateTime) continue; // 파싱 실패 행은 건너뜀

    if (arrivalDateTime.getTime() < now.getTime()) {
      // 실제 시트의 행번호(헤더 1행 + i 오프셋)
      rowsToDelete.push(i + 2);
    }
  }

  // 아래에서 위로 삭제(행번호 틀어짐 방지)
  rowsToDelete.sort((a, b) => b - a).forEach((rowNum) => {
    sheet.deleteRow(rowNum);
  });
}

/**
 * 날짜(A열) + 도착시간(E열)을 Date로 파싱
 * 지원 포맷:
 * - 날짜: Date 객체 또는 "YYYY-MM-DD"
 * - 시간: Date 객체 또는 "HH:mm"
 */
function parseArrivalDateTime(dateCell, timeCell, tz) {
  let y, m, d;

  if (Object.prototype.toString.call(dateCell) === '[object Date]' && !isNaN(dateCell)) {
    y = Number(Utilities.formatDate(dateCell, tz, 'yyyy'));
    m = Number(Utilities.formatDate(dateCell, tz, 'MM'));
    d = Number(Utilities.formatDate(dateCell, tz, 'dd'));
  } else if (typeof dateCell === 'string' && dateCell.trim()) {
    const mDate = dateCell.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!mDate) return null;
    y = Number(mDate[1]);
    m = Number(mDate[2]);
    d = Number(mDate[3]);
  } else {
    return null;
  }

  let hh, mm;
  if (Object.prototype.toString.call(timeCell) === '[object Date]' && !isNaN(timeCell)) {
    hh = Number(Utilities.formatDate(timeCell, tz, 'HH'));
    mm = Number(Utilities.formatDate(timeCell, tz, 'mm'));
  } else if (typeof timeCell === 'string' && timeCell.trim()) {
    const mTime = timeCell.trim().match(/^(\d{1,2}):(\d{2})$/);
    if (!mTime) return null;
    hh = Number(mTime[1]);
    mm = Number(mTime[2]);
  } else {
    return null;
  }

  // 스크립트 타임존 기준 Date 생성
  return new Date(y, m - 1, d, hh, mm, 0, 0);
}

/**
 * 10분 간격 트리거 생성 (최초 1회 실행)
 */
function setupCleanupTrigger() {
  // 중복 트리거 방지: 기존 동일 함수 트리거 제거
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((t) => {
    if (t.getHandlerFunction() === 'cleanupCompletedTrainRows') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('cleanupCompletedTrainRows')
    .timeBased()
    .everyMinutes(10)
    .create();
}

/**
 * 트리거 삭제용(필요 시)
 */
function removeCleanupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((t) => {
    if (t.getHandlerFunction() === 'cleanupCompletedTrainRows') {
      ScriptApp.deleteTrigger(t);
    }
  });
}
