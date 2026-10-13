/**
 * 방문·장바구니·주문 기록 (골목상회 공동구매)
 *
 * 사이트가 방문, 장바구니 담기, 주문하기 시작, 주문 완료, 공유를 할 때마다
 * 이 스크립트로 한 줄씩 보내고, 'stats_events' 탭에 쌓습니다.
 * 이름·전화번호·주소 같은 개인정보는 받지 않습니다.
 *
 * 설치
 *   1) 주문 시트의 Apps Script 프로젝트에 이 파일을 새 파일(stats.gs)로 붙여 넣기
 *   2) payapp-feedback.gs 의 doPost 첫 줄에 아래 한 줄이 있는지 확인
 *        if (e && e.parameter && e.parameter.ev) return statsLog_(e.parameter);
 *   3) 배포 > 배포 관리 > (기존 웹 앱) 수정 > 버전: 새 버전 > 배포
 *      처음이면: 배포 > 새 배포 > 웹 앱 / 실행: 나 / 액세스: 모든 사용자
 *   4) 웹 앱 주소(.../exec)를 index.html 의 CONFIG.STATS_URL 에 넣기
 *      (페이앱 결제 결과 주소 PAYAPP_FEEDBACK_URL 과 같은 주소입니다)
 */

const STATS_SPREADSHEET_ID = '1uLy9s_xysOv6Bheri7htnSEP8D0VxV_pxLWqqSNgjE4';   // 주문 시트(orders_template) 주소의 /d/ 뒤 부분
const STATS_SHEET = 'stats_events';
const STATS_HEAD = ['time', 'date', 'event', 'visitor', 'session', 'item_id', 'item_name', 'qty', 'value', 'order_no', 'source', 'device'];
const STATS_EVENTS = ['page_view', 'add_to_cart', 'begin_checkout', 'purchase', 'share'];

function statsLog_(p) {
  const ev = String(p.ev || '');
  if (STATS_EVENTS.indexOf(ev) < 0) return statsReply_('IGNORED');
  const clip = (v, n) => String(v == null ? '' : v).slice(0, n || 80);
  const num = v => { const x = Number(v); return isFinite(x) ? x : ''; };
  const now = new Date();
  const row = [
    Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd'),
    ev, clip(p.vid, 24), clip(p.sid, 24), clip(p.item_id, 20), clip(p.item_name, 60),
    num(p.qty), num(p.value), clip(p.order_no, 30), clip(p.src, 40), clip(p.dev, 10),
  ];
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(STATS_SPREADSHEET_ID);
    let sh = ss.getSheetByName(STATS_SHEET);
    if (!sh) {
      sh = ss.insertSheet(STATS_SHEET);
      sh.getRange(1, 1, 1, STATS_HEAD.length).setValues([STATS_HEAD]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow(row);
  } catch (err) {
    return statsReply_('ERROR: ' + err.message);   // 사이트에는 영향 없음. 원인 확인용
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
  return statsReply_('OK');
}

function statsReply_(s) {
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT);
}
