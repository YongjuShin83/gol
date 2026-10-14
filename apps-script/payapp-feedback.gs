/**
 * 페이앱 결제 결과 받기 (골목상회 공동구매)
 *
 * 주문 시트(구글시트)에 연결된 Apps Script 프로젝트에 새 파일로 붙여 넣으세요.
 * 페이앱이 결제가 끝날 때마다 이 웹 앱으로 결과를 보내고, 이 스크립트가
 * 주문번호(order_no)가 같은 줄의 status 를 '결제완료' 등으로 바꿉니다.
 *
 * 설정 (프로젝트 설정 > 스크립트 속성)
 *   PAYAPP_USERID   페이앱 판매자 아이디
 *   PAYAPP_LINKKEY  페이앱 연동 KEY   (페이앱 판매자 사이트 > 설정 > 연동정보)
 *   PAYAPP_LINKVAL  페이앱 연동 VALUE (같은 곳)
 *
 * 배포: 배포 > 새 배포 > 웹 앱 / 실행: 나 / 액세스 권한: 모든 사용자
 *       → 나온 주소(.../exec)를 index.html 의 CONFIG.PAYAPP_FEEDBACK_URL 에 넣기
 *
 * 주의: 이 프로젝트에 이미 doPost 함수가 있으면 이름이 겹칩니다. 그때는 알려주세요.
 */

const PAYAPP_ORDER_SHEET = 'orders_template';   // 주문이 쌓이는 탭 이름

function doPost(e) {
  if (e && e.parameter && e.parameter.ev) return statsLog_(e.parameter);   // 방문·장바구니 기록 (stats.gs)
  const p = (e && e.parameter) || {};
  const props = PropertiesService.getScriptProperties();

  // 1) 페이앱에서 온 요청인지 확인
  if (p.userid !== props.getProperty('PAYAPP_USERID') ||
      p.linkkey !== props.getProperty('PAYAPP_LINKKEY') ||
      p.linkval !== props.getProperty('PAYAPP_LINKVAL')) {
    return payappReply_('FAIL');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(STATS_SPREADSHEET_ID).getSheetByName(PAYAPP_ORDER_SHEET);   // 주문 시트 (ID는 stats.gs)
    const col = payappColumns_(sh, ['order_no', 'order_total', 'status', 'pay_method', 'pay_no', 'paid_at']);
    const values = sh.getDataRange().getValues();
    const orderNo = String(p.var1 || '').trim();

    // 같은 주문번호의 줄들 (상품마다 한 줄)
    const rows = [];
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][col.order_no - 1]).trim() === orderNo) rows.push(r + 1);
    }
    if (!orderNo || !rows.length) return payappReply_('FAIL');   // 아직 시트에 없으면 페이앱이 다시 보냄

    const first = values[rows[0] - 1];
    const total = Number(String(first[col.order_total - 1]).replace(/[^\d]/g, ''));
    const current = String(first[col.status - 1]).trim();
    const state = String(p.pay_state);
    const set = (c, v) => rows.forEach(r => sh.getRange(r, c).setValue(v));

    if (state === '4') {                                   // 결제 완료
      if (current === '결제완료') return payappReply_('SUCCESS');   // 같은 알림이 여러 번 올 수 있음
      if (Number(p.price) !== total) {                     // 금액이 주문 금액과 다르면 사람이 확인
        set(col.status, '금액확인필요');
        set(col.pay_no, p.mul_no || '');
        return payappReply_('SUCCESS');
      }
      set(col.status, '결제완료');
      set(col.pay_method, '카드결제');
      set(col.pay_no, p.mul_no || '');
      set(col.paid_at, p.pay_date || Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'));
      payappOnPaid_(orderNo, sh, rows);
    } else if (state === '9' || state === '64') {         // 승인 취소
      set(col.status, '결제취소');
    } else if (state === '70' || state === '71') {        // 부분 취소
      set(col.status, '부분취소');
    }
    // 1(요청), 8·32(요청 취소), 10(입금 대기)은 '결제대기' 그대로 둠
    return payappReply_('SUCCESS');
  } finally {
    lock.releaseLock();
  }
}

/** 결제 완료 시 추가로 할 일 — 예: 기존 솔라피 문자 함수 호출 */
function payappOnPaid_(orderNo, sheet, rows) {
  // 예) sendOrderSms_(orderNo);
}

/** 헤더 이름으로 열 번호 찾기. 없는 열은 맨 오른쪽에 새로 만듦 */
function payappColumns_(sh, names) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
  const out = {};
  let next = lastCol + 1;
  names.forEach(n => {
    let i = head.indexOf(n) + 1;
    if (!i) { sh.getRange(1, next).setValue(n); i = next++; }
    out[n] = i;
  });
  return out;
}

function payappReply_(s) {
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT);
}
