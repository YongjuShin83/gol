/**
 * 홍보글 발행 대기열 (골목상회 홍보글 작업대)
 *
 * 작업대(sns.html)에서 채널별 글을 확정하고 '발행'을 누르면, 확정한 글이 이 스크립트로 와서
 * '골목상회 발행 대기열' 시트의 queue 탭에 채널마다 한 줄씩 쌓입니다 (status: 대기).
 * Claude가 대기열을 읽어 밴드·당근에 올리고 status·posted_at·post_url 을 채웁니다.
 *
 * 설치
 *   1) 주문 시트의 Apps Script 프로젝트에 이 파일을 새 파일(publish-queue.gs)로 붙여 넣기
 *   2) payapp-feedback.gs 의 doPost 첫 줄 위에 아래 한 줄 넣기
 *        if (e && e.parameter && e.parameter.pub) return pubQueue_(e.parameter);
 *   3) 함수 '발행키만들기' 를 한 번 실행 (권한 승인) → 실행 기록에 나온 발행 키 확인
 *   4) 배포 > 배포 관리 > (기존 웹 앱) 수정 > 버전: 새 버전 > 배포   (주소는 그대로)
 *   5) 작업대 '기본 설정 > 발행 키' 에 3)의 발행 키를 넣기 (기기마다 한 번)
 *
 * 발행 키는 스크립트 속성 PUB_KEY 에 저장됩니다. 키가 새어 나가면 '발행키만들기' 를 다시 실행하세요.
 */

const PUB_SPREADSHEET_ID = '1K-gDkcO8LQ8ZCDYcH8gE5PcVuVsheFVaztUHm9h1Ki8';   // 골목상회 발행 대기열
const PUB_SHEET = 'queue';
const PUB_HEAD = ['batch_id', 'created_at', 'channel', 'target', 'title', 'body', 'link', 'status', 'posted_at', 'post_url', 'note'];
const PUB_TARGETS = {
  band:   '밴드 · 골목상회 (band.us/band/98565221)',
  daangn: '당근 · 골목상회 비즈프로필 (4086589)',
  blog:   '네이버 블로그',
  insta:  '인스타그램',
};

function pubQueue_(p) {
  const key = PropertiesService.getScriptProperties().getProperty('PUB_KEY');
  if (!key) return statsReply_('ERROR: 발행 키가 아직 없어요 (발행키만들기 실행 필요)');
  if (String(p.key || '') !== key) return statsReply_('DENIED');

  let posts;
  try { posts = JSON.parse(String(p.pub || '')); } catch (err) { return statsReply_('ERROR: 글 형식이 올바르지 않아요'); }
  if (!Array.isArray(posts)) return statsReply_('ERROR: 글 형식이 올바르지 않아요');
  posts = posts.filter(x => x && PUB_TARGETS[x.channel] && String(x.body || '').trim()).slice(0, Object.keys(PUB_TARGETS).length);
  if (!posts.length) return statsReply_('ERROR: 발행할 글이 없어요');

  // '=' 로 시작하는 글이 수식으로 바뀌지 않게 앞에 ' 를 붙임
  const cell = (v, n) => { const s = String(v == null ? '' : v).slice(0, n); return s.charAt(0) === '=' ? "'" + s : s; };
  const now = new Date();
  const batch = Utilities.formatDate(now, 'Asia/Seoul', 'yyMMdd-HHmmss');
  const time = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  const rows = posts.map(x => [batch, time, x.channel, PUB_TARGETS[x.channel], cell(x.title, 200), cell(x.body, 5000), cell(x.link, 300), '대기', '', '', '']);

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = SpreadsheetApp.openById(PUB_SPREADSHEET_ID).getSheetByName(PUB_SHEET);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, PUB_HEAD.length).setNumberFormat('@').setValues(rows);
  } catch (err) {
    return statsReply_('ERROR: ' + err.message);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
  return statsReply_('OK ' + batch + ' ' + rows.length);
}

/** 처음 한 번 실행: 권한 승인 + 발행 키 만들기. 다시 실행하면 새 키로 바뀜 (작업대에도 새 키를 넣어야 함) */
function 발행키만들기() {
  const k = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('PUB_KEY', k);
  Logger.log('대기열 시트: ' + SpreadsheetApp.openById(PUB_SPREADSHEET_ID).getUrl());
  Logger.log('발행 키: ' + k);
}
