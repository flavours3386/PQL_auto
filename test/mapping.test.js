require('./gas');
const test = require('node:test');
const assert = require('node:assert');

const deal = (id, shop, extra) => Object.assign({ id: id, title: 't' + id, owner_id: 7, stage_id: 3, label_ids: [301, 299], custom_fields: { [PD_FIELD_SHOP_ID]: shop } }, extra || {});

test('딜 분리: 숫자만 매핑된 딜', () => {
  const s = splitDeals_([deal(1, '123'), deal(2, ' 456 '), deal(3, '채널톡'), deal(4, ''), deal(5, null), { id: 6 }]);
  assert.deepStrictEqual([...s.shopIds], ['123', '456']);
  assert.deepStrictEqual(s.unmapped.map((d) => d.id), [3, 4, 5, 6]);
});

test('deal list 행: 레이블 id 순, 소유자·단계 이름, 이번에 반영한 shop_id', () => {
  const rows = dealListRows_([deal(1, '123'), deal(2, '채널톡')], { 7: '한서연' }, { 3: '컨택전' }, { 299: '알파리뷰', 301: '알파푸시' }, { 2: '108903' });
  assert.deepStrictEqual(rows, [
    ['거래 - 레이블', '거래 - shop_id', '거래 - 이름', '거래 - 소유자', '거래 - 단계'],
    ['알파리뷰, 알파푸시', 123, 't1', '한서연', '컨택전'],
    ['알파리뷰, 알파푸시', 108903, 't2', '한서연', '컨택전'],
  ]);
});

const mrow = (deal, shop, verdict, status) => [deal, 'n', '채널톡', shop, 's', 'email, phone', '확인 필요', verdict, status, '2026-09-30'];

test('매핑 탭 상태 읽기', () => {
  const st = readMappingState_([
    mrow('1', '10', '승인', '대기'), mrow('2', '20', '거절', '대기'), mrow('3', '30', '', '대기'),
    mrow('4', '40', '', '거절됨'), mrow('5', '50', '', '반영됨'), mrow('6', '60', '', '실패: x'), mrow('', '', '', ''),
  ]);
  assert.deepStrictEqual(st.approved.map((x) => x[0]), ['1']);
  assert.deepStrictEqual(st.pending.map((x) => x[0]), ['3']);
  assert.deepStrictEqual([...st.rejectedPairs].sort(), ['2:20', '4:40']);
  assert.deepStrictEqual(st.keep.map((x) => x[0] + x[8]), ['2거절됨', '4거절됨', '5반영됨', '6실패: x']);
});

test('한 딜에 승인이 둘이면 반영하지 않는다', () => {
  const s = splitApprovals_([mrow('1', '10', '승인', '대기'), mrow('1', '11', '승인', '대기'), mrow('2', '20', '승인', '대기')]);
  assert.deepStrictEqual(s.apply.map((x) => x[0]), ['2']);
  assert.deepStrictEqual(s.duplicate.map((x) => x[3]), ['10', '11']);
});

test('자동 반영 계획: 상한 이하면 자동, 초과면 전부 대기', () => {
  const m = (tier) => ({ tier: tier, deal: {}, candidates: [] });
  const matches = [m('high'), m('high'), m('review'), m('none')];
  let p = planMappings_(matches, true, 2);
  assert.strictEqual(p.auto.length, 2);
  assert.strictEqual(p.pendingHigh.length, 0);
  assert.strictEqual(p.review.length, 1);
  assert.strictEqual(p.overLimit, false);
  p = planMappings_(matches, true, 1);
  assert.strictEqual(p.auto.length, 0);
  assert.strictEqual(p.pendingHigh.length, 2);
  assert.strictEqual(p.overLimit, true);
  p = planMappings_(matches, false, 100);
  assert.strictEqual(p.auto.length, 0);
  assert.strictEqual(p.pendingHigh.length, 2);
  assert.strictEqual(p.overLimit, false);
});

test('매핑 행·노트·파일명', () => {
  const match = { deal: { id: 9, title: '서브마켓', raw: '채널톡' } };
  assert.deepStrictEqual(mappingRow_(match, { shopId: '108903', keys: ['email', 'url'], shopName: '서브마켓' }, '높음', '', '반영됨', '2026-10-01'),
    ['9', '서브마켓', '채널톡', '108903', '서브마켓', 'email, url', '높음', '', '반영됨', '2026-10-01']);
  assert.strictEqual(mappingNote_('자동', '채널톡', '108903', ['email', 'url'], '2026-10-01'),
    "[PQL 자동매핑 2026-10-01] shop_id '채널톡' → 108903 (근거: email, url)");
  assert.strictEqual(uploadFileName_('all_subscription_1001.csv', '0930'), 'pipedrive_up(1001).xlsx');
  assert.strictEqual(uploadFileName_('all_subscription_004141.csv', '0930'), 'pipedrive_up(0930).xlsx');
});

test('요약 문구에 단계별 수치가 들어간다', () => {
  const lines = summaryLines_({
    fileName: 'all_subscription_1001.csv', fileUpdated: '2026-09-30 10:23',
    counts: { total: 100, orders: 50, review: 1, site: 2, pro: 3, phone: 4, deal: 5, mapped: 6, noTarget: 7 },
    targetCounts: { 업셀: 10, '업셀, 푸시': 8, 푸시: 4 }, cleanCount: 22, uploadCount: 20,
    approvedApplied: 1, autoApplied: 6, autoFailed: 0, overLimit: false, pending: 3, elapsedSec: 42,
  }).join('\n');
  for (const s of ['원천 행 100', '-50', '(역매핑) -6', '결과 22곳', '업셀 10', '업로드 xlsx 20행', '자동 반영 6건', '승인 대기 3건', '42초']) {
    assert.ok(lines.indexOf(s) >= 0, s);
  }
});
