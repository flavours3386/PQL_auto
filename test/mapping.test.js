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

test('매핑 행·노트', () => {
  const match = { deal: { id: 9, title: '서브마켓', raw: '채널톡' } };
  assert.deepStrictEqual(mappingRow_(match, { shopId: '108903', keys: ['email', 'url'], shopName: '서브마켓' }, '높음', '', '반영됨', '2026-10-01'),
    ['9', '서브마켓', '채널톡', '108903', '서브마켓', 'email, url', '높음', '', '반영됨', '2026-10-01']);
  assert.strictEqual(mappingNote_('자동', '채널톡', '108903', ['email', 'url'], '2026-10-01'),
    "[PQL 자동매핑 2026-10-01] shop_id '채널톡' → 108903 (근거: email, url)");
});

test('요약 문구에 단계별 수치와 업로드 결과가 들어간다', () => {
  const base = {
    fileName: 'all_subscription_1001.csv', fileUpdated: '2026-09-30 10:23',
    counts: { total: 100, orders: 50, review: 1, site: 2, pro: 3, phone: 4, deal: 5, mapped: 6, noTarget: 7 },
    targetCounts: { 업셀: 10, '업셀, 푸시': 8, 푸시: 4 }, cleanCount: 22, suspectCount: 2,
    approvedApplied: 1, autoApplied: 6, autoFailed: 0, overLimit: false, pending: 3, elapsedSec: 42,
    upload: { total: 20, created: 18, failed: 1, skipped: 1, blocked: false, off: false },
  };
  const lines = summaryLines_(base).join('\n');
  for (const s of ['원천 행 100', '-50', '(역매핑) -6', '결과 22곳', '업셀 10', 'Pipedrive 업로드 18/20건', '실패 1', '남음 1', '딜 의심 2곳', '자동 반영 6건', '승인 대기 3건', '42초']) {
    assert.ok(lines.indexOf(s) >= 0, s);
  }
  const blocked = summaryLines_(Object.assign({}, base, { upload: { total: 600, created: 0, failed: 0, skipped: 0, blocked: true, off: false } })).join('\n');
  assert.ok(blocked.indexOf('500곳 초과') >= 0);
});

test('업로드 ID 해석: 소유자·단계·라벨 이름 → id, 없으면 멈춘다', () => {
  const ids = resolveUploadIds_({ 24324011: '한서연', 1: '김혜빈' }, { 71: '컨택전', 89: '라이트 ' }, { 299: '알파리뷰', 303: 'null' });
  assert.deepStrictEqual(ids, { ownerId: 24324011, stageId: 71, labelIds: { 알파리뷰: 299, null: 303 } });
  assert.throws(() => resolveUploadIds_({ 1: '김혜빈' }, { 71: '컨택전' }, {}), /소유자 '한서연'/);
  assert.throws(() => resolveUploadIds_({ 24324011: '한서연' }, { 89: '라이트' }, {}), /단계 '컨택전'/);
});

test('업로드 계획: 끔·상한 초과면 한 건도 올리지 않는다', () => {
  const items = [{ row: 1 }, { row: 2 }, { row: 3 }];
  assert.deepStrictEqual(planUpload_(items, true, 3), { items: items, blocked: false, off: false });
  assert.deepStrictEqual(planUpload_(items, true, 2), { items: [], blocked: true, off: false });
  assert.deepStrictEqual(planUpload_(items, false, 100), { items: [], blocked: false, off: true });
});

test('업로드 결과를 clean 탭 업로드 열 값으로', () => {
  const head = OUTPUT_HEADERS.slice();
  const idx = head.indexOf('업로드');
  const row = (v) => { const r = head.map(() => ''); r[idx] = v; return r; };
  const cleanRows = [head, row(''), row('업로드 안 함(딜 의심)'), row(''), row('')];
  const items = [{ row: 1 }, { row: 3 }, { row: 4 }];
  const results = [{ id: 501 }, { error: '400 bad', warn: '' }, { skipped: true }];
  assert.deepStrictEqual(uploadColumn_(cleanRows, items, results, { blocked: false, off: false }),
    [['501'], ['업로드 안 함(딜 의심)'], ['실패: 400 bad'], ['남음(시간 초과) — PQL 생성을 다시 누르면 이어서 올라감']]);
  assert.deepStrictEqual(uploadColumn_(cleanRows, [], [], { blocked: true, off: false })[0], ['업로드 안 함(대상 500곳 초과)']);
  assert.deepStrictEqual(uploadColumn_(cleanRows, [], [], { blocked: false, off: true })[0], ['업로드 꺼짐']);
  assert.deepStrictEqual(uploadColumn_(cleanRows, [{ row: 1 }], [{ id: 7, warn: '담당자 생성 실패' }], { blocked: false, off: false })[0], ['7 (담당자 생성 실패)']);
});

// 업로드 이력: 실행마다 (타겟 × 세일즈티어)별 업로드 수와 전체 합계를 누적한다. 월 = 원천 파일 기준 PQL 월.
test('PQL 월: 원천 파일 MM, 연도는 업로드일 기준(연말·연초 넘김 처리)', () => {
  assert.strictEqual(pqlMonth_('all_subscription_1001.csv', '2026-09-30'), '2026-10');
  assert.strictEqual(pqlMonth_('all_subscription_0101.csv', '2026-12-31'), '2027-01');
  assert.strictEqual(pqlMonth_('all_subscription_1231.csv', '2027-01-02'), '2026-12');
  assert.strictEqual(pqlMonth_('all_subscription_004141.csv', '2026-06-24'), '2026-06'); // MMDD가 아니면 업로드일의 달
});

test('업로드 이력 행: 실제 생성된 딜만, 타겟·세일즈티어 순으로 집계하고 전체 합계', () => {
  const it = (target, tier) => ({ target: target, tier: tier });
  const items = [it('업셀', '베이직'), it('업셀', '베이직'), it('업셀', '그로스'), it('푸시, 리뷰', '비즈니스'), it('업셀', '라이트')];
  const results = [{ id: 1 }, { id: 2 }, { id: 3 }, { id: 4 }, { error: 'x' }];
  assert.deepStrictEqual(uploadSummaryRows_(items, results, '2026-10', '2026-09-30'), [
    ['2026-10', '2026-09-30', '업셀', '베이직', 2],
    ['2026-10', '2026-09-30', '업셀', '그로스', 1],
    ['2026-10', '2026-09-30', '푸시, 리뷰', '비즈니스', 1],
    ['2026-10', '2026-09-30', '전체', '전체', 4],
  ]);
  assert.deepStrictEqual(uploadSummaryRows_(items, items.map(() => ({ skipped: true })), '2026-10', '2026-09-30'), []);
});
