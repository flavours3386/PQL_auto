require('./gas');
const test = require('node:test');
const assert = require('node:assert');

const H = ['shop_id', 'mall_id', 'shop_name', '플랫폼', '최근 30일 플랫폼 주문수', '알파리뷰 상태', '알파업셀 상태', '알파푸시 상태', '사이트 상태',
  '회사명', '쇼핑몰명', '담당자명', '담당자전화번호', '담당자이메일', '대표도메인'];
const row = (id, o) => {
  const v = Object.assign({ shop_id: id, mall_id: 'm' + id, shop_name: '몰' + id, 플랫폼: 'cafe24', '최근 30일 플랫폼 주문수': '600',
    '알파리뷰 상태': '라이브(과금중)', '알파업셀 상태': '구독없음', '알파푸시 상태': '구독없음', '사이트 상태': '라이브',
    회사명: '회사' + id, 쇼핑몰명: '몰' + id, 담당자명: '홍', 담당자전화번호: '010-0000-00' + id.padStart(2, '0'), 담당자이메일: id + '@m.com', 대표도메인: id + '.com' }, o || {});
  return H.map((h) => v[h]);
};

const IDS = { ownerId: 24324011, stageId: 71, labelIds: { 알파리뷰: 299, 알파업셀: 300, 알파푸시: 301, null: 303 } };

function build(rows, opts) {
  const b = createPqlBuilder_(Object.assign({ dealShopIds: new Set(), unmappedDeals: [], rejectedPairs: new Set(), uploadIds: IDS }, opts || {}));
  [H].concat(rows).forEach((r) => b.onRow(r));
  return b.finish();
}

test('단계별 탈락·역매핑 제외·딜 의심·타겟 집계', () => {
  const unmappedDeals = [
    { id: 900, title: '몰5', raw: '채널톡', keys: { email: ['5@m.com'], phone: [], name: ['몰5'], url: [] } }, // high → shop 5 제외
    { id: 901, title: 'x', raw: '', keys: { email: [], phone: [], name: [], url: ['6.com'] } }, // review → shop 6 딜 의심
  ];
  const out = build([
    row('1', { '최근 30일 플랫폼 주문수': '50' }),
    row('2', { '알파리뷰 상태': '서비스중단' }),
    row('3', { 담당자전화번호: '' }),
    row('4'), // 숫자 shop_id 딜
    row('5'), // 역매핑 high
    row('6'), // 딜 의심 → clean에는 있고 xlsx에는 없음
    row('7', { '알파업셀 상태': '라이브(과금중)', '알파푸시 상태': '라이브(과금중)' }), // 타겟 없음
    row('8', { 플랫폼: 'imweb' }), // 타겟 없음
    row('9', { '최근 30일 플랫폼 주문수': '200' }), // 업셀만
    [''], // 빈 줄
  ], { dealShopIds: new Set(['4']), unmappedDeals: unmappedDeals });
  assert.deepStrictEqual(out.counts, { total: 9, orders: 1, review: 1, site: 0, pro: 0, phone: 1, deal: 1, mapped: 1, noTarget: 2 });
  assert.deepStrictEqual(out.cleanRows.slice(1).map((r) => r[1]), ['6', '9']);
  assert.strictEqual(out.cleanRows[1][7], '901');
  assert.strictEqual(out.cleanRows[1][8], '업로드 안 함(딜 의심)');
  assert.strictEqual(out.cleanRows[2][8], '');
  assert.deepStrictEqual(out.uploadItems.map((it) => [it.row, it.shopId, it.deal.title]), [[2, '9', '몰9']]);
  assert.deepStrictEqual(out.targetCounts, { '업셀, 푸시': 1, 업셀: 1 });
  assert.deepStrictEqual(out.matches.map((m) => m.tier), ['high', 'review']);
});

test('필수 열이 없으면 없는 열 이름을 모두 띄우고 멈춘다', () => {
  const b = createPqlBuilder_({ dealShopIds: new Set(), unmappedDeals: [], rejectedPairs: new Set(), uploadIds: IDS });
  assert.throws(() => b.onRow(H.filter((h) => h !== '알파업셀 상태' && h !== '담당자명')), /필수 열이 없습니다: 알파업셀 상태, 담당자명/);
});

test('빈 CSV는 멈춘다', () => {
  const b = createPqlBuilder_({ dealShopIds: new Set(), unmappedDeals: [], rejectedPairs: new Set(), uploadIds: IDS });
  assert.throws(() => b.finish(), /CSV가 비어 있습니다/);
});
