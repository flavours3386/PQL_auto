require('./gas');
const test = require('node:test');
const assert = require('node:assert');

const H = ['shop_id', 'mall_id', 'shop_name', '플랫폼', '회사명', '최근 30일 플랫폼 주문수', '알파리뷰 상태', '알파업셀 상태', '알파푸시 상태', '사이트 상태',
  '쇼핑몰명', '회사명', '주소1', '주소2', '담당자명', '담당자전화번호', '담당자이메일', '대표도메인', '최근 30일 플랫폼 주문수(API)'];
const hi = headerIndex_(H);
const BASE = {
  shop_id: '1', mall_id: 'm1', shop_name: '몰', 플랫폼: 'cafe24', '최근 30일 플랫폼 주문수': '600',
  '알파리뷰 상태': '라이브(과금중)', '알파업셀 상태': '구독없음', '알파푸시 상태': '구독없음', '사이트 상태': '라이브',
  쇼핑몰명: '몰', 주소1: '서울', 주소2: '1층', 담당자명: '홍길동', 담당자전화번호: '1012345678', 담당자이메일: 'a@b.com',
  대표도메인: 'mall.com', '최근 30일 플랫폼 주문수(API)': '610',
};
function rec(over) {
  const v = Object.assign({}, BASE, over || {});
  const row = H.map((h) => (v[h] === undefined ? '' : v[h]));
  row[4] = '첫회사';
  row[11] = '둘째회사';
  return toRecord_(row, hi);
}

test('전화번호 정규화', () => {
  assert.deepStrictEqual(normalizePhone_('1012345678'), { display: '010-1234-5678', digits: '01012345678' });
  assert.deepStrictEqual(normalizePhone_('070-1234-5678'), { display: '070-1234-5678', digits: '07012345678' });
  assert.deepStrictEqual(normalizePhone_(''), { display: '', digits: '' });
});

test('같은 이름 열이 둘이면 get은 마지막 열', () => {
  assert.strictEqual(rec().get('회사명'), '둘째회사');
});

test('공통 클렌징 사유', () => {
  const none = new Set();
  assert.strictEqual(cleanseReason_(rec(), none), '');
  assert.strictEqual(cleanseReason_(rec({ '최근 30일 플랫폼 주문수': '' }), none), 'orders');
  assert.strictEqual(cleanseReason_(rec({ '최근 30일 플랫폼 주문수': '99' }), none), 'orders');
  assert.strictEqual(cleanseReason_(rec({ '최근 30일 플랫폼 주문수': 'abc' }), none), 'orders');
  assert.strictEqual(cleanseReason_(rec({ '최근 30일 플랫폼 주문수': '100' }), none), '');
  assert.strictEqual(cleanseReason_(rec({ '알파리뷰 상태': '서비스 중단' }), none), 'review');
  assert.strictEqual(cleanseReason_(rec({ '알파리뷰 상태': '서비스중단' }), none), 'review');
  assert.strictEqual(cleanseReason_(rec({ '사이트 상태': '계정활성화' }), none), 'site');
  assert.strictEqual(cleanseReason_(rec({ 담당자명: '프로', 담당자전화번호: '070-1234-5678' }), none), 'pro');
  assert.strictEqual(cleanseReason_(rec({ 담당자명: '프로', 담당자전화번호: '010-1234-5678' }), none), '');
  assert.strictEqual(cleanseReason_(rec({ 담당자전화번호: '' }), none), 'phone');
  assert.strictEqual(cleanseReason_(rec(), new Set(['1'])), 'deal');
});

test('타겟: 업셀', () => {
  assert.deepStrictEqual(matchTargets_(rec({ '최근 30일 플랫폼 주문수': '200' })), ['업셀']);
  assert.deepStrictEqual(matchTargets_(rec({ '최근 30일 플랫폼 주문수': '200', '알파업셀 상태': '라이브(무료구독중)' })), []);
  assert.deepStrictEqual(matchTargets_(rec({ '최근 30일 플랫폼 주문수': '200', '알파업셀 상태': '제거중' })), []);
  assert.deepStrictEqual(matchTargets_(rec({ 플랫폼: 'imweb' })), []);
});

test('타겟: 푸시는 cafe24·주문 500 이상·유료 미사용', () => {
  const up = { '알파업셀 상태': '라이브(과금중)' };
  assert.deepStrictEqual(matchTargets_(rec(Object.assign({ '알파푸시 상태': '라이브(무료구독중)' }, up))), ['푸시']);
  assert.deepStrictEqual(matchTargets_(rec(Object.assign({ '알파푸시 상태': '라이브(과금중)' }, up))), []);
  assert.deepStrictEqual(matchTargets_(rec(Object.assign({ '알파푸시 상태': '라이브(체험중)' }, up))), []);
  assert.deepStrictEqual(matchTargets_(rec(Object.assign({ '최근 30일 플랫폼 주문수': '499' }, up))), []);
  assert.deepStrictEqual(matchTargets_(rec()), ['업셀', '푸시']);
});

test('서비스 라벨', () => {
  assert.strictEqual(serviceLabel_(rec({ '알파리뷰 상태': '구독없음' })), 'null');
  assert.strictEqual(serviceLabel_(rec({ '알파업셀 상태': '라이브(과금중)', '알파푸시 상태': '라이브(무료구독중)' })), '알파리뷰, 알파업셀, 알파푸시');
  assert.strictEqual(serviceLabel_(rec({ '알파리뷰 상태': '제거중' })), '');
});

test('clean 행 33열, 타겟·라벨·딜 의심 위치', () => {
  const row = cleanRow_(rec(), ['업셀', '푸시'], '알파리뷰', ['77', '88']);
  assert.strictEqual(row.length, 33);
  assert.deepStrictEqual(row.slice(0, 8), ['몰', '1', 'm1', 'cafe24', '610', '업셀, 푸시', '알파리뷰', '77, 88']);
  assert.strictEqual(row[11], '010-1234-5678');
  assert.strictEqual(row[14], '서울 1층');
});

test('업로드 행 15열 = 0901 양식', () => {
  assert.deepStrictEqual(uploadRow_(rec(), '알파리뷰'), [
    '몰', 1, 'm1', 'cafe24', 600, '한서연', '컨택전', '알파리뷰', '둘째회사', '홍길동', '몰', '010-1234-5678', 'a@b.com', 'mall.com', '서울 1층',
  ]);
});
