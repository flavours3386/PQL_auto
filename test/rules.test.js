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

// 업셀은 주문 150건 이상 (2026-09-30 사용자 결정: 101건 통과는 너무 느슨함)
test('타겟: 업셀', () => {
  assert.deepStrictEqual(matchTargets_(rec({ '최근 30일 플랫폼 주문수': '200' })), ['업셀']);
  assert.deepStrictEqual(matchTargets_(rec({ '최근 30일 플랫폼 주문수': '150' })), ['업셀']);
  assert.deepStrictEqual(matchTargets_(rec({ '최근 30일 플랫폼 주문수': '149' })), []);
  assert.deepStrictEqual(matchTargets_(rec({ '최근 30일 플랫폼 주문수': '200', '알파업셀 상태': '라이브(무료구독중)' })), []);
  assert.deepStrictEqual(matchTargets_(rec({ '최근 30일 플랫폼 주문수': '200', '알파업셀 상태': '제거중' })), []);
  assert.deepStrictEqual(matchTargets_(rec({ 플랫폼: 'imweb' })), []);
});

// 푸시 무료(카페24 PRO 번들)도 라이브로 본다: 세 제품이 모두 라이브인 몰(널핏·척도)은 올리지 않는다 (2026-09-30 사용자 결정)
test('타겟: 푸시는 cafe24·주문 500 이상·라이브 아님(무료 포함)', () => {
  const up = { '알파업셀 상태': '라이브(과금중)' };
  assert.deepStrictEqual(matchTargets_(rec(Object.assign({ '알파푸시 상태': '라이브(무료구독중)' }, up))), []);
  assert.deepStrictEqual(matchTargets_(rec(Object.assign({ '알파푸시 상태': '서비스중단' }, up))), ['푸시']);
  assert.deepStrictEqual(matchTargets_(rec(Object.assign({ '알파푸시 상태': '프로덕트온보딩중' }, up))), ['푸시']);
  assert.deepStrictEqual(matchTargets_(rec(Object.assign({ '알파푸시 상태': '라이브(과금중)' }, up))), []);
  assert.deepStrictEqual(matchTargets_(rec(Object.assign({ '알파푸시 상태': '라이브(체험중)' }, up))), []);
  assert.deepStrictEqual(matchTargets_(rec(Object.assign({ '최근 30일 플랫폼 주문수': '499' }, up))), []);
  assert.deepStrictEqual(matchTargets_(rec()), ['업셀', '푸시']);
});

test('타겟: 리뷰는 주문 1,000건 이상·리뷰 라이브 아님, 아임웹 포함', () => {
  const others = { '알파업셀 상태': '라이브(과금중)', '알파푸시 상태': '라이브(무료구독중)' };
  const r = (o) => matchTargets_(rec(Object.assign({}, others, o)));
  assert.deepStrictEqual(r({ '알파리뷰 상태': '구독없음', '최근 30일 플랫폼 주문수': '1000' }), ['리뷰']);
  assert.deepStrictEqual(r({ '알파리뷰 상태': '구독없음', '최근 30일 플랫폼 주문수': '999' }), []);
  assert.deepStrictEqual(r({ '알파리뷰 상태': '프로덕트온보딩중', '최근 30일 플랫폼 주문수': '1500', 플랫폼: 'imweb' }), ['리뷰']);
  assert.deepStrictEqual(r({ '알파리뷰 상태': '라이브(무료구독중)', '최근 30일 플랫폼 주문수': '1500' }), []);
  assert.deepStrictEqual(matchTargets_(rec({ '알파리뷰 상태': '구독없음', '최근 30일 플랫폼 주문수': '1200' })), ['업셀', '푸시', '리뷰']);
});

test('서비스 라벨', () => {
  assert.strictEqual(serviceLabel_(rec({ '알파리뷰 상태': '구독없음' })), 'null');
  assert.strictEqual(serviceLabel_(rec({ '알파업셀 상태': '라이브(과금중)', '알파푸시 상태': '라이브(무료구독중)' })), '알파리뷰, 알파업셀, 알파푸시');
  assert.strictEqual(serviceLabel_(rec({ '알파리뷰 상태': '제거중' })), '');
});

test('clean 행 34열, 타겟·라벨·딜 의심·업로드 위치', () => {
  const row = cleanRow_(rec(), ['업셀', '푸시'], '알파리뷰', ['77', '88'], '업로드 안 함(딜 의심)');
  assert.strictEqual(row.length, 34);
  assert.deepStrictEqual(row.slice(0, 9), ['몰', '1', 'm1', 'cafe24', '610', '업셀, 푸시', '알파리뷰', '77, 88', '업로드 안 함(딜 의심)']);
  assert.strictEqual(row[12], '010-1234-5678');
  assert.strictEqual(row[15], '서울 1층');
});

const IDS = { ownerId: 24324011, stageId: 71, labelIds: { 알파리뷰: 299, 알파업셀: 300, 알파푸시: 301, null: 303 } };

// 0901 가져오기로 만든 딜과 같은 배치 (딜 필드·담당자·조직)
test('업로드 재료: 딜·담당자·조직', () => {
  const it = uploadItem_(rec(), '알파리뷰, 알파푸시', IDS, 5);
  assert.strictEqual(it.row, 5);
  assert.deepStrictEqual(it.org, { name: '둘째회사', address: '서울 1층' });
  assert.deepStrictEqual(it.person, { name: '홍길동', email: 'a@b.com', phone: '010-1234-5678' });
  assert.deepStrictEqual(it.deal, {
    title: '몰', owner_id: 24324011, pipeline_id: 9, stage_id: 71, label_ids: [299, 301],
    custom_fields: {
      [PD_FIELD_SHOP_ID]: '1', [PD_FIELD_MALL_ID]: 'm1', [PD_FIELD_HOSTING]: 388, [PD_FIELD_MONTHLY_ORDERS]: 600,
      [PD_FIELD_MALL_NAME]: '몰', [PD_FIELD_URL]: 'mall.com', [PD_FIELD_SALES_TIER]: 234,
    },
  });
});

test('업로드 재료: 빈 값은 빼고, 라벨 null·빈 라벨·아임웹 처리', () => {
  const it = uploadItem_(rec({ 플랫폼: 'imweb', 대표도메인: '', 담당자명: '', 담당자이메일: '' }), 'null', IDS, 1);
  assert.deepStrictEqual(it.deal.label_ids, [303]);
  assert.strictEqual(it.deal.custom_fields[PD_FIELD_HOSTING], 389);
  assert.ok(!(PD_FIELD_URL in it.deal.custom_fields));
  assert.strictEqual(it.person.name, '몰'); // 담당자명이 없으면 shop_name
  assert.deepStrictEqual(personPayload_(it.person, 7), { name: '몰', org_id: 7, phones: [{ value: '010-1234-5678', primary: true, label: 'work' }] });
  assert.deepStrictEqual(uploadItem_(rec(), '', IDS, 1).deal.label_ids, []);
});

// 플랜이 선택지 이름이면 그대로, '-'·빈값·모르는 값이면 월 주문수 구간(상한 이하)으로 판정
test('세일즈티어: 플랜 우선, 없으면 주문수 구간', () => {
  const t = (plan, orders) => { const x = salesTier_(plan, orders); return x && x.name; };
  assert.strictEqual(t('베이직', 5000), '베이직');
  assert.strictEqual(t(' 엔터프라이즈6 ', 10), '엔터프라이즈6');
  assert.strictEqual(t('-', 100), '라이트');
  assert.strictEqual(t('-', 101), '베이직');
  assert.strictEqual(t('-', 1000), '베이직');
  assert.strictEqual(t('-', 1001), '그로스');
  assert.strictEqual(t('', 5000), '비즈니스');
  assert.strictEqual(t('-', 8195), '엔터프라이즈2');
  assert.strictEqual(t('-', 50000), '엔터프라이즈5');
  assert.strictEqual(t('-', 50001), '엔터프라이즈6');
  assert.strictEqual(t('모르는값', 150), '베이직');
  assert.strictEqual(salesTier_('-', NaN), null);
  assert.strictEqual(salesTier_('엔터프라이즈1', 1).id, 237);
});

test('업로드 재료: 세일즈티어는 CSV 플랜 열을 쓴다', () => {
  const H2 = H.concat(['플랜']);
  const hi2 = headerIndex_(H2);
  const row = H2.map((h) => (BASE[h] === undefined ? '' : BASE[h]));
  row[H2.length - 1] = '엔터프라이즈3';
  assert.strictEqual(uploadItem_(toRecord_(row, hi2), '', IDS, 1).deal.custom_fields[PD_FIELD_SALES_TIER], 239);
});

test('조직·딜 요청 본문', () => {
  assert.deepStrictEqual(orgPayload_({ name: '회사', address: '서울' }), { name: '회사', address: { value: '서울' } });
  assert.deepStrictEqual(orgPayload_({ name: '회사', address: '' }), { name: '회사' });
  const it = uploadItem_(rec(), '알파리뷰', IDS, 1);
  const d = dealPayload_(it, 11, 22);
  assert.strictEqual(d.person_id, 11);
  assert.strictEqual(d.org_id, 22);
  assert.ok(!('person_id' in dealPayload_(it, null, 22)));
  assert.ok(!('person_id' in it.deal)); // 원본 재료는 건드리지 않는다
});
