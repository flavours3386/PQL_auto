require('./gas');
const test = require('node:test');
const assert = require('node:assert');

const cf = (o) => ({ custom_fields: o });

test('키 정규화', () => {
  assert.strictEqual(nameKey_('(주) 서브 마켓!'), '서브마켓');
  assert.strictEqual(nameKey_('A'), '');
  assert.strictEqual(domainKey_('https://www.ServeMarket.kr/index.html?x=1'), 'servemarket.kr');
  assert.strictEqual(domainKey_('채널톡'), '');
  assert.strictEqual(dealUrlKey_('blackteddys'), 'blackteddys.cafe24.com');
  assert.strictEqual(dealUrlKey_('채널톡'), '');
  assert.strictEqual(emailKey_(' A@B.com '), 'a@b.com');
  assert.strictEqual(phoneKey_('1012345678'), '01012345678');
  assert.strictEqual(phoneKey_('1234'), '');
});

test('딜 키: 제목·조직·쇼핑몰명, URL 필드, 텍스트 shop_id', () => {
  const deal = Object.assign({ title: '서브마켓' }, cf({
    [PD_FIELD_SHOP_ID]: 'blackteddys', [PD_FIELD_URL]: 'https://www.servemarket.kr/', [PD_FIELD_MALL_NAME]: '서브 마켓몰',
  }));
  const person = { emails: [{ value: 'CEO@servemarket.kr' }], phones: [{ value: '010-1111-2222' }] };
  const keys = dealMatchKeys_(deal, person, { name: '(주)서브' });
  assert.deepStrictEqual(keys.email, ['ceo@servemarket.kr']);
  assert.deepStrictEqual(keys.phone, ['01011112222']);
  assert.deepStrictEqual(keys.name, ['서브마켓', '서브', '서브마켓몰']);
  assert.deepStrictEqual(keys.url, ['servemarket.kr', 'blackteddys.cafe24.com']);
  assert.strictEqual(rawShopId_(deal), 'blackteddys');
  assert.strictEqual(rawShopId_({}), '');
});

test('CSV 행 키: 회사명 두 열 모두, mall_id는 cafe24 도메인', () => {
  const H = ['shop_id', 'mall_id', 'shop_name', '회사명', '회사명', '대표도메인', '담당자이메일'];
  const keys = rowMatchKeys_(['10', 'abc', '몰A', '회사가', '회사나', 'www.a.com', 'x@y.com'], headerIndex_(H));
  assert.deepStrictEqual(keys.name, ['몰a', '회사가', '회사나']);
  assert.deepStrictEqual(keys.url, ['a.com', 'abc.cafe24.com']);
  assert.deepStrictEqual(keys.email, ['x@y.com']);
});

test('판정: 장신몰형(전화 후보 2 ∩ 이름 1) → high', () => {
  const c = classifyMatch_({ phone: new Set(['126757', '6332']), name: new Set(['126757']) });
  assert.strictEqual(c.tier, 'high');
  assert.deepStrictEqual(c.candidates, [{ shopId: '126757', keys: ['phone', 'name'] }]);
});

test('판정: 키 1개 → review, 키끼리 다른 shop → review, 없음 → none', () => {
  assert.deepStrictEqual(classifyMatch_({ email: new Set(['1']) }), { tier: 'review', candidates: [{ shopId: '1', keys: ['email'] }] });
  assert.deepStrictEqual(classifyMatch_({ email: new Set(['1']), name: new Set(['2']) }), {
    tier: 'review', candidates: [{ shopId: '1', keys: ['email'] }, { shopId: '2', keys: ['name'] }],
  });
  assert.deepStrictEqual(classifyMatch_({}), { tier: 'none', candidates: [] });
});

test('매처: CSV를 흘려 후보를 모으고, 거절한 쌍은 뺀다', () => {
  const H = ['shop_id', 'shop_name', '담당자이메일', '담당자전화번호'];
  const hi = headerIndex_(H);
  const deals = [
    { id: 1, title: 'd1', raw: '채널톡', keys: { email: ['e@x.com'], phone: ['01011112222'], name: [], url: [] } },
    { id: 2, title: 'd2', raw: '', keys: { email: [], phone: [], name: ['같은이름'], url: [] } },
  ];
  const run = (rejected) => {
    const m = createDealMatcher_(deals);
    [['10', '몰10', 'e@x.com', '010-1111-2222'], ['20', '같은이름', '', ''], ['21', '같은 이름', '', '']].forEach((row) => m.onRow(row, hi, row[0], row[1]));
    return m.results(rejected);
  };
  const r = run(new Set());
  assert.strictEqual(r[0].tier, 'high');
  assert.deepStrictEqual(r[0].candidates, [{ shopId: '10', keys: ['email', 'phone'], shopName: '몰10' }]);
  assert.strictEqual(r[1].tier, 'review');
  assert.deepStrictEqual(r[1].candidates.map((c) => c.shopId), ['20', '21']);
  assert.strictEqual(run(new Set(['1:10']))[0].tier, 'none');
});
