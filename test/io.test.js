require('./gas');
const test = require('node:test');
const assert = require('node:assert');

// 가짜 Pipedrive: 경로별 응답을 흉내 낸다 (네트워크 없음). fail(route, body) → 상태 코드를 주면 그 요청을 실패시킨다.
function fakePipedrive(opts) {
  opts = opts || {};
  const db = { orgs: Object.assign({}, opts.orgs), deals: JSON.parse(JSON.stringify(opts.deals || {})), persons: {}, notes: [], nextId: 1000 };
  const calls = [];
  const sleeps = [];
  const respond = (code, body) => ({ getResponseCode: () => code, getContentText: () => JSON.stringify(body) });
  function handle(method, url, payload) {
    const path = url.replace('https://api.pipedrive.com', '');
    const route = method.toUpperCase() + ' ' + path.split('?')[0].replace(/\/\d+$/, '/:id');
    const body = payload ? JSON.parse(payload) : null;
    calls.push({ route: route, path: path, body: body });
    const code = opts.fail && opts.fail(route, body);
    if (code) return respond(code, { error: 'fake ' + code });
    const idOf = () => /\/(\d+)$/.exec(path.split('?')[0])[1];
    if (route === 'GET /api/v2/organizations/search') {
      const term = decodeURIComponent(/term=([^&]*)/.exec(path)[1]);
      return respond(200, { data: { items: db.orgs[term] ? [{ item: { id: db.orgs[term], name: term } }] : [] } });
    }
    if (route === 'POST /api/v2/organizations') { const id = db.nextId++; db.orgs[body.name] = id; return respond(201, { data: { id: id } }); }
    if (route === 'POST /api/v2/persons') { const id = db.nextId++; db.persons[id] = body; return respond(201, { data: { id: id } }); }
    if (route === 'POST /api/v2/deals') { const id = db.nextId++; db.deals[id] = body; return respond(201, { data: { id: id } }); }
    if (route === 'GET /api/v2/deals/:id') {
      const d = db.deals[idOf()] || {};
      return respond(200, { data: { id: Number(idOf()), custom_fields: d.custom_fields || {} } });
    }
    if (route === 'PATCH /api/v2/deals/:id') {
      const d = (db.deals[idOf()] = db.deals[idOf()] || {});
      d.custom_fields = Object.assign({}, d.custom_fields, body.custom_fields);
      return respond(200, { data: { id: Number(idOf()) } });
    }
    if (route === 'POST /api/v1/notes') { db.notes.push(body); return respond(201, { data: { id: 1 } }); }
    if (route === 'GET /api/v2/deals') {
      const all = opts.listDeals || [];
      const start = Number((/cursor=(\d+)/.exec(path) || [0, 0])[1]);
      const page = all.slice(start, start + 2);
      return respond(200, { data: page, additional_data: { next_cursor: start + 2 < all.length ? String(start + 2) : null } });
    }
    if (route === 'GET /api/v2/persons' || route === 'GET /api/v2/organizations') {
      const ids = decodeURIComponent(/ids=([^&]*)/.exec(path)[1]).split(',').map(Number);
      return respond(200, { data: ids.map((id) => ({ id: id, name: 'n' + id, emails: [{ value: 'e' + id + '@x.com' }], phones: [], extra: 'x'.repeat(50) })) });
    }
    if (route === 'GET /api/v1/users') return respond(200, { data: [{ id: 24324011, name: '한서연' }] });
    if (route === 'GET /api/v1/stages') return respond(200, { data: [{ id: 71, name: '컨택전' }] });
    if (route === 'GET /api/v1/dealFields') return respond(200, { data: [{ key: 'label', options: [{ id: 299, label: '알파리뷰' }] }] });
    return respond(404, { error: 'no route ' + route });
  }
  global.UrlFetchApp = {
    fetch: (url, o) => handle(o.method || 'get', url, o.payload),
    fetchAll: (reqs) => (opts.fetchAllThrows && opts.fetchAllThrows() ? (() => { throw new Error('fetchAll 장애'); })() : reqs.map((r) => handle(r.method || 'get', r.url, r.payload))),
  };
  global.Utilities = { sleep: (ms) => sleeps.push(ms) };
  return { db: db, calls: calls, sleeps: sleeps, count: (route) => calls.filter((c) => c.route === route).length };
}

const item = (n, org) => ({
  row: n, shopId: String(n),
  org: { name: org || '조직' + n, address: '' },
  person: { name: 'p' + n, email: '', phone: '010-0000-000' + n },
  deal: { title: 't' + n, owner_id: 1, pipeline_id: 9, stage_id: 71, label_ids: [], custom_fields: { [PD_FIELD_SHOP_ID]: String(n) } },
});
const FUTURE = () => Date.now() + 600000;

test('업로드 정상: 조직 생성 → 담당자 → 딜, 같은 이름 조직은 한 번만 만든다', () => {
  const pd = fakePipedrive({ orgs: { 기존조직: 7 } });
  const res = pdCreateDeals_('t', [item(1, '기존조직'), item(2, '새조직'), item(3, '새조직')], FUTURE());
  assert.deepStrictEqual(res.map((r) => !!r.id), [true, true, true]);
  assert.strictEqual(pd.count('POST /api/v2/organizations'), 1);
  assert.strictEqual(pd.db.deals[res[0].id].org_id, 7);
});

// 리뷰 #5: 검색이 실패하면 없는 것으로 보고 동명 조직을 새로 만들면 안 된다
test('조직 검색 실패면 그 몰은 조직·담당자·딜을 만들지 않는다', () => {
  const pd = fakePipedrive({ fail: (route, body) => (route === 'GET /api/v2/organizations/search' ? 500 : 0) });
  const res = pdCreateDeals_('t', [item(1)], FUTURE());
  assert.match(res[0].error, /조직 검색 실패/);
  assert.strictEqual(pd.count('POST /api/v2/organizations'), 0);
  assert.strictEqual(pd.count('POST /api/v2/persons'), 0);
  assert.strictEqual(pd.count('POST /api/v2/deals'), 0);
});

// 리뷰 #6: 선행 단계가 실패했는데 딜만 만들면 다음 실행에서 제외돼 영영 복구되지 않는다
test('조직·담당자 생성이 실패하면 딜을 만들지 않는다 (다음 실행에서 다시 시도)', () => {
  let pd = fakePipedrive({ fail: (route, body) => (route === 'POST /api/v2/persons' && body.name === 'p2' ? 400 : 0) });
  let res = pdCreateDeals_('t', [item(1), item(2)], FUTURE());
  assert.ok(res[0].id);
  assert.match(res[1].error, /담당자 생성 실패/);
  assert.strictEqual(pd.count('POST /api/v2/deals'), 1);
  pd = fakePipedrive({ fail: (route) => (route === 'POST /api/v2/organizations' ? 400 : 0) });
  res = pdCreateDeals_('t', [item(1)], FUTURE());
  assert.match(res[0].error, /조직 생성 실패/);
  assert.strictEqual(pd.count('POST /api/v2/persons') + pd.count('POST /api/v2/deals'), 0);
});

// 리뷰 #4: 시간 예산이 지나면 새 묶음을 시작하지 않고, 429 재시도도 예산을 넘겨 기다리지 않는다
test('시간 예산: 지나면 남은 곳은 skipped, 429 재시도는 예산 안에서만', () => {
  let pd = fakePipedrive();
  let res = pdCreateDeals_('t', [item(1), item(2)], Date.now() - 1);
  assert.deepStrictEqual(res, [{ skipped: true }, { skipped: true }]);
  assert.strictEqual(pd.calls.length, 0);
  pd = fakePipedrive({ fail: (route) => (route === 'GET /api/v2/organizations/search' ? 429 : 0) });
  const r = pdFetchAll_('t', [{ method: 'get', path: '/api/v2/organizations/search?term=x' }], Date.now() + 1000);
  assert.strictEqual(r[0].ok, false);
  assert.deepStrictEqual(pd.sleeps, []); // 2초를 기다리면 예산을 넘으므로 재시도하지 않는다
});

test('한 묶음에서 예외가 나도 그 묶음만 실패로 적고 다음 묶음은 진행한다', () => {
  let first = true;
  const pd = fakePipedrive({ fetchAllThrows: () => { const t = first; first = false; return t; } });
  const items = Array.from({ length: UPLOAD_BATCH + 1 }, (_, i) => item(i + 1));
  const res = pdCreateDeals_('t', items, FUTURE());
  assert.strictEqual(res.length, items.length);
  assert.ok(res.slice(0, UPLOAD_BATCH).every((r) => /예외/.test(r.error)));
  assert.ok(res[UPLOAD_BATCH].id);
  assert.strictEqual(pd.count('POST /api/v2/deals'), 1);
});

// 리뷰 #2: 처음 조회 뒤 누가 숫자 shop_id를 채웠으면 덮어쓰지 않는다
test('shop_id 반영 직전에 다시 읽는다: 숫자면 건너뜀, 값이 바뀌었으면 보류, 후보가 숫자 아니면 실패', () => {
  const pd = fakePipedrive({ deals: { 5: { custom_fields: { [PD_FIELD_SHOP_ID]: '202' } }, 6: { custom_fields: { [PD_FIELD_SHOP_ID]: '소개' } }, 7: { custom_fields: { [PD_FIELD_SHOP_ID]: '랜딩' } } } });
  const note = (cur) => 'note ' + cur;
  assert.deepStrictEqual(pdApplyShopId_('t', 5, '101', '', note), { applied: false, status: '건너뜀: 이미 shop_id 202' });
  assert.deepStrictEqual(pdApplyShopId_('t', 6, '101', '채널톡', note), { applied: false, status: '보류: shop_id가 바뀜(소개)' });
  assert.deepStrictEqual(pdApplyShopId_('t', 7, 'abc', '랜딩', note), { applied: false, status: '실패: 후보 shop_id가 숫자가 아님' });
  assert.strictEqual(pd.count('PATCH /api/v2/deals/:id'), 0);
  assert.deepStrictEqual(pdApplyShopId_('t', 7, '303', '랜딩', note), { applied: true, status: '반영됨' });
  assert.strictEqual(pd.db.deals[7].custom_fields[PD_FIELD_SHOP_ID], '303');
  assert.deepStrictEqual(pd.db.notes, [{ deal_id: 7, content: 'note 랜딩' }]);
});

test('자동 반영도 같은 확인을 거친다', () => {
  const pd = fakePipedrive({ deals: { 5: { custom_fields: { [PD_FIELD_SHOP_ID]: '202' } } } });
  const m = { deal: { id: 5, title: 'd', raw: '' }, tier: 'high', candidates: [{ shopId: '101', keys: ['email', 'name'], shopName: 's' }] };
  const out = autoApply_('t', [m], '2026-10-01', FUTURE());
  assert.deepStrictEqual(out.overrides, {});
  assert.strictEqual(out.rows[0][8], '건너뜀: 이미 shop_id 202');
  assert.strictEqual(pd.count('PATCH /api/v2/deals/:id'), 0);
});

// 리뷰 #1: 두 실행이 겹치면 같은 몰 딜을 두 번 만든다 → 잠금을 못 얻으면 아무것도 하지 않는다
test('다른 실행이 잠금을 쥐고 있으면 PQL 생성·승인 반영은 쓰기 없이 멈춘다', () => {
  const pd = fakePipedrive();
  const alerts = [];
  let released = 0;
  global.SpreadsheetApp = {
    getUi: () => ({ alert: (...a) => alerts.push(a.join(' ')), ButtonSet: { OK: 'OK' } }),
    getActiveSpreadsheet: () => ({}),
  };
  global.PropertiesService = { getScriptProperties: () => ({ getProperty: () => 'tok' }) };
  global.LockService = { getScriptLock: () => ({ tryLock: () => false, releaseLock: () => released++ }) };
  runPql();
  applyApprovedMappings();
  assert.strictEqual(pd.calls.length, 0);
  assert.strictEqual(alerts.length, 2);
  assert.ok(alerts.every((a) => /다른 실행이 진행 중/.test(a)));
  assert.strictEqual(released, 0);
});

// 리뷰 #9: 기존 시트를 먼저 지우고 쓰다 실패하면 거절·반영 이력을 잃는다
test('매핑 탭은 새 값을 먼저 쓰고 남는 아래 행만 지운다', () => {
  const ops = [];
  const range = (r, c, nr, nc) => ({
    setNumberFormat: () => ops.push(['format', r, nr]), setValues: (v) => ops.push(['values', r, v.length]),
    clearContent: () => ops.push(['clear', r, nr]), setFontWeight: () => {}, setDataValidation: () => {},
  });
  const sheet = { getRange: range, getLastRow: () => 50, setFrozenRows: () => {}, clearContents: () => ops.push(['clearAll']) };
  global.SpreadsheetApp = { newDataValidation: () => ({ requireValueInList: function () { return this; }, setAllowInvalid: function () { return this; }, build: () => ({}) }) };
  writeMappingRows_({ getSheetByName: () => sheet }, [MAPPING_HEADERS.map(() => 'x'), MAPPING_HEADERS.map(() => 'y')]);
  assert.ok(!ops.some((o) => o[0] === 'clearAll'));
  const vi = ops.findIndex((o) => o[0] === 'values');
  const ci = ops.findIndex((o) => o[0] === 'clear');
  assert.ok(vi >= 0 && ci > vi, JSON.stringify(ops));
  assert.deepStrictEqual(ops[ci], ['clear', 4, 47]);
});

// 라이브에서 5천 행 쓰기가 반영되기 전에 탭을 추가하다 '스프레드시트 서비스 타임아웃'이 세 번 났다(2026-09-30).
// 진단에서는 단계마다 flush하면 모두 3초 안이었다 → 탭을 먼저 만들고, 탭마다 바로 반영한다. deal list는 참고용이라 실패해도 멈추지 않는다.
function fakeSpreadsheet(opts) {
  opts = opts || {};
  const ops = [];
  const sheets = {};
  const makeSheet = (name) => ({
    name: name,
    getRange: () => ({
      setNumberFormat: () => ops.push(['format', name]),
      setValues: () => { if (opts.failValues === name) throw new Error('타임아웃'); ops.push(['values', name]); },
      clearContent: () => ops.push(['clear', name]),
      setFontWeight: () => {}, setDataValidation: () => {},
    }),
    getLastRow: () => 0, getMaxRows: () => 1000, setFrozenRows: () => {}, setColumnWidths: () => {},
  });
  (opts.existing || []).forEach((n) => { sheets[n] = makeSheet(n); });
  global.SpreadsheetApp = {
    flush: () => ops.push(['flush']),
    newDataValidation: () => ({ requireValueInList: function () { return this; }, setAllowInvalid: function () { return this; }, build: () => ({}) }),
  };
  global.Utilities = { formatDate: () => '20261001_090000', sleep: () => {} };
  global.Session = { getScriptTimeZone: () => 'Asia/Seoul' };
  const ss = {
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => { ops.push(['insert', n]); sheets[n] = makeSheet(n); return sheets[n]; },
  };
  return { ss: ss, ops: ops };
}

test('시트 쓰기: 탭을 먼저 다 만들고, 탭마다 바로 반영한다', () => {
  const f = fakeSpreadsheet({ existing: ['deal list'] });
  const out = writeOutputs_(f.ss, [['m']], [OUTPUT_HEADERS], [['d']]);
  assert.strictEqual(out.cleanName, 'clean_20261001_090000');
  assert.strictEqual(out.warn, '');
  const firstValues = f.ops.findIndex((o) => o[0] === 'values');
  const lastInsert = f.ops.map((o) => o[0]).lastIndexOf('insert');
  assert.ok(lastInsert < firstValues, JSON.stringify(f.ops));
  // 각 탭 쓰기 뒤에는 다음 탭 쓰기 전에 flush가 있다
  const seq = f.ops.filter((o) => o[0] === 'values' || o[0] === 'flush').map((o) => o[0] === 'flush' ? 'F' : o[1]);
  assert.deepStrictEqual(seq, ['F', 'shop_id 매핑', 'F', 'clean_20261001_090000', 'F', 'deal list', 'F']);
});

test('deal list 쓰기가 실패해도 매핑·clean은 남고 경고만 돌려준다', () => {
  const f = fakeSpreadsheet({ failValues: 'deal list' });
  const out = writeOutputs_(f.ss, [['m']], [OUTPUT_HEADERS], [['d']]);
  assert.match(out.warn, /deal list 갱신 실패: 타임아웃/);
  assert.ok(f.ops.some((o) => o[0] === 'values' && o[1] === 'shop_id 매핑'));
  assert.ok(f.ops.some((o) => o[0] === 'values' && o[1] === 'clean_20261001_090000'));
});

// 라이브에서 Sales 딜 5,476건을 사용자 필드 94개째 들고 있으니(응답 32.7MB) 메모리 압박으로 시트 쓰기가 타임아웃됐다(2026-09-30).
// 참조를 놓자 같은 쓰기가 1~2초 → 딜은 쓰는 사용자 필드 3개만 요청하고, 받자마자 쓰는 값만 남긴다.
test('딜 조회는 사용자 필드 3개만 요청하고 필요한 값만 남긴다', () => {
  const big = {}; for (let i = 0; i < 90; i++) big['k' + i] = 'v'.repeat(100);
  const deal = (id, shop) => ({ id: id, title: 't' + id, owner_id: 24324011, stage_id: 71, label_ids: [299], person_id: id + 100, org_id: null,
    status: 'open', value: 0, add_time: 'x', custom_fields: Object.assign({ [PD_FIELD_SHOP_ID]: shop, [PD_FIELD_URL]: 'u.com', [PD_FIELD_MALL_NAME]: 'm' }, big) });
  const pd = fakePipedrive({ listDeals: [deal(1, '101'), deal(2, '채널톡'), deal(3, '')] });
  const r = fetchPipedrive_('t');
  const listCalls = pd.calls.filter((c) => c.route === 'GET /api/v2/deals');
  assert.strictEqual(listCalls.length, 2);
  listCalls.forEach((c) => {
    const cf = decodeURIComponent(/custom_fields=([^&]*)/.exec(c.path)[1]).split(',').sort();
    assert.deepStrictEqual(cf, [PD_FIELD_SHOP_ID, PD_FIELD_URL, PD_FIELD_MALL_NAME].sort());
  });
  assert.deepStrictEqual(Object.keys(r.deals[0]).sort(), ['custom_fields', 'id', 'label_ids', 'org_id', 'owner_id', 'person_id', 'stage_id', 'title']);
  assert.deepStrictEqual(Object.keys(r.deals[0].custom_fields).sort(), [PD_FIELD_SHOP_ID, PD_FIELD_URL, PD_FIELD_MALL_NAME].sort());
  assert.deepStrictEqual([...r.dealShopIds], ['101']);
  assert.deepStrictEqual(r.unmappedDeals.map((d) => [d.id, d.raw, d.keys.email]), [[2, '채널톡', ['e102@x.com']], [3, '', ['e103@x.com']]]);
});
