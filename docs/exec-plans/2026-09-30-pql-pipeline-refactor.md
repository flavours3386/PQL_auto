# PQL 파이프라인 재구축 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 최신 구독 CSV(45MB)를 6분 안에 직접 읽어 Sales 딜 없는 업셀·푸시 타겟만 골라 clean 탭과 Pipedrive 업로드 xlsx를 만들고, shop_id가 비거나 텍스트인 Sales 딜은 CSV와 대조해 shop_id를 자동 반영한다.

**Architecture:** Apps Script 바운드 프로젝트를 `src/` 4개 파일로 나눈다. `Config.js`(설정·타겟 규칙), `Core.js`(Apps Script 서비스를 쓰지 않는 순수 로직, node 테스트 대상), `Io.js`(Drive·Pipedrive·Sheets), `Main.js`(메뉴·실행 흐름). Pipedrive를 먼저 조회해 매핑 안 된 딜 키만 작은 역색인으로 만든 뒤, CSV를 20MB Range 조각으로 받아 스트리밍 파서로 한 번 훑으며 클렌징·타겟·역매핑을 같이 처리한다.

**Tech Stack:** Google Apps Script(V8), UrlFetchApp(Drive v3·Pipedrive v1/v2 REST), clasp 3.4, node 24 내장 `node:test`

**Spec:** `docs/design-docs/2026-09-30-pql-pipeline-refactor-design.md`

## Global Constraints

- Advanced Service 사용 금지 (`appsscript.json`의 `enabledAdvancedServices: []` 유지). Drive·Pipedrive 모두 UrlFetch
- 원천 폴더 `05. PQL`(`1PjCz9YxLLqGLYOZLffPO97tk7UKEGEaF`)은 읽기만 한다. 파일 생성·수정 금지
- UrlFetch 응답 50MB/회 한도 → 다운로드 조각 `DOWNLOAD_CHUNK_BYTES = 20 * 1024 * 1024`
- 실행 한도 6분/회
- Pipedrive 인증은 헤더 `x-api-token`, 토큰은 Script Properties `PIPEDRIVE_API_TOKEN`에만 둔다. 코드·로그에 토큰 금지
- Pipedrive 필드 키: shop_id `9d4ea1fcf0bde157910e96a2e0354e76c220e6c8`, URL `5a7464db665cc9fb3cebc7530c536f39205768ca`, 쇼핑몰명 `4cf3a83ff7316bb926dbf2c7f9c7b92308bad7bc`. Sales 파이프라인 id `9`
- 숫자 shop_id가 이미 있는 딜은 절대 덮어쓰지 않는다
- 주석은 한국어, 식별자는 영어. 커밋 메시지 `[카테고리] 한국어 요약` + `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`
- 고객 데이터(CSV·Pipedrive 응답)는 저장소에 넣지 않는다. 실데이터 검증은 scratchpad에서만: `SCRATCH=/private/tmp/claude-501/-Users-sales-Desktop-02-PJT-Work-PQL-auto/9707c6f1-bd9b-4a4e-8c63-d0b7f56bfa81/scratchpad`

## Review Focus

1. 따옴표 필드 안의 줄바꿈이 다운로드 조각 경계에 걸림 → 행이 쪼개지지 않아야 한다 (Task 1 테스트)
2. 한글·이모지 멀티바이트 글자가 Range 바이트 경계에 걸림 → 글자가 깨지거나 바이트가 밀리지 않아야 한다 (Task 1 테스트)
3. 상류 CSV 헤더 이름이 바뀜(2026-08 `회사명2`→`회사명` 전례) → 조용히 필터를 건너뛰지 말고 없는 열 이름을 모두 띄우고 멈춰야 한다 (Task 5 테스트)
4. 사람이 `거절`한 (딜, shop) 쌍 → 다시 후보로 올라오거나 자동 반영되면 안 된다 (Task 3·4 테스트)
5. 원천 이상으로 높은 확신 매칭이 100건을 넘음 → Pipedrive에 한 건도 쓰지 않고 전부 대기로 돌려야 한다 (Task 4 테스트)

---

## File Structure

| 파일 | 책임 |
|---|---|
| `.clasp.json` | scriptId, `rootDir: src` |
| `src/appsscript.json` | 라이브와 동일한 매니페스트 |
| `src/Config.js` | 설정 상수, 상태값 집합, 타겟 규칙, 출력 열 |
| `src/Core.js` | CSV 스트리밍 파서·분할 다운로드 경계, 정규화, 클렌징·타겟·라벨, 행 구성, 역매핑, 매핑 탭 상태, 빌더, 요약 문구 |
| `src/Io.js` | Pipedrive 요청·조회·쓰기, Drive CSV 찾기·스트리밍·xlsx 생성, 시트 탭 쓰기 |
| `src/Main.js` | 메뉴, `runPql`, `applyApprovedMappings`, 요약 창 |
| `test/gas.js` | `Config.js`+`Core.js`를 Apps Script처럼 한 전역 스코프에 올리는 테스트 로더 |
| `test/*.test.js` | Core 단위·통합 테스트 |
| `PQL.md` | 사용법만 (코드블록 제거) |

---

### Task 1: 저장소 구조 + CSV 스트리밍 코어

**Files:**
- Create: `.clasp.json`, `src/appsscript.json`, `src/Config.js`, `src/Core.js`, `test/gas.js`, `test/csv.test.js`

**Interfaces:**
- Produces: `createCsvParser_(onRow) → { feed(text), end() }`, `utf8ByteLength_(str) → number`, `streamChunks_(size, chunkBytes, fetchRange(start, end) → string, onText(text))`

- [ ] **Step 1: clasp 설정과 매니페스트 작성**

`.clasp.json`:
```json
{ "scriptId": "1bDdQ0oWl-rtXv7z1YNMfez8LhkHwhja0J9hvzo7KbVdCk19ofLYh1_dP", "rootDir": "src" }
```

`src/appsscript.json` (라이브에서 pull한 값 그대로):
```json
{
  "timeZone": "Asia/Seoul",
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "dependencies": {
    "enabledAdvancedServices": []
  }
}
```

- [ ] **Step 2: `src/Config.js` 작성 (전체)**

```javascript
/***************************************
 * 설정값 — 운영 중 바꿀 값은 이 파일에만 둔다
 ***************************************/
const SOURCE_FOLDER_ID = '1PjCz9YxLLqGLYOZLffPO97tk7UKEGEaF'; // Drive '05. PQL' (crema·alphareview-ref도 읽는 공유 폴더 — 읽기만 한다)
const SOURCE_NAME_PREFIX = 'all_subscription_';
const DOWNLOAD_CHUNK_BYTES = 20 * 1024 * 1024; // UrlFetch 응답 한도 50MB/회

const MIN_ORDERS = 100;
const PUSH_MIN_ORDERS = 500;

const SALES_PIPELINE_ID = 9;
const PD_FIELD_SHOP_ID = '9d4ea1fcf0bde157910e96a2e0354e76c220e6c8';
const PD_FIELD_URL = '5a7464db665cc9fb3cebc7530c536f39205768ca';
const PD_FIELD_MALL_NAME = '4cf3a83ff7316bb926dbf2c7f9c7b92308bad7bc';
const PD_TOKEN_PROPERTY = 'PIPEDRIVE_API_TOKEN';

const AUTO_APPLY = true; // 높은 확신 역매핑을 Pipedrive에 자동 반영
const AUTO_APPLY_MAX = 100; // 한 실행에서 이보다 많으면 자동 반영을 멈추고 전부 대기로

const DEAL_OWNER = '한서연';
const DEAL_STAGE = '컨택전';

const TAB_DEAL_LIST = 'deal list';
const TAB_MAPPING = 'shop_id 매핑';
const CLEAN_TAB_PREFIX = 'clean_';

// 상태값은 공백을 뺀 형태로 적는다 (비교 전에 원천 값의 공백도 뺀다)
const REVIEW_EXCLUDE = new Set(['제거중', '해지완료', '서비스중단']);
const SITE_EXCLUDE = new Set(['구독종료', '해지완료', '계정활성화']);
const NOT_USED = new Set(['구독없음', '서비스중단', '프로덕트온보딩중', '']);
const PAID_LIVE = new Set(['라이브(과금중)', '라이브(계약구독중)', '라이브(체험중)']);

// 타겟 규칙: 하나라도 맞으면 PQL에 남는다. 새 타겟은 항목 하나를 추가한다.
const TARGETS = [
  { name: '업셀', test: (r) => r.platform === 'cafe24' && !isLive_(r.upsell) && r.upsell !== '제거중' },
  { name: '푸시', test: (r) => r.platform === 'cafe24' && r.orders >= PUSH_MIN_ORDERS && !PAID_LIVE.has(r.push) && r.push !== '제거중' },
];

const REQUIRED_COLUMNS = ['shop_id', '플랫폼', '최근 30일 플랫폼 주문수', '알파리뷰 상태', '알파업셀 상태', '알파푸시 상태', '사이트 상태', '담당자명', '담당자전화번호'];

const OUTPUT_HEADERS = [
  'shop_name', 'shop_id', 'mall_id', '플랫폼', '최근 30일 플랫폼 주문수(API)', '타겟', '서비스 라벨', '딜 의심',
  '회사명', '담당자명', '쇼핑몰명', '담당자전화번호', '담당자이메일', '대표도메인', '주소',
  'shop_no', '플랜', '사이트 상태', '알파리뷰 상태', '알파업셀 상태', '알파푸시 상태',
  '최근 30일 플랫폼 주문수', '최근 30일 전체 주문수', '설치시점 플랫폼 주문수(API)',
  '최근 30일 UV(방문자수)', '최근 30일 PV(페이지뷰)', '임직원 수', '이메일', '사업자', '고객센터', '전화번호', '담당자직책', '결제담당이메일',
];
const UPLOAD_HEADERS = ['거래 제목', 'shop_id', '상점아이디', '호스팅사', '월 주문 수', '거래 소유자', '단계 (파이프라인)', '거래 라벨', '조직 이름', '이름', '쇼핑몰명', '전화', '이메일', 'URL', '주소'];
const MAPPING_HEADERS = ['딜 ID', '딜 이름', '원래 shop_id', '후보 shop_id', '후보 shop_name', '일치 키', '신뢰도', '판정', '상태', '기록일'];
```

- [ ] **Step 3: 테스트 로더 `test/gas.js` 작성**

```javascript
// Apps Script처럼 src 파일들을 하나의 전역 스코프에 올린다 (테스트 전용)
const fs = require('fs');
const path = require('path');
const vm = require('vm');

for (const f of ['Config.js', 'Core.js']) {
  vm.runInThisContext(fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8'), { filename: f });
}
```

- [ ] **Step 4: 실패하는 테스트 `test/csv.test.js` 작성**

```javascript
require('./gas');
const test = require('node:test');
const assert = require('node:assert');

function parseAll(chunks) {
  const rows = [];
  const p = createCsvParser_((r) => rows.push(r));
  chunks.forEach((c) => p.feed(c));
  p.end();
  return rows;
}

test('따옴표 안 쉼표·줄바꿈·이스케이프, BOM, CRLF', () => {
  const csv = '﻿a,b,c\r\n1,"x, y","he said ""hi"""\r\n2,"line1\nline2",\r\n';
  assert.deepStrictEqual(parseAll([csv]), [['a', 'b', 'c'], ['1', 'x, y', 'he said "hi"'], ['2', 'line1\nline2', '']]);
});

test('조각 경계가 따옴표 필드·이스케이프 한가운데여도 결과가 같다', () => {
  const csv = 'a,b\n1,"he said ""hi""\nbye"\n2,z';
  const whole = parseAll([csv]);
  for (let i = 1; i < csv.length; i++) {
    assert.deepStrictEqual(parseAll([csv.slice(0, i), csv.slice(i)]), whole, 'cut at ' + i);
  }
});

test('마지막 줄에 줄바꿈이 없어도 행이 나온다', () => {
  assert.deepStrictEqual(parseAll(['a,b\n1,2']), [['a', 'b'], ['1', '2']]);
});

test('utf8ByteLength_는 Buffer 바이트 수와 같다', () => {
  for (const s of ['abc', '한글', 'é', '😀', '﻿shop_id,회사명\n']) {
    assert.strictEqual(utf8ByteLength_(s), Buffer.byteLength(s, 'utf8'), s);
  }
});

test('streamChunks_: 멀티바이트 글자가 바이트 경계에 걸려도 원문이 복원된다', () => {
  const csv = '﻿shop_id,회사명,주소\n1,"주식회사 가나","서울, 강남"\n2,다라😀,"부산\n해운대"\n3,마바,대구\n';
  const buf = Buffer.from(csv, 'utf8');
  // Apps Script getContentText처럼 BOM을 지우지 않고, 잘린 끝 글자는 U+FFFD로 둔다
  const dec = new TextDecoder('utf-8', { ignoreBOM: true });
  const minChunk = Math.max(...csv.split('\n').map((l) => Buffer.byteLength(l + '\n')));
  for (let chunk = minChunk; chunk <= buf.length; chunk++) {
    let out = '';
    streamChunks_(buf.length, chunk, (s, e) => dec.decode(buf.subarray(s, e + 1)), (t) => { out += t; });
    assert.strictEqual(out, csv, 'chunk ' + chunk);
  }
});

test('한 행이 분할 크기보다 크면 멈춘다', () => {
  const buf = Buffer.from('aaaaaaaaaa\nb\n');
  assert.throws(() => streamChunks_(buf.length, 4, (s, e) => buf.subarray(s, e + 1).toString(), () => {}), /분할 크기/);
});
```

- [ ] **Step 5: 실패 확인**

Run: `touch src/Core.js && node --test test/`
Expected: FAIL — `createCsvParser_ is not defined`

- [ ] **Step 6: `src/Core.js`에 CSV 섹션 작성**

```javascript
/***************************************
 * 순수 로직 — Apps Script 서비스를 쓰지 않는다 (node 테스트 대상)
 ***************************************/

/* ---------- CSV ---------- */

// 한 행씩 onRow(fields)로 넘기는 스트리밍 파서. feed()를 여러 번 불러도 따옴표·행 상태가 이어진다.
function createCsvParser_(onRow) {
  let field = '';
  let row = [];
  let inQuotes = false;
  let justClosed = false; // 직전 문자가 닫는 따옴표였는지 ("" 이스케이프 판별)
  let started = false;
  return {
    feed: function (text) {
      let i = 0;
      if (!started) {
        started = true;
        if (text.charCodeAt(0) === 0xfeff) i = 1; // BOM
      }
      for (; i < text.length; i++) {
        const ch = text[i];
        if (inQuotes) {
          if (ch === '"') {
            inQuotes = false;
            justClosed = true;
          } else field += ch;
          continue;
        }
        if (ch === '"') {
          if (justClosed) {
            field += '"';
            inQuotes = true;
          } else if (field === '') inQuotes = true;
          else field += ch;
          justClosed = false;
          continue;
        }
        justClosed = false;
        if (ch === ',') {
          row.push(field);
          field = '';
        } else if (ch === '\n') {
          row.push(field);
          onRow(row);
          row = [];
          field = '';
        } else if (ch !== '\r') field += ch;
      }
    },
    end: function () {
      if (field !== '' || row.length) {
        row.push(field);
        onRow(row);
      }
      row = [];
      field = '';
    },
  };
}

// UTF-8 바이트 수 (분할 다운로드에서 다음 Range 시작점 계산용)
function utf8ByteLength_(s) {
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x80) n += 1;
    else if (c < 0x800) n += 2;
    else if (c >= 0xd800 && c <= 0xdbff) {
      n += 4;
      i++;
    } else n += 3;
  }
  return n;
}

// 파일을 Range 요청으로 나눠 받아 onText로 넘긴다. fetchRange(start, end)는 그 바이트 구간을 UTF-8로 디코드한 문자열.
// 조각마다 마지막 '\n'에서 자르고 나머지는 다음 요청에서 다시 받는다. '\n'(0x0A)은 UTF-8 멀티바이트 안에 나오지 않아 자른 곳이 항상 글자 경계다.
// ponytail: 원천에 깨진 UTF-8 바이트가 있으면 U+FFFD 바이트 수가 달라 오프셋이 밀린다. 원천은 정상 UTF-8(BOM) CSV라 검사하지 않는다.
function streamChunks_(size, chunkBytes, fetchRange, onText) {
  let offset = 0;
  while (offset < size) {
    const end = Math.min(offset + chunkBytes, size) - 1;
    const text = fetchRange(offset, end);
    if (end === size - 1) {
      onText(text);
      return;
    }
    const cut = text.lastIndexOf('\n');
    if (cut < 0) throw new Error('CSV 한 행이 분할 크기(' + chunkBytes + '바이트)보다 큽니다');
    const head = text.slice(0, cut + 1);
    onText(head);
    offset += utf8ByteLength_(head);
  }
}
```

- [ ] **Step 7: 통과 확인**

Run: `node --test test/`
Expected: PASS 6/6

- [ ] **Step 8: 커밋**

```bash
git add .clasp.json src/ test/
git commit -m "[refactor] Apps Script 코드를 src/로 옮기고 CSV 스트리밍 파서 추가

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: 레코드·클렌징·타겟·라벨·행 구성

**Files:**
- Modify: `src/Core.js` (끝에 추가)
- Test: `test/rules.test.js`

**Interfaces:**
- Consumes: Config 상수 (`MIN_ORDERS`, `REVIEW_EXCLUDE`, `SITE_EXCLUDE`, `NOT_USED`, `TARGETS`, `OUTPUT_HEADERS`, `DEAL_OWNER`, `DEAL_STAGE`)
- Produces:
  - `normStatus_(s) → string`, `isLive_(s) → boolean`, `normalizePhone_(raw) → { display, digits }`, `toNumberOr_(s) → number|string`
  - `headerIndex_(headerRow) → { last: {name: idx}, all: {name: [idx]} }`
  - `toRecord_(row, hi) → { get(name), shopId, platform, ordersRaw, orders, review, upsell, push, site, manager, phone, phoneDigits, address }`
  - `cleanseReason_(r, dealShopIds: Set<string>) → '' | 'orders' | 'review' | 'site' | 'pro' | 'phone' | 'deal'`
  - `matchTargets_(r) → string[]`, `serviceLabel_(r) → string`
  - `cleanRow_(r, targets: string[], label, suspectDealIds: string[]) → string[33]`, `uploadRow_(r, label) → Array(15)`

- [ ] **Step 1: 실패하는 테스트 `test/rules.test.js` 작성**

```javascript
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
```

- [ ] **Step 2: 실패 확인**

Run: `node --test test/rules.test.js`
Expected: FAIL — `headerIndex_ is not defined`

- [ ] **Step 3: `src/Core.js` 끝에 추가**

```javascript
/* ---------- 정규화 ---------- */

function normStatus_(s) {
  return String(s == null ? '' : s).replace(/\s/g, '');
}

function isLive_(s) {
  return s.indexOf('라이브') === 0;
}

function normalizePhone_(raw) {
  let display = String(raw == null ? '' : raw).trim();
  if (display === '') return { display: '', digits: '' };
  let digits = display.replace(/\D/g, '');
  if (digits.length === 10 && digits.indexOf('10') === 0) digits = '0' + digits;
  if (digits.length === 11 && digits.indexOf('010') === 0) display = '010-' + digits.slice(3, 7) + '-' + digits.slice(7);
  return { display: display, digits: digits };
}

function toNumberOr_(s) {
  const n = Number(s);
  return s !== '' && isFinite(n) ? n : s;
}

/* ---------- 레코드 ---------- */

// 같은 이름 열(회사명 등)이 둘이면 last는 마지막 열, all은 전부
function headerIndex_(header) {
  const last = {};
  const all = {};
  header.forEach(function (n, i) {
    const k = String(n).trim();
    if (!k) return;
    last[k] = i;
    (all[k] = all[k] || []).push(i);
  });
  return { last: last, all: all };
}

function toRecord_(row, hi) {
  const get = function (name) {
    const i = hi.last[name];
    return i === undefined || row[i] == null ? '' : String(row[i]).trim();
  };
  const phone = normalizePhone_(get('담당자전화번호'));
  const ordersRaw = get('최근 30일 플랫폼 주문수');
  return {
    get: get,
    shopId: get('shop_id'),
    platform: get('플랫폼').toLowerCase(),
    ordersRaw: ordersRaw,
    orders: ordersRaw === '' ? NaN : Number(ordersRaw),
    review: normStatus_(get('알파리뷰 상태')),
    upsell: normStatus_(get('알파업셀 상태')),
    push: normStatus_(get('알파푸시 상태')),
    site: normStatus_(get('사이트 상태')),
    manager: get('담당자명'),
    phone: phone.display,
    phoneDigits: phone.digits,
    address: (get('주소1') + ' ' + get('주소2')).trim(),
  };
}

/* ---------- 규칙 ---------- */

// 공통 클렌징 ①~⑤ + ⑥(숫자 shop_id 딜). 제외 사유 코드, 통과면 ''.
function cleanseReason_(r, dealShopIds) {
  if (!(r.orders >= MIN_ORDERS)) return 'orders';
  if (REVIEW_EXCLUDE.has(r.review)) return 'review';
  if (SITE_EXCLUDE.has(r.site)) return 'site';
  if (r.manager === '프로' && !/^010/.test(r.phoneDigits)) return 'pro';
  if (r.phone === '') return 'phone';
  if (dealShopIds.has(r.shopId)) return 'deal';
  return '';
}

function matchTargets_(r) {
  return TARGETS.filter(function (t) { return t.test(r); }).map(function (t) { return t.name; });
}

// 지금 라이브인 제품(무료 포함). SDR은 라벨에 없는 제품을 제안한다.
function serviceLabel_(r) {
  if (NOT_USED.has(r.review) && NOT_USED.has(r.upsell) && NOT_USED.has(r.push)) return 'null';
  const live = [];
  if (isLive_(r.review)) live.push('알파리뷰');
  if (isLive_(r.upsell)) live.push('알파업셀');
  if (isLive_(r.push)) live.push('알파푸시');
  return live.join(', ');
}

/* ---------- 출력 행 ---------- */

function cleanRow_(r, targets, label, suspectDealIds) {
  return OUTPUT_HEADERS.map(function (h) {
    if (h === '타겟') return targets.join(', ');
    if (h === '서비스 라벨') return label;
    if (h === '딜 의심') return suspectDealIds.join(', ');
    if (h === '담당자전화번호') return r.phone;
    if (h === '주소') return r.address;
    return r.get(h);
  });
}

// Pipedrive 가져오기 양식 (pipedrive_up(0901).xlsx와 같은 15열)
function uploadRow_(r, label) {
  return [
    r.get('shop_name'), toNumberOr_(r.shopId), r.get('mall_id'), r.get('플랫폼'), toNumberOr_(r.ordersRaw),
    DEAL_OWNER, DEAL_STAGE, label, r.get('회사명'), r.get('담당자명'), r.get('쇼핑몰명'),
    r.phone, r.get('담당자이메일'), r.get('대표도메인'), r.address,
  ];
}
```

- [ ] **Step 4: 통과 확인**

Run: `node --test test/`
Expected: PASS (csv 6 + rules 9)

- [ ] **Step 5: 커밋**

```bash
git add src/Core.js test/rules.test.js
git commit -m "[feat] PQL 공통 클렌징·업셀/푸시 타겟·서비스 라벨 규칙 추가

리뷰 '서비스중단' 띄어쓰기 오타로 새던 필터 수정, 라벨에 알파업셀 표기.

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: shop_id 역매핑 판정

**Files:**
- Modify: `src/Core.js` (끝에 추가)
- Test: `test/match.test.js`

**Interfaces:**
- Consumes: `normalizePhone_`, `PD_FIELD_SHOP_ID`, `PD_FIELD_URL`, `PD_FIELD_MALL_NAME`
- Produces:
  - `emailKey_`, `phoneKey_`, `nameKey_`, `domainKey_`, `dealUrlKey_` (각각 `raw → string`, 키가 안 되면 `''`)
  - `rawShopId_(deal) → string` (custom_fields의 shop_id 원래 값, trim)
  - `dealMatchKeys_(deal, person, org) → { email: [], phone: [], name: [], url: [] }` (deal·person·org는 Pipedrive v2 객체)
  - `rowMatchKeys_(row, hi) → 같은 형태`
  - `classifyMatch_(hitsByKey: {key: Set<shopId>}) → { tier: 'high'|'review'|'none', candidates: [{ shopId, keys: string[] }] }`
  - `createDealMatcher_(unmappedDeals: [{ id, title, raw, keys }]) → { onRow(row, hi, shopId, shopName), results(rejectedPairs: Set<'dealId:shopId'>) → [{ deal, tier, candidates: [{ shopId, keys, shopName }] }] }`

- [ ] **Step 1: 실패하는 테스트 `test/match.test.js` 작성**

```javascript
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
```

- [ ] **Step 2: 실패 확인**

Run: `node --test test/match.test.js`
Expected: FAIL — `nameKey_ is not defined`

- [ ] **Step 3: `src/Core.js` 끝에 추가**

```javascript
/* ---------- shop_id 역매핑 ---------- */

const MATCH_KEYS = ['email', 'phone', 'name', 'url'];

function emailKey_(raw) {
  const s = String(raw == null ? '' : raw).trim().toLowerCase();
  return s.indexOf('@') > 0 ? s : '';
}

function phoneKey_(raw) {
  const d = normalizePhone_(raw).digits;
  return d.length >= 9 ? d : '';
}

function nameKey_(raw) {
  const s = String(raw == null ? '' : raw).replace(/\(주\)|주식회사|㈜/g, '').replace(/[^0-9A-Za-z가-힣]/g, '').toLowerCase();
  return s.length >= 2 ? s : '';
}

function domainKey_(raw) {
  let s = String(raw == null ? '' : raw).trim().toLowerCase().replace(/^[a-z]+:\/\//, '');
  s = s.split('/')[0].split('?')[0].split(':')[0].replace(/^(www|m)\./, '');
  return s.indexOf('.') > 0 ? s : '';
}

// 딜 쪽 URL 값: 도메인이면 도메인, 영숫자만이면 mall_id로 보고 cafe24 기본 도메인
function dealUrlKey_(raw) {
  const d = domainKey_(raw);
  if (d) return d;
  const s = String(raw == null ? '' : raw).trim().toLowerCase();
  return /^[a-z0-9]+$/.test(s) ? s + '.cafe24.com' : '';
}

function rawShopId_(deal) {
  const v = ((deal && deal.custom_fields) || {})[PD_FIELD_SHOP_ID];
  return v == null ? '' : String(v).trim();
}

function pushKey_(arr, v) {
  if (v && arr.indexOf(v) < 0) arr.push(v);
}

function dealMatchKeys_(deal, person, org) {
  const cf = deal.custom_fields || {};
  const keys = { email: [], phone: [], name: [], url: [] };
  ((person && person.emails) || []).forEach(function (e) { pushKey_(keys.email, emailKey_(e.value)); });
  ((person && person.phones) || []).forEach(function (p) { pushKey_(keys.phone, phoneKey_(p.value)); });
  [deal.title, org && org.name, cf[PD_FIELD_MALL_NAME]].forEach(function (v) { pushKey_(keys.name, nameKey_(v)); });
  pushKey_(keys.url, dealUrlKey_(cf[PD_FIELD_URL]));
  pushKey_(keys.url, dealUrlKey_(rawShopId_(deal)));
  return keys;
}

function rowMatchKeys_(row, hi) {
  const vals = function (name) { return (hi.all[name] || []).map(function (i) { return row[i]; }); };
  const keys = { email: [], phone: [], name: [], url: [] };
  const add = function (key, cols, fn) {
    cols.forEach(function (c) { vals(c).forEach(function (v) { pushKey_(keys[key], fn(v)); }); });
  };
  add('email', ['담당자이메일', '이메일', '결제담당이메일'], emailKey_);
  add('phone', ['담당자전화번호', '전화번호', '고객센터'], phoneKey_);
  add('name', ['shop_name', '쇼핑몰명', '회사명'], nameKey_);
  add('url', ['대표도메인', '기본제공 도메인'], domainKey_);
  add('url', ['mall_id'], function (v) {
    const m = String(v == null ? '' : v).trim().toLowerCase();
    return m ? m + '.cafe24.com' : '';
  });
  return keys;
}

// high: 걸린 키가 2개 이상이고 교집합이 shop 1개 / review: 그 밖에 후보가 있음 / none
function classifyMatch_(hitsByKey) {
  const keys = MATCH_KEYS.filter(function (k) { return hitsByKey[k] && hitsByKey[k].size; });
  if (!keys.length) return { tier: 'none', candidates: [] };
  const keysOf = function (id) { return keys.filter(function (k) { return hitsByKey[k].has(id); }); };
  let inter = Array.from(hitsByKey[keys[0]]);
  keys.slice(1).forEach(function (k) { inter = inter.filter(function (id) { return hitsByKey[k].has(id); }); });
  if (keys.length >= 2 && inter.length === 1) return { tier: 'high', candidates: [{ shopId: inter[0], keys: keys }] };
  const union = new Set();
  keys.forEach(function (k) { hitsByKey[k].forEach(function (id) { union.add(id); }); });
  return { tier: 'review', candidates: Array.from(union).sort().map(function (id) { return { shopId: id, keys: keysOf(id) }; }) };
}

// 매핑 안 된 딜 키를 역색인으로 만들고, CSV 행을 흘려 보내며 후보 shop을 모은다 (전체 CSV 색인은 만들지 않는다)
function createDealMatcher_(unmappedDeals) {
  const index = {};
  unmappedDeals.forEach(function (d, pos) {
    MATCH_KEYS.forEach(function (k) {
      d.keys[k].forEach(function (v) { (index[k + ':' + v] = index[k + ':' + v] || []).push(pos); });
    });
  });
  const hits = unmappedDeals.map(function () { return {}; });
  const shopNames = {};
  return {
    onRow: function (row, hi, shopId, shopName) {
      const keys = rowMatchKeys_(row, hi);
      MATCH_KEYS.forEach(function (k) {
        keys[k].forEach(function (v) {
          (index[k + ':' + v] || []).forEach(function (pos) {
            (hits[pos][k] = hits[pos][k] || new Set()).add(shopId);
            shopNames[shopId] = shopName;
          });
        });
      });
    },
    results: function (rejectedPairs) {
      return unmappedDeals.map(function (d, pos) {
        const h = {};
        Object.keys(hits[pos]).forEach(function (k) {
          const s = new Set();
          hits[pos][k].forEach(function (id) { if (!rejectedPairs.has(d.id + ':' + id)) s.add(id); });
          if (s.size) h[k] = s;
        });
        const c = classifyMatch_(h);
        return {
          deal: d,
          tier: c.tier,
          candidates: c.candidates.map(function (x) { return { shopId: x.shopId, keys: x.keys, shopName: shopNames[x.shopId] || '' }; }),
        };
      });
    },
  };
}
```

- [ ] **Step 4: 통과 확인**

Run: `node --test test/`
Expected: PASS (csv 6 + rules 9 + match 6)

- [ ] **Step 5: 커밋**

```bash
git add src/Core.js test/match.test.js
git commit -m "[feat] Sales 딜 shop_id 역매핑 판정 추가 (이메일·전화·이름·URL)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: 딜 분리·deal list·매핑 탭 상태·요약 문구

**Files:**
- Modify: `src/Core.js` (끝에 추가)
- Test: `test/mapping.test.js`

**Interfaces:**
- Consumes: `rawShopId_`, `toNumberOr_`, `MAPPING_HEADERS` 순서(0 딜 ID, 1 딜 이름, 2 원래 shop_id, 3 후보 shop_id, 4 후보 shop_name, 5 일치 키, 6 신뢰도, 7 판정, 8 상태, 9 기록일)
- Produces:
  - `splitDeals_(deals) → { shopIds: Set<string>, unmapped: deal[] }`
  - `dealListRows_(deals, users: {id: name}, stages: {id: name}, labels: {id: label}, overrides: {dealId: shopId}) → rows(헤더 포함, 5열)`
  - `readMappingState_(rows) → { approved: row[], pending: row[], rejectedPairs: Set<string>, keep: row[] }`
  - `splitApprovals_(approved) → { apply: row[], duplicate: row[] }`
  - `planMappings_(matches, autoApply, autoMax) → { auto: match[], pendingHigh: match[], review: match[], overLimit: boolean }`
  - `mappingRow_(match, candidate, confidence, verdict, status, today) → row(10열)`
  - `mappingNote_(source, raw, shopId, keys, today) → string`
  - `uploadFileName_(csvName, fallbackMmdd) → string`
  - `summaryLines_(s) → string[]`

- [ ] **Step 1: 실패하는 테스트 `test/mapping.test.js` 작성**

```javascript
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
```

- [ ] **Step 2: 실패 확인**

Run: `node --test test/mapping.test.js`
Expected: FAIL — `splitDeals_ is not defined`

- [ ] **Step 3: `src/Core.js` 끝에 추가**

```javascript
/* ---------- Pipedrive 데이터 가공 ---------- */

function splitDeals_(deals) {
  const shopIds = new Set();
  const unmapped = [];
  deals.forEach(function (d) {
    const v = rawShopId_(d);
    if (/^\d+$/.test(v)) shopIds.add(v);
    else unmapped.push(d);
  });
  return { shopIds: shopIds, unmapped: unmapped };
}

function dealListRows_(deals, users, stages, labels, overrides) {
  const rows = [['거래 - 레이블', '거래 - shop_id', '거래 - 이름', '거래 - 소유자', '거래 - 단계']];
  deals.forEach(function (d) {
    const shop = overrides[d.id] !== undefined ? String(overrides[d.id]) : rawShopId_(d);
    const lab = (d.label_ids || []).slice().sort(function (a, b) { return a - b; })
      .map(function (id) { return labels[id] || String(id); }).join(', ');
    rows.push([lab, toNumberOr_(shop), d.title || '', users[d.owner_id] || '', stages[d.stage_id] || '']);
  });
  return rows;
}

/* ---------- shop_id 매핑 탭 ---------- */

// 열: 0 딜 ID, 1 딜 이름, 2 원래 shop_id, 3 후보 shop_id, 4 후보 shop_name, 5 일치 키, 6 신뢰도, 7 판정, 8 상태, 9 기록일
function readMappingState_(rows) {
  const approved = [];
  const pending = [];
  const keep = [];
  const rejectedPairs = new Set();
  rows.forEach(function (x) {
    const dealId = String(x[0]).trim();
    if (!dealId) return;
    const pair = dealId + ':' + String(x[3]).trim();
    const verdict = String(x[7]).trim();
    const status = String(x[8]).trim();
    if (status === '대기' && verdict === '승인') approved.push(x);
    else if (status === '대기' && verdict === '거절') {
      const y = x.slice();
      y[8] = '거절됨';
      keep.push(y);
      rejectedPairs.add(pair);
    } else if (status === '대기') pending.push(x);
    else {
      if (status === '거절됨') rejectedPairs.add(pair);
      keep.push(x); // 반영됨·거절됨·실패는 누적
    }
  });
  return { approved: approved, pending: pending, rejectedPairs: rejectedPairs, keep: keep };
}

function splitApprovals_(approved) {
  const byDeal = {};
  approved.forEach(function (x) {
    const id = String(x[0]).trim();
    (byDeal[id] = byDeal[id] || []).push(x);
  });
  const apply = [];
  const duplicate = [];
  Object.keys(byDeal).forEach(function (id) {
    const g = byDeal[id];
    if (g.length === 1) apply.push(g[0]);
    else g.forEach(function (x) { duplicate.push(x); });
  });
  return { apply: apply, duplicate: duplicate };
}

function planMappings_(matches, autoApply, autoMax) {
  const high = matches.filter(function (m) { return m.tier === 'high'; });
  const review = matches.filter(function (m) { return m.tier === 'review'; });
  const overLimit = autoApply && high.length > autoMax;
  const auto = autoApply && !overLimit ? high : [];
  return { auto: auto, pendingHigh: auto.length ? [] : high, review: review, overLimit: overLimit };
}

function mappingRow_(m, cand, confidence, verdict, status, today) {
  return [String(m.deal.id), m.deal.title, m.deal.raw, cand.shopId, cand.shopName, cand.keys.join(', '), confidence, verdict, status, today];
}

function mappingNote_(source, raw, shopId, keys, today) {
  return '[PQL ' + source + '매핑 ' + today + "] shop_id '" + raw + "' → " + shopId + ' (근거: ' + keys.join(', ') + ')';
}

/* ---------- 파일명·요약 ---------- */

function uploadFileName_(csvName, fallbackMmdd) {
  const m = /^all_subscription_(\d{4})\.csv$/i.exec(csvName);
  return 'pipedrive_up(' + (m ? m[1] : fallbackMmdd) + ').xlsx';
}

function summaryLines_(s) {
  const c = s.counts;
  const t = Object.keys(s.targetCounts).sort().map(function (k) { return k + ' ' + s.targetCounts[k]; }).join(' / ');
  const auto = 'shop_id 자동 반영 ' + s.autoApplied + '건' + (s.autoFailed ? ' (실패 ' + s.autoFailed + ')' : '') +
    (s.overLimit ? ' — ' + AUTO_APPLY_MAX + '건 초과라 자동 반영을 멈추고 전부 대기로 돌림' : '');
  return [
    '원천: ' + s.fileName + ' (수정 ' + s.fileUpdated + ')',
    '원천 행 ' + c.total,
    '① 주문수 100 미만·빈값 -' + c.orders,
    '② 알파리뷰 제거중·해지완료·서비스중단 -' + c.review,
    '③ 사이트 구독종료·해지완료·계정활성화 -' + c.site,
    '④ 프로 담당자 + 비핸드폰 -' + c.pro,
    '⑤ 담당자 전화 없음 -' + c.phone,
    '⑥ Sales 딜 있음(shop_id) -' + c.deal,
    '⑥ Sales 딜 있음(역매핑) -' + c.mapped,
    '타겟 해당 없음 -' + c.noTarget,
    '결과 ' + s.cleanCount + '곳 (' + t + ')',
    '업로드 xlsx ' + s.uploadCount + '행 (딜 의심 ' + (s.cleanCount - s.uploadCount) + '곳 제외)',
    '승인 매핑 반영 ' + s.approvedApplied + '건',
    auto,
    '승인 대기 ' + s.pending + "건 → '" + TAB_MAPPING + "' 탭",
    '소요 ' + s.elapsedSec + '초',
  ];
}
```

- [ ] **Step 4: 통과 확인**

Run: `node --test test/`
Expected: PASS (이전 21 + mapping 7)

- [ ] **Step 5: 커밋**

```bash
git add src/Core.js test/mapping.test.js
git commit -m "[feat] deal list·shop_id 매핑 탭 상태와 자동 반영 상한 로직 추가

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: 빌더 통합 + 실데이터 패리티·성능

**Files:**
- Modify: `src/Core.js` (끝에 추가)
- Test: `test/builder.test.js`
- Scratch(커밋 안 함): `$SCRATCH/parity.js`

**Interfaces:**
- Consumes: Task 1~4 함수 전부
- Produces: `createPqlBuilder_({ dealShopIds: Set, unmappedDeals: [{ id, title, raw, keys }], rejectedPairs: Set }) → { onRow(row), finish() → { cleanRows, uploadRows, matches, counts, targetCounts } }` — `cleanRows`·`uploadRows`는 헤더 포함

- [ ] **Step 1: 실패하는 테스트 `test/builder.test.js` 작성**

```javascript
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

function build(rows, opts) {
  const b = createPqlBuilder_(Object.assign({ dealShopIds: new Set(), unmappedDeals: [], rejectedPairs: new Set() }, opts || {}));
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
  assert.deepStrictEqual(out.uploadRows.slice(1).map((r) => r[1]), [9]);
  assert.deepStrictEqual(out.targetCounts, { '업셀, 푸시': 1, 업셀: 1 });
  assert.deepStrictEqual(out.matches.map((m) => m.tier), ['high', 'review']);
});

test('필수 열이 없으면 없는 열 이름을 모두 띄우고 멈춘다', () => {
  const b = createPqlBuilder_({ dealShopIds: new Set(), unmappedDeals: [], rejectedPairs: new Set() });
  assert.throws(() => b.onRow(H.filter((h) => h !== '알파업셀 상태' && h !== '담당자명')), /필수 열이 없습니다: 알파업셀 상태, 담당자명/);
});

test('빈 CSV는 멈춘다', () => {
  const b = createPqlBuilder_({ dealShopIds: new Set(), unmappedDeals: [], rejectedPairs: new Set() });
  assert.throws(() => b.finish(), /CSV가 비어 있습니다/);
});
```

- [ ] **Step 2: 실패 확인**

Run: `node --test test/builder.test.js`
Expected: FAIL — `createPqlBuilder_ is not defined`

- [ ] **Step 3: `src/Core.js` 끝에 추가**

```javascript
/* ---------- 빌더 ---------- */

// CSV 행을 받아 클렌징·역매핑 대조를 한 번에 하고, 끝나면 역매핑 제외·타겟 판정 후 출력 행을 만든다. 첫 행은 헤더.
function createPqlBuilder_(opts) {
  const matcher = createDealMatcher_(opts.unmappedDeals);
  const counts = { total: 0, orders: 0, review: 0, site: 0, pro: 0, phone: 0, deal: 0, mapped: 0, noTarget: 0 };
  const kept = [];
  let hi = null;
  return {
    onRow: function (row) {
      if (!hi) {
        hi = headerIndex_(row);
        const missing = REQUIRED_COLUMNS.filter(function (c) { return hi.last[c] === undefined; });
        if (missing.length) throw new Error('CSV에 필수 열이 없습니다: ' + missing.join(', '));
        return;
      }
      if (row.length === 1 && row[0] === '') return; // 빈 줄
      counts.total++;
      const r = toRecord_(row, hi);
      matcher.onRow(row, hi, r.shopId, r.get('shop_name'));
      const reason = cleanseReason_(r, opts.dealShopIds);
      if (reason) counts[reason]++;
      else kept.push(r);
    },
    finish: function () {
      if (!hi) throw new Error('CSV가 비어 있습니다');
      const matches = matcher.results(opts.rejectedPairs);
      const mapped = new Set();
      const suspect = {};
      matches.forEach(function (m) {
        if (m.tier === 'high') mapped.add(m.candidates[0].shopId);
        else if (m.tier === 'review') {
          m.candidates.forEach(function (c) { (suspect[c.shopId] = suspect[c.shopId] || []).push(String(m.deal.id)); });
        }
      });
      const cleanRows = [OUTPUT_HEADERS];
      const uploadRows = [UPLOAD_HEADERS];
      const targetCounts = {};
      kept.forEach(function (r) {
        if (mapped.has(r.shopId)) {
          counts.mapped++;
          return;
        }
        const targets = matchTargets_(r);
        if (!targets.length) {
          counts.noTarget++;
          return;
        }
        const label = serviceLabel_(r);
        const sus = suspect[r.shopId] || [];
        cleanRows.push(cleanRow_(r, targets, label, sus));
        if (!sus.length) uploadRows.push(uploadRow_(r, label));
        const key = targets.join(', ');
        targetCounts[key] = (targetCounts[key] || 0) + 1;
      });
      return { cleanRows: cleanRows, uploadRows: uploadRows, matches: matches, counts: counts, targetCounts: targetCounts };
    },
  };
}
```

- [ ] **Step 4: 통과 확인**

Run: `node --test test/`
Expected: PASS (이전 28 + builder 3)

- [ ] **Step 5: 실데이터 패리티·성능 스크립트 작성 (scratch, 커밋 안 함)**

`$SCRATCH/parity.js`:
```javascript
// 10/1 CSV + 현재 Sales 딜로 빌더를 돌려 브레인스토밍 추정치와 대조하고 소요 시간을 잰다 (읽기 전용)
require('/Users/sales/Desktop/02.PJT/Work/PQL_auto/test/gas.js');
const fs = require('fs');
const SCRATCH = __dirname;
const token = process.env.PIPEDRIVE_API_TOKEN;
async function pd(path) {
  const r = await fetch('https://api.pipedrive.com' + path, { headers: { 'x-api-token': token } });
  if (!r.ok) throw new Error(path + ' ' + r.status);
  return r.json();
}
(async () => {
  const deals = [];
  let cursor = '';
  do {
    const j = await pd('/api/v2/deals?pipeline_id=9&limit=500' + (cursor ? '&cursor=' + encodeURIComponent(cursor) : ''));
    deals.push(...j.data);
    cursor = (j.additional_data && j.additional_data.next_cursor) || '';
  } while (cursor);
  const split = splitDeals_(deals);
  const byIds = async (entity, ids) => {
    const out = {};
    for (let i = 0; i < ids.length; i += 100) (await pd('/api/v2/' + entity + '?ids=' + ids.slice(i, i + 100).join(',') + '&limit=100')).data.forEach((x) => { out[x.id] = x; });
    return out;
  };
  const persons = await byIds('persons', [...new Set(split.unmapped.map((d) => d.person_id).filter(Boolean))]);
  const orgs = await byIds('organizations', [...new Set(split.unmapped.map((d) => d.org_id).filter(Boolean))]);
  const unmappedDeals = split.unmapped.map((d) => ({ id: d.id, title: d.title || '', raw: rawShopId_(d), keys: dealMatchKeys_(d, persons[d.person_id], orgs[d.org_id]) }));

  const t0 = Date.now();
  const b = createPqlBuilder_({ dealShopIds: split.shopIds, unmappedDeals, rejectedPairs: new Set() });
  const buf = fs.readFileSync(SCRATCH + '/src_1001.csv');
  const dec = new TextDecoder('utf-8', { ignoreBOM: true });
  const p = createCsvParser_(b.onRow);
  streamChunks_(buf.length, 20 * 1024 * 1024, (s, e) => dec.decode(buf.subarray(s, e + 1)), p.feed);
  p.end();
  const out = b.finish();
  const tiers = {};
  out.matches.forEach((m) => { tiers[m.tier] = (tiers[m.tier] || 0) + 1; });
  console.log('deals', deals.length, 'unmapped', split.unmapped.length, 'ms', Date.now() - t0);
  console.log('counts', out.counts);
  console.log('tiers', tiers, 'clean', out.cleanRows.length - 1, 'upload', out.uploadRows.length - 1, out.targetCounts);
  console.log('서브마켓·장신몰 제외 여부', out.cleanRows.some((r) => r[1] === '108903' || r[1] === '126757') ? 'FAIL' : 'OK');
})();
```

- [ ] **Step 6: 패리티·성능 실행**

Run:
```bash
cd /Users/sales/Desktop/02.PJT/Work/pipedrive_auto && set -a && source .env && set +a && node $SCRATCH/parity.js
```
Expected (2026-09-30 브레인스토밍 추정치, 이후 생긴 딜만큼 소폭 차이 허용):
- `deals` 약 5,472 / `unmapped` 약 253
- `counts`: total 65,786 / orders 60,564 / review 110 / site 10 / pro 307 / phone 1,571 / deal 약 2,799
- `tiers`: high 약 49 / review 약 50 / none 약 154
- `clean` 약 311 (업셀 223 · 업셀, 푸시 56 · 푸시 32), `upload` 약 304
- `서브마켓·장신몰 제외 여부 OK`
- `ms`: node에서 10초 미만 (Apps Script는 이보다 느리다. 30초를 넘으면 파서의 문자 단위 `field +=`를 구간 slice 방식으로 바꾼다)

추정치와 5건 넘게 다르면 다른 shop_id를 뽑아 원인(정규화 정규식 차이 등)을 확인하고, 설명되지 않으면 Task 2~3으로 돌아간다.

- [ ] **Step 7: 커밋**

```bash
git add src/Core.js test/builder.test.js
git commit -m "[feat] PQL 빌더 추가 (CSV 1회 순회로 클렌징·역매핑·타겟 처리)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: I/O 계층 (Drive·Pipedrive·Sheets)

**Files:**
- Create: `src/Io.js`

**Interfaces:**
- Consumes: `splitDeals_`, `rawShopId_`, `dealMatchKeys_`, `createCsvParser_`, `streamChunks_`, Config 상수
- Produces:
  - `pdToken_() → string` (없으면 입력창으로 받아 저장)
  - `pdRequest_(token, method, path, body?) → object` (429 3회 재시도, 300 이상이면 throw)
  - `fetchPipedrive_(token) → { deals, dealShopIds: Set, unmappedDeals, users, stages, labels }`
  - `pdGetShopId_(token, dealId) → string`
  - `pdSetShopId_(token, dealId, shopId, note) → ''|'노트 실패'`
  - `findLatestCsv_() → DriveApp.File`
  - `streamCsvFile_(file, onRow)`
  - `readMappingRows_(ss) → row[]`, `writeMappingRows_(ss, rows)`, `writeDealList_(ss, rows)`, `writeCleanTab_(ss, rows) → tabName`
  - `exportUploadXlsx_(rows, fileName) → DriveApp.File`

Apps Script 서비스 전용이라 node 단위 테스트 대상이 아니다. Task 7의 테스트 시트 실행이 검증이다.

- [ ] **Step 1: `src/Io.js` 작성 (전체)**

```javascript
/***************************************
 * 외부 I/O — Drive · Pipedrive · Sheets
 ***************************************/

/* ---------- Pipedrive ---------- */

function pdToken_() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty(PD_TOKEN_PROPERTY);
  if (token) return token;
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Pipedrive API 토큰', '처음 한 번만 입력합니다. 스크립트 속성에 저장됩니다.', ui.ButtonSet.OK_CANCEL);
  token = res.getResponseText().trim();
  if (res.getSelectedButton() !== ui.Button.OK || !token) throw new Error('Pipedrive 토큰이 없습니다');
  props.setProperty(PD_TOKEN_PROPERTY, token);
  return token;
}

function pdRequest_(token, method, path, body) {
  const opt = { method: method, muteHttpExceptions: true, headers: { 'x-api-token': token } };
  if (body) {
    opt.contentType = 'application/json';
    opt.payload = JSON.stringify(body);
  }
  for (let attempt = 0; ; attempt++) {
    const res = UrlFetchApp.fetch('https://api.pipedrive.com' + path, opt);
    const code = res.getResponseCode();
    if (code === 429 && attempt < 3) {
      Utilities.sleep(2000 * (attempt + 1));
      continue;
    }
    if (code >= 300) throw new Error('Pipedrive ' + method.toUpperCase() + ' ' + path.split('?')[0] + ' → ' + code + ' ' + res.getContentText().slice(0, 200));
    return JSON.parse(res.getContentText());
  }
}

// v2 cursor 페이지네이션
function pdList_(token, path) {
  const out = [];
  let cursor = '';
  do {
    const res = pdRequest_(token, 'get', path + (path.indexOf('?') < 0 ? '?' : '&') + 'limit=500' + (cursor ? '&cursor=' + encodeURIComponent(cursor) : ''));
    (res.data || []).forEach(function (x) { out.push(x); });
    cursor = (res.additional_data && res.additional_data.next_cursor) || '';
  } while (cursor);
  return out;
}

// v2 persons·organizations를 id 100개씩 조회
function pdByIds_(token, entity, ids) {
  const out = {};
  for (let i = 0; i < ids.length; i += 100) {
    const res = pdRequest_(token, 'get', '/api/v2/' + entity + '?ids=' + ids.slice(i, i + 100).join(',') + '&limit=100');
    (res.data || []).forEach(function (x) { out[x.id] = x; });
  }
  return out;
}

function uniqueIds_(arr) {
  return Array.from(new Set(arr.filter(function (x) { return x; })));
}

function fetchPipedrive_(token) {
  const deals = pdList_(token, '/api/v2/deals?pipeline_id=' + SALES_PIPELINE_ID);
  const split = splitDeals_(deals);
  const persons = pdByIds_(token, 'persons', uniqueIds_(split.unmapped.map(function (d) { return d.person_id; })));
  const orgs = pdByIds_(token, 'organizations', uniqueIds_(split.unmapped.map(function (d) { return d.org_id; })));
  const unmappedDeals = split.unmapped.map(function (d) {
    return { id: d.id, title: d.title || '', raw: rawShopId_(d), keys: dealMatchKeys_(d, persons[d.person_id], orgs[d.org_id]) };
  });
  const users = {};
  (pdRequest_(token, 'get', '/api/v1/users').data || []).forEach(function (u) { users[u.id] = u.name; });
  const stages = {};
  (pdRequest_(token, 'get', '/api/v1/stages?pipeline_id=' + SALES_PIPELINE_ID).data || []).forEach(function (s) { stages[s.id] = s.name; });
  const labels = {};
  const labelField = (pdRequest_(token, 'get', '/api/v1/dealFields?limit=500').data || []).filter(function (f) { return f.key === 'label'; })[0];
  ((labelField && labelField.options) || []).forEach(function (o) { labels[o.id] = o.label; });
  return { deals: deals, dealShopIds: split.shopIds, unmappedDeals: unmappedDeals, users: users, stages: stages, labels: labels };
}

function pdGetShopId_(token, dealId) {
  return rawShopId_(pdRequest_(token, 'get', '/api/v2/deals/' + dealId).data);
}

// shop_id를 쓰고 원래 값·근거를 노트로 남긴다. 노트만 실패하면 '노트 실패'를 돌려준다.
function pdSetShopId_(token, dealId, shopId, note) {
  const cf = {};
  cf[PD_FIELD_SHOP_ID] = String(shopId);
  pdRequest_(token, 'patch', '/api/v2/deals/' + dealId, { custom_fields: cf });
  try {
    pdRequest_(token, 'post', '/api/v1/notes', { deal_id: Number(dealId), content: note });
    return '';
  } catch (e) {
    return '노트 실패';
  }
}

/* ---------- Drive ---------- */

function findLatestCsv_() {
  const it = DriveApp.getFolderById(SOURCE_FOLDER_ID).getFiles();
  let best = null;
  while (it.hasNext()) {
    const f = it.next();
    const name = f.getName();
    if (name.indexOf(SOURCE_NAME_PREFIX) !== 0 || !/\.csv$/i.test(name)) continue;
    if (!best || f.getLastUpdated() > best.getLastUpdated()) best = f;
  }
  if (!best) throw new Error("'05. PQL' 폴더에 " + SOURCE_NAME_PREFIX + '*.csv 파일이 없습니다');
  return best;
}

function streamCsvFile_(file, onRow) {
  const parser = createCsvParser_(onRow);
  const token = ScriptApp.getOAuthToken();
  const url = 'https://www.googleapis.com/drive/v3/files/' + file.getId() + '?alt=media&supportsAllDrives=true';
  streamChunks_(file.getSize(), DOWNLOAD_CHUNK_BYTES, function (start, end) {
    const res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token, Range: 'bytes=' + start + '-' + end },
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code !== 206 && code !== 200) throw new Error('CSV 다운로드 실패 ' + code + ': ' + res.getContentText().slice(0, 200));
    return res.getContentText('UTF-8');
  }, parser.feed);
  parser.end();
}

// 임시 스프레드시트에 쓰고 xlsx로 내보낸 뒤 임시본은 휴지통으로. 실행한 사람의 내 드라이브에 저장된다.
function exportUploadXlsx_(rows, fileName) {
  const tmp = SpreadsheetApp.create('[tmp] ' + fileName);
  try {
    const sh = tmp.getSheets()[0];
    const range = sh.getRange(1, 1, rows.length, rows[0].length);
    range.setNumberFormat('@'); // 전화번호 앞자리 0 보존
    if (rows.length > 1) {
      sh.getRange(2, 2, rows.length - 1, 1).setNumberFormat('0'); // shop_id
      sh.getRange(2, 5, rows.length - 1, 1).setNumberFormat('0'); // 월 주문 수
    }
    range.setValues(rows);
    SpreadsheetApp.flush();
    const res = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) throw new Error('xlsx 변환 실패 ' + res.getResponseCode());
    return DriveApp.createFile(res.getBlob().setName(fileName));
  } finally {
    DriveApp.getFileById(tmp.getId()).setTrashed(true);
  }
}

/* ---------- Sheets ---------- */

function readMappingRows_(ss) {
  const sh = ss.getSheetByName(TAB_MAPPING);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, MAPPING_HEADERS.length).getValues();
}

function writeMappingRows_(ss, rows) {
  const sh = ss.getSheetByName(TAB_MAPPING) || ss.insertSheet(TAB_MAPPING);
  sh.clearContents();
  const all = [MAPPING_HEADERS].concat(rows);
  const range = sh.getRange(1, 1, all.length, MAPPING_HEADERS.length);
  range.setNumberFormat('@');
  range.setValues(all);
  sh.getRange(1, 1, 1, MAPPING_HEADERS.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (rows.length) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['승인', '거절'], true).setAllowInvalid(false).build();
    sh.getRange(2, 8, rows.length, 1).setDataValidation(rule);
  }
}

// 기존 테이블(표2) 범위를 넘는 행은 테이블 밖에 쓰인다. 제외 판정은 코드가 하므로 결과에는 영향 없다.
function writeDealList_(ss, rows) {
  const sh = ss.getSheetByName(TAB_DEAL_LIST) || ss.insertSheet(TAB_DEAL_LIST);
  sh.getRange(1, 1, sh.getMaxRows(), 5).clearContent();
  sh.getRange(1, 1, rows.length, 5).setValues(rows);
}

function writeCleanTab_(ss, rows) {
  const name = CLEAN_TAB_PREFIX + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const sh = ss.insertSheet(name);
  const range = sh.getRange(1, 1, rows.length, rows[0].length);
  range.setNumberFormat('@'); // 전화·shop_id 앞자리 0과 날짜 오인 방지
  range.setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sh.setColumnWidths(1, rows[0].length, 120);
  sh.setFrozenRows(1);
  return name;
}
```

- [ ] **Step 2: 문법 확인**

Run: `node --check src/Io.js && node --test test/`
Expected: 출력 없음(문법 OK) + 기존 테스트 전부 PASS

- [ ] **Step 3: 커밋**

```bash
git add src/Io.js
git commit -m "[feat] Drive 분할 다운로드·Pipedrive 조회/쓰기·시트 탭·xlsx 생성 I/O 추가

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 7: 실행 흐름·메뉴 + 테스트 시트 검증

**Files:**
- Create: `src/Main.js`

**Interfaces:**
- Consumes: Task 4~6 전부
- Produces: 메뉴 함수 `onOpen`, `runPql`, `applyApprovedMappings`

- [ ] **Step 1: `src/Main.js` 작성 (전체)**

```javascript
/***************************************
 * 메뉴 · 실행 흐름
 ***************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PQL 자동화')
    .addItem('PQL 생성', 'runPql')
    .addItem('승인된 shop_id 반영', 'applyApprovedMappings')
    .addToUi();
}

function step_(name, fn) {
  try {
    return fn();
  } catch (e) {
    throw new Error('[' + name + '] ' + e.message);
  }
}

function todayStr_(fmt) {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), fmt);
}

// 매핑 탭에서 '승인'된 대기 행을 Pipedrive에 반영한다. 반영 직전에 shop_id가 이미 숫자면 건너뛴다.
function applyApproved_(token, state, today) {
  const split = splitApprovals_(state.approved);
  split.duplicate.forEach(function (x) {
    const y = x.slice();
    y[8] = '실패: 중복 승인';
    state.keep.push(y);
  });
  let applied = 0;
  split.apply.forEach(function (x) {
    const y = x.slice();
    try {
      const cur = pdGetShopId_(token, y[0]);
      if (/^\d+$/.test(cur)) y[8] = '실패: 이미 shop_id ' + cur;
      else {
        const warn = pdSetShopId_(token, y[0], y[3], mappingNote_('승인', y[2], y[3], String(y[5]).split(', '), today));
        y[8] = warn ? '반영됨 (' + warn + ')' : '반영됨';
        applied++;
      }
    } catch (e) {
      y[8] = '실패: ' + e.message.slice(0, 80);
    }
    y[9] = today;
    state.keep.push(y);
  });
  state.applied = applied;
  return state;
}

function autoApply_(token, matches, today) {
  const overrides = {};
  const rows = [];
  let failed = 0;
  matches.forEach(function (m) {
    const c = m.candidates[0];
    let status;
    try {
      const warn = pdSetShopId_(token, m.deal.id, c.shopId, mappingNote_('자동', m.deal.raw, c.shopId, c.keys, today));
      overrides[m.deal.id] = c.shopId;
      status = warn ? '반영됨 (' + warn + ')' : '반영됨';
    } catch (e) {
      status = '실패: ' + e.message.slice(0, 80);
      failed++;
    }
    rows.push(mappingRow_(m, c, '높음', '', status, today));
  });
  return { overrides: overrides, rows: rows, failed: failed };
}

function runPql() {
  const ui = SpreadsheetApp.getUi();
  const started = Date.now();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const token = pdToken_();
    const today = todayStr_('yyyy-MM-dd');
    const state = step_('승인 매핑 반영', function () { return applyApproved_(token, readMappingState_(readMappingRows_(ss)), today); });
    const pd = step_('Pipedrive 조회', function () { return fetchPipedrive_(token); });
    const file = step_('CSV 찾기', findLatestCsv_);
    const builder = createPqlBuilder_({ dealShopIds: pd.dealShopIds, unmappedDeals: pd.unmappedDeals, rejectedPairs: state.rejectedPairs });
    step_('CSV 읽기', function () { streamCsvFile_(file, builder.onRow); });
    const out = builder.finish();

    const plan = planMappings_(out.matches, AUTO_APPLY, AUTO_APPLY_MAX);
    const auto = step_('shop_id 자동 반영', function () { return autoApply_(token, plan.auto, today); });
    const pending = [];
    plan.pendingHigh.forEach(function (m) { pending.push(mappingRow_(m, m.candidates[0], '높음', '', '대기', today)); });
    plan.review.forEach(function (m) {
      m.candidates.forEach(function (c) { pending.push(mappingRow_(m, c, '확인 필요', '', '대기', today)); });
    });

    const cleanTab = step_('시트 쓰기', function () {
      writeDealList_(ss, dealListRows_(pd.deals, pd.users, pd.stages, pd.labels, auto.overrides));
      writeMappingRows_(ss, pending.concat(auto.rows, state.keep));
      return writeCleanTab_(ss, out.cleanRows);
    });
    const xlsx = step_('xlsx 생성', function () { return exportUploadXlsx_(out.uploadRows, uploadFileName_(file.getName(), todayStr_('MMdd'))); });

    showSummary_(summaryLines_({
      fileName: file.getName(),
      fileUpdated: Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      counts: out.counts,
      targetCounts: out.targetCounts,
      cleanCount: out.cleanRows.length - 1,
      uploadCount: out.uploadRows.length - 1,
      approvedApplied: state.applied,
      autoApplied: plan.auto.length - auto.failed,
      autoFailed: auto.failed,
      overLimit: plan.overLimit,
      pending: pending.length,
      elapsedSec: Math.round((Date.now() - started) / 1000),
    }).concat(['clean 탭: ' + cleanTab]), xlsx);
  } catch (e) {
    ui.alert('PQL 생성 실패', e.message, ui.ButtonSet.OK);
  }
}

function applyApprovedMappings() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const token = pdToken_();
    const state = applyApproved_(token, readMappingState_(readMappingRows_(ss)), todayStr_('yyyy-MM-dd'));
    writeMappingRows_(ss, state.pending.concat(state.keep));
    ui.alert('승인된 shop_id 반영', state.applied + '건 반영했습니다. 결과는 상태 열을 보세요.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('반영 실패', e.message, ui.ButtonSet.OK);
  }
}

function showSummary_(lines, xlsx) {
  const esc = function (s) {
    return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
  };
  const html = '<div style="font:13px/1.7 sans-serif">' + lines.map(esc).join('<br>') +
    '<p><a href="' + esc(xlsx.getDownloadUrl()) + '" target="_blank">' + esc(xlsx.getName()) + ' 다운로드</a></p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(480).setHeight(520), 'PQL 생성 완료');
}
```

- [ ] **Step 2: 문법·테스트 확인**

Run: `for f in src/*.js; do node --check "$f" || exit 1; done && node --test test/`
Expected: 문법 오류 없음 + 전부 PASS

- [ ] **Step 3: 커밋**

```bash
git add src/Main.js
git commit -m "[feat] PQL 생성·승인 반영 메뉴와 실행 흐름 추가

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

- [ ] **Step 4: 테스트 시트에 배포 (자동 반영 끔)**

라이브 PQL 시트를 건드리지 않도록 빈 테스트 스프레드시트에 새 바운드 스크립트를 만들고, `AUTO_APPLY=false`로 바꾼 사본을 올린다. Pipedrive에는 읽기만 한다.

```bash
T=$SCRATCH/test-deploy && rm -rf $T && mkdir -p $T/src && cp src/* $T/src/
sed -i '' 's/^const AUTO_APPLY = true;/const AUTO_APPLY = false;/' $T/src/Config.js && grep -n "^const AUTO_APPLY =" $T/src/Config.js
cd $T && clasp create-script --type sheets --title "PQL_test_$(date +%m%d)" --rootDir src
cd $T && clasp push -f && clasp show-file-status
```
Expected: `AUTO_APPLY = false` 확인, 새 스프레드시트 URL과 scriptId 출력, push 파일 목록에 `appsscript.json Config.js Core.js Io.js Main.js`

- [ ] **Step 5: 사용자 실행 (사람 작업)**

사용자에게 요청: 테스트 스프레드시트 열기 → 새로고침 → 메뉴 `PQL 자동화 > PQL 생성` → 권한 승인 → Pipedrive 토큰 입력 → 요약 창 스크린샷 또는 문구 공유. `승인된 shop_id 반영`은 누르지 않는다(누르면 Pipedrive에 씀).

- [ ] **Step 6: 결과 대조**

```bash
# 탭 목록·행 수
gws sheets spreadsheets get --params '{"spreadsheetId":"<테스트 시트 ID>","fields":"sheets.properties(title,gridProperties)"}'
# clean 탭 앞 3행
gws sheets spreadsheets values get --params '{"spreadsheetId":"<테스트 시트 ID>","range":"<clean 탭 이름>!A1:H3"}'
# 내 드라이브의 xlsx
gws drive files list --params '{"q":"name contains '"'"'pipedrive_up('"'"' and trashed=false","orderBy":"createdTime desc","pageSize":1,"fields":"files(id,name,size,createdTime)"}'
```
xlsx를 `$SCRATCH`로 받아 확인:
```bash
python3 -c "import openpyxl,sys;ws=openpyxl.load_workbook(sys.argv[1]).worksheets[0];r=list(ws.iter_rows(values_only=True));print(len(r)-1,r[0]);print(r[1]);print({x[5] for x in r[1:]},{x[6] for x in r[1:]})" $SCRATCH/pipedrive_up.xlsx
```
Expected:
- 요약 창 수치가 Task 5 Step 6 패리티 결과와 일치(그 사이 생긴 딜만큼 차이 허용), 소요 6분 미만. **소요 시간을 기록한다**
- clean 탭 행 수 = 결과 곳 수 + 1, 5~7열이 `타겟`·`서비스 라벨`·`딜 의심`
- `shop_id 매핑` 탭: `AUTO_APPLY=false`라 높은 확신도 전부 `대기`/`높음`, 판정 열 드롭다운
- `deal list` 탭 행 수 = 딜 수 + 1
- xlsx: 헤더가 0901 양식 15열과 같음, 소유자 `{'한서연'}`, 단계 `{'컨택전'}`, 행 수 = 요약의 업로드 행 수, 전화번호가 `010-` 형식 문자열
- 원천 폴더 `05. PQL`에 새 파일이 생기지 않음 (`gws drive files list`로 최신 수정일 확인)

6분을 넘기면 Task 5 Step 6의 파서 최적화로 돌아간다.

---

### Task 8: 문서 동기화

**Files:**
- Modify: `PQL.md`, `CLAUDE.md`, `ARCHITECTURE.md`, `AGENTS.md`, `docs/PLANS.md`
- Modify: `/Users/sales/Desktop/02.PJT/ARCHITECTURE.md` (워크스페이스 저장소, 별도 커밋)

- [ ] **Step 1: `PQL.md`를 사용법 문서로 교체**

코드블록을 지우고 다음 내용으로 바꾼다: 코드 위치(`src/`), 배포(`clasp push -f`, `.clasp.json`), 메뉴 2개 사용법, 월간 절차(CSV 업로드 → `PQL 생성` → 요약 확인 → xlsx 다운로드 → Pipedrive 가져오기 → `shop_id 매핑` 탭 판정), 설정 위치(`src/Config.js`), 토큰 재설정 방법(Apps Script 편집기 > 프로젝트 설정 > 스크립트 속성 `PIPEDRIVE_API_TOKEN`).

- [ ] **Step 2: `CLAUDE.md`·`ARCHITECTURE.md`·`AGENTS.md`·`docs/PLANS.md` 갱신**

- CLAUDE.md: 개요(목적 = Sales 누락 타겟 발굴), 기술 스택(Pipedrive REST 추가), 주요 명령어(`node --test test/`, `clasp push -f`), 폴더 구조(`src/`, `test/`, `.clasp.json`), 최근 변경사항을 이번 재구축 1세대로 교체하고 지난 세대(2026-06-30 프로 필터, 2026-03-09 변경들)는 `CHANGELOG.md`로 옮긴다. 트러블슈팅에 "CSV 45MB·시트 셀 한도로 raw 적재 방식 폐기" 추가
- ARCHITECTURE.md: 데이터 흐름을 spec 4장 흐름으로 교체, 규칙 표(spec 5장), 제약(6분·50MB/회·원천 폴더 읽기 전용)
- AGENTS.md: Golden Principle 4(시트 API 3회)를 "시트 쓰기는 모든 계산 뒤 한 번에, 원천 폴더에는 쓰지 않는다"로, 5를 "Pipedrive shop_id 필드에 쓰는 주체다(빈칸·텍스트만)"로 갱신. Key Files를 새 구조로
- docs/PLANS.md: Phase 3 "자동 실행·clean 정리·중복 감지" 중 딜 중복 감지 완료 표시, 기술 부채 표에서 해소된 항목(필터 하드코딩→Config.js, 에러 메시지→요약 창) 갱신

- [ ] **Step 3: 워크스페이스 교차 영향 표 갱신**

`/Users/sales/Desktop/02.PJT/ARCHITECTURE.md` 62행·124행: PQL_auto를 "하류 데이터, API 미연동"에서 "Pipedrive API 소비자 — Sales(9) 딜 읽기 + 빈칸·텍스트 shop_id에 한해 shop_id 쓰기·노트 작성, 공유 토큰"으로 바꾸고 API 직접 소비자 수를 갱신한다. Drive `05. PQL` 폴더 공유 항목에는 "PQL_auto는 읽기만 함(임시 파일 생성 제거)"을 적는다.

- [ ] **Step 4: 커밋**

```bash
git add PQL.md CLAUDE.md ARCHITECTURE.md AGENTS.md docs/PLANS.md CHANGELOG.md
git commit -m "[docs] PQL 재구축에 맞춰 사용법·아키텍처·원칙 문서 갱신

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
cd /Users/sales/Desktop/02.PJT && git add ARCHITECTURE.md && git commit -m "[docs] PQL_auto를 Pipedrive shop_id 쓰기 주체로 교차 영향 표 갱신

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 9: 라이브 배포·검증·정리 (4단계 독립 리뷰 통과 후)

**게이트:** Task 7 완료 후 `/orchestration`으로 Orca 워커(Codex)에게 브랜치 diff 독립 리뷰를 맡기고, 살아남은 Critical 0개일 때만 진행한다.

- [ ] **Step 1: 라이브 배포**

```bash
cd /Users/sales/Desktop/02.PJT/Work/PQL_auto && grep -n "^const AUTO_APPLY =" src/Config.js   # true 확인
clasp push -f && clasp pull && git status --short src/   # 원격 = 로컬 (변경 없음)
```
Expected: `AUTO_APPLY = true`, push 5개 파일, 재pull 후 `src/` 변경 없음. 원격 `Code.js`가 사라졌는지 `clasp show-file-status`로 확인

- [ ] **Step 2: 사용자 라이브 실행 (사람 작업)**

사용자에게 요청: 라이브 PQL 시트 새로고침 → `PQL 생성` → 권한 승인 → 토큰 입력 → 요약 공유. 첫 실행에서 높은 확신 약 49건이 Pipedrive에 자동 반영된다.

- [ ] **Step 3: Pipedrive 반영 대조 (읽기 전용)**

`shop_id 매핑` 탭에서 상태 `반영됨` 딜 ID를 뽑아 Pipedrive에서 shop_id와 노트를 읽는다:
```bash
cd /Users/sales/Desktop/02.PJT/Work/pipedrive_auto && set -a && source .env && set +a
for id in <반영됨 딜 ID 목록>; do
  curl -s -H "x-api-token: $PIPEDRIVE_API_TOKEN" "https://api.pipedrive.com/api/v2/deals/$id" | python3 -c "import json,sys;d=json.load(sys.stdin)['data'];print(d['id'],d['title'],d['custom_fields']['9d4ea1fcf0bde157910e96a2e0354e76c220e6c8'])"
  curl -s -H "x-api-token: $PIPEDRIVE_API_TOKEN" "https://api.pipedrive.com/api/v1/notes?deal_id=$id&limit=1&sort=add_time%20DESC" | python3 -c "import json,sys;print((json.load(sys.stdin)['data'] or [{}])[0].get('content'))"
done
```
Expected: 각 딜 shop_id가 매핑 탭 `후보 shop_id`와 같고, 최신 노트가 `[PQL 자동매핑 ...] shop_id '원래값' → ...`. 서브마켓 → `108903`, 장신몰 → `126757`

- [ ] **Step 4: 일회성 정리 (실행 직전 사용자 재확인)**

사용자 확인 후에만 실행:
- 라이브 시트 `raw`·`raw222` 탭 삭제 (Sheets API `deleteSheet`, sheetId는 `gws sheets spreadsheets get`으로 조회)
- 폴더 `05. PQL`의 `[Temp] all_subscription_0405`, `[Temp XLSX] all_subscription_0405.xlsx` 휴지통 이동 (`gws drive files update --params '{"fileId":"<id>","supportsAllDrives":true}' --json '{"trashed":true}'`)
- 테스트 스프레드시트 휴지통 이동

- [ ] **Step 5: PR 생성**

```bash
git push -u origin refactor/csv-direct-pipeline
gh pr create --title "[refactor] PQL 파이프라인 재구축 (CSV 직접 처리·딜 자동 제외·shop_id 역매핑)" --body "<spec 요약·검증 결과·라이브 실행 수치>

🤖 Generated with [Claude Code](https://claude.com/claude-code)"
```
머지는 4단계 리뷰 통과 기록이 PR에 남은 뒤에만 한다.
