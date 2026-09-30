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
