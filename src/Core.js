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

// Apps Script getContentText는 파일 첫머리의 UTF-8 BOM을 지운다(2026-09-30 실측). 그러면 첫 조각 바이트 수가 3 적게 계산돼
// 다음 조각이 3바이트 앞당겨진다. 첫 조각이 BOM 없이 오면 3바이트 뒤부터 받은 글자(probeFrom3)와 첫머리를 비교해,
// 같으면 BOM이 지워진 것이므로 되돌린다. (3바이트 Range 요청은 Apps Script에서 빈 본문이 와서 바이트로는 판별하지 않는다)
function restoreBom_(fetchRange, probeFrom3) {
  return function (start, end) {
    const text = fetchRange(start, end);
    if (start !== 0 || text.charCodeAt(0) === 0xfeff) return text;
    const probe = probeFrom3();
    if (!probe) throw new Error('CSV 첫머리 확인 실패(빈 응답)');
    const n = Math.min(32, probe.length - 1, text.length); // probe 끝 글자는 잘렸을 수 있다
    return n > 0 && text.slice(0, n) === probe.slice(0, n) ? '\ufeff' + text : text;
  };
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

// Pipedrive 가져오기 양식 (pipedrive_up(0901).xlsx와 같은 15열). shop_id는 Pipedrive 텍스트 필드라 문자열, 월 주문 수는 숫자 필드라 숫자
function uploadRow_(r, label) {
  return [
    r.get('shop_name'), r.shopId, r.get('mall_id'), r.get('플랫폼'), toNumberOr_(r.ordersRaw),
    DEAL_OWNER, DEAL_STAGE, label, r.get('회사명'), r.get('담당자명'), r.get('쇼핑몰명'),
    r.phone, r.get('담당자이메일'), r.get('대표도메인'), r.address,
  ];
}

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
