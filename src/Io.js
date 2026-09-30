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

// deadlineMs를 주면 429 재시도 대기가 그 시각을 넘길 때 기다리지 않고 실패한다
function pdRequest_(token, method, path, body, deadlineMs) {
  const opt = { method: method, muteHttpExceptions: true, headers: { 'x-api-token': token } };
  if (body) {
    opt.contentType = 'application/json';
    opt.payload = JSON.stringify(body);
  }
  for (let attempt = 0; ; attempt++) {
    const res = UrlFetchApp.fetch('https://api.pipedrive.com' + path, opt);
    const code = res.getResponseCode();
    if (code === 429 && attempt < 3) {
      const wait = 2000 * (attempt + 1);
      if (deadlineMs && Date.now() + wait > deadlineMs) throw new Error('Pipedrive 요청 과다(429) — 시간 예산 안에 재시도할 수 없음');
      Utilities.sleep(wait);
      continue;
    }
    if (code >= 300) throw new Error('Pipedrive ' + method.toUpperCase() + ' ' + path.split('?')[0] + ' → ' + code + ' ' + res.getContentText().slice(0, 200));
    return JSON.parse(res.getContentText());
  }
}

// v2 cursor 페이지네이션. map을 주면 페이지마다 바로 줄여서 담는다(원본 객체를 쌓아 두지 않는다).
function pdList_(token, path, map) {
  const out = [];
  let cursor = '';
  do {
    const res = pdRequest_(token, 'get', path + (path.indexOf('?') < 0 ? '?' : '&') + 'limit=500' + (cursor ? '&cursor=' + encodeURIComponent(cursor) : ''));
    (res.data || []).forEach(function (x) { out.push(map ? map(x) : x); });
    cursor = (res.additional_data && res.additional_data.next_cursor) || '';
  } while (cursor);
  return out;
}

// v2 persons·organizations를 id 100개씩 조회. 역매핑에 쓰는 이름·이메일·전화만 남긴다.
function pdByIds_(token, entity, ids) {
  const out = {};
  for (let i = 0; i < ids.length; i += 100) {
    const res = pdRequest_(token, 'get', '/api/v2/' + entity + '?ids=' + ids.slice(i, i + 100).join(',') + '&limit=100');
    (res.data || []).forEach(function (x) { out[x.id] = { name: x.name, emails: x.emails, phones: x.phones }; });
  }
  return out;
}

function uniqueIds_(arr) {
  return Array.from(new Set(arr.filter(function (x) { return x; })));
}

function fetchPipedrive_(token) {
  const deals = pdList_(token, '/api/v2/deals?pipeline_id=' + SALES_PIPELINE_ID + '&custom_fields=' + PD_DEAL_CUSTOM_FIELDS.join(','), slimDeal_);
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

// 모든 shop_id 쓰기의 공통 관문. 쓰기 직전에 현재 값을 다시 읽어, 숫자면 쓰지 않고 처음 본 값(expectedRaw)과 달라졌으면 보류한다.
// 노트에는 실제로 덮어쓴 직전 값이 남는다. (다른 도구가 GET과 PATCH 사이에 쓰는 경쟁까지 막지는 못한다 — Pipedrive에 조건부 갱신이 없다)
function pdApplyShopId_(token, dealId, shopId, expectedRaw, makeNote) {
  if (!/^\d+$/.test(String(shopId))) return { applied: false, status: '실패: 후보 shop_id가 숫자가 아님' };
  const cur = pdGetShopId_(token, dealId);
  if (/^\d+$/.test(cur)) return { applied: false, status: '건너뜀: 이미 shop_id ' + cur };
  if (cur !== String(expectedRaw == null ? '' : expectedRaw).trim()) return { applied: false, status: '보류: shop_id가 바뀜(' + cur + ')' };
  const warn = pdSetShopId_(token, dealId, shopId, makeNote(cur));
  return { applied: true, status: warn ? '반영됨 (' + warn + ')' : '반영됨' };
}

// shop_id를 쓰고 원래 값·근거를 노트로 남긴다. 노트만 실패하면 '노트 실패'를 돌려준다. 직접 부르지 말고 pdApplyShopId_를 거친다.
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

// 여러 요청을 동시에 보낸다. 429(요청 과다)는 하나씩 재시도한다 (429는 처리되지 않은 요청이라 다시 보내도 중복되지 않는다).
function pdFetchAll_(token, reqs, deadlineMs) {
  if (!reqs.length) return [];
  const responses = UrlFetchApp.fetchAll(reqs.map(function (r) {
    const o = { url: 'https://api.pipedrive.com' + r.path, method: r.method, muteHttpExceptions: true, headers: { 'x-api-token': token } };
    if (r.body) {
      o.contentType = 'application/json';
      o.payload = JSON.stringify(r.body);
    }
    return o;
  }));
  return responses.map(function (res, k) {
    const code = res.getResponseCode();
    if (code === 429) {
      try {
        return { ok: true, body: pdRequest_(token, reqs[k].method, reqs[k].path, reqs[k].body, deadlineMs) };
      } catch (e) {
        return { ok: false, error: e.message.slice(0, 120) };
      }
    }
    if (code >= 300) return { ok: false, error: code + ' ' + res.getContentText().slice(0, 120) };
    return { ok: true, body: JSON.parse(res.getContentText()) };
  });
}

// UPLOAD_BATCH곳씩 조직(같은 이름 재사용, 없으면 생성) → 담당자 → 딜 순으로 만든다.
// deadlineMs가 지나면 새 묶음을 시작하지 않고(남은 곳은 skipped), 진행 중 묶음의 429 재시도는 그 뒤 30초까지만 기다린다.
// 한 묶음에서 예외가 나면 그 묶음만 실패로 적고 다음 묶음으로 간다.
function pdCreateDeals_(token, items, deadlineMs) {
  const results = [];
  const orgIds = {}; // 이번 실행 안에서 조직 이름 → id
  const hardMs = deadlineMs + 30000;
  for (let i = 0; i < items.length; i += UPLOAD_BATCH) {
    const batch = items.slice(i, i + UPLOAD_BATCH);
    if (Date.now() > deadlineMs) {
      for (let j = i; j < items.length; j++) results.push({ skipped: true });
      break;
    }
    let res;
    try {
      res = pdUploadBatch_(token, batch, orgIds, hardMs);
    } catch (e) {
      res = batch.map(function () { return { error: '예외: ' + String(e.message).slice(0, 100) }; });
    }
    res.forEach(function (r) { results.push(r); });
  }
  return results;
}

// 선행 단계(조직 검색·생성, 담당자)가 실패한 몰은 딜을 만들지 않는다: 딜이 생기면 다음 실행에서 빠져 복구할 기회가 없다
function pdUploadBatch_(token, batch, orgIds, hardMs) {
  const res = batch.map(function () { return null; });
  const orgOf = {};
  batch.forEach(function (it) { orgOf[it.org.name] = it.org; });
  const names = Object.keys(orgOf).filter(function (n) { return !(n in orgIds); });
  const found = pdFetchAll_(token, names.map(function (n) {
    return { method: 'get', path: '/api/v2/organizations/search?term=' + encodeURIComponent(n) + '&fields=name&exact_match=true&limit=1' };
  }), hardMs);
  const searchError = {};
  const missing = [];
  names.forEach(function (n, k) {
    if (!found[k].ok) {
      searchError[n] = found[k].error;
      return;
    }
    const hits = found[k].body.data && found[k].body.data.items;
    if (hits && hits.length) orgIds[n] = hits[0].item.id;
    else missing.push(n);
  });
  const created = pdFetchAll_(token, missing.map(function (n) { return { method: 'post', path: '/api/v2/organizations', body: orgPayload_(orgOf[n]) }; }), hardMs);
  const createError = {};
  missing.forEach(function (n, k) {
    if (created[k].ok) orgIds[n] = created[k].body.data.id;
    else createError[n] = created[k].error;
  });
  batch.forEach(function (it, k) {
    const n = it.org.name;
    if (searchError[n]) res[k] = { error: '조직 검색 실패: ' + searchError[n] };
    else if (!orgIds[n]) res[k] = { error: '조직 생성 실패: ' + (createError[n] || '') };
  });

  const needPerson = batch.map(function (it, k) { return k; }).filter(function (k) { return !res[k]; });
  const persons = pdFetchAll_(token, needPerson.map(function (k) {
    return { method: 'post', path: '/api/v2/persons', body: personPayload_(batch[k].person, orgIds[batch[k].org.name]) };
  }), hardMs);
  const personOf = {};
  needPerson.forEach(function (k, j) {
    if (persons[j].ok) personOf[k] = persons[j].body.data.id;
    else res[k] = { error: '담당자 생성 실패: ' + persons[j].error };
  });

  const needDeal = needPerson.filter(function (k) { return !res[k]; });
  const deals = pdFetchAll_(token, needDeal.map(function (k) {
    return { method: 'post', path: '/api/v2/deals', body: dealPayload_(batch[k], personOf[k], orgIds[batch[k].org.name]) };
  }), hardMs);
  needDeal.forEach(function (k, j) {
    res[k] = deals[j].ok ? { id: deals[j].body.data.id } : { error: '딜 생성 실패(담당자 ' + personOf[k] + ' 생성됨): ' + deals[j].error };
  });
  return res;
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
  const fetchRange = function (start, end) {
    const res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token, Range: 'bytes=' + start + '-' + end },
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code !== 206 && code !== 200) throw new Error('CSV 다운로드 실패 ' + code + ': ' + res.getContentText().slice(0, 200));
    return res;
  };
  const size = file.getSize();
  const fetchText = function (start, end) { return fetchRange(start, end).getContentText('UTF-8'); };
  const probeFrom3 = function () { return size > 3 ? fetchText(3, Math.min(size, 131) - 1) : 'x'; };
  streamChunks_(size, DOWNLOAD_CHUNK_BYTES, restoreBom_(fetchText, probeFrom3), parser.feed);
  parser.end();
}

/* ---------- Sheets ---------- */

function readMappingRows_(ss) {
  const sh = ss.getSheetByName(TAB_MAPPING);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, MAPPING_HEADERS.length).getValues();
}

function writeMappingRows_(ss, rows) {
  const sh = ss.getSheetByName(TAB_MAPPING) || ss.insertSheet(TAB_MAPPING);
  const all = [MAPPING_HEADERS].concat(rows);
  const range = sh.getRange(1, 1, all.length, MAPPING_HEADERS.length);
  range.setNumberFormat('@');
  range.setValues(all); // 새 값을 먼저 쓴다: 도중에 실패해도 기존 거절·반영 이력이 먼저 지워지지 않는다
  const extra = sh.getLastRow() - all.length;
  if (extra > 0) sh.getRange(all.length + 1, 1, extra, MAPPING_HEADERS.length).clearContent();
  sh.getRange(1, 1, 1, MAPPING_HEADERS.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (rows.length) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['승인', '거절'], true).setAllowInvalid(false).build();
    sh.getRange(2, 8, rows.length, 1).setDataValidation(rule);
  }
}

// 시트 쓰기. 필요한 탭을 먼저 다 만들고, 탭을 하나 쓸 때마다 바로 반영(flush)한다.
// 5천 행 쓰기가 반영되기 전에 탭을 추가하다 라이브 문서에서 '스프레드시트 서비스 타임아웃'이 세 번 났다(2026-09-30,
// 같은 작업을 단계마다 flush하면 모두 3초 안). deal list는 참고용이라 실패해도 멈추지 않고 경고만 돌려준다.
function writeOutputs_(ss, mappingRows, cleanRows, dealRows) {
  const cleanName = CLEAN_TAB_PREFIX + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  [TAB_MAPPING, cleanName, TAB_DEAL_LIST].forEach(function (n) { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
  SpreadsheetApp.flush();
  writeMappingRows_(ss, mappingRows);
  SpreadsheetApp.flush();
  writeCleanTab_(ss, cleanName, cleanRows);
  SpreadsheetApp.flush();
  let warn = '';
  try {
    writeDealList_(ss, dealRows);
    SpreadsheetApp.flush();
  } catch (e) {
    warn = 'deal list 갱신 실패: ' + e.message;
  }
  return { cleanName: cleanName, warn: warn };
}

// 새 값을 먼저 쓰고 남는 아래 행만 지운다
function writeDealList_(ss, rows) {
  const sh = ss.getSheetByName(TAB_DEAL_LIST) || ss.insertSheet(TAB_DEAL_LIST);
  sh.getRange(1, 1, rows.length, 5).setValues(rows);
  const extra = sh.getLastRow() - rows.length;
  if (extra > 0) sh.getRange(rows.length + 1, 1, extra, 5).clearContent();
}

function writeUploadColumn_(ss, tabName, col) {
  if (!col.length) return;
  ss.getSheetByName(tabName).getRange(2, OUTPUT_HEADERS.indexOf('업로드') + 1, col.length, 1).setValues(col);
}

function writeCleanTab_(ss, name, rows) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  const range = sh.getRange(1, 1, rows.length, rows[0].length);
  range.setNumberFormat('@'); // 전화·shop_id 앞자리 0과 날짜 오인 방지
  range.setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sh.setColumnWidths(1, rows[0].length, 120);
  sh.setFrozenRows(1);
  return name;
}
