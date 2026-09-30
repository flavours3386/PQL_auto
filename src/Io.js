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

// 여러 요청을 동시에 보낸다. 429(요청 과다)는 하나씩 재시도한다 (429는 처리되지 않은 요청이라 다시 보내도 중복되지 않는다).
function pdFetchAll_(token, reqs) {
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
        return { ok: true, body: pdRequest_(token, reqs[k].method, reqs[k].path, reqs[k].body) };
      } catch (e) {
        return { ok: false, error: e.message.slice(0, 120) };
      }
    }
    if (code >= 300) return { ok: false, error: code + ' ' + res.getContentText().slice(0, 120) };
    return { ok: true, body: JSON.parse(res.getContentText()) };
  });
}

// UPLOAD_BATCH곳씩 조직(같은 이름 재사용, 없으면 생성) → 담당자 → 딜 순으로 만든다. deadlineMs가 지나면 남은 곳은 skipped.
function pdCreateDeals_(token, items, deadlineMs) {
  const results = [];
  const orgIds = {}; // 이번 실행 안에서 조직 이름 → id
  for (let i = 0; i < items.length; i += UPLOAD_BATCH) {
    if (Date.now() > deadlineMs) {
      for (let j = i; j < items.length; j++) results.push({ skipped: true });
      break;
    }
    const batch = items.slice(i, i + UPLOAD_BATCH);
    const orgOf = {};
    batch.forEach(function (it) { orgOf[it.org.name] = it.org; });
    const names = Object.keys(orgOf).filter(function (n) { return !(n in orgIds); });
    const found = pdFetchAll_(token, names.map(function (n) {
      return { method: 'get', path: '/api/v2/organizations/search?term=' + encodeURIComponent(n) + '&fields=name&exact_match=true&limit=1' };
    }));
    const missing = [];
    names.forEach(function (n, k) {
      const items0 = found[k].ok && found[k].body.data && found[k].body.data.items;
      if (items0 && items0.length) orgIds[n] = items0[0].item.id;
      else missing.push(n);
    });
    const created = pdFetchAll_(token, missing.map(function (n) { return { method: 'post', path: '/api/v2/organizations', body: orgPayload_(orgOf[n]) }; }));
    missing.forEach(function (n, k) { if (created[k].ok) orgIds[n] = created[k].body.data.id; });

    const persons = pdFetchAll_(token, batch.map(function (it) {
      return { method: 'post', path: '/api/v2/persons', body: personPayload_(it.person, orgIds[it.org.name]) };
    }));
    const deals = pdFetchAll_(token, batch.map(function (it, k) {
      return { method: 'post', path: '/api/v2/deals', body: dealPayload_(it, persons[k].ok ? persons[k].body.data.id : null, orgIds[it.org.name]) };
    }));
    batch.forEach(function (it, k) {
      if (!deals[k].ok) {
        results.push({ error: deals[k].error });
        return;
      }
      const warn = [];
      if (!orgIds[it.org.name]) warn.push('조직 생성 실패');
      if (!persons[k].ok) warn.push('담당자 생성 실패');
      results.push({ id: deals[k].body.data.id, warn: warn.join(', ') });
    });
  }
  return results;
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

function writeUploadColumn_(ss, tabName, col) {
  if (!col.length) return;
  ss.getSheetByName(tabName).getRange(2, OUTPUT_HEADERS.indexOf('업로드') + 1, col.length, 1).setValues(col);
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
