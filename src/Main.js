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
    '<p><a href="' + esc(xlsx.getUrl()) + '" target="_blank">' + esc(xlsx.getName()) + ' 열기</a> (내 드라이브 · 열린 화면에서 다운로드)</p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(480).setHeight(520), 'PQL 생성 완료');
}
