# PQL 자동화 스크립트

## 개요
구글 드라이브에 업로드된 엑셀 파일을 Google Sheets로 가져와 필요한 컬럼만 남기고 가공하는 자동화 스크립트

## 사용법
1. Google Sheets > 확장 프로그램 > Apps Script에 아래 코드 붙여넣기
2. 스프레드시트 새로고침 후 메뉴 `PQL 자동화` 사용
3. Advanced Drive Service 추가 불필요 (REST API 직접 호출 방식)

## 변경 이력
- 2026-03-09: Drive API v2 → v3 마이그레이션 대응
  - `Drive.Files.insert` → `Drive.Files.create`
  - `title` → `name`
  - `parents: [{id: '...'}]` → `parents: ['...']`
  - Advanced Drive Service(`Drive.Files.create`) 대신 `UrlFetchApp` REST API 직접 호출로 변경
  - 임시 파일 삭제: `Drive.Files.remove` → `DriveApp.getFileById().setTrashed(true)`

---

## 전체 코드

```javascript
/***************************************
 * 설정값
 ***************************************/
const TARGET_FOLDER_ID = '1PjCz9YxLLqGLYOZLffPO97tk7UKEGEaF'; // 리드 파일 폴더 ID
const RAW_SHEET_NAME = 'raw';
const OUTPUT_SHEET_PREFIX = 'clean_';

/**
 * [설정] 최종 결과 시트에 표시할 헤더 순서 (재정렬됨)
 * 1. 요청하신 중요 컬럼들을 맨 앞에 배치
 * 2. 나머지 컬럼들을 뒤에 이어서 배치
 * 3. 삭제된 컬럼: '카테고리', '기본제공 도메인', '주소1', '주소2'(통합됨)
 */
const OUTPUT_HEADERS = [
  // --- [1] 중요 컬럼 (요청하신 순서) ---
  'shop_name',
  'shop_id',
  'mall_id',
  '최근 30일 카페24 주문수(API)',
  '서비스 라벨',
  'Cafe24-회사명',
  '담당자명',
  '쇼핑몰명',
  '담당자전화번호',
  '담당자이메일',
  '대표도메인',
  '주소',

  // --- [2] 나머지 컬럼들 (뒤에 붙임) ---
  'shop_no',
  '플랜',
  '사이트 상태',
  '알파리뷰 상태',
  '알파업셀 상태',
  '알파푸시 상태',
  '최근 30일 카페24 주문수',
  '최근 30일 전체 주문수',
  '설치시점 카페24 주문수(API)',
  '최근 30일 UV(방문자수)',
  '최근 30일 PV(페이지뷰)',
  '임직원 수',
  '회사명',
  '이메일',
  '사업자',
  '고객센터',
  '전화번호',
  '담당자직책',
  '결제담당이메일'
];

/***************************************
 * 메뉴 생성
 ***************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PQL 자동화')
    .addItem('최신 파일 가져오기 & 가공 실행', 'runOneStopProcess')
    .addSeparator()
    .addItem('(수동) 1. 최신 파일 가져오기', 'importLatestDataToRaw')
    .addItem('(수동) 2. 데이터 가공하기', 'createCleanSheetFromRaw')
    .addToUi();
}

/***************************************
 * [메인] 원스톱 실행 함수
 ***************************************/
function runOneStopProcess() {
  const importSuccess = importLatestDataToRaw();
  if (importSuccess) {
    Utilities.sleep(1000);
    createCleanSheetFromRaw();
  }
}

/***************************************
 * [1단계] 드라이브 최신 파일 -> Raw 시트
 ***************************************/
function importLatestDataToRaw() {
  const ui = SpreadsheetApp.getUi();
  try {
    const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    const files = folder.getFiles();
    let latestFile = null;
    let latestTime = 0;

    while (files.hasNext()) {
      const file = files.next();
      const mime = file.getMimeType();
      if (
        mime === MimeType.GOOGLE_SHEETS ||
        mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      ) {
        if (file.getLastUpdated().getTime() > latestTime) {
          latestFile = file;
          latestTime = file.getLastUpdated().getTime();
        }
      }
    }

    if (!latestFile) {
      ui.alert('폴더에 가져올 파일이 없습니다.');
      return false;
    }

    let values = [];
    if (latestFile.getMimeType() === MimeType.GOOGLE_SHEETS) {
      const sourceSheet = SpreadsheetApp.openById(latestFile.getId()).getSheets()[0];
      values = sourceSheet.getDataRange().getValues();
    } else {
      const blob = latestFile.getBlob();
      blob.setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      try {
        // Advanced Service 대신 REST API 직접 호출 (v3 호환 보장)
        const metadata = {
          name: '[Temp] ' + latestFile.getName(),
          parents: [TARGET_FOLDER_ID],
          mimeType: 'application/vnd.google-apps.spreadsheet'
        };
        const boundary = Utilities.getUuid();
        const header = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n';
        const middle = '\r\n--' + boundary + '\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n';
        const ending = '\r\n--' + boundary + '--';

        var payload = [];
        payload = payload.concat(Utilities.newBlob(header + JSON.stringify(metadata) + middle).getBytes());
        payload = payload.concat(blob.getBytes());
        payload = payload.concat(Utilities.newBlob(ending).getBytes());

        var options = {
          method: 'post',
          contentType: 'multipart/related; boundary=' + boundary,
          payload: payload,
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
          muteHttpExceptions: true
        };
        var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', options);
        var tempFileId = JSON.parse(res.getContentText()).id;

        if (!tempFileId) {
          ui.alert('엑셀 변환 실패: ' + res.getContentText());
          return false;
        }

        var tempSs = SpreadsheetApp.openById(tempFileId);
        values = tempSs.getSheets()[0].getDataRange().getValues();
        DriveApp.getFileById(tempFileId).setTrashed(true);
      } catch (e) {
        ui.alert('엑셀 변환 실패: ' + e.message);
        return false;
      }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let rawSheet = ss.getSheetByName(RAW_SHEET_NAME);
    if (!rawSheet) rawSheet = ss.insertSheet(RAW_SHEET_NAME);
    else rawSheet.clear();

    if (values.length > 0) {
      rawSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
      rawSheet.getDataRange().setNumberFormat('@');
    }
    return true;
  } catch (e) {
    ui.alert('오류 발생: ' + e.message);
    return false;
  }
}

/***************************************
 * [2단계] Raw 시트 -> 가공
 ***************************************/
function createCleanSheetFromRaw() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAW_SHEET_NAME);

  if (!rawSheet) throw new Error(`'${RAW_SHEET_NAME}' 시트가 없습니다.`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < 1) throw new Error('데이터가 없습니다.');

  // 데이터 가져오기
  const rawValues = rawSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headerRow = rawValues[0];

  // Raw 시트의 헤더 위치 매핑
  const rawHeaderMap = {};
  headerRow.forEach((name, idx) => {
    if (name) rawHeaderMap[String(name).trim()] = idx;
  });

  // 주요 로직 처리를 위한 인덱스 찾기
  const idxRecentOrder30 = rawHeaderMap['최근 30일 카페24 주문수'];
  const idxUpsellStatus = rawHeaderMap['알파업셀 상태'];
  const idxReviewStatus = rawHeaderMap['알파리뷰 상태'];
  const idxPushStatus = rawHeaderMap['알파푸시 상태'];
  const idxSiteStatus = rawHeaderMap['사이트 상태'];
  const idxManagerName = rawHeaderMap['담당자명'];
  const idxManagerPhone = rawHeaderMap['담당자전화번호'];
  const idxAddr1 = rawHeaderMap['주소1'];
  const idxAddr2 = rawHeaderMap['주소2'];

  // 결과 데이터 저장소
  const outputValues = [];
  outputValues.push(OUTPUT_HEADERS); // 설정한 순서대로 헤더 삽입

  // 데이터 반복 처리
  for (let r = 1; r < rawValues.length; r++) {
    const row = rawValues[r];
    let shouldDelete = false;

    // --- 1. 삭제 조건 체크 (Raw 데이터 기준) ---
    if (idxRecentOrder30 !== undefined) {
      const v = row[idxRecentOrder30];
      if (v === '' || v === null || Number(v) < 100) shouldDelete = true;
    }
    if (!shouldDelete && idxUpsellStatus !== undefined) {
      if (['라이브', '제거중'].includes(String(row[idxUpsellStatus]).trim()))
        shouldDelete = true;
    }
    if (!shouldDelete && idxReviewStatus !== undefined) {
      if (['제거중', '해지완료', '서비스 중단'].includes(String(row[idxReviewStatus]).trim()))
        shouldDelete = true;
    }
    if (!shouldDelete && idxSiteStatus !== undefined) {
      if (['구독종료', '해지완료', '계정활성화'].includes(String(row[idxSiteStatus]).trim()))
        shouldDelete = true;
    }
    if (!shouldDelete && idxManagerName !== undefined) {
      if (String(row[idxManagerName]).trim() === '프로') shouldDelete = true;
    }

    if (shouldDelete) continue;

    // --- 2. 값 가공 ---

    // [라벨링]
    let labelValue = '';
    const vReview =
      idxReviewStatus !== undefined ? String(row[idxReviewStatus]).trim() : '';
    const vUpsell =
      idxUpsellStatus !== undefined ? String(row[idxUpsellStatus]).trim() : '';
    const vPush =
      idxPushStatus !== undefined ? String(row[idxPushStatus]).trim() : '';
    const badStatuses = ['구독없음', '서비스중단', '프로덕트온보딩중', ''];

    const isReviewBad = badStatuses.includes(vReview);
    const isUpsellBad = badStatuses.includes(vUpsell);
    const isPushBad = badStatuses.includes(vPush);

    if (isReviewBad && isUpsellBad && isPushBad) labelValue = 'null';
    else if (vReview === '라이브' && vPush === '라이브')
      labelValue = '알파리뷰, 알파푸시';
    else if (vReview === '라이브') labelValue = '알파리뷰';
    else if (vPush === '라이브') labelValue = '알파푸시';

    // [전화번호 복구]
    let phoneValue = '';
    if (idxManagerPhone !== undefined) {
      let display = String(row[idxManagerPhone] || '').trim();
      if (display !== '') {
        let digits = display.replace(/\D/g, '');
        if (digits.length === 10 && digits.startsWith('10'))
          digits = '0' + digits;
        if (digits.length === 11 && digits.startsWith('010')) {
          display = `010-${digits.slice(3, 7)}-${digits.slice(7)}`;
        }
        phoneValue = display;
      }
    }

    // [주소 통합]
    let mergedAddress = '';
    const addr1 =
      idxAddr1 !== undefined ? String(row[idxAddr1] || '').trim() : '';
    const addr2 =
      idxAddr2 !== undefined ? String(row[idxAddr2] || '').trim() : '';
    mergedAddress = (addr1 + ' ' + addr2).trim();

    // --- 3. 최종 행 구성 (OUTPUT_HEADERS 순서대로 배치) ---
    const newRow = OUTPUT_HEADERS.map((headerName) => {
      // 커스텀 컬럼 처리
      if (headerName === '서비스 라벨') return labelValue;
      if (headerName === '주소') return mergedAddress;
      if (headerName === '담당자전화번호') return phoneValue;

      // 일반 컬럼 처리 (Raw 시트에서 매핑)
      const rawIdx = rawHeaderMap[headerName];
      if (rawIdx !== undefined) {
        return row[rawIdx];
      }
      return ''; // 매핑 안되면 빈칸
    });

    // 전화번호 없는 행 제외
    if (idxManagerPhone !== undefined && phoneValue === '') continue;

    outputValues.push(newRow);
  }

  // 결과 시트 생성
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd_HHmmss'
  );
  const newSheetName = OUTPUT_SHEET_PREFIX + timestamp;
  const newSheet = ss.insertSheet(newSheetName);

  if (outputValues.length > 0) {
    // 텍스트 서식 적용 (0 잘림 방지)
    newSheet
      .getRange(1, 1, outputValues.length, outputValues[0].length)
      .setNumberFormat('@');
    // 값 넣기
    newSheet
      .getRange(1, 1, outputValues.length, outputValues[0].length)
      .setValues(outputValues);

    // 스타일링
    newSheet
      .getRange(1, 1, 1, outputValues[0].length)
      .setFontWeight('bold');
    newSheet.autoResizeColumns(1, outputValues[0].length);
  }

  SpreadsheetApp.getUi().alert(
    `가공 완료!\n- 중요 컬럼 앞으로 정렬 완료\n- 생성된 시트: ${newSheetName}`
  );
}
```
