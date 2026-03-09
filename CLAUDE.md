# PQL 자동화 프로젝트

## 개요
구글 드라이브에 업로드된 엑셀(리드) 파일을 Google Sheets로 가져와 필요한 컬럼만 남기고 가공하는 Google Apps Script 자동화 스크립트

## 기술 스택
- Google Apps Script
- Google Drive API v3 (REST, `DriveApp` + `UrlFetchApp` `files.copy`)
- Google Sheets API

## 주요 파일
- `PQL.md` - 전체 스크립트 코드 및 사용법

## 주요 기능
1. 지정 폴더에서 최신 엑셀 파일 자동 감지
2. 엑셀 → Google Sheets 변환 후 raw 시트에 데이터 적재
3. 필터링 조건에 따라 불필요한 행 제거
4. 서비스 라벨링, 전화번호 포맷팅, 주소 통합 등 데이터 가공
5. 중요 컬럼 우선 배치된 clean 시트 생성

## 트러블슈팅

### Drive API v2 → v3 마이그레이션 오류 (2026-03-09)

**증상:**
```
엑셀 변환 실패 (Drive API v2 확인): 다음 오류로 인해 drive.files.insert API를 호출하지 못했습니다. Bad Request
```

**원인:**
Google이 Apps Script Advanced Drive Service의 기본 버전을 v2에서 v3로 자동 변경함. 기존 코드가 v2 문법(`Drive.Files.insert`, `title`, `parents: [{id}]`)을 사용하고 있어 호환성 오류 발생.

**시도한 해결 방법:**

| 시도 | 내용 | 결과 |
|------|------|------|
| 1차 | v3 문법으로 변경 (`insert`→`create`, `title`→`name`, `parents` 형식 변경) | Bad Request 지속 |
| 2차 | blob 콘텐츠 타입 명시 + mimeType 문자열 직접 지정 + `{fields:'id'}` 옵션 추가 | Bad Request 지속 |
| 3차 (최종) | Advanced Drive Service 완전 제거, `UrlFetchApp`으로 REST API 직접 호출 | 해결 |

**최종 해결:**
- `Drive.Files.insert` / `Drive.Files.create` (Advanced Service) → `DriveApp.createFile` + `files.copy` API
- `Drive.Files.remove` → `DriveApp.getFileById().setTrashed(true)`
- Advanced Drive Service 의존성 완전 제거로 향후 버전 변경 영향 없음

**핵심 코드 (현재 방식):**
```javascript
// 1. DriveApp으로 xlsx 업로드
var tempXlsx = DriveApp.getFolderById(TARGET_FOLDER_ID).createFile(blob);
// 2. files.copy API로 Google Sheets 변환
var copyRes = UrlFetchApp.fetch(
  'https://www.googleapis.com/drive/v3/files/' + tempXlsx.getId() + '/copy?fields=id&supportsAllDrives=true',
  { method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ name: '[Temp]', mimeType: 'application/vnd.google-apps.spreadsheet' }),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
);
```

**교훈:**
- Google Advanced Service는 버전 자동 업그레이드로 예고 없이 깨질 수 있음
- REST API multipart/resumable 업로드보다 `DriveApp` + `files.copy` 분리가 더 안정적
- Advanced Service 제거 시 서비스 목록에서도 삭제 가능

### 엑셀 변환 400 Bad Request (2026-03-09)

**증상:**
```
엑셀 변환 실패: {"error":{"code":400,"message":"Bad Request"...}}
```

**원인:**
1. REST API multipart/resumable 업로드 + 변환 시 대용량/암호화 파일에서 400 발생
2. 비밀번호(열기 암호)가 걸린 xlsx는 Google Drive API가 변환 불가

**시도한 해결 방법:**

| 시도 | 내용 | 결과 |
|------|------|------|
| 1차 | `uploadType=multipart` REST API 직접 호출 | 400 Bad Request |
| 2차 | `uploadType=resumable` 2단계 업로드 | 400 Bad Request |
| 3차 (최종) | `DriveApp.createFile` + `files.copy` 분리 방식 | 해결 |

**최종 해결:**
- `DriveApp.createFile(blob)`으로 xlsx를 Drive에 업로드 (REST API 불필요)
- `files.copy` API로 xlsx → Google Sheets 변환 (mimeType 지정)
- 비밀번호가 걸린 xlsx는 업로드 전 비밀번호 제거 필요

**교훈:**
- REST API multipart/resumable 업로드 + 변환은 파일 조건에 따라 불안정
- `DriveApp` 업로드 + `files.copy` 변환 분리가 더 안정적
- 비밀번호(열기 암호)가 걸린 xlsx는 Google API로 변환 불가 (시트 보호는 가능)

## 최근 변경사항

### 컬럼명 변경 및 플랫폼 컬럼 추가 (2026-03-09)
- `카페24` → `플랫폼`으로 컬럼명 변경 (주문수, 회사명 등)
- 중요 컬럼에 `플랫폼` 추가 (mall_id 다음)
- `Cafe24-회사명` → `회사명` 통합, 중복 `회사명` 컬럼 제거
- 필터링 로직 참조 컬럼명도 동기화

### 엑셀 변환 방식 변경 (2026-03-09)
- REST API multipart/resumable 업로드 → `DriveApp.createFile` + `files.copy` 분리 방식
- 대용량 파일, 공유 드라이브 등 다양한 환경에서 안정적 동작

### createCleanSheetFromRaw() 성능 최적화 (2026-03-09)
- `autoResizeColumns(1, 28)` 제거 → `setColumnWidths(1, colCount, 120)` 1회 호출로 대체 (28회 → 1회)
- `setNumberFormat('@')` 제거 → JS에서 `String()` 변환으로 대체 (API 호출 1회 감소)
- 필터링 조건에 `Set.has()` 사용 (`Array.includes()` O(n) → O(1))
- String 변환 1회만 수행 후 재사용, OUTPUT_HEADERS 매핑 함수 사전 생성
- Sheets API 호출: 기존 ~33회 → 최적화 후 3회 (setValues, setFontWeight, setColumnWidths)

## 폴더 구조
```
PQL_auto/
  CLAUDE.md    - 프로젝트 문서 (이 파일)
  PQL.md       - 스크립트 코드 및 사용법
```
