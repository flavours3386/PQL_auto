# PQL 자동화 프로젝트

## 개요
구글 드라이브에 업로드된 엑셀(리드) 파일을 Google Sheets로 가져와 필요한 컬럼만 남기고 가공하는 Google Apps Script 자동화 스크립트

## 기술 스택
- Google Apps Script
- Google Drive API v3 (REST, UrlFetchApp 직접 호출)
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
- `Drive.Files.insert` / `Drive.Files.create` (Advanced Service) → `UrlFetchApp.fetch` (REST API 직접 호출)
- `Drive.Files.remove` → `DriveApp.getFileById().setTrashed(true)`
- Advanced Drive Service 의존성 완전 제거로 향후 버전 변경 영향 없음

**핵심 코드 (변경 부분):**
```javascript
// REST API multipart 업로드로 엑셀 → Google Sheets 변환
const metadata = {
  name: '[Temp] ' + latestFile.getName(),
  parents: [TARGET_FOLDER_ID],
  mimeType: 'application/vnd.google-apps.spreadsheet'
};
const boundary = Utilities.getUuid();
// multipart body 구성 후 UrlFetchApp.fetch로 직접 호출
var res = UrlFetchApp.fetch(
  'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id',
  options
);
```

**교훈:**
- Google Advanced Service는 버전 자동 업그레이드로 예고 없이 깨질 수 있음
- 안정성이 중요한 스크립트는 `UrlFetchApp` + REST API 직접 호출이 더 안전
- Advanced Service 제거 시 서비스 목록에서도 삭제 가능

## 폴더 구조
```
PQL_auto/
  CLAUDE.md    - 프로젝트 문서 (이 파일)
  PQL.md       - 스크립트 코드 및 사용법
```
