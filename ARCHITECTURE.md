# PQL 자동화 - Architecture

## 시스템 개요

Google Drive 특정 폴더에 업로드된 PQL 리드 엑셀 파일을 Google Apps Script로 가져와서 불필요한 행을 필터링하고, 서비스 라벨링/전화번호 포맷팅/주소 통합 등 가공을 수행한 뒤 clean 시트를 생성하는 자동화 스크립트.

## 데이터 흐름

```
Google Drive 폴더 (1PjCz9YxLLqGLYOZLffPO97tk7UKEGEaF)
  |
  v
[importLatestDataToRaw()]
  폴더에서 최신 엑셀/시트 파일 탐색 (최종 수정일 기준)
  |
  +-- Google Sheets 파일? --> SpreadsheetApp.openById() --> 데이터 읽기
  +-- XLSX 파일? --> DriveApp.createFile(blob) --> files.copy API (변환) --> 데이터 읽기
  |
  v
  raw 시트에 전체 데이터 적재
  |
  v
[createCleanSheetFromRaw()]
  |
  +-- [1단계] 행 필터링 (삭제 조건)
  |     - 최근 30일 주문수 < 100 또는 빈값
  |     - 알파업셀 상태: 라이브, 제거중
  |     - 알파리뷰 상태: 제거중, 해지완료, 서비스 중단
  |     - 사이트 상태: 구독종료, 해지완료, 계정활성화
  |     - 담당자명: 프로 + 전화번호가 핸드폰(010) 아님 (070/지역번호 등만 제외, 010이면 유지)
  |     - 담당자전화번호: 빈값
  |
  +-- [2단계] 값 가공
  |     - 전화번호 포맷팅 (010-XXXX-XXXX)
  |     - 서비스 라벨링 (알파리뷰, 알파푸시, null)
  |     - 주소 통합 (주소1 + 주소2)
  |
  +-- [3단계] 컬럼 재배치
  |     - OUTPUT_HEADERS 순서대로 28개 컬럼 배치
  |     - 중요 컬럼 13개가 앞쪽
  |
  v
  clean_{timestamp} 시트 생성
```

## 처리 단계 상세

### 1단계: 파일 가져오기 (importLatestDataToRaw)

| 단계 | 처리 | 비고 |
|------|------|------|
| 폴더 탐색 | DriveApp.getFolderById -> files 순회 | Google Sheets 또는 XLSX만 대상 |
| 최신 파일 선택 | getLastUpdated().getTime() 비교 | 가장 최근 수정된 파일 |
| XLSX 변환 | DriveApp.createFile + files.copy API | mimeType: application/vnd.google-apps.spreadsheet |
| 데이터 읽기 | getDataRange().getValues() | 첫 번째 시트만 |
| raw 적재 | setValues() + setNumberFormat('@') | 기존 데이터 clear 후 덮어쓰기 |
| 임시 파일 정리 | setTrashed(true) | XLSX 원본 + 변환 파일 모두 삭제 |

### 2단계: 데이터 가공 (createCleanSheetFromRaw)

| 처리 | 조건/로직 | 성능 최적화 |
|------|-----------|-------------|
| 행 삭제 | 6개 필터 조건 (주문수, 상태, 담당자, 전화번호) | Set.has() O(1) 검색 |
| 전화번호 포맷 | 10자리(10으로 시작) -> 0 앞에 추가, 11자리 -> 010-XXXX-XXXX | 정규식 + 조건 분기 |
| 서비스 라벨 | 리뷰+푸시=양쪽, 리뷰만, 푸시만, 전부 bad=null | badStatusSet 사전 정의 |
| 주소 통합 | 주소1 + " " + 주소2, trim | |
| 컬럼 매핑 | headerMappers 사전 생성 (루프 내 if 제거) | 매핑 함수 배열 |

### Sheets API 호출 최적화

| 호출 | 용도 | 횟수 |
|------|------|------|
| setValues | 전체 데이터 쓰기 | 1회 |
| setFontWeight | 헤더 볼드 | 1회 |
| setColumnWidths | 고정 너비 120px | 1회 |
| **합계** | | **3회** |

기존 ~33회(autoResizeColumns 28회 + setNumberFormat 등)에서 3회로 최적화됨.

## 에러 처리 전략

| 단계 | 에러 유형 | 처리 방식 |
|------|-----------|-----------|
| 파일 탐색 | 폴더 접근 실패, 파일 없음 | UI alert 표시 + return false |
| XLSX 변환 | files.copy API 400 Bad Request | UI alert + 임시 파일 정리 + return false |
| 비밀번호 엑셀 | Google API 변환 불가 | 사전에 비밀번호 제거 필요 (수동) |
| raw 시트 없음 | createCleanSheetFromRaw 호출 시 | Error throw |
| 데이터 없음 | raw 시트 빈 상태 | Error throw |

## 설정 상수

| 상수 | 값 | 설명 |
|------|-----|------|
| TARGET_FOLDER_ID | 1PjCz9YxLLqGLYOZLffPO97tk7UKEGEaF | 리드 파일 업로드 폴더 |
| RAW_SHEET_NAME | raw | 원본 데이터 적재 시트 |
| OUTPUT_SHEET_PREFIX | clean_ | 가공 결과 시트 접두사 |
| OUTPUT_HEADERS | 28개 컬럼 | 중요 13개 + 나머지 15개 |

## 제약사항

- **Google Apps Script(ES5)**: 최신 JS 문법 사용에 제한이 있다. const/let은 사용 가능하나 async/await, optional chaining 등은 불가.
- **비밀번호 엑셀**: 열기 암호가 걸린 XLSX는 Google Drive API로 변환 불가. 업로드 전 비밀번호 제거 필요.
- **단일 시트만 처리**: 엑셀의 첫 번째 시트만 읽는다. 멀티 시트 파일은 첫 시트 기준.
- **실행 시간 제한**: Apps Script 실행 시간은 6분(무료) / 30분(Workspace)으로 제한된다. 대용량 데이터 시 주의.
