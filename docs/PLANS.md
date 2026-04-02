# PQL 자동화 - Plans

> 최종 업데이트: 2026-04-02

## 현재 우선순위

1. **필터링 조건 검증** -- 현재 필터링 기준(주문수 100, 상태값 등)이 최신 비즈니스 룰과 일치하는지 세일즈팀과 확인한다.
2. **자동 실행 트리거 검토** -- 폴더에 새 파일 업로드 시 자동 실행하는 Apps Script 트리거 설정 가능 여부를 확인한다.
3. **이전 clean 시트 정리** -- 오래된 clean_{timestamp} 시트가 누적되므로 자동 정리 로직을 추가한다.

## 로드맵

### Phase 1: 기반 구축 (완료)

- importLatestDataToRaw(): 폴더 최신 파일 -> raw 시트
- createCleanSheetFromRaw(): 필터링 + 가공 + 컬럼 재배치
- runOneStopProcess(): 원스톱 실행
- UI 메뉴 등록

### Phase 2: 안정화 (완료)

- Drive API v2 -> v3 마이그레이션 대응 (DriveApp + files.copy)
- 성능 최적화 (Sheets API 33회 -> 3회)
- 컬럼명 변경 반영 (카페24 -> 플랫폼)

### Phase 3: 운영 개선 (계획)

- 자동 실행 트리거 (폴더 감시)
- 오래된 clean 시트 자동 정리
- 중복 리드 감지 (이전 시트 대비)

### Phase 4: 고도화 (계획)

- Pipedrive 딜 자동 생성 연동 (clean 리드 -> Pipedrive 딜)
- 아웃바운드 결과 추적 (콜 성공/실패 기록)

## 기술적 의사결정

### 왜 Apps Script인가?

SDR(비개발자)이 Google Sheets 메뉴에서 직접 실행할 수 있어야 한다. Python 스크립트는 별도 실행 환경이 필요하지만, Apps Script는 스프레드시트에 내장되어 원클릭으로 동작한다.

### 왜 Advanced Drive Service를 제거했는가?

Google이 Apps Script Advanced Drive Service의 기본 버전을 v2에서 v3로 예고 없이 변경하여 기존 코드가 깨졌다. `Drive.Files.insert` -> `Drive.Files.create` 마이그레이션도 400 Bad Request로 실패했다. DriveApp + UrlFetchApp REST API 직접 호출로 전환하여 향후 버전 변경 영향을 완전히 차단했다.

### 왜 headerMappers 사전 생성인가?

데이터 행 반복(수천 행) 내에서 매번 if/else로 컬럼별 처리를 분기하면 비효율적이다. OUTPUT_HEADERS에 대한 매핑 함수 배열을 사전에 생성하고, 반복문에서는 `fn(row, ctx)`만 호출하여 루프 내 분기를 제거했다.

## 기술 부채

| 항목 | 심각도 | 설명 | 해결 방안 |
|------|--------|------|-----------|
| clean 시트 누적 | 중간 | 실행할 때마다 clean_{timestamp} 시트가 생성되어 누적됨 | 30일 이상 오래된 시트 자동 삭제 |
| 필터링 조건 하드코딩 | 중간 | upsellDeleteSet, reviewDeleteSet 등이 코드에 하드코딩 | 설정 시트 또는 상수 영역으로 분리 |
| 에러 메시지 | 낮음 | UI alert으로만 표시, 로그 없음 | 별도 log 시트에 실행 이력 기록 |
| 테스트 부재 | 낮음 | 유닛 테스트 없음 | Apps Script 환경에서는 수동 검증에 의존 |
| TARGET_FOLDER_ID 하드코딩 | 낮음 | 폴더 ID가 코드에 직접 기재 | PropertiesService로 설정 외부화 |
