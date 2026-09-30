# PQL 자동화 - Plans

> 최종 업데이트: 2026-09-30

## 현재 우선순위

1. **라이브 첫 실행 모니터링** -- 2026-09-30 재구축 후 첫 라이브 실행에서 역매핑 자동 반영(약 49건)·딜 자동 업로드(약 200곳) 결과를 Pipedrive에서 대조한다.
2. **`shop_id 매핑` 탭 대기 행 정리** -- 확인 필요 후보(약 50딜)를 AE가 승인/거절해 딜 의심 몰을 줄인다.
3. **오래된 clean 탭 정리** -- 실행마다 `clean_{시각}` 탭이 쌓인다.

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

### Phase 3: 운영 개선 (2026-09-30 대부분 완료)

- 완료: CSV 직접 처리, Sales 딜 자동 제외(중복 리드 감지 대체), shop_id 역매핑, Pipedrive 자동 업로드
- 남음: 자동 실행 트리거(폴더 감시), 오래된 clean 탭 정리

### Phase 4: 고도화 (계획)

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
|---|---|---|---|
| clean 탭 누적 | 중간 | 실행마다 `clean_{시각}` 탭 생성 | 30일 지난 탭 자동 삭제 |
| 업로드 시간 예산 | 중간 | 대상이 수백 곳을 넘으면 5분 예산에서 끊겨 다시 눌러야 함 | 실측 후 `UPLOAD_BATCH`·예산 조정 |
| 원천 50MB 근접 | 낮음 | 분할 다운로드로 대응됨. 한 행이 20MB를 넘는 경우만 멈춤 | - |
| 토큰 공유 | 낮음 | 스크립트 편집 권한자는 Script Properties의 Pipedrive 토큰을 볼 수 있음 | 전용 토큰 발급 검토 |

해소: 필터 하드코딩(→ `src/Config.js`), 에러 로그 없음(→ 요약 창·clean 탭 업로드 열), 테스트 부재(→ `test/`), 폴더 ID 하드코딩(→ Config).
