# CHANGELOG

현재 세대는 [CLAUDE.md](CLAUDE.md) "최근 변경사항". 이 파일은 지난 세대 기록이다. 아래 항목의 코드(raw 탭 적재·xlsx 변환·`createCleanSheetFromRaw`)는 2026-09-30 재구축으로 대체됐다.

## 2026-06-30 담당자명 '프로' 필터 기준 변경
- 과거: 담당자명이 `프로`인 행 전체 제외 (프로 담당자 연락처가 고객사에 일괄 등록돼 있었기 때문)
- 변경: 연락처 정상화 완료 → `프로` 중 핸드폰(010)이 아닌 번호(070/지역번호 등)만 제외, 010이면 유지
- 판정은 "010 화이트리스트" 방식(`/^010/`)
- 부수효과: 프로 + 핸드폰 행이 살아남아 clean 시트 출력 행 수 증가 (의도된 변화)

## 2026-03-09 컬럼명 변경 및 플랫폼 컬럼 추가
- `카페24` → `플랫폼`으로 컬럼명 변경 (주문수, 회사명 등)
- 중요 컬럼에 `플랫폼` 추가 (mall_id 다음)
- `Cafe24-회사명` → `회사명` 통합, 중복 `회사명` 컬럼 제거

## 2026-03-09 엑셀 변환 방식 변경
- REST API multipart/resumable 업로드 → `DriveApp.createFile` + `files.copy` 분리 방식

## 2026-03-09 createCleanSheetFromRaw() 성능 최적화
- `autoResizeColumns` 28회 → `setColumnWidths` 1회, `setNumberFormat('@')` 제거, 필터링 `Set.has()`
- Sheets API 호출 ~33회 → 3회
