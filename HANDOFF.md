# HANDOFF

## 2026-10-27 — 11월 CSV 업로드 확인, 구독 CSV 소비처·적재 시점 정리 (코드 변경 없음)

- 11월분 `all_subscription_1101.csv`가 10/27 09:26 KST에 올라왔고, 사용자가 11월 `PQL 생성`을 이미 실행했다. crema BQ 적재는 11/2 06:00, alphareview-ref 반영은 11/3 06:00.
- 11월 PQL은 PR #3 배포 전 코드로 돌았다(라이브 pull 비교로 확인). 영향 없음 — 그 시점 업셀 파이프라인(11) 딜은 `[QA] 업셀 미리보기` 1건이고 shop_id가 비어 있었다. 이후 PR #3을 `clasp push -f`로 배포했고 다시 pull해 라이브 = 로컬을 확인했다(테스트 59개 통과). 10/21 항목의 "미배포"·"ARCHITECTURE ①·⑥ 미반영"은 모두 해소(①·⑥은 8e86069).
- 구독 CSV 소비처 정리: 직접 = PQL_auto·crema 로더(매월 2일 06:00 → BQ `leadgen.subscriptions`, 이번 달 파일이 아니면 실패). BQ 경유 = SQL v3(새 문의마다 최신 스냅샷 조회, 캐시 없음)·alphareview-ref(매월 3일 06:00)·upsell-lead-gen(매일 08:00, 그쪽 HANDOFF 기준 스케줄러 미가동)·crema 데일리 잡. AGENTS.md 교차 영향·Golden Principle 2와 ARCHITECTURE 제약의 "alphareview-ref가 폴더를 읽는다"를 정정했다(BQ를 읽는다).
- 결정: crema 적재를 업로드 시점 트리거·재시도 cron으로 바꾸지 않는다 — 사용자가 말일 전에 다음 달 CSV를 올리므로 불필요. 업로드가 2일 이후로 늦어지는 일이 생기면 `0 6 2-10 * *` 재시도 안을 다시 꺼낸다.
- 다음: 11월 실행 결과 확인(소요 시간·업로드 이력·`남음` 여부, 9/30 항목 2번), `shop_id 매핑` 대기 행 판정(대기 행은 매 실행 다시 만들어지므로 지워도 되지만 `거절됨` 행은 지우지 말 것).

## 2026-10-21 — 업셀 파이프라인(11) 딜 shop_id도 중복 제외 (PR #3, 미배포)

- `fetchPipedrive_`가 업셀 전용 pipeline 11 `Upsell Lead Gen` 딜의 숫자 shop_id를 제외 집합(사유 `deal`)에 합친다(`UPSELL_PIPELINE_ID`). 역매핑·deal list에는 넣지 않는다. 처음엔 Lead Gen Auto(10) 기준으로 만들었으나 업셀 딜을 별도 파이프라인에 만들기로 바뀌어(upsell-lead-gen D9) 10/22 11로 교체했다. 11은 10/21 생성이라 아직 딜이 거의 없다(upsell-lead-gen이 shop_id를 채운 딜을 만들기 시작하면 늘어남). ID 근거 = `Work/upsell-lead-gen/docs/research/2026-10-21-pipedrive-upsell-ids.md`.
- 운영 반영은 `clasp push -f` 필요(미실행, 사용자 승인 대기). `ARCHITECTURE.md` ①·⑥ 설명은 아직 Sales만 적혀 있다.

## 2026-09-30 — PQL 파이프라인 재구축 완료, 새 시트에서 라이브 가동

**한 줄 상태**: 새 시트 `PQL_auto`에서 CSV 업로드 → `PQL 생성` 한 번으로 클렌징·Sales 딜 제외·Pipedrive 딜 업로드·업로드 이력 누적까지 끝난다. 2026-10 PQL 194건 업로드 완료. 남은 건 매핑 탭 사람 확인과 다음 달 첫 정기 실행 확인.

**이번 세션 결정**
- 원천 CSV를 스크립트가 직접 스트리밍 파싱, raw 탭 폐지 (설계: `docs/design-docs/2026-09-30-pql-pipeline-refactor-design.md`)
- 타겟: 업셀 cafe24·150+ / 푸시 cafe24·500+·푸시 라이브 아님(무료 포함) / 리뷰 1,000+(아임웹 포함). 세 제품 모두 라이브면 올리지 않는다
- Sales 딜 제외는 매 실행 Pipedrive 전체 조회(증분 동기화 안 함 — 단계·shop_id 변경·삭제를 놓치지 않기 위해)
- 업로드는 `PQL 생성` 때 자동, xlsx는 만들지 않음(수동 가져오기와 중복 방지). 소유자 한서연·Sales/컨택전·세일즈티어(플랜 우선, `-`면 주문수 구간)
- 이전 시트 `PQL_cleansing_auto`는 편집 이력이 무거워(같은 쓰기 61.5초 vs 새 시트 8.3초) 사용 중지. 스크립트는 "새 시트 안내" 메뉴만 남김
- `업로드 이력` 탭: 월(원천 파일 기준)·업로드일·타겟·세일즈티어별 업로드 수 + 전체 합계 누적

**다음 세션 우선 작업**
1. `shop_id 매핑` 탭 확인 필요 178행 — AE 승인/거절 (승인분은 다음 실행 시작 때 Pipedrive 반영)
2. 2026-11 PQL 첫 정기 실행 결과 확인(소요 시간, 업로드 이력 누적, `남음` 여부)
3. 보류: Pipedrive 레이블 이름이 바뀌면 라벨 없는 딜이 조용히 생김(`resolveUploadIds_`는 소유자·단계만 검사) — 필요 시 라벨 누락도 멈추게
4. 테스트로 올린 아코이 딜(주문 105건, 새 기준 150 미달)은 사용자 판단으로 유지 중

**핵심 경로**: 코드 `src/` · 배포 `clasp push -f`(scriptId `1y8duYMO…`) · 사용법 `PQL.md` · 흐름 `ARCHITECTURE.md` · 트러블슈팅 `CLAUDE.md`
