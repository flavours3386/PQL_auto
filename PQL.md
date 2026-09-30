# PQL 자동화 사용법

코드는 `src/`에 있다(시트 `PQL_auto`의 바운드 스크립트, 2026-09-30 이전 시트 `PQL_cleansing_auto`에서 이전 — 이전 시트는 쓰지 않는다). 이 문서에는 코드를 두지 않는다. 규칙과 흐름은 [ARCHITECTURE.md](ARCHITECTURE.md).

## 월간 절차

1. 전사 구독 시스템에서 CSV를 내보내 `all_subscription_MMDD.csv`로 이름을 바꾸고 Drive `05. PQL` 폴더에 올린다
2. PQL 시트 메뉴 `PQL 자동화 > PQL 생성`
3. 요약 창을 확인한다: 원천 파일명·단계별 탈락·결과·Pipedrive 업로드·승인 대기
4. `shop_id 매핑` 탭의 `대기` 행에서 `판정` 열을 `승인`/`거절`로 고른다. `승인된 shop_id 반영`을 누르거나 다음 `PQL 생성` 때 자동 반영된다

## 산출물

| 위치 | 내용 |
|---|---|
| Pipedrive 딜 | 결과 몰마다 딜·담당자·조직 생성. 소유자 한서연, Sales/컨택전, 라벨=서비스 라벨, shop_id·상점아이디·호스팅사·월 주문 수·쇼핑몰명·URL·세일즈티어 |
| `clean_{시각}` 탭 | 결과 몰. `타겟`·`서비스 라벨`·`딜 의심`·`업로드`(딜 ID·실패 사유) 열 |
| `deal list` 탭 | Sales 파이프라인 딜 전체(실행마다 갱신) |
| `shop_id 매핑` 탭 | shop_id가 빈칸·텍스트인 Sales 딜의 후보 shop과 반영 이력 (지우지 말 것 — 승인·거절 기록) |
| `업로드 이력` 탭 | 실행마다 월(원천 파일 기준 PQL 월)·업로드일·타겟·세일즈티어별 업로드 수와 전체 합계를 누적. PQL 추세용 (지우지 말 것) |

## 요약 창에 이런 문구가 뜨면

| 문구 | 할 일 |
|---|---|
| `남음 N (다시 누르면 이어서)` | 5분 예산을 넘겨 업로드를 멈춘 것. `PQL 생성`을 다시 누르면 남은 곳만 올라간다(올라간 곳은 딜이 생겨 자동 제외) |
| `대상 N곳이 500곳 초과` | 원천 CSV가 이상할 수 있다. 확인 후 `src/Config.js`의 `UPLOAD_MAX` 조정 |
| `자동 반영을 멈추고 전부 대기로` | 높은 확신 역매핑이 100건을 넘음. 매핑 탭에서 확인 |
| `CSV에 필수 열이 없습니다: …` | 상류 CSV 헤더가 바뀜. 열 이름 확인 |
| `[Pipedrive 조회] …` | 토큰·네트워크 문제. clean 탭은 만들지 않는다 |
| `deal list 갱신 실패: …` | 참고용 탭만 못 쓴 것. 업로드는 정상, 다음 실행에서 다시 쓴다 |
| `업로드 이력 기록 실패: …` | 업로드는 정상. `업로드 이력` 탭에 이번 실행 건수가 빠졌으니 clean 탭 업로드 열로 확인 |
| `[시트 쓰기] … 스프레드시트 서비스가 타임아웃` | 문서가 무거워진 신호. 이전 시트에서 이 문제로 새 시트로 옮겼다(CLAUDE.md 트러블슈팅) |

## 설정

`src/Config.js`에서만 바꾼다: 주문 기준(`MIN_ORDERS` 100·`UPSELL_MIN_ORDERS` 150·`PUSH_MIN_ORDERS` 500·`REVIEW_MIN_ORDERS` 1000), 타겟 규칙(`TARGETS`), 소유자·단계(`DEAL_OWNER`·`DEAL_STAGE`), `AUTO_UPLOAD`·`UPLOAD_MAX`·`UPLOAD_TIME_BUDGET_SEC`, `AUTO_APPLY`·`AUTO_APPLY_MAX`, Pipedrive 필드 키·세일즈티어 구간.

Pipedrive 토큰은 첫 실행 때 입력창으로 받는다. 바꾸려면 Apps Script 편집기 > 프로젝트 설정 > 스크립트 속성 `PIPEDRIVE_API_TOKEN`.

## 배포

```bash
node --test 'test/*.test.js'   # 순수 로직 테스트
clasp push -f                   # 루트 .clasp.json (rootDir: src)
clasp pull                      # 원격 = 로컬 확인 (git status로 변경 없음)
```

clasp 토큰이 만료되면(`invalid_grant`) `! clasp login`으로 다시 로그인한다.
