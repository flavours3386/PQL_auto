# PQL 자동화 - Agent Guide

> 이 파일은 목차다. 상세 내용은 각 링크를 따라가라.

## Quick Start

1. `node --test 'test/*.test.js'` — 순수 로직 테스트
2. `clasp push -f` — 라이브 배포 (`.clasp.json` rootDir: src). 토큰 만료 시 `! clasp login`
3. 시트 메뉴 `PQL 자동화 > PQL 생성` (사람이 실행)

라이브를 건드리지 않고 확인하려면 빈 스프레드시트에 `clasp create-script --type sheets`로 테스트 스크립트를 만들고, `AUTO_APPLY`·`AUTO_UPLOAD`를 `false`로 바꾼 사본을 올린다(생성 직후 `appsscript.json`이 기본값으로 덮어써지므로 저장소 것으로 되돌릴 것).

## Golden Principles

1. **Advanced Service를 쓰지 않는다** — Drive v2→v3 자동 전환으로 깨진 전력이 있다. Drive·Pipedrive 모두 UrlFetch REST.
2. **원천 폴더 `05. PQL`에는 쓰지 않는다** — crema BQ 로더·alphareview-ref가 같은 폴더에서 최신 CSV를 읽는다.
3. **순수 로직은 Core.js에, 테스트 먼저** — Apps Script 서비스를 쓰는 코드는 Io.js·Main.js로 격리한다.
4. **시트 쓰기는 계산이 끝난 뒤, Pipedrive 업로드는 시트 쓰기 뒤** — 중간 실패로 반쯤 만든 탭이 남지 않고, 업로드가 끊겨도 다음 실행이 이어받는다.
5. **Pipedrive 쓰기는 조건부** — 숫자 shop_id는 덮어쓰지 않는다. 자동 반영 100건·업로드 500곳 상한을 넘으면 멈춘다. 필드 키·단계·소유자는 `src/Config.js`에만 둔다.
6. **예시·검증은 이번 결과 안의 몰로** — 필터 대상 몰을 견본으로 쓰지 않는다.

## Docs Map

| 문서 | 용도 |
|---|---|
| [CLAUDE.md](CLAUDE.md) | 개요, 명령어, 최근 변경, 트러블슈팅 |
| [ARCHITECTURE.md](ARCHITECTURE.md) | 흐름·규칙·Pipedrive 쓰기·제약 |
| [PQL.md](PQL.md) | 사용법 (월간 절차·요약 문구·설정·배포) |
| [CHANGELOG.md](CHANGELOG.md) | 지난 세대 변경 |
| [docs/design-docs/](docs/design-docs/) | 설계(spec) |
| [docs/exec-plans/](docs/exec-plans/) | 구현 계획 |
| [docs/PLANS.md](docs/PLANS.md) | 우선순위·기술 부채 |

## 교차 영향

| 공유 자원 | 함께 쓰는 프로젝트 | 주의 |
|---|---|---|
| Drive `05. PQL` (`1PjCz9YxLLqGLYOZLffPO97tk7UKEGEaF`) | crema `bq_load_customers.py`, alphareview-ref | CSV 헤더가 바뀌면 PQL은 필수 열 누락으로 멈춘다 |
| Pipedrive 공유 토큰 | pipedrive_auto, team-agent, pipedrive-mcp, crema 등 | PQL이 Sales 딜 생성·shop_id 쓰기를 한다. shop_id·세일즈티어 등 필드 키, 컨택전 단계, 한서연 소유자 변경 시 `src/Config.js` 수정 |

정본은 워크스페이스 [ARCHITECTURE.md](../../ARCHITECTURE.md) 교차 영향 표.
