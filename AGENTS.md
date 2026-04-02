# PQL 자동화 - Agent Guide

> 이 파일은 목차다. 상세 내용은 각 링크를 따라가라.

## Quick Start

1. Google Sheets > 확장 프로그램 > Apps Script에 PQL.md의 코드 붙여넣기
2. 스프레드시트 새로고침 후 메뉴 `PQL 자동화` 클릭
3. `최신 파일 가져오기 & 가공 실행` 선택 (원스톱)

```
수동 단계별 실행:
  (수동) 1. 최신 파일 가져오기  -> importLatestDataToRaw()
  (수동) 2. 데이터 가공하기     -> createCleanSheetFromRaw()
```

## Golden Principles

1. **Advanced Drive Service를 사용하지 않는다** -- Google이 v2->v3 자동 업그레이드로 예고 없이 깨뜨린 전력이 있다. DriveApp + UrlFetchApp REST API 직접 호출 방식을 유지한다.
2. **OUTPUT_HEADERS 순서가 곧 비즈니스 우선순위다** -- 중요 컬럼(shop_name, shop_id, mall_id, 플랫폼, 주문수, 서비스 라벨)이 맨 앞에 배치된다. 순서 변경 시 사용자 워크플로에 직접 영향.
3. **필터링 조건은 Set.has()를 사용한다** -- O(n)인 Array.includes() 대신 O(1)인 Set.has()로 성능을 보장한다. 새 필터 조건 추가 시 이 패턴을 따를 것.
4. **Sheets API 호출을 최소화한다** -- 현재 3회(setValues, setFontWeight, setColumnWidths)로 최적화되어 있다. autoResizeColumns, setNumberFormat 등 호출을 추가하지 말 것.
5. **pipedrive_auto, pipedrive-dashboard와 Pipedrive 인스턴스를 공유한다** -- 직접적 API 연동은 없지만, PQL 리드 데이터는 Pipedrive 딜의 하류 데이터이므로 필드명/구조 변경 시 영향을 받는다.

## Key Files

```
PQL_auto/
├── CLAUDE.md                   # 프로젝트 문서
├── AGENTS.md                   # 목차 + 핵심 원칙 (이 파일)
├── ARCHITECTURE.md             # 데이터 흐름, 모듈 경계
├── PQL.md                      # Apps Script 전체 코드 + 사용법
├── docs/
│   ├── PRODUCT_SENSE.md        # 제품 방향
│   ├── PLANS.md                # 우선순위, 로드맵
│   ├── design-docs/            # 설계 문서
│   └── exec-plans/             # 실행 계획
└── .gitignore
```

## Docs Map

| 문서 | 용도 | 변경 빈도 |
|------|------|-----------|
| [CLAUDE.md](CLAUDE.md) | 프로젝트 개요, 기술 스택, 트러블슈팅 | 기능/필드 변경 시 |
| [AGENTS.md](AGENTS.md) | 목차 + 핵심 원칙 (이 파일) | 구조 변경 시 |
| [ARCHITECTURE.md](ARCHITECTURE.md) | 데이터 흐름, 처리 단계 | 로직 변경 시 |
| [docs/PRODUCT_SENSE.md](docs/PRODUCT_SENSE.md) | 제품 방향, 사용자 | 분기별 |
| [docs/PLANS.md](docs/PLANS.md) | 우선순위, 기술 부채 | 스프린트마다 |
| [PQL.md](PQL.md) | Apps Script 전체 코드 + 변경 이력 | 코드 수정 시 |

## Pipedrive CRM 연관 프로젝트

| 프로젝트 | 역할 | 공유 자원 |
|----------|------|-----------|
| [pipedrive_auto](../pipedrive_auto/) | Pipedrive 딜 -> Google Sheets/Drive 일일 동기화 | Pipedrive API, 딜 데이터 원천 |
| [pipedrive-dashboard](../pipedrive-dashboard/) | 세일즈 인사이트 HTML 대시보드 | Pipedrive API, 딜 분석 |
| **PQL_auto** (이 프로젝트) | PQL 리드 엑셀 -> Google Sheets 가공 | Google Drive 폴더, 리드 데이터 |

**교차 영향 주의사항:**
- PQL 리드 엑셀의 컬럼명(`알파리뷰 상태`, `알파업셀 상태`, `알파푸시 상태` 등)은 Pipedrive의 프로덕트/서비스 구분과 동일한 체계를 따른다. 프로덕트명 변경 시 필터링 조건과 라벨링 로직 수정 필요.
- `필요 서비스` 옵션(알파리뷰, 알파업셀, 알파푸시)은 pipedrive_auto/pipedrive-dashboard의 옵션 매핑과 동일한 체계다.
- Google Drive 폴더(TARGET_FOLDER_ID: 1PjCz9YxLLqGLYOZLffPO97tk7UKEGEaF)에 엑셀이 업로드되는 프로세스가 변경되면 이 스크립트에 영향.
