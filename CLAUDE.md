# PQL 자동화 프로젝트

> 문서 목차 및 핵심 원칙: [AGENTS.md](AGENTS.md)

## 개요
Sales 파이프라인에서 누락된 세일즈 타겟 재고를 찾는 Apps Script. 매월 Drive `05. PQL` 폴더에 올라오는 전사 구독 CSV(`all_subscription_MMDD.csv`, 약 45MB·6.6만 행)를 직접 읽어 공통 클렌징과 업셀·푸시·리뷰 타겟 규칙을 적용하고, Sales 딜이 있는 몰은 빼고, 남은 몰을 Pipedrive 딜로 자동 생성한다. shop_id가 비거나 텍스트인 Sales 딜은 CSV와 대조해 shop_id를 채운다. SDR이 시트 메뉴 한 번으로 실행한다.

## 기술 스택
- Google Apps Script (V8), 바운드 스크립트 (scriptId `1y8duYMOG6reCDg_2fPSeRhh0_rAQuFqTDNt8Ml7eGDQK0wmM3-yqRirW`, 시트 `PQL_auto` `1G4vYWS_AVRdFouTBB0eD9PwkFUz1kCZCfDDQ4n2j6Lk`). 2026-09-30 이전 시트 `PQL_cleansing_auto`(`13vgWMZ5…`)는 편집 이력이 무거워 시트 쓰기가 7배 느리고 타임아웃이 나서 사용 중지
- Drive API v3 REST (`alt=media` Range 분할 다운로드), Pipedrive API v1/v2 REST — 모두 `UrlFetchApp`, Advanced Service 미사용
- 테스트: node 24 내장 `node:test` (순수 로직만)
- 배포: clasp 3.4

## 주요 명령어
```bash
node --test 'test/*.test.js'   # 테스트 (디렉터리 인자 'test/'는 Node 24에서 동작하지 않음)
clasp push -f                   # 배포 (.clasp.json rootDir: src)
clasp pull                      # 원격 확인
```

## 폴더 구조
```
PQL_auto/
├── .clasp.json            # scriptId, rootDir: src
├── src/
│   ├── appsscript.json    # 매니페스트 (timeZone Asia/Seoul)
│   ├── Config.js          # 설정·타겟 규칙·Pipedrive 필드 키 (운영 중 바꾸는 값은 여기만)
│   ├── Core.js            # 순수 로직: CSV 스트리밍·클렌징·타겟·라벨·역매핑·업로드 재료·빌더
│   ├── Io.js              # Drive·Pipedrive·Sheets I/O
│   └── Main.js            # 메뉴·실행 흐름
├── test/                  # gas.js(로더) + *.test.js (io.test.js = 가짜 Pipedrive·시트로 쓰기 경로)
├── PQL.md                 # 사용법
├── ARCHITECTURE.md        # 흐름·규칙·제약
├── CHANGELOG.md           # 지난 세대 변경 기록
└── docs/                  # PRODUCT_SENSE, PLANS, design-docs(spec), exec-plans(계획)
```

## 최근 변경사항

### 2026-09-30 파이프라인 재구축
- 원인: 원천이 7월부터 CSV 45MB(789만 셀)로 바뀌었는데 스크립트는 xlsx·Sheets만 찾아 옛 파일을 최신으로 잡았고, 전체를 raw 탭에 옮기다 6분 한도·1,000만 셀 한도에 걸렸다. 알파리뷰 `서비스 중단` 띄어쓰기 오타로 82행이 새고 있었다
- CSV를 20MB Range 조각으로 받아 스트리밍 파서로 한 번 훑는다. raw 탭 폐지, 원천 폴더에는 쓰지 않는다. 약 36초
- 공통 클렌징 ①~⑥ + 타겟 3종(업셀 cafe24·주문 150+·업셀 라이브 아님 / 푸시 cafe24·500+·푸시 라이브 아님 / 리뷰 1,000+·리뷰 라이브 아님). 라이브는 무료 포함 4종
- Sales 딜 제외를 코드가 한다(수동 deal list·XLOOKUP 대체). shop_id가 빈칸·텍스트인 딜은 이메일·전화·이름·URL로 CSV와 대조해 키 2개 이상 일치 시 Pipedrive shop_id 자동 반영(노트로 원래 값 보존), 나머지는 `shop_id 매핑` 탭에서 승인/거절
- 결과를 Pipedrive 딜로 자동 업로드(0901 수동 가져오기와 같은 필드 배치 + 세일즈티어). xlsx는 만들지 않는다(수동 가져오기와 겹치면 중복 딜)
- 배포 전 Orca(Codex) 독립 리뷰에서 Critical 2·Important 7 수정: 실행 잠금, shop_id 쓰기 직전 재확인, 이메일+전화만으로 자동 반영 금지, 업로드 선행 실패 시 딜 보류, 시간 예산·429 상한, 필수 열 23개, shop_id 검증·중복 중단, 매핑 탭 쓰기 순서. 테스트 52개(가짜 Pipedrive·시트 포함)
- `업로드 이력` 탭: 실행마다 월(원천 파일 기준)·업로드일·타겟·세일즈티어별 업로드 수와 전체 합계를 누적(PQL 추세용). 2026-10 PQL 194건부터 기록
- 설계 = `docs/design-docs/2026-09-30-pql-pipeline-refactor-design.md`, 계획 = `docs/exec-plans/2026-09-30-pql-pipeline-refactor.md`

## 트러블슈팅

### 이전 시트에서 '스프레드시트 서비스 타임아웃' (2026-09-30)
- 증상: `[시트 쓰기] … 스프레드시트 서비스가 타임아웃되었습니다`가 실행마다 다른 탭·작업에서 발생
- 가설 기각 순서: 표 객체(표2) 삭제 → 동일, 탭 선생성·단계별 flush → 동일, 딜 필드 축소(메모리) → 동일
- 확정: 같은 실행에서 같은 결과를 새 스프레드시트에 쓰면 8.3초, 이전 시트는 61.5초(7배). 매달 790만 셀 raw 덮어쓰기로 쌓인 편집 이력 탓으로 추정(탭을 지워도 이력은 남음)
- 해결: 새 스프레드시트 `PQL_auto`로 이전(`.clasp.json` scriptId 교체). 원천 전체를 시트에 쓰지 않으니 다시 무거워지지 않는다
- 참고: 이 스크립트의 기본 GCP 프로젝트에는 Sheets REST API가 꺼져 있어(403) UrlFetch로 Sheets API를 쓰는 우회는 불가

### Apps Script가 CSV BOM을 지워 분할 다운로드 경계가 밀림 (2026-09-30)
- 증상: 원천 행이 1행 많게 나오고 빈 조각 행(`["",""]`)이 끼어듦
- 원인: `HTTPResponse.getContentText('UTF-8')`이 파일 첫머리 UTF-8 BOM(3바이트)을 지운다. 첫 조각 바이트 수가 3 적게 계산돼 다음 Range가 3바이트 앞당겨짐. 3바이트 Range 요청(`bytes=0-2`)은 206인데 본문이 비어서 바이트로 판별할 수도 없다
- 해결: `restoreBom_` — 첫 조각이 BOM 없이 오면 3바이트 뒤부터 받은 글자와 첫머리를 비교해 같으면 BOM을 되돌린다. 빈 응답이면 멈춘다
- 교훈: node `TextDecoder` 기본값도 BOM을 지운다. 테스트는 행 수가 아니라 행 내용 전체를 비교해야 3바이트 누락을 잡는다

### Google xlsx 내보내기가 숫자를 `600.0`으로 저장 (2026-09-30)
- Pipedrive shop_id는 텍스트 필드라 `600.0`으로 들어가면 다음 달 숫자 매칭이 깨진다. shop_id는 문자열로 쓴다(현재는 API 업로드라 해당 경로 없음)

### Drive API v2 → v3 자동 전환으로 Advanced Drive Service가 깨짐 (2026-03-09)
- `Drive.Files.insert` Bad Request → Advanced Service를 완전히 걷어내고 UrlFetch REST 직접 호출로 전환. 이후 Advanced Service는 쓰지 않는다(Golden Principle 1)

### 비밀번호 걸린 xlsx는 Google API로 변환 불가 (2026-03-09)
- 현재 원천은 CSV라 해당 없음. xlsx를 다시 쓰게 되면 업로드 전 암호 제거 필요
