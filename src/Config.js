/***************************************
 * 설정값 — 운영 중 바꿀 값은 이 파일에만 둔다
 ***************************************/
const SOURCE_FOLDER_ID = '1PjCz9YxLLqGLYOZLffPO97tk7UKEGEaF'; // Drive '05. PQL' (crema·alphareview-ref도 읽는 공유 폴더 — 읽기만 한다)
const SOURCE_NAME_PREFIX = 'all_subscription_';
const DOWNLOAD_CHUNK_BYTES = 20 * 1024 * 1024; // UrlFetch 응답 한도 50MB/회

const MIN_ORDERS = 100;
const PUSH_MIN_ORDERS = 500;
const REVIEW_MIN_ORDERS = 1000;

const SALES_PIPELINE_ID = 9;
const PD_FIELD_SHOP_ID = '9d4ea1fcf0bde157910e96a2e0354e76c220e6c8';
const PD_FIELD_URL = '5a7464db665cc9fb3cebc7530c536f39205768ca';
const PD_FIELD_MALL_NAME = '4cf3a83ff7316bb926dbf2c7f9c7b92308bad7bc';
const PD_TOKEN_PROPERTY = 'PIPEDRIVE_API_TOKEN';

const AUTO_APPLY = true; // 높은 확신 역매핑을 Pipedrive에 자동 반영
const AUTO_APPLY_MAX = 100; // 한 실행에서 이보다 많으면 자동 반영을 멈추고 전부 대기로

const DEAL_OWNER = '한서연';
const DEAL_STAGE = '컨택전';

const TAB_DEAL_LIST = 'deal list';
const TAB_MAPPING = 'shop_id 매핑';
const CLEAN_TAB_PREFIX = 'clean_';

// 상태값은 공백을 뺀 형태로 적는다 (비교 전에 원천 값의 공백도 뺀다)
const REVIEW_EXCLUDE = new Set(['제거중', '해지완료', '서비스중단']);
const SITE_EXCLUDE = new Set(['구독종료', '해지완료', '계정활성화']);
const NOT_USED = new Set(['구독없음', '서비스중단', '프로덕트온보딩중', '']);

// 타겟 규칙: 하나라도 맞으면 PQL에 남는다. 새 타겟은 항목 하나를 추가한다.
const TARGETS = [
  { name: '업셀', test: (r) => r.platform === 'cafe24' && !isLive_(r.upsell) && r.upsell !== '제거중' },
  // 푸시 무료(카페24 PRO 번들)도 라이브로 본다 — 세 제품이 모두 라이브인 몰은 올리지 않는다
  { name: '푸시', test: (r) => r.platform === 'cafe24' && r.orders >= PUSH_MIN_ORDERS && !isLive_(r.push) && r.push !== '제거중' },
  // 리뷰는 아임웹도 지원한다 (플랫폼 조건 없음)
  { name: '리뷰', test: (r) => r.orders >= REVIEW_MIN_ORDERS && !isLive_(r.review) && r.review !== '제거중' },
];

const REQUIRED_COLUMNS = ['shop_id', '플랫폼', '최근 30일 플랫폼 주문수', '알파리뷰 상태', '알파업셀 상태', '알파푸시 상태', '사이트 상태', '담당자명', '담당자전화번호'];

const OUTPUT_HEADERS = [
  'shop_name', 'shop_id', 'mall_id', '플랫폼', '최근 30일 플랫폼 주문수(API)', '타겟', '서비스 라벨', '딜 의심',
  '회사명', '담당자명', '쇼핑몰명', '담당자전화번호', '담당자이메일', '대표도메인', '주소',
  'shop_no', '플랜', '사이트 상태', '알파리뷰 상태', '알파업셀 상태', '알파푸시 상태',
  '최근 30일 플랫폼 주문수', '최근 30일 전체 주문수', '설치시점 플랫폼 주문수(API)',
  '최근 30일 UV(방문자수)', '최근 30일 PV(페이지뷰)', '임직원 수', '이메일', '사업자', '고객센터', '전화번호', '담당자직책', '결제담당이메일',
];
const UPLOAD_HEADERS = ['거래 제목', 'shop_id', '상점아이디', '호스팅사', '월 주문 수', '거래 소유자', '단계 (파이프라인)', '거래 라벨', '조직 이름', '이름', '쇼핑몰명', '전화', '이메일', 'URL', '주소'];
const MAPPING_HEADERS = ['딜 ID', '딜 이름', '원래 shop_id', '후보 shop_id', '후보 shop_name', '일치 키', '신뢰도', '판정', '상태', '기록일'];
