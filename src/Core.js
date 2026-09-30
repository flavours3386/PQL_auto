/***************************************
 * 순수 로직 — Apps Script 서비스를 쓰지 않는다 (node 테스트 대상)
 ***************************************/

/* ---------- CSV ---------- */

// 한 행씩 onRow(fields)로 넘기는 스트리밍 파서. feed()를 여러 번 불러도 따옴표·행 상태가 이어진다.
function createCsvParser_(onRow) {
  let field = '';
  let row = [];
  let inQuotes = false;
  let justClosed = false; // 직전 문자가 닫는 따옴표였는지 ("" 이스케이프 판별)
  let started = false;
  return {
    feed: function (text) {
      let i = 0;
      if (!started) {
        started = true;
        if (text.charCodeAt(0) === 0xfeff) i = 1; // BOM
      }
      for (; i < text.length; i++) {
        const ch = text[i];
        if (inQuotes) {
          if (ch === '"') {
            inQuotes = false;
            justClosed = true;
          } else field += ch;
          continue;
        }
        if (ch === '"') {
          if (justClosed) {
            field += '"';
            inQuotes = true;
          } else if (field === '') inQuotes = true;
          else field += ch;
          justClosed = false;
          continue;
        }
        justClosed = false;
        if (ch === ',') {
          row.push(field);
          field = '';
        } else if (ch === '\n') {
          row.push(field);
          onRow(row);
          row = [];
          field = '';
        } else if (ch !== '\r') field += ch;
      }
    },
    end: function () {
      if (field !== '' || row.length) {
        row.push(field);
        onRow(row);
      }
      row = [];
      field = '';
    },
  };
}

// UTF-8 바이트 수 (분할 다운로드에서 다음 Range 시작점 계산용)
function utf8ByteLength_(s) {
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x80) n += 1;
    else if (c < 0x800) n += 2;
    else if (c >= 0xd800 && c <= 0xdbff) {
      n += 4;
      i++;
    } else n += 3;
  }
  return n;
}

// 파일을 Range 요청으로 나눠 받아 onText로 넘긴다. fetchRange(start, end)는 그 바이트 구간을 UTF-8로 디코드한 문자열.
// 조각마다 마지막 '\n'에서 자르고 나머지는 다음 요청에서 다시 받는다. '\n'(0x0A)은 UTF-8 멀티바이트 안에 나오지 않아 자른 곳이 항상 글자 경계다.
// ponytail: 원천에 깨진 UTF-8 바이트가 있으면 U+FFFD 바이트 수가 달라 오프셋이 밀린다. 원천은 정상 UTF-8(BOM) CSV라 검사하지 않는다.
function streamChunks_(size, chunkBytes, fetchRange, onText) {
  let offset = 0;
  while (offset < size) {
    const end = Math.min(offset + chunkBytes, size) - 1;
    const text = fetchRange(offset, end);
    if (end === size - 1) {
      onText(text);
      return;
    }
    const cut = text.lastIndexOf('\n');
    if (cut < 0) throw new Error('CSV 한 행이 분할 크기(' + chunkBytes + '바이트)보다 큽니다');
    const head = text.slice(0, cut + 1);
    onText(head);
    offset += utf8ByteLength_(head);
  }
}
