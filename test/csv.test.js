require('./gas');
const test = require('node:test');
const assert = require('node:assert');

function parseAll(chunks) {
  const rows = [];
  const p = createCsvParser_((r) => rows.push(r));
  chunks.forEach((c) => p.feed(c));
  p.end();
  return rows;
}

test('따옴표 안 쉼표·줄바꿈·이스케이프, BOM, CRLF', () => {
  const csv = '﻿a,b,c\r\n1,"x, y","he said ""hi"""\r\n2,"line1\nline2",\r\n';
  assert.deepStrictEqual(parseAll([csv]), [['a', 'b', 'c'], ['1', 'x, y', 'he said "hi"'], ['2', 'line1\nline2', '']]);
});

test('조각 경계가 따옴표 필드·이스케이프 한가운데여도 결과가 같다', () => {
  const csv = 'a,b\n1,"he said ""hi""\nbye"\n2,z';
  const whole = parseAll([csv]);
  for (let i = 1; i < csv.length; i++) {
    assert.deepStrictEqual(parseAll([csv.slice(0, i), csv.slice(i)]), whole, 'cut at ' + i);
  }
});

test('마지막 줄에 줄바꿈이 없어도 행이 나온다', () => {
  assert.deepStrictEqual(parseAll(['a,b\n1,2']), [['a', 'b'], ['1', '2']]);
});

test('utf8ByteLength_는 Buffer 바이트 수와 같다', () => {
  for (const s of ['abc', '한글', 'é', '😀', '﻿shop_id,회사명\n']) {
    assert.strictEqual(utf8ByteLength_(s), Buffer.byteLength(s, 'utf8'), s);
  }
});

test('streamChunks_: 멀티바이트 글자가 바이트 경계에 걸려도 원문이 복원된다', () => {
  const csv = '﻿shop_id,회사명,주소\n1,"주식회사 가나","서울, 강남"\n2,다라😀,"부산\n해운대"\n3,마바,대구\n';
  const buf = Buffer.from(csv, 'utf8');
  // Apps Script getContentText처럼 BOM을 지우지 않고, 잘린 끝 글자는 U+FFFD로 둔다
  const dec = new TextDecoder('utf-8', { ignoreBOM: true });
  const minChunk = Math.max(...csv.split('\n').map((l) => Buffer.byteLength(l + '\n')));
  for (let chunk = minChunk; chunk <= buf.length; chunk++) {
    let out = '';
    streamChunks_(buf.length, chunk, (s, e) => dec.decode(buf.subarray(s, e + 1)), (t) => { out += t; });
    assert.strictEqual(out, csv, 'chunk ' + chunk);
  }
});

test('한 행이 분할 크기보다 크면 멈춘다', () => {
  const buf = Buffer.from('aaaaaaaaaa\nb\n');
  assert.throws(() => streamChunks_(buf.length, 4, (s, e) => buf.subarray(s, e + 1).toString(), () => {}), /분할 크기/);
});

test('디코더가 BOM을 지워도(Apps Script getContentText) restoreBom_로 감싸면 조각 경계가 밀리지 않는다', () => {
  const body = 'a,b\n' + Array.from({ length: 50 }, (_, i) => i + ',"값' + i + '"').join('\n') + '\n';
  const count = (csv, wrap) => {
    const buf = Buffer.from(csv, 'utf8');
    const strip = new TextDecoder('utf-8'); // 기본값: BOM 삭제
    const fetch = (s, e) => strip.decode(buf.subarray(s, e + 1));
    const rows = [];
    const p = createCsvParser_((r) => rows.push(r));
    streamChunks_(buf.length, 64, wrap ? wrap(fetch) : fetch, p.feed);
    p.end();
    return rows.length;
  };
  assert.notStrictEqual(count('﻿' + body), 51); // 감싸지 않으면 조각 행이 끼어든다 (재현)
  assert.strictEqual(count('﻿' + body, (f) => restoreBom_(f, true)), 51);
  assert.strictEqual(count(body, (f) => restoreBom_(f, false)), 51);
});

test('isUtf8Bom_: 부호 있는 바이트(Apps Script)와 없는 바이트 모두 판별', () => {
  assert.strictEqual(isUtf8Bom_([-17, -69, -65]), true);
  assert.strictEqual(isUtf8Bom_([0xef, 0xbb, 0xbf]), true);
  assert.strictEqual(isUtf8Bom_([0x73, 0x68, 0x6f]), false);
  assert.strictEqual(isUtf8Bom_([]), false);
});
