// Apps Script처럼 src 파일들을 하나의 전역 스코프에 올린다 (테스트 전용)
const fs = require('fs');
const path = require('path');
const vm = require('vm');

// Io.js·Main.js는 불러오기만 해서는 Apps Script 서비스를 부르지 않는다. 테스트가 전역 가짜(UrlFetchApp 등)를 넣고 호출한다.
for (const f of ['Config.js', 'Core.js', 'Io.js', 'Main.js']) {
  vm.runInThisContext(fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8'), { filename: f });
}
