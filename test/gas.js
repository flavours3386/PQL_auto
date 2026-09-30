// Apps Script처럼 src 파일들을 하나의 전역 스코프에 올린다 (테스트 전용)
const fs = require('fs');
const path = require('path');
const vm = require('vm');

for (const f of ['Config.js', 'Core.js']) {
  vm.runInThisContext(fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8'), { filename: f });
}
