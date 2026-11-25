// dashboard.html 의 parser 스크립트와 daily_fetch_tail.js 를 이어 daily_fetch.js 를 만든다.
// 실행: node scripts/build_daily_fetch.js   (dashboard.html 의 parser 를 고친 뒤 꼭 다시 실행)
const fs = require("fs");
const path = require("path");
const dir = __dirname;
const html = fs.readFileSync(path.join(dir, "..", "dashboard.html"), "utf8");
const parser = html.match(/<script id="parser">([\s\S]*?)<\/script>/)[1];
const tail = fs.readFileSync(path.join(dir, "daily_fetch_tail.js"), "utf8");
fs.writeFileSync(path.join(dir, "daily_fetch.js"), "/* 자동 생성 파일: build_daily_fetch.js 로 만든다. 직접 고치지 말 것 */\n" + parser + tail);
console.log("daily_fetch.js written");
