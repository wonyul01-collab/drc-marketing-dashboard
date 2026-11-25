
/* ===== 매일 가져오기: 실적센터(로그인된 탭)에서 최신 통합실적 엑셀을 읽어 대시보드 저장소 문서로 바꾼다 =====
   dashboard.html 의 parser 스크립트 + 이 파일을 이어 붙인 것이 daily_fetch.js 다 (build_daily_fetch.js 가 만든다).
   실적센터 탭에서 javascript 로 실행하면 { ok, base, file, docs } 를 돌려준다. */
await new Promise(function (ok, no) {
  if (window.XLSX) return ok();
  var s = document.createElement("script");
  s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
  s.onload = ok; s.onerror = function () { no(new Error("xlsx_load_failed")); };
  document.head.append(s);
});
var __out = await (async function () {
  if (/login/.test(location.pathname)) return { ok: false, reason: "login" };
  var stRes = await fetch("/data/status.json", { cache: "no-store" });
  if (stRes.status !== 200) return { ok: false, reason: "login", status: stRes.status };
  var st = await stRes.json();
  var rep = (st.reports || []).filter(function (r) { return r.file; })[0];
  if (!rep) return { ok: false, reason: "no_report" };
  var buf = await (await fetch("/data/" + rep.file, { cache: "no-store" })).arrayBuffer();
  var P = LEDGER.parseWorkbook(XLSX.read(buf, { type: "array", cellDates: false }));
  if (!P || P.kind !== "center") return { ok: false, reason: "shape" };
  var cur = P.base.slice(0, 7);
  var d = new Date(cur + "-01T00:00:00"); d.setMonth(d.getMonth() - 1);
  var prev = d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
  var lm = LEDGER.byMonth(P.days), pm = LEDGER.byMonth(P.products), now = new Date().toISOString();
  var docs = {};
  [prev, cur].forEach(function (m) {
    if (lm[m]) docs["ledger/center_" + m] = { source: "center", month: m, days: lm[m], updatedAt: now, file: rep.name };
    if (pm[m]) docs["products/center_" + m] = { source: "center", month: m, days: pm[m], updatedAt: now };
  });
  if (P.adMonth) docs["admonth/" + P.adMonth.month] = Object.assign({ updatedAt: now, file: rep.name }, P.adMonth);
  var ds = Object.keys(P.days).sort();
  var monthOnline = 0;
  Object.keys(lm[cur] || {}).forEach(function (k) { var s = lm[cur][k].sales; Object.keys(s).forEach(function (c) { if (c !== "B2B·오프라인") monthOnline += s[c]; }); });
  return { ok: true, base: P.base, file: rep.name, madeAt: rep.at, months: [prev, cur], monthOnline: monthOnline, upload: { file: rep.name, kind: "실적센터 매일 자동 가져오기", from: ds[0], to: ds[ds.length - 1], days: ds.length, products: 0, at: now }, docs: docs };
})();
__out;
