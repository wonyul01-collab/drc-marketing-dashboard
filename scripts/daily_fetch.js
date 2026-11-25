/* 자동 생성 파일: build_daily_fetch.js 로 만든다. 직접 고치지 말 것 */

/* ===== 엑셀 → 통합 원장 변환 (화면과 시험 스크립트가 함께 쓰는 순수 함수) ===== */
var LEDGER = (function () {
  var CHANNELS = ["스마트스토어", "쿠팡 로켓배송", "쿠팡 판매자(윙)", "자사몰(카페24)", "기타 온라인", "B2B·오프라인"];
  var OFFLINE = "B2B·오프라인";
  function num(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return isFinite(v) ? v : 0;
    var s = String(v).replace(/[,원\s]/g, "");
    var n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function serialToISO(serial) {
    var ms = Math.round((serial - 25569) * 86400000);
    var d = new Date(ms);
    return d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate());
  }
  function toISO(v) {
    if (typeof v === "number" && v > 20000 && v < 80000) return serialToISO(v);
    if (v instanceof Date && !isNaN(v)) return v.getFullYear() + "-" + pad(v.getMonth() + 1) + "-" + pad(v.getDate());
    var s = String(v || "").trim();
    var m = s.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
    if (m) return m[1] + "-" + pad(+m[2]) + "-" + pad(+m[3]);
    m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m) return m[1] + "-" + m[2] + "-" + m[3];
    return null;
  }
  function rows(wb, name) {
    var ws = wb.Sheets[name];
    if (!ws) return null;
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  }
  function cell(r, i) { return r && r[i] !== undefined ? r[i] : null; }
  function str(v) { return v === null || v === undefined ? "" : String(v).trim(); }

  function isCenter(wb) { return !!(wb.Sheets["일실적·추이"] && wb.Sheets["광고"]); }

  function baseDateOf(wb) {
    var names = ["요약", "일실적·추이"];
    for (var i = 0; i < names.length; i++) {
      var rs = rows(wb, names[i]) || [];
      for (var j = 0; j < Math.min(rs.length, 5); j++) {
        var m = str(cell(rs[j], 0)).match(/(\d{4})-(\d{2})-(\d{2})/);
        if (m) return { y: +m[1], m: +m[2], d: +m[3], iso: m[0] };
      }
    }
    return null;
  }
  function mdToISO(label, base) {
    var m = str(label).match(/^(\d{1,2})\/(\d{1,2})/);
    if (!m) return null;
    var mm = +m[1], dd = +m[2], y = base.y;
    if (mm > base.m) y -= 1;
    return y + "-" + pad(mm) + "-" + pad(dd);
  }

  /* 실적센터 '디알컴퍼니 통합 실적' 엑셀 */
  function parseCenter(wb) {
    var base = baseDateOf(wb);
    if (!base) throw new Error("기준일을 찾지 못했습니다. '요약' 시트 둘째 줄에 '기준일 2026-11-16' 같은 글이 있어야 합니다.");
    var out = { kind: "center", base: base.iso, days: {}, products: {}, adMonth: null };
    // 1) 일실적·추이: 날짜별 채널 매출·광고비 합계
    var rs = rows(wb, "일실적·추이");
    var h = -1;
    for (var i = 0; i < rs.length; i++) { if (str(cell(rs[i], 0)) === "구간" && rs[i].indexOf("스마트스토어") >= 0) { h = i; break; } }
    if (h < 0) throw new Error("'일실적·추이' 시트에서 '구간' 머리줄을 찾지 못했습니다.");
    var head = rs[h].map(str);
    var col = function (name) { return head.indexOf(name); };
    for (i = h + 1; i < rs.length; i++) {
      var iso = mdToISO(cell(rs[i], 0), base);
      if (!iso) { if (str(cell(rs[i], 0)) === "") break; continue; }
      var sales = {};
      CHANNELS.forEach(function (ch) { var c = col(ch); if (c >= 0) sales[ch] = num(cell(rs[i], c)); });
      out.days[iso] = {
        sales: sales,
        orders: num(cell(rs[i], col("주문건수"))),
        adCost: num(cell(rs[i], col("광고비"))),
        adRev: num(cell(rs[i], col("광고매출"))),
        asOf: base.iso
      };
    }
    // 2) 광고: 이번 달 채널별 광고비·광고매출
    rs = rows(wb, "광고") || [];
    for (i = 0; i < rs.length; i++) {
      var t = str(cell(rs[i], 0)).match(/^이번 달 \((\d{1,2})\/(\d{1,2})~(\d{1,2})\/(\d{1,2})\)/);
      if (!t) continue;
      var hd = (rs[i + 1] || []).map(str);
      var ci = function (n) { return hd.indexOf(n); };
      var fromISO = mdToISO(t[1] + "/" + t[2], base), toISOd = mdToISO(t[3] + "/" + t[4], base);
      var am = { month: fromISO.slice(0, 7), from: fromISO, to: toISOd, rows: [] };
      for (var k = i + 2; k < rs.length; k++) {
        var ch = str(cell(rs[k], 0));
        if (!ch) break;
        am.rows.push({ channel: ch, source: str(cell(rs[k], ci("원천"))), cost: num(cell(rs[k], ci("광고비"))), rev: num(cell(rs[k], ci("광고매출"))), imp: num(cell(rs[k], ci("노출"))), click: num(cell(rs[k], ci("클릭"))) });
      }
      out.adMonth = am;
      break;
    }
    // 3) 상품원본: 날짜·채널·제품별 수량·매출
    rs = rows(wb, "상품원본");
    if (rs && rs.length) {
      var ph = rs[0].map(str);
      var pc = function (n) { return ph.indexOf(n); };
      for (i = 1; i < rs.length; i++) {
        var pd = toISO(cell(rs[i], pc("일자")));
        if (!pd) continue;
        var brand = str(cell(rs[i], pc("브랜드"))), prod = str(cell(rs[i], pc("상품")));
        (out.products[pd] = out.products[pd] || []).push([str(cell(rs[i], pc("채널"))), brand, prod, num(cell(rs[i], pc("수량"))), num(cell(rs[i], pc("매출")))]);
      }
    }
    return out;
  }

  /* 원장 양식 CSV: 날짜,채널,제품,매출,주문수,광고비,광고매출 */
  var CSV_HEAD = ["날짜", "채널", "제품", "매출", "주문수", "광고비", "광고매출"];
  function isLedgerCsv(wb) {
    var rs = rows(wb, wb.SheetNames[0]) || [];
    var hd = (rs[0] || []).map(str);
    return hd.indexOf("날짜") >= 0 && hd.indexOf("채널") >= 0 && (hd.indexOf("매출") >= 0 || hd.indexOf("광고비") >= 0);
  }
  function parseLedgerCsv(wb) {
    var rs = rows(wb, wb.SheetNames[0]);
    var hd = rs[0].map(str);
    var c = function (n) { return hd.indexOf(n); };
    var out = { kind: "csv", base: null, days: {}, products: {}, adMonth: null, skipped: 0 };
    for (var i = 1; i < rs.length; i++) {
      var d = toISO(cell(rs[i], c("날짜")));
      var ch = str(cell(rs[i], c("채널")));
      if (!d || !ch) { if (rs[i].some(function (v) { return str(v) !== ""; })) out.skipped++; continue; }
      var day = out.days[d] = out.days[d] || { sales: {}, orders: 0, adCost: 0, adRev: 0, adByCh: {} };
      var s = num(cell(rs[i], c("매출"))), ac = num(cell(rs[i], c("광고비"))), ar = num(cell(rs[i], c("광고매출")));
      day.sales[ch] = (day.sales[ch] || 0) + s;
      day.orders += num(cell(rs[i], c("주문수")));
      day.adCost += ac; day.adRev += ar;
      if (ac || ar) { var a = day.adByCh[ch] = day.adByCh[ch] || { cost: 0, rev: 0 }; a.cost += ac; a.rev += ar; }
      var p = str(cell(rs[i], c("제품")));
      if (p && s) (out.products[d] = out.products[d] || []).push([ch, "", p, 0, s]);
      if (!out.base || d > out.base) out.base = d;
    }
    return out;
  }

  function parseWorkbook(wb) {
    if (isCenter(wb)) return parseCenter(wb);
    if (isLedgerCsv(wb)) return parseLedgerCsv(wb);
    return null;
  }

  /* 파싱 결과 → 월별 문서로 묶기 (같은 날짜는 덮어씀) */
  function byMonth(obj) {
    var m = {};
    Object.keys(obj).forEach(function (d) { (m[d.slice(0, 7)] = m[d.slice(0, 7)] || {})[d] = obj[d]; });
    return m;
  }
  function mergeDays(existing, incoming) {
    var out = {};
    Object.keys(existing || {}).forEach(function (d) { out[d] = existing[d]; });
    Object.keys(incoming).forEach(function (d) { out[d] = incoming[d]; });
    return out;
  }

  return { CHANNELS: CHANNELS, OFFLINE: OFFLINE, CSV_HEAD: CSV_HEAD, parseWorkbook: parseWorkbook, byMonth: byMonth, mergeDays: mergeDays, num: num, toISO: toISO };
})();

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
