/**
 * Task 28 — LIVE smoke run for the WhatsApp/OpenWA integration.
 *
 *   node google-apps-script/tests/node/wa-live-test.cjs
 *
 * 1. Boots a local mock of the OpenWA gateway on 127.0.0.1:29785 implementing
 *    the rmyndharis/OpenWA REST surface (POST /api/sessions/{id}/messages/send-text,
 *    X-API-Key auth — same shape as the real gateway's README).
 * 2. Wires Apps Script's UrlFetchApp.fetch stub to REAL HTTP (sync curl) so the
 *    production code path in SalesDashboard.gs runs unchanged.
 * 3. Exercises:
 *      - sendWhatsAppAnswer()  : chat-bot query -> gateway send-text
 *      - doPost()              : inbound message.received webhook -> auto-reply
 *      - wrong API key         : surfaces a clean gateway-side error
 * 4. Dumps every HTTP exchange the gateway received.
 *
 * To go real: set WA_GATEWAY_URL / WA_API_KEY / WA_SESSION (/WA_CHATID) in the
 * Apps Script project Script Properties and re-run — no code changes.
 */
const http = require("http");
const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const PORT = 29785;
const RAW = fs.readFileSync(path.join(__dirname, "..", "..", "SalesDashboard.gs"), "utf8");
const gatewayScript = path.join(__dirname, "wa-mock-gateway.cjs");

// captured exchanges (gateway writes one JSON line per hit; we keep tails)
let seen = [];
function collectGatewayLog(logFile) {
  try {
    seen = fs.readFileSync(logFile, "utf8").trim().split("\n")
      .filter(Boolean).map((l) => JSON.parse(l)).filter((o) => o.hit);
  } catch (e) { seen = []; }
}

/* ---------- Apps Script overlays ---------- *//* ---------- Apps Script overlays ---------- */
global.Utilities = { formatDate(d, tz, f) {
  if (f && f.indexOf("HH") > -1) return "30 Sep 2026, 10:30";
  const p = (n) => (n < 10 ? "0" : "") + n;
  return p(d.getDate()) + "-" + p(d.getMonth() + 1) + "-" + d.getFullYear();
}};
global.Session = { getScriptTimeZone: () => "Asia/Calcutta" };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {} }) };
global.HtmlService = { createHtmlOutput: () => ({ addMetaTag() { return this; }, setTitle() { return this; } }) };
global.ContentService = { createTextOutput(t) { const o = { t, setMimeType() { return o; } }; return o; },
  MimeType: { JSON: "application/json" } };

const PROPS = {
  WA_GATEWAY_URL: `http://127.0.0.1:${PORT}`,
  WA_API_KEY: "k-live-123",
  WA_SESSION: "sales",
  WA_CHATID: "919791994639@c.us",
};
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => (k in PROPS ? PROPS[k] : null) }) };

function mockSheet(name, values) {
  return { getName: () => name, getDataRange: () => ({
    getValues: () => values,
    getDisplayValues: () => values.map((r) => r.map((v) => (v == null || v instanceof Date) ? "" : String(v))),
  }) };
}
const crm = [["InvoiceCode","InvoiceDate","SalesPersonName","RevenueType","ItemDescription","SubCategory","SchemeCode","TaxableAmount","GrossProfit","TranType","ItemCode","PartType"],
  ["INV1","23-09-2026 10:15","Subha","Service","iPhone 13 display","Display","","5000","1500","SALES","IP13D","Display"]];
const service = [["CaseId","FirstName","LastName","MobileNo","ProductDescription","Status"],
  ["C1001","Ravi","Kumar","9988776655","Earpods Pro","Closed"]];
const ALL = [mockSheet("crm", crm), mockSheet("service", service)];
const SS = { getSheets: () => ALL,
  getSheetByName: (n) => ALL.find((s) => s.getName() === n) || null };
global.SpreadsheetApp = { getActiveSpreadsheet: () => SS };

// UrlFetchApp shim -> REAL http via sync curl (headers via -H, json via --data)
global.UrlFetchApp = { fetch(url, opts) {
  opts = opts || {};
  const hargs = Object.keys(opts.headers || {})
    .map((k) => `-H ${JSON.stringify(k + ": " + opts.headers[k])}`).join(" ");
  const cmd = `curl -s -o /tmp/walive_resp.json -w "%{http_code}" -X ${(opts.method || "get").toUpperCase()} ` +
    `-H ${JSON.stringify("Content-Type: " + (opts.contentType || "application/json"))} ${hargs} ` +
    (opts.payload ? `--data-binary '${String(opts.payload).replace(/'/g, `'\\''`)}' ` : "") +
    JSON.stringify(url) + " 2>/dev/null";
  let code = "0";
  try { code = String(execSync(cmd, { shell: "/bin/bash" }).toString()).trim(); } catch (e) { code = "000"; }
  let body = "";
  try { body = fs.readFileSync("/tmp/walive_resp.json", "utf8"); } catch (e) {}
  return { getResponseCode: () => parseInt(code, 10) || 0, getContentText: () => body };
}};

/* ---------- run the real SalesDashboard.gs ---------- */
eval(RAW);
getSpreadsheet_ = () => SS;

function banner(t) { console.log("\n=== " + t + " " + "=".repeat(Math.max(0, 58 - t.length))); }

const LOGF = "/tmp/walive_gateway.log";
const { spawn } = require("child_process");
// direct fd — child's writes land in the file with no pipe to flush
const outFd = fs.openSync(LOGF, "w");
const child = spawn("node", [gatewayScript], { stdio: ["ignore", outFd, "inherit"] });

function waitForGateway(cb) {
  http.get(`http://127.0.0.1:${PORT}/api/docs`, (res) => {
    if (res.statusCode === 401 || res.statusCode === 404) return cb();
    setTimeout(() => waitForGateway(cb), 100);
  }).on("error", () => setTimeout(() => waitForGateway(cb), 100));
}

waitForGateway(() => {
  console.log("mock OpenWA gateway up on http://127.0.0.1:" + PORT);

  banner("1. sendWhatsAppAnswer — sheet-bot answer pushed to gateway");
  const r1 = JSON.parse(sendWhatsAppAnswer(JSON.stringify({ q: "Ravi Kumar", chat: "9988776655" })));
  console.log("result:", JSON.stringify(r1, null, 2));

  banner("2. sendWhatsAppAnswer — default chat from WA_CHATID");
  const r2 = JSON.parse(sendWhatsAppAnswer(JSON.stringify({ q: "sale date from 23.09.26" })));
  console.log("result:", JSON.stringify(r2, null, 2));

  banner("3. doPost webhook — inbound WhatsApp message auto-answered");
  doPost({ postData: { contents: JSON.stringify({
    type: "message.received",
    data: { chatId: "919988776655@c.us", body: "part number for display", fromMe: false },
  }) } });
  console.log("webhook handled");

  banner("4. wrong API key surfaces cleanly");
  PROPS.WA_API_KEY = "wrong-key";
  const r4 = JSON.parse(sendWhatsAppAnswer(JSON.stringify({ q: "Ravi Kumar", chat: "9988776655" })));
  console.log("result:", JSON.stringify(r4, null, 2));
  PROPS.WA_API_KEY = "k-live-123";

  collectGatewayLog(LOGF);
  banner("GATEWAY LOG — every HTTP exchange received");
  seen.forEach((s, i) => console.log(`  [${i + 1}] session=${s.session} chatId=${s.chatId}\n      "${s.headline}" (${s.lines} lines)`));
  console.log("\nTotal sends observed: " + seen.length + " (expected 3 ok; the bad-key call was 401-rejected)");

  const pass = seen.length === 3
    && r1.ok === true && r2.ok === true
    && r4.ok === false && r4.code === 401;
  console.log(pass ? "\nLIVE SMOKE RUN: PASS" : "\nLIVE SMOKE RUN: FAIL");
  child.kill();
  process.exit(pass ? 0 : 1);
});
