/* Mock OpenWA gateway used by wa-live-test.cjs — runs as its OWN process so a
 * synchronous UrlFetchApp stub can block the test process without stalling the
 * server. Implements the rmyndharis/OpenWA REST surface subset:
 *   POST /api/sessions/{sessionId}/messages/send-text  (X-API-Key: …)
 * Exchanges are logged line-by-line to stdout as JSON. */
const http = require("http");
const PORT = parseInt(process.env.WA_MOCK_PORT || "29785", 10);
const KEY = process.env.WA_MOCK_KEY || "k-live-123";

http.createServer((req, res) => {
  let body = "";
  req.on("data", (c) => (body += c));
  req.on("end", () => {
    if ((req.headers["x-api-key"] || "") !== KEY) {
      res.writeHead(401, { "Content-Type": "application/json" });
      return res.end(JSON.stringify({ error: "Invalid API key" }));
    }
    const m = req.url.match(/^\/api\/sessions\/([^/]+)\/messages\/send-text$/);
    if (m) {
      let p = {}; try { p = JSON.parse(body || "{}"); } catch (e) {}
      process.stdout.write(JSON.stringify({
        hit: "send-text", session: m[1], chatId: p.chatId,
        headline: (p.text || "").split("\n")[0],
        lines: (p.text || "").split("\n").length,
      }) + "\n");
      res.writeHead(201, { "Content-Type": "application/json" });
      return res.end(JSON.stringify({ messageId: "wamid.live", queued: true }));
    }
    res.writeHead(404, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "route not found" }));
  });
}).listen(PORT, "127.0.0.1", () => process.stdout.write(JSON.stringify({ ready: true, port: PORT }) + "\n"));
