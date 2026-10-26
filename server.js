// 테스트용 HTTPS 서버
//  1) docs 폴더(추가기능 파일)를 https://localhost:3000 으로 제공
//  2) POST /api/hook : 추가기능이 보낸 데이터를 받아 콘솔에 출력하는 가짜 API
const https = require("https");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

const PORT = 3000;
const PUBLIC_DIR = path.join(__dirname, "docs");

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".png": "image/png",
};

function sendCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

function handleHook(req, res) {
  let raw = "";
  req.on("data", (chunk) => (raw += chunk));
  req.on("end", () => {
    console.log(`\n[${new Date().toLocaleString()}] 메일에서 감지됨`);
    try {
      console.log(JSON.stringify(JSON.parse(raw), null, 2));
    } catch {
      console.log(raw);
    }
    sendCors(res);
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ ok: true }));
  });
}

function handleStatic(req, res) {
  const urlPath = decodeURIComponent(req.url.split("?")[0]);
  const filePath = path.join(PUBLIC_DIR, urlPath === "/" ? "taskpane.html" : urlPath);
  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403).end();
    return;
  }
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404).end("Not found");
      return;
    }
    res.writeHead(200, { "Content-Type": MIME[path.extname(filePath)] || "application/octet-stream" });
    res.end(data);
  });
}

devCerts.getHttpsServerOptions().then((options) => {
  https
    .createServer(options, (req, res) => {
      console.log(`[${new Date().toLocaleTimeString()}] ${req.method} ${req.url}`);
      if (req.url.startsWith("/api/hook")) {
        if (req.method === "OPTIONS") {
          sendCors(res);
          res.writeHead(204).end();
        } else {
          handleHook(req, res);
        }
        return;
      }
      handleStatic(req, res);
    })
    .listen(PORT, () => {
      console.log(`추가기능 서버 실행 중: https://localhost:${PORT}`);
      console.log("메일을 보내면 감지된 내용이 여기에 출력됩니다.");
    });
});
