const http = require("http");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const PORT = Number(process.env.PORT) || 8080;
const ROOT = __dirname;
const BASIC_AUTH_USER = String(process.env.BASIC_AUTH_USER || "").trim();
const BASIC_AUTH_PASS = String(process.env.BASIC_AUTH_PASS || "").trim();
const BASIC_AUTH_ENABLED = BASIC_AUTH_USER.length > 0 && BASIC_AUTH_PASS.length > 0;

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".ico": "image/x-icon",
  ".webp": "image/webp"
};

function writeSecurityHeaders(res, contentType = "text/plain; charset=utf-8") {
  res.setHeader("Content-Type", contentType);
  res.setHeader("X-Content-Type-Options", "nosniff");
  res.setHeader("X-Frame-Options", "DENY");
  res.setHeader("Referrer-Policy", "no-referrer");
  res.setHeader("Permissions-Policy", "camera=(), microphone=(), geolocation=(), payment=()");
  res.setHeader("Cross-Origin-Opener-Policy", "same-origin");
  res.setHeader("Cross-Origin-Resource-Policy", "same-origin");
  res.setHeader(
    "Content-Security-Policy",
    "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; connect-src 'self'; object-src 'none'; frame-ancestors 'none'; base-uri 'self'; form-action 'self'"
  );
  res.setHeader("Cache-Control", "no-store, max-age=0");
}

function safeEqual(a, b) {
  const left = Buffer.from(String(a));
  const right = Buffer.from(String(b));
  if (left.length !== right.length) {
    return false;
  }
  return crypto.timingSafeEqual(left, right);
}

function requestIsAuthorized(req) {
  if (!BASIC_AUTH_ENABLED) {
    return true;
  }
  const raw = String(req.headers.authorization || "");
  if (!raw.startsWith("Basic ")) {
    return false;
  }
  let decoded = "";
  try {
    decoded = Buffer.from(raw.slice(6), "base64").toString("utf8");
  } catch (error) {
    return false;
  }
  const separatorIndex = decoded.indexOf(":");
  if (separatorIndex <= 0) {
    return false;
  }
  const user = decoded.slice(0, separatorIndex);
  const pass = decoded.slice(separatorIndex + 1);
  return safeEqual(user, BASIC_AUTH_USER) && safeEqual(pass, BASIC_AUTH_PASS);
}

function sendFile(res, filePath) {
  fs.readFile(filePath, (error, content) => {
    if (error) {
      writeSecurityHeaders(res, "text/plain; charset=utf-8");
      res.writeHead(error.code === "ENOENT" ? 404 : 500);
      res.end(error.code === "ENOENT" ? "404 Not Found" : "500 Internal Server Error");
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const mimeType = MIME_TYPES[ext] || "application/octet-stream";
    writeSecurityHeaders(res, mimeType);
    res.writeHead(200);
    res.end(content);
  });
}

const server = http.createServer((req, res) => {
  if (req.method !== "GET" && req.method !== "HEAD") {
    writeSecurityHeaders(res, "text/plain; charset=utf-8");
    res.writeHead(405);
    res.end("405 Method Not Allowed");
    return;
  }

  if (!requestIsAuthorized(req)) {
    writeSecurityHeaders(res, "text/plain; charset=utf-8");
    res.setHeader("WWW-Authenticate", "Basic realm=\"MIDP\", charset=\"UTF-8\"");
    res.writeHead(401);
    res.end("401 Unauthorized");
    return;
  }

  const urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
  const requested = urlPath === "/" ? "/index.html" : urlPath;

  // Prevent path traversal outside the app root.
  const safePath = path.normalize(requested).replace(/^(\.\.[/\\])+/, "");
  const absolutePath = path.join(ROOT, safePath);

  if (!absolutePath.startsWith(ROOT)) {
    writeSecurityHeaders(res, "text/plain; charset=utf-8");
    res.writeHead(403);
    res.end("403 Forbidden");
    return;
  }

  sendFile(res, absolutePath);
});

server.listen(PORT, () => {
  // eslint-disable-next-line no-console
  console.log(`MIDP web running on port ${PORT}`);
});
