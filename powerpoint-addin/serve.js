#!/usr/bin/env node
/**
 * powerpoint-addin/serve.js
 *
 * Minimal static file server for the DataLink PowerPoint add-in.
 * Serves the powerpoint-addin/ directory over http://localhost:3000.
 *
 * No external dependencies — pure Node.js built-ins only.
 * Office JS desktop client allows plain HTTP from localhost for sideloaded add-ins.
 *
 * Usage:
 *   node serve.js            # start on default port 3000
 *   PORT=4000 node serve.js  # override port
 */

"use strict";

const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = parseInt(process.env.PORT || "3000", 10);
const ROOT = __dirname; // serves from powerpoint-addin/

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js":   "application/javascript; charset=utf-8",
  ".css":  "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png":  "image/png",
  ".jpg":  "image/jpeg",
  ".ico":  "image/x-icon",
  ".txt":  "text/plain; charset=utf-8",
  ".xml":  "application/xml; charset=utf-8",
};

const server = http.createServer((req, res) => {
  // Strip query string and decode
  let urlPath;
  try {
    urlPath = decodeURIComponent(req.url.split("?")[0]);
  } catch {
    urlPath = req.url.split("?")[0];
  }

  // Default route
  if (urlPath === "/" || urlPath === "") {
    urlPath = "/taskpane.html";
  }

  // Resolve to an absolute file path
  const filePath = path.normalize(path.join(ROOT, urlPath));

  // Prevent directory traversal outside ROOT
  if (!filePath.startsWith(ROOT + path.sep) && filePath !== ROOT) {
    res.writeHead(403, { "Content-Type": "text/plain" });
    res.end("403 Forbidden");
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      if (err.code === "ENOENT") {
        res.writeHead(404, { "Content-Type": "text/plain" });
        res.end(`404 Not Found: ${urlPath}`);
      } else {
        res.writeHead(500, { "Content-Type": "text/plain" });
        res.end("500 Internal Server Error");
      }
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || "application/octet-stream";

    res.writeHead(200, {
      "Content-Type": contentType,
      // Allow the Office JS taskpane iframe to load resources
      "Access-Control-Allow-Origin": "*",
      // Disable caching so updates to the bundle take effect immediately
      "Cache-Control": "no-store",
    });
    res.end(data);
  });
});

server.on("error", (err) => {
  if (err.code === "EADDRINUSE") {
    console.error(
      `\nPort ${PORT} is already in use.\n` +
      "DataLink server may already be running — check your system tray or taskbar.\n" +
      `If not, kill the process using port ${PORT} and try again.`
    );
  } else {
    console.error("Server error:", err.message);
  }
  process.exit(1);
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`\nDataLink PowerPoint server running at http://localhost:${PORT}`);
  console.log(`Serving: ${ROOT}`);
  console.log("\nKeep this window open while using the add-in.");
  console.log("Press Ctrl+C to stop.\n");
});
