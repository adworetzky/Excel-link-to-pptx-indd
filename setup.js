#!/usr/bin/env node
/**
 * DataLink cross-platform setup script.
 *
 * Works on Windows, macOS, and Linux with no dependencies beyond Node.js itself.
 *
 *   node setup.js          ← the one command
 *   npm run setup          ← alias via package.json
 */

"use strict";

const { execSync }  = require("child_process");
const fs            = require("fs");
const os            = require("os");
const path          = require("path");

// ── ANSI colours (disabled on Windows unless the terminal supports them) ─────
const isWin    = process.platform === "win32";
const isMac    = process.platform === "darwin";
const useColor = !isWin || process.env.WT_SESSION || process.env.TERM_PROGRAM;

const C = useColor
  ? { cyan: "\x1b[36m", green: "\x1b[32m", yellow: "\x1b[33m",
      red: "\x1b[31m", gray: "\x1b[90m", bold: "\x1b[1m", reset: "\x1b[0m" }
  : Object.fromEntries(
      ["cyan","green","yellow","red","gray","bold","reset"].map(k => [k, ""])
    );

const step  = (msg) => console.log(`\n${C.cyan}${msg}${C.reset}`);
const ok    = (msg) => console.log(`  ${C.green}OK${C.reset}  ${msg}`);
const fail  = (msg) => { console.error(`  ${C.red}!!${C.reset}  ${msg}`); process.exit(1); };
const info  = (msg) => console.log(`      ${C.gray}${msg}${C.reset}`);
const print = (msg) => console.log(msg);

// ── Resolve project root (wherever this script lives) ────────────────────────
const ROOT = __dirname;

// ── Helper: run a shell command, streaming output ────────────────────────────
function run(cmd, opts = {}) {
  try {
    execSync(cmd, { stdio: "inherit", cwd: ROOT, ...opts });
  } catch {
    fail(`Command failed: ${cmd}`);
  }
}

// ── Helper: ensure a directory exists ────────────────────────────────────────
function mkdirp(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

// ─────────────────────────────────────────────────────────────────────────────
// 1. Check Node.js version
// ─────────────────────────────────────────────────────────────────────────────
step("Checking prerequisites...");

const [major] = process.versions.node.split(".").map(Number);
if (major < 16) {
  fail(`Node.js 16+ required. You are running ${process.version}. Get it from https://nodejs.org`);
}
ok(`Node.js ${process.version}`);

// ─────────────────────────────────────────────────────────────────────────────
// 2. Install npm dependencies
// ─────────────────────────────────────────────────────────────────────────────
step("Installing npm dependencies...");
run("npm install");
ok("Dependencies installed");

// ─────────────────────────────────────────────────────────────────────────────
// 3. Build both plugins
// ─────────────────────────────────────────────────────────────────────────────
step("Building InDesign plugin...");
run("npm run build:indesign");
ok("indesign-plugin/dist/bundle.js built");

step("Building PowerPoint add-in...");
run("npm run build:powerpoint");
ok("powerpoint-addin/dist/taskpane.bundle.js built");

// ─────────────────────────────────────────────────────────────────────────────
// 4. Platform-specific install
// ─────────────────────────────────────────────────────────────────────────────
const serveJs = path.join(ROOT, "powerpoint-addin", "serve.js");
const nodeExe = process.execPath; // absolute path to the running node binary

if (isWin) {
  // ── Windows ────────────────────────────────────────────────────────────────

  step("Installing PowerPoint add-in (Windows)...");

  const wefDir = path.join(
    os.homedir(), "AppData", "Roaming", "Microsoft", "Office", "16", "Wef"
  );
  mkdirp(wefDir);

  const manifestSrc  = path.join(ROOT, "powerpoint-addin", "manifest.xml");
  const manifestDest = path.join(wefDir, "DataLink.xml");
  fs.copyFileSync(manifestSrc, manifestDest);
  ok(`Manifest → ${manifestDest}`);

  // Silent VBScript launcher in Startup folder (no admin needed)
  step("Adding server to Windows Startup...");

  const startupDir = path.join(
    os.homedir(),
    "AppData", "Roaming", "Microsoft", "Windows",
    "Start Menu", "Programs", "Startup"
  );
  mkdirp(startupDir);

  const vbs = [
    `' DataLink server auto-start — delete to disable`,
    `Dim oShell`,
    `Set oShell = CreateObject("WScript.Shell")`,
    `oShell.Run """${nodeExe}"" ""${serveJs}""", 0, False`,
  ].join("\r\n");

  const vbsDest = path.join(startupDir, "DataLink-Server.vbs");
  fs.writeFileSync(vbsDest, vbs, "utf8");
  ok(`Auto-start script → ${vbsDest}`);
  info("The server will start silently at each Windows login.");

  // Desktop shortcut (.lnk) via WScript.Shell COM
  step("Creating Desktop shortcut...");
  try {
    const desktopPath = path.join(os.homedir(), "Desktop", "Start DataLink Server.lnk");
    // PowerShell one-liner to create a .lnk — no separate .ps1 file needed
    const ps = [
      `$s=(New-Object -COM WScript.Shell).CreateShortcut('${desktopPath}');`,
      `$s.TargetPath='${nodeExe}';`,
      `$s.Arguments='"${serveJs}"';`,
      `$s.Description='Start DataLink local server';`,
      `$s.Save()`,
    ].join(" ");
    execSync(`powershell -NoProfile -Command "${ps}"`, { stdio: "ignore" });
    ok(`Desktop shortcut → ${desktopPath}`);
  } catch {
    info("Could not create desktop shortcut (non-critical).");
  }

  // Start server now (detached, hidden)
  step("Starting DataLink server...");
  const { spawn } = require("child_process");
  const child = spawn(nodeExe, [serveJs], {
    detached: true,
    stdio:    "ignore",
    windowsHide: true,
  });
  child.unref();
  ok("Server started on http://localhost:3000");

} else if (isMac) {
  // ── macOS ──────────────────────────────────────────────────────────────────

  step("Installing launchd agent (macOS)...");

  const launchAgentsDir = path.join(os.homedir(), "Library", "LaunchAgents");
  const logDir          = path.join(os.homedir(), "Library", "Logs", "DataLink");
  mkdirp(launchAgentsDir);
  mkdirp(logDir);

  const plistDest = path.join(launchAgentsDir, "com.datalink.server.plist");

  const plist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>com.datalink.server</string>

  <key>ProgramArguments</key>
  <array>
    <string>${nodeExe}</string>
    <string>${serveJs}</string>
  </array>

  <key>RunAtLoad</key>
  <true/>

  <key>KeepAlive</key>
  <true/>

  <key>StandardOutPath</key>
  <string>${logDir}/server.log</string>

  <key>StandardErrorPath</key>
  <string>${logDir}/server-error.log</string>

  <key>WorkingDirectory</key>
  <string>${path.join(ROOT, "powerpoint-addin")}</string>
</dict>
</plist>`;

  fs.writeFileSync(plistDest, plist, "utf8");

  // Unload first in case it was already registered
  try { execSync(`launchctl unload "${plistDest}" 2>/dev/null`, { stdio: "ignore" }); } catch {}
  try {
    execSync(`launchctl load "${plistDest}"`, { stdio: "ignore" });
    ok(`launchd agent installed → ${plistDest}`);
    ok("Server started and will auto-start at every login");
    info(`Logs: ${logDir}/server.log`);
  } catch {
    info("Could not load launchd agent (try: launchctl load \"" + plistDest + "\")");
  }

} else {
  // ── Linux ──────────────────────────────────────────────────────────────────

  step("Starting DataLink server (Linux)...");

  const logFile = path.join(ROOT, "powerpoint-addin", "server.log");
  const { spawn } = require("child_process");
  const child = spawn(nodeExe, [serveJs], {
    detached: true,
    stdio:    ["ignore", fs.openSync(logFile, "a"), fs.openSync(logFile, "a")],
  });
  child.unref();

  // Give it a moment to start or fail
  setTimeout(() => {
    try {
      process.kill(child.pid, 0); // throws if not running
      ok(`Server started (PID ${child.pid}) on http://localhost:3000`);
      info(`Log: ${logFile}`);
    } catch {
      fail("Server failed to start. Check: " + logFile);
    }

    printSummary();
  }, 600);

  return; // async path — printSummary called in callback above
}

printSummary();

// ─────────────────────────────────────────────────────────────────────────────
function printSummary() {
  const manifestPath = path.join(ROOT, "indesign-plugin", "manifest.json");

  console.log(`\n${C.green}===========================================${C.reset}`);
  console.log(`${C.green}  SETUP COMPLETE${C.reset}`);
  console.log(`${C.green}===========================================${C.reset}`);

  console.log(`\n${C.cyan}INDESIGN PLUGIN${C.reset}`);
  print("  Load the plugin once via UXP Developer Tools:");
  print("");
  print("    1. Open InDesign CC 2021+");
  print("    2. Plugins → UXP Developer Tools → Load Plugin");
  print(`    3. Select: ${C.yellow}${manifestPath}${C.reset}`);
  print("    4. The DataLink panel will appear in your panels list.");
  print("");

  if (isWin) {
    console.log(`${C.cyan}POWERPOINT ADD-IN${C.reset}`);
    print("  Manifest installed. Server is running on http://localhost:3000.");
    print("  It will start automatically at every Windows login.");
    print("");
    print("  In PowerPoint: Insert → My Add-ins → DataLink");
    print("");
    console.log(`${C.cyan}UNINSTALL${C.reset}`);
    print("  PowerPoint: delete %APPDATA%\\Microsoft\\Office\\16\\Wef\\DataLink.xml");
    print("  Auto-start: delete the DataLink-Server.vbs from your Startup folder");
  } else if (isMac) {
    console.log(`${C.cyan}POWERPOINT${C.reset}`);
    print("  PowerPoint for Mac does not support the DesktopRuntime requirement.");
    print("  Use a Windows machine for the PowerPoint add-in.");
    print("");
    console.log(`${C.cyan}UNINSTALL SERVER${C.reset}`);
    print("  launchctl unload ~/Library/LaunchAgents/com.datalink.server.plist");
    print("  rm ~/Library/LaunchAgents/com.datalink.server.plist");
  }

  print("");
}
