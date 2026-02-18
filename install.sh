#!/usr/bin/env bash
# DataLink one-click installer for macOS / Linux.
#
# 1. Verifies Node.js is installed
# 2. Installs npm dependencies and builds both plugins
# 3. Installs a launchd agent (macOS) so the PowerPoint server
#    starts automatically at login — no admin rights required
# 4. Prints InDesign load instructions
#
# Usage (from the project root):
#   bash install.sh

set -euo pipefail

# ── Colours ───────────────────────────────────────────────────────────────────
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'
CYAN='\033[0;36m'; GRAY='\033[0;37m'; BOLD='\033[1m'; NC='\033[0m'

ok()   { echo -e "  ${GREEN}OK${NC}  $*"; }
fail() { echo -e "  ${RED}!!${NC}  $*" >&2; exit 1; }
info() { echo -e "      ${GRAY}$*${NC}"; }
step() { echo -e "\n${CYAN}$*${NC}"; }

# ── Project root ──────────────────────────────────────────────────────────────
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo ""
echo -e "${CYAN}==========================================${NC}"
echo -e "${CYAN}  DataLink Installer${NC}"
echo -e "${CYAN}==========================================${NC}"

# ── 1. Check prerequisites ────────────────────────────────────────────────────
step "Checking prerequisites..."

if ! command -v node &>/dev/null; then
    fail "Node.js not found. Install Node.js 16+ from https://nodejs.org and re-run."
fi
NODE_VERSION=$(node --version)
ok "Node.js $NODE_VERSION found"

if ! command -v npm &>/dev/null; then
    fail "npm not found (should come with Node.js)."
fi
ok "npm found"

# ── 2. Install dependencies ───────────────────────────────────────────────────
step "Installing npm dependencies..."
npm install --silent
ok "Dependencies installed"

# ── 3. Build both plugins ─────────────────────────────────────────────────────
step "Building InDesign plugin..."
npm run build:indesign --silent
ok "indesign-plugin/dist/bundle.js built"

step "Building PowerPoint add-in..."
npm run build:powerpoint --silent
ok "powerpoint-addin/dist/taskpane.bundle.js built"

# ── 4. macOS auto-start via launchd ──────────────────────────────────────────
IS_MACOS=false
if [[ "$(uname)" == "Darwin" ]]; then
    IS_MACOS=true
fi

NODE_BIN="$(command -v node)"
SERVE_JS="$SCRIPT_DIR/powerpoint-addin/serve.js"

if $IS_MACOS; then
    step "Installing launchd agent for auto-start (macOS)..."

    LAUNCH_AGENTS_DIR="$HOME/Library/LaunchAgents"
    PLIST_DEST="$LAUNCH_AGENTS_DIR/com.datalink.server.plist"
    LOG_DIR="$HOME/Library/Logs/DataLink"
    mkdir -p "$LAUNCH_AGENTS_DIR" "$LOG_DIR"

    cat > "$PLIST_DEST" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>com.datalink.server</string>

  <key>ProgramArguments</key>
  <array>
    <string>$NODE_BIN</string>
    <string>$SERVE_JS</string>
  </array>

  <!-- Start at login -->
  <key>RunAtLoad</key>
  <true/>

  <!-- Restart if it crashes -->
  <key>KeepAlive</key>
  <true/>

  <key>StandardOutPath</key>
  <string>$LOG_DIR/server.log</string>

  <key>StandardErrorPath</key>
  <string>$LOG_DIR/server-error.log</string>

  <key>WorkingDirectory</key>
  <string>$SCRIPT_DIR/powerpoint-addin</string>
</dict>
</plist>
PLIST

    # Load it now (equivalent to "start immediately + enable auto-start")
    # Unload first in case it was already registered
    launchctl unload "$PLIST_DEST" 2>/dev/null || true
    launchctl load "$PLIST_DEST"

    ok "launchd agent installed: $PLIST_DEST"
    ok "Server started and will auto-start at every login"
    info "Logs: $LOG_DIR/server.log"
    info "To stop and remove: launchctl unload \"$PLIST_DEST\" && rm \"$PLIST_DEST\""

else
    # Linux: just start it in the background and print instructions
    step "Starting DataLink server in the background..."
    nohup "$NODE_BIN" "$SERVE_JS" >"$SCRIPT_DIR/powerpoint-addin/server.log" 2>&1 &
    SERVER_PID=$!
    sleep 0.5
    if kill -0 "$SERVER_PID" 2>/dev/null; then
        ok "Server started (PID $SERVER_PID) on http://localhost:3000"
        info "Log: $SCRIPT_DIR/powerpoint-addin/server.log"
    else
        fail "Server failed to start. Check server.log for details."
    fi

    echo ""
    echo -e "${YELLOW}Linux auto-start:${NC}"
    echo "  Add the following line to your ~/.profile or ~/.bashrc to auto-start"
    echo "  the server at login:"
    echo ""
    echo "    nohup $NODE_BIN $SERVE_JS &"
    echo ""
fi

# ── 5. Summary ────────────────────────────────────────────────────────────────
echo ""
echo -e "${GREEN}==========================================${NC}"
echo -e "${GREEN}  SETUP COMPLETE${NC}"
echo -e "${GREEN}==========================================${NC}"

echo ""
echo -e "${CYAN}INDESIGN PLUGIN${NC}"
echo "  Load the plugin once via UXP Developer Tools:"
echo ""
echo "    1. Open InDesign CC 2021+"
echo "    2. Plugins menu > UXP Developer Tools"
echo "    3. Click 'Load Plugin'"
echo -e "    4. Select: ${YELLOW}$SCRIPT_DIR/indesign-plugin/manifest.json${NC}"
echo "    5. The DataLink panel will appear in your panels list."
echo ""

if $IS_MACOS; then
    echo -e "${CYAN}POWERPOINT ADD-IN (macOS — requires Microsoft 365 desktop)${NC}"
    echo "  The server is running at http://localhost:3000 and will auto-start at login."
    echo ""
    echo "  To install the manifest:"
    echo "    1. Open PowerPoint"
    echo "    2. Insert > Get Add-ins > Manage My Add-ins > Upload My Add-in"
    echo -e "    3. Upload: ${YELLOW}$SCRIPT_DIR/powerpoint-addin/manifest.xml${NC}"
    echo ""
else
    echo -e "${CYAN}POWERPOINT ADD-IN (Windows required for sideload)${NC}"
    echo "  The server is running. On Windows, run install.ps1 to complete sideloading."
    echo ""
fi
