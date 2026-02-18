# DataLink

DataLink links text frames in Adobe InDesign and text shapes in Microsoft PowerPoint to individual cells in a local Excel (`.xlsx`) file. When the Excel data changes, the linked text updates automatically — with no visible tokens, underlines, or formatting changes in the documents.

## Features

- Link InDesign text frames or PowerPoint text shapes to Excel cells
- Auto-update when returning to the application
- Number formatting from Excel is preserved (currency, percentages, dates)
- Highlight mode shows all linked items with a semi-transparent overlay
- All link metadata stored invisibly inside the document (no external files)

## Requirements

| Requirement | Details |
|---|---|
| **Node.js 16+** | Build tool and local server — [nodejs.org](https://nodejs.org) |
| **Adobe InDesign CC 2021+** | UXP API 6.0+ required |
| **Microsoft PowerPoint** | **Windows desktop only** (Microsoft 365, Office 2019+). The add-in uses `DesktopRuntime` and cannot run in PowerPoint for Mac or PowerPoint Online. |

---

## Installation

### Windows — one command

Open PowerShell **in the project folder** and run:

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1
```

The installer will:

1. Verify Node.js is installed
2. Run `npm install` and build both plugins
3. Copy `manifest.xml` to `%APPDATA%\Microsoft\Office\16\Wef\DataLink.xml` (the Office sideload folder)
4. Add a silent launcher to your **Windows Startup folder** — the server starts automatically at every login
5. Create a **"Start DataLink Server"** shortcut on your Desktop for manual starts
6. Start the server immediately on `http://localhost:3000`
7. Print the InDesign plugin load path

No admin rights required at any step.

---

### macOS — one command

```bash
bash install.sh
```

The installer will:

1. Verify Node.js is installed
2. Run `npm install` and build both plugins
3. Install a **launchd user agent** (`~/Library/LaunchAgents/com.datalink.server.plist`) — the server starts automatically at every login
4. Start the server immediately on `http://localhost:3000`
5. Print the InDesign plugin load path

> **Note:** The PowerPoint add-in requires Windows desktop. On macOS, only the InDesign plugin is usable.

---

## First Use After Installation

### InDesign (Windows and macOS)

Loading the plugin via UXP Developer Tools is a **one-time manual step**:

1. Open InDesign CC 2021+
2. Go to **Plugins → UXP Developer Tools**
3. Click **Load Plugin**
4. Navigate to and select `indesign-plugin/manifest.json`
5. The **DataLink** panel will appear in your panels list

> To use the plugin in future sessions, repeat steps 2–4, or keep UXP Developer Tools open. The plugin does not auto-load between application relaunches when loaded this way.

### PowerPoint (Windows only)

The DataLink server (`http://localhost:3000`) must be running before you open PowerPoint.

- **Automatic (after running `install.ps1`):** The server starts silently at every Windows login.
- **Manual start:** Double-click **"Start DataLink Server"** on your Desktop, or run:
  ```
  node powerpoint-addin/serve.js
  ```
- **If port 3000 is in use by another app:**
  ```
  PORT=3001 node powerpoint-addin/serve.js
  ```
  Then edit `manifest.xml` to replace `localhost:3000` with `localhost:3001` before re-copying it.

To open the add-in in PowerPoint: **Insert → My Add-ins → DataLink**

---

## Developer Build Reference

```bash
# Install all dependencies (run once, or after pulling changes)
npm install

# Build both plugins
npm run build

# Build individually
npm run build:indesign    # → indesign-plugin/dist/bundle.js
npm run build:powerpoint  # → powerpoint-addin/dist/taskpane.bundle.js

# Watch mode (rebuilds on file save)
npm run dev --workspace=indesign-plugin
npm run dev --workspace=powerpoint-addin

# Run shared module tests (14 assertions, no framework)
npm test

# Start the PowerPoint server manually
node powerpoint-addin/serve.js
```

---

## Uninstall

**Windows:**

| What | How |
|---|---|
| Stop server auto-start | Delete `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\DataLink-Server.vbs` |
| Remove PowerPoint add-in | Delete `%APPDATA%\Microsoft\Office\16\Wef\DataLink.xml` |
| Remove Desktop shortcut | Delete `%USERPROFILE%\Desktop\Start DataLink Server.lnk` |
| Remove InDesign plugin | Plugins → UXP Developer Tools → unload DataLink |

**macOS:**

```bash
launchctl unload ~/Library/LaunchAgents/com.datalink.server.plist
rm ~/Library/LaunchAgents/com.datalink.server.plist
```

Then unload the InDesign plugin from Plugins → UXP Developer Tools.

---

## Project Structure

See `CLAUDE.md` for the full technical specification and architecture.
