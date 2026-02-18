# DataLink

DataLink links text frames in Adobe InDesign and text shapes in Microsoft PowerPoint to individual cells in a local Excel (`.xlsx`) file. When the Excel data changes, the linked text updates automatically — with no visible tokens, underlines, or formatting changes in the documents.

## Features

- Link InDesign text frames or PowerPoint text shapes to Excel cells
- Auto-update when returning to the application
- Number formatting from Excel is preserved (currency, percentages, dates)
- Highlight mode shows all linked items with a semi-transparent overlay
- All link metadata stored invisibly inside the document (no external files)
- Broken-link detection with one-click repoint when an Excel file is moved

## Requirements

- Node.js 16+ (for build and local server) — [nodejs.org](https://nodejs.org)
- Adobe InDesign CC 2021+ (UXP API 6.0+)
- Microsoft PowerPoint desktop (Windows, Microsoft 365)

---

## Quick Setup

### Windows (one command)

Open PowerShell in the project folder and run:

```powershell
powershell -ExecutionPolicy Bypass -File install.ps1
```

This will:
1. Check that Node.js is installed
2. Install npm dependencies and build both plugins
3. Copy the PowerPoint manifest to the Office sideload folder
4. Add the DataLink server to your Windows **Startup** folder so it runs automatically at every login
5. Start the server immediately and print InDesign load instructions

No admin rights required.

### macOS (one command)

```bash
bash install.sh
```

This will build both plugins and install a **launchd agent** that starts the server automatically at login. Follow the printed instructions to load the InDesign plugin.

---

## After Installation

### InDesign

Load the plugin once via UXP Developer Tools — this is a one-time step:

1. Open InDesign CC 2021+
2. **Plugins → UXP Developer Tools → Load Plugin**
3. Navigate to `indesign-plugin/manifest.json`
4. The **DataLink** panel appears in your panels list

The plugin reloads automatically on subsequent launches.

### PowerPoint (Windows)

The add-in server must be running at `http://localhost:3000` before opening PowerPoint.

- **Automatic:** The server starts at Windows login via the Startup script installed by `install.ps1`
- **Manual:** Double-click **"Start DataLink Server"** on your Desktop, or run `node powerpoint-addin/serve.js`

Then in PowerPoint: **Insert → My Add-ins → DataLink**

---

## Manual Build (developers)

```bash
# Install dependencies
npm install

# Build InDesign plugin  → indesign-plugin/dist/bundle.js
npm run build:indesign

# Build PowerPoint add-in → powerpoint-addin/dist/taskpane.bundle.js
npm run build:powerpoint

# Run shared module tests (14 assertions)
npm test

# Start the PowerPoint server manually
node powerpoint-addin/serve.js
```

---

## Uninstall

**PowerPoint add-in:**
- Delete `%APPDATA%\Microsoft\Office\16\Wef\DataLink.xml`
- Delete `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\DataLink-Server.vbs`

**InDesign plugin:**
- Plugins → UXP Developer Tools → unload the DataLink plugin

---

## Project Structure

See `CLAUDE.md` for the full project specification and architecture.
