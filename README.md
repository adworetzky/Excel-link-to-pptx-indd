# DataLink

DataLink links text frames in Adobe InDesign and text shapes in Microsoft PowerPoint to individual cells in a local Excel (`.xlsx`) file. When the Excel data changes, the linked text updates automatically — with no visible tokens, underlines, or formatting changes in the documents.

## Features

- Link InDesign text frames or PowerPoint text shapes to Excel cells
- Auto-update when returning to the application
- Number formatting from Excel is preserved (currency, percentages, dates)
- Highlight mode shows all linked items with a semi-transparent overlay
- All link metadata stored invisibly inside the document (no external files)

## Requirements

- Adobe InDesign CC 2021+ (UXP API 6.0+)
- Microsoft PowerPoint (desktop, Windows) with Office JS 1.3+
- Node.js 16+ for building

## Build

```bash
# Install all dependencies
npm install

# Build InDesign plugin
cd indesign-plugin && npm run build

# Build PowerPoint add-in
cd powerpoint-addin && npm run build

# Run shared module tests
cd shared && node test/excel-reader.test.js
```

## Installation

### InDesign Plugin

1. Open InDesign CC 2021+
2. Plugins → UXP Developer Tools → Load Plugin
3. Point to `indesign-plugin/manifest.json`

### PowerPoint Add-in (Windows)

1. Copy `powerpoint-addin/manifest.xml` to `%APPDATA%\Microsoft\Office\16\Wef\`
2. Open PowerPoint → Insert → My Add-ins → DataLink

## Project Structure

See `CLAUDE.md` for the full project specification and architecture.
