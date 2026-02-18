# CLAUDE.md — DataLink Project Briefing

> Drop this file in the root of the `datalink/` repository. Claude Code will read it automatically on startup.

-----

## What This Project Is

**DataLink** is a two-part productivity tool that lets users link text frames in Adobe InDesign and text shapes in Microsoft PowerPoint to individual cells in a local Excel (`.xlsx`) file. When the Excel data changes, the linked text in InDesign or PowerPoint updates automatically — with no visible tokens, underlines, or formatting changes in the documents.

The project consists of:

1. A **shared Excel reader module** (Node.js, used by both plugins)
1. An **InDesign UXP plugin** (task pane panel)
1. A **PowerPoint Office JS add-in** (task pane panel, sideloaded)

-----

## Absolute Constraints — Read Before Writing Any Code

- **Never modify formatting when updating linked text.** Only the text content changes. Font, size, color, paragraph style, character style — all untouched. In InDesign this means using `frame.parentStory.contents`. In PowerPoint this means `shape.textFrame.textRange.text`. Do not use any API that replaces or re-creates the text run.
- **Linked items must look identical to non-linked items** in both applications. No underlines, no colored borders, no tokens, no visible indicators of any kind in normal view.
- **Highlight mode is a temporary overlay only**, not a modification to the document. It is toggled on/off by the user and never persists to the saved file.
- **Links are stored as invisible metadata** embedded in the document itself — InDesign frame labels, PowerPoint shape tags. No external database, no sidecar files.
- **Excel files are referenced by absolute local file path.** No cloud sync, no Microsoft Graph API. SheetJS reads the `.xlsx` file directly from disk.
- **Number formatting from Excel is carried through.** If a cell is formatted as `$#,##0.00`, the string written to InDesign/PowerPoint is `$1,234.56`, not `1234.56`. Use `XLSX.utils.format_cell()` for this.
- **The PowerPoint add-in must run in the desktop Office runtime.** It is sideloaded via a local manifest. It must not require any installation, admin rights, or AppSource publishing. Enforce this with `Requirements: DesktopRuntime` in the manifest.
- **Both plugins are single-bundle webpack builds.** No `node_modules` folder is shipped. All dependencies are bundled.

-----

## Repository Structure

```
datalink/
├── CLAUDE.md                        ← this file
├── package.json                     ← npm workspaces root
├── README.md
│
├── shared/
│   ├── package.json
│   ├── index.js                     ← main export
│   ├── excel-reader.js              ← xlsx open, cell read, format
│   └── test/
│       └── excel-reader.test.js     ← CLI test harness
│
├── indesign-plugin/
│   ├── manifest.json                ← UXP plugin manifest
│   ├── package.json
│   ├── webpack.config.js
│   ├── index.html                   ← plugin panel shell
│   └── src/
│       ├── index.js                 ← UXP entry point, mounts React
│       ├── App.jsx                  ← root component, panel state machine
│       ├── ui/
│       │   ├── PanelNoSelection.jsx
│       │   ├── PanelNotLinked.jsx
│       │   ├── PanelLinked.jsx
│       │   └── LinkedItemsBar.jsx   ← bottom bar: count + highlight toggle
│       ├── linker.js                ← create, read, remove link on a frame
│       ├── updater.js               ← read Excel value, write to frame text
│       ├── storage.js               ← encode/decode link metadata in frame.label
│       └── highlight.js            ← draw/clear temporary overlay on linked frames
│
└── powerpoint-addin/
    ├── manifest.xml                 ← Office JS sideload manifest
    ├── package.json
    ├── webpack.config.js
    ├── taskpane.html                ← add-in panel shell
    └── src/
        ├── taskpane.js              ← Office JS entry point, mounts React
        ├── App.jsx                  ← root component, panel state machine
        ├── ui/
        │   ├── PanelNoSelection.jsx
        │   ├── PanelNotLinked.jsx
        │   ├── PanelLinked.jsx
        │   └── LinkedItemsBar.jsx
        ├── linker.js
        ├── updater.js
        ├── storage.js               ← encode/decode link metadata in shape.tags
        └── highlight.js            ← apply/clear temporary highlight on shapes
```

-----

## Data Model

Every link is described by this object. It is serialized to JSON and stored invisibly in the document.

```typescript
interface DataLink {
  linkId: string;          // UUID v4, unique per link
  excelFile: string;       // absolute path e.g. "C:/Users/adam/data.xlsx"
  sheet: string;           // sheet name e.g. "Sheet1"
  cell: string;            // A1 notation e.g. "B4"
  lastValue: string;       // last formatted string written to the document
  lastUpdated: string;     // ISO 8601 timestamp
  autoUpdate: boolean;     // whether to update on document activation
}
```

### InDesign storage

```js
// Write
frame.label = JSON.stringify(dataLink);

// Read (scan all page items)
app.activeDocument.pages.everyItem().getElements().forEach(page => {
  page.allPageItems.forEach(item => {
    try {
      const meta = JSON.parse(item.label);
      if (meta.linkId) { /* it's a DataLink */ }
    } catch {}
  });
});
```

### PowerPoint storage

```js
// Write
await shape.tags.add("DATALINK", JSON.stringify(dataLink));

// Read (scan all shapes on all slides)
const slides = context.presentation.slides;
slides.load("items");
await context.sync();
for (const slide of slides.items) {
  slide.shapes.load("items/tags");
  await context.sync();
  for (const shape of slide.shapes.items) {
    const tag = shape.tags.getItemOrNullObject("DATALINK");
    await context.sync();
    if (!tag.isNullObject) { /* it's a DataLink */ }
  }
}
```

-----

## Core Logic: Update a Linked Item

```
UpdateLinkedItem(linkMeta, forceUpdate = false):

  1. Open linkMeta.excelFile with SheetJS
  2. Get the cell at linkMeta.sheet + linkMeta.cell
  3. Apply the cell's number format string via XLSX.utils.format_cell()
  4. If newValue === linkMeta.lastValue AND forceUpdate === false → return (no change)
  5. Find the frame/shape in the document by scanning for linkMeta.linkId
  6. Write newValue to the text content ONLY:
     - InDesign:   frame.parentStory.contents = newValue
     - PowerPoint: shape.textFrame.textRange.text = newValue
  7. Update linkMeta.lastValue = newValue
  8. Update linkMeta.lastUpdated = new Date().toISOString()
  9. Write updated linkMeta back to frame.label / shape.tags
```

-----

## Panel UI — State Machine

|State            |Condition                                  |Panel Shows                                                                       |
|-----------------|-------------------------------------------|----------------------------------------------------------------------------------|
|`NO_SELECTION`   |Nothing selected                           |Instructional message                                                             |
|`NOT_LINKED`     |Frame/shape selected, no DataLink metadata |File browse, sheet picker, cell input, Link button                                |
|`LINKED`         |Frame/shape selected, has DataLink metadata|Source info, last value, last updated, Update Now, Remove Link, Auto-update toggle|
|`MULTI_SELECTION`|Multiple items selected                    |"Select a single item to manage its link"                                         |

-----

## Technology Stack

|Layer                  |Technology                           |
|-----------------------|-------------------------------------|
|Excel parsing          |SheetJS (`xlsx` npm package)         |
|InDesign plugin        |UXP, React 18, webpack               |
|PowerPoint add-in      |Office JS 1.3+, React 18, webpack    |
|Unique IDs             |`uuid` v4                            |
|Styling                |CSS Modules, no external UI framework|
|Testing (shared module)|Node.js assert, no framework needed  |
