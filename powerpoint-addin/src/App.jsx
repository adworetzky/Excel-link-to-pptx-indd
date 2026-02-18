/**
 * powerpoint-addin/src/App.jsx
 *
 * Root component for the PowerPoint task pane add-in.
 * Manages panel state machine:
 *   NO_SELECTION | NOT_LINKED | LINKED | MULTI_SELECTION
 *
 * Auto-update is triggered by DocumentSelectionChanged events with a
 * 2-second debounce.
 */
import React, { useState, useEffect, useCallback, useRef } from "react";

import PanelNoSelection from "./ui/PanelNoSelection.jsx";
import PanelNotLinked from "./ui/PanelNotLinked.jsx";
import PanelLinked from "./ui/PanelLinked.jsx";
import LinkedItemsBar from "./ui/LinkedItemsBar.jsx";

import { createLink, readLink, removeLink, setAutoUpdate } from "./linker.js";
import { updateLinkedItem, updateAllAutoLinks, repointExcelFile } from "./updater.js";
import { getAllLinkedShapes } from "./storage.js";
import { setHighlightActive } from "./highlight.js";
import { readCell } from "@datalink/shared";

const PANEL_STATE = {
  NO_SELECTION: "NO_SELECTION",
  NOT_LINKED: "NOT_LINKED",
  LINKED: "LINKED",
  MULTI_SELECTION: "MULTI_SELECTION",
};

export default function App() {
  const [panelState, setPanelState] = useState(PANEL_STATE.NO_SELECTION);
  const [currentLink, setCurrentLink] = useState(null);
  const [linkedCount, setLinkedCount] = useState(0);
  const [highlightActive, setHighlightActiveState] = useState(false);
  const [currentShapeId, setCurrentShapeId] = useState(null);
  const autoUpdateDebounceRef = useRef(null);
  const [ready, setReady] = useState(false);

  // ── Selection refresh ──────────────────────────────────────────────────────

  const refreshSelection = useCallback(async () => {
    try {
      await PowerPoint.run(async (context) => {
        const selection = context.presentation.getSelectedShapes();
        selection.load("items");
        await context.sync();

        const items = selection.items;

        if (!items || items.length === 0) {
          setPanelState(PANEL_STATE.NO_SELECTION);
          setCurrentLink(null);
          setCurrentShapeId(null);
        } else if (items.length > 1) {
          setPanelState(PANEL_STATE.MULTI_SELECTION);
          setCurrentLink(null);
          setCurrentShapeId(null);
        } else {
          const shape = items[0];
          shape.tags.load("items");
          await context.sync();

          const link = await readLink(shape, context);
          setCurrentShapeId(shape.id);
          if (link) {
            // Check for broken file
            const { validateFile } = await import("@datalink/shared");
            const fileOk = await validateFile(link.excelFile);
            setCurrentLink({ ...link, _fileBroken: !fileOk });
            setPanelState(PANEL_STATE.LINKED);
          } else {
            setCurrentLink(null);
            setPanelState(PANEL_STATE.NOT_LINKED);
          }
        }

        // Count all linked shapes
        const linked = await getAllLinkedShapes(context);
        setLinkedCount(linked.length);
      });
    } catch {
      // getSelectedShapes may not be available in all requirement sets
      setPanelState(PANEL_STATE.NO_SELECTION);
    }
  }, []);

  // ── Auto-update (debounced) ────────────────────────────────────────────────

  const handleSelectionChanged = useCallback(() => {
    // Debounce: refresh selection + potentially run auto-updates
    if (autoUpdateDebounceRef.current) {
      clearTimeout(autoUpdateDebounceRef.current);
    }
    autoUpdateDebounceRef.current = setTimeout(async () => {
      await refreshSelection();
      // Run auto-updates
      try {
        await PowerPoint.run(async (context) => {
          await updateAllAutoLinks(context);
        });
      } catch {}
    }, 2000);
  }, [refreshSelection]);

  // ── Init ───────────────────────────────────────────────────────────────────

  useEffect(() => {
    if (!ready) return;

    refreshSelection();

    // Register selection-change handler for auto-update trigger
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      handleSelectionChanged
    );

    return () => {
      Office.context.document.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        { handler: handleSelectionChanged }
      );
      if (autoUpdateDebounceRef.current) {
        clearTimeout(autoUpdateDebounceRef.current);
      }
    };
  }, [ready, refreshSelection, handleSelectionChanged]);

  // Called from taskpane.js once Office.onReady() resolves
  useEffect(() => {
    window.__datalinkSetReady = () => setReady(true);
    return () => { delete window.__datalinkSetReady; };
  }, []);

  // ── Handlers ──────────────────────────────────────────────────────────────

  async function handleLink({ filePath, sheet, cell }) {
    const { value } = await readCell(filePath, sheet, cell);
    await PowerPoint.run(async (context) => {
      const selection = context.presentation.getSelectedShapes();
      selection.load("items");
      await context.sync();
      if (!selection.items || selection.items.length !== 1)
        throw new Error("Select exactly one shape to link.");
      const shape = selection.items[0];
      await createLink(shape, filePath, sheet, cell, value, false, context);
    });
    await refreshSelection();
  }

  async function handleUpdateNow() {
    if (!currentLink) throw new Error("No link");
    let result;
    await PowerPoint.run(async (context) => {
      result = await updateLinkedItem(context, currentLink, true);
    });
    await refreshSelection();
    return result;
  }

  async function handleRemoveLink() {
    await PowerPoint.run(async (context) => {
      const selection = context.presentation.getSelectedShapes();
      selection.load("items");
      await context.sync();
      if (!selection.items || selection.items.length !== 1) return;
      const shape = selection.items[0];
      shape.tags.load("items");
      await context.sync();
      await removeLink(shape, context);
    });
    await refreshSelection();
  }

  async function handleToggleAutoUpdate(enabled) {
    await PowerPoint.run(async (context) => {
      const selection = context.presentation.getSelectedShapes();
      selection.load("items");
      await context.sync();
      if (!selection.items || selection.items.length !== 1) return;
      const shape = selection.items[0];
      shape.tags.load("items");
      await context.sync();
      await setAutoUpdate(shape, enabled, context);
    });
    await refreshSelection();
  }

  async function handleRepointFile() {
    // In Office JS taskpane we use an HTML file input
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".xlsx";
    input.onchange = async () => {
      const file = input.files[0];
      if (!file) return;
      const newPath = file.path || file.name;
      if (!currentLink) return;
      // Rewrite stored paths for every link that used the old file
      await PowerPoint.run(async (context) => {
        await repointExcelFile(context, currentLink.excelFile, newPath);
      });
      // Force-fetch the new value immediately so the shape is up to date
      const updatedLink = { ...currentLink, excelFile: newPath, _fileBroken: false };
      await PowerPoint.run(async (context) => {
        await updateLinkedItem(context, updatedLink, true);
      });
      await refreshSelection();
    };
    input.click();
  }

  async function handleToggleHighlight() {
    const next = !highlightActive;
    setHighlightActiveState(next);
    try {
      await PowerPoint.run(async (context) => {
        await setHighlightActive(context, next);
      });
    } catch {}
  }

  // ── Render ─────────────────────────────────────────────────────────────────

  if (!ready) {
    return (
      <div style={{ padding: 16, fontSize: 12, color: "#767676" }}>
        Initializing…
      </div>
    );
  }

  function renderPanel() {
    switch (panelState) {
      case PANEL_STATE.NO_SELECTION:
        return <PanelNoSelection />;
      case PANEL_STATE.MULTI_SELECTION:
        return (
          <div style={{ padding: 16 }}>
            <p style={{ fontSize: 12, color: "#3c3c3c" }}>
              Select a single item to manage its link.
            </p>
          </div>
        );
      case PANEL_STATE.NOT_LINKED:
        return <PanelNotLinked onLink={handleLink} />;
      case PANEL_STATE.LINKED:
        return (
          <PanelLinked
            link={currentLink}
            onUpdateNow={handleUpdateNow}
            onRemoveLink={handleRemoveLink}
            onToggleAutoUpdate={handleToggleAutoUpdate}
            onRepointFile={handleRepointFile}
          />
        );
      default:
        return null;
    }
  }

  return (
    <div style={styles.root}>
      <div style={styles.scrollArea}>{renderPanel()}</div>
      <LinkedItemsBar
        count={linkedCount}
        highlightActive={highlightActive}
        onToggleHighlight={handleToggleHighlight}
      />
    </div>
  );
}

const styles = {
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    background: "#f4f4f4",
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
    fontSize: 12,
    color: "#1e1e1e",
  },
  scrollArea: {
    flex: 1,
    overflowY: "auto",
    paddingBottom: 44,
  },
};
