/**
 * indesign-plugin/src/App.jsx
 *
 * Root component. Manages the panel state machine:
 *   NO_SELECTION | NOT_LINKED | LINKED | MULTI_SELECTION
 *
 * Wires up the InDesign selection-change listener and the auto-update
 * afterActivate listener.
 */
import React, { useState, useEffect, useCallback, useRef } from "react";

import PanelNoSelection from "./ui/PanelNoSelection.jsx";
import PanelNotLinked from "./ui/PanelNotLinked.jsx";
import PanelLinked from "./ui/PanelLinked.jsx";
import LinkedItemsBar from "./ui/LinkedItemsBar.jsx";

import { createLink, readLink, removeLink, setAutoUpdate } from "./linker.js";
import { updateLinkedItem, updateAllAutoLinks, repointExcelFile } from "./updater.js";
import { getAllLinkedFrames, readLinkFromFrame } from "./storage.js";
import { setHighlightActive } from "./highlight.js";
import { readCell, getSheetNames } from "@datalink/shared";

const uxp = typeof require !== "undefined" ? require("uxp") : null;

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
  const [currentFrame, setCurrentFrame] = useState(null);

  // Keep refs for cleanup
  const selectionListenerRef = useRef(null);
  const activateListenerRef = useRef(null);

  // Get the InDesign app global
  function getApp() {
    try {
      return require("indesign").app;
    } catch {
      return typeof app !== "undefined" ? app : null;
    }
  }

  const refreshLinkedCount = useCallback(() => {
    const appObj = getApp();
    if (!appObj || !appObj.activeDocument) {
      setLinkedCount(0);
      return;
    }
    const all = getAllLinkedFrames(appObj);
    setLinkedCount(all.length);
  }, []);

  const refreshSelection = useCallback(() => {
    const appObj = getApp();
    if (!appObj || !appObj.activeDocument) {
      setPanelState(PANEL_STATE.NO_SELECTION);
      setCurrentFrame(null);
      setCurrentLink(null);
      return;
    }

    const selection = appObj.selection;
    if (!selection || selection.length === 0) {
      setPanelState(PANEL_STATE.NO_SELECTION);
      setCurrentFrame(null);
      setCurrentLink(null);
    } else if (selection.length > 1) {
      setPanelState(PANEL_STATE.MULTI_SELECTION);
      setCurrentFrame(null);
      setCurrentLink(null);
    } else {
      const frame = selection[0];
      setCurrentFrame(frame);
      const link = readLinkFromFrame(frame);
      if (link) {
        setCurrentLink(link);
        setPanelState(PANEL_STATE.LINKED);
      } else {
        setCurrentLink(null);
        setPanelState(PANEL_STATE.NOT_LINKED);
      }
    }

    refreshLinkedCount();
  }, [refreshLinkedCount]);

  // Auto-update handler (afterActivate)
  const handleActivate = useCallback(async () => {
    const appObj = getApp();
    if (!appObj) return;
    await updateAllAutoLinks(appObj);
    refreshSelection();
  }, [refreshSelection]);

  useEffect(() => {
    const appObj = getApp();
    if (!appObj) return;

    // Initial state
    refreshSelection();

    // Selection change listener
    try {
      selectionListenerRef.current = appObj.eventListeners.add(
        "afterSelectionChanged",
        refreshSelection
      );
    } catch {
      // Some UXP environments may differ
    }

    // afterActivate for auto-update
    try {
      activateListenerRef.current = appObj.eventListeners.add(
        "afterActivate",
        handleActivate
      );
    } catch {}

    return () => {
      try { selectionListenerRef.current?.remove(); } catch {}
      try { activateListenerRef.current?.remove(); } catch {}
    };
  }, [refreshSelection, handleActivate]);

  // ── Handlers ──────────────────────────────────────────────────────────────

  async function handleLink({ filePath, sheet, cell }) {
    const appObj = getApp();
    if (!appObj || !currentFrame) throw new Error("No frame selected");

    const { value } = await readCell(filePath, sheet, cell);
    createLink(currentFrame, filePath, sheet, cell, value, false);
    refreshSelection();
  }

  async function handleUpdateNow() {
    const appObj = getApp();
    if (!appObj || !currentLink) throw new Error("No link");
    const result = await updateLinkedItem(appObj, currentLink, true);
    refreshSelection();
    return result;
  }

  function handleRemoveLink() {
    if (!currentFrame) return;
    removeLink(currentFrame);
    refreshSelection();
  }

  function handleToggleAutoUpdate(enabled) {
    if (!currentFrame) return;
    setAutoUpdate(currentFrame, enabled);
    refreshSelection();
  }

  async function handleRepointFile() {
    const appObj = getApp();
    if (!appObj || !currentLink) return;
    if (!uxp) return;

    try {
      const result = await uxp.storage.localFileSystem.getFileForOpening({
        types: ["xlsx"],
      });
      if (!result) return;
      const newPath = result.nativePath || result.name;
      repointExcelFile(appObj, currentLink.excelFile, newPath);
      refreshSelection();
    } catch {}
  }

  function handleToggleHighlight() {
    const appObj = getApp();
    const next = !highlightActive;
    setHighlightActiveState(next);
    if (appObj) setHighlightActive(appObj, next);
  }

  // ── Render ─────────────────────────────────────────────────────────────────

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
    paddingBottom: 44, // leave room for the fixed bottom bar
  },
};
