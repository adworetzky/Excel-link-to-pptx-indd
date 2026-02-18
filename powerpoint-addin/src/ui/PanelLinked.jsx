/**
 * PanelLinked — shown when a linked shape is selected in PowerPoint.
 */
import React, { useState } from "react";

export default function PanelLinked({ link, onUpdateNow, onRemoveLink, onToggleAutoUpdate, onRepointFile }) {
  const [updating, setUpdating] = useState(false);
  const [updateResult, setUpdateResult] = useState(null);
  const [removing, setRemoving] = useState(false);

  const isFileBroken = link._fileBroken;

  async function handleUpdateNow() {
    setUpdating(true);
    setUpdateResult(null);
    try {
      const result = await onUpdateNow();
      setUpdateResult(result.updated ? "Updated." : "Already up to date.");
    } catch (err) {
      setUpdateResult(`Error: ${err.message}`);
    } finally {
      setUpdating(false);
    }
  }

  async function handleRemove() {
    if (!window.confirm("Remove the DataLink from this shape? The text content will not change.")) return;
    setRemoving(true);
    try {
      await onRemoveLink();
    } finally {
      setRemoving(false);
    }
  }

  function formatTimestamp(iso) {
    if (!iso) return "—";
    try {
      return new Date(iso).toLocaleString();
    } catch {
      return iso;
    }
  }

  return (
    <div style={styles.container}>
      {/* Status indicator */}
      <div style={styles.statusRow}>
        <span style={{ ...styles.dot, background: isFileBroken ? "#c00" : "#2e7d32" }} />
        <span style={styles.statusText}>
          {isFileBroken ? "Broken — file not found" : "Linked"}
        </span>
      </div>

      {isFileBroken && (
        <div style={styles.warningBox}>
          <p style={styles.warningText}>
            The Excel file could not be found. The link is broken.
          </p>
          <button style={styles.secondaryBtn} onClick={onRepointFile}>
            Repoint Excel File…
          </button>
        </div>
      )}

      {/* Source info */}
      <div style={styles.infoBlock}>
        <p style={styles.label}>File</p>
        <p style={styles.value} title={link.excelFile}>
          {shortPath(link.excelFile)}
        </p>

        <p style={styles.label}>Sheet</p>
        <p style={styles.value}>{link.sheet}</p>

        <p style={styles.label}>Cell</p>
        <p style={styles.value}>{link.cell}</p>

        <p style={styles.label}>Current Value</p>
        <p style={styles.valueEm}>{link.lastValue || "—"}</p>

        <p style={styles.label}>Last Updated</p>
        <p style={styles.value}>{formatTimestamp(link.lastUpdated)}</p>
      </div>

      {/* Auto-update toggle */}
      <div style={styles.toggleRow}>
        <label style={styles.toggleLabel}>
          <input
            type="checkbox"
            checked={!!link.autoUpdate}
            onChange={(e) => onToggleAutoUpdate(e.target.checked)}
            style={{ marginRight: 6 }}
          />
          Auto-update on document activity
        </label>
      </div>

      {/* Update Now */}
      <button
        style={styles.primaryBtn}
        onClick={handleUpdateNow}
        disabled={updating || isFileBroken}
      >
        {updating ? "Updating…" : "Update Now"}
      </button>
      {updateResult && <p style={styles.resultText}>{updateResult}</p>}

      {/* Remove Link */}
      <button
        style={styles.destructiveBtn}
        onClick={handleRemove}
        disabled={removing}
      >
        {removing ? "Removing…" : "Remove Link"}
      </button>
    </div>
  );
}

function shortPath(p) {
  if (!p) return "—";
  const parts = p.replace(/\\/g, "/").split("/");
  return parts.length > 3 ? "…/" + parts.slice(-2).join("/") : p;
}

const styles = {
  container: {
    padding: 16,
    display: "flex",
    flexDirection: "column",
    gap: 10,
  },
  statusRow: {
    display: "flex",
    alignItems: "center",
    gap: 6,
  },
  dot: {
    width: 8,
    height: 8,
    borderRadius: "50%",
    flexShrink: 0,
  },
  statusText: {
    fontSize: 12,
    color: "#3c3c3c",
  },
  warningBox: {
    background: "#fff3f3",
    border: "1px solid #c00",
    padding: "10px 12px",
    display: "flex",
    flexDirection: "column",
    gap: 8,
  },
  warningText: {
    fontSize: 11,
    color: "#c00",
    lineHeight: 1.5,
  },
  infoBlock: {
    display: "grid",
    gridTemplateColumns: "72px 1fr",
    gap: "4px 8px",
    alignItems: "start",
  },
  label: {
    fontSize: 11,
    color: "#767676",
    paddingTop: 1,
  },
  value: {
    fontSize: 12,
    color: "#1e1e1e",
    wordBreak: "break-all",
  },
  valueEm: {
    fontSize: 13,
    fontWeight: "600",
    color: "#1e1e1e",
  },
  toggleRow: {
    marginTop: 4,
  },
  toggleLabel: {
    display: "flex",
    alignItems: "center",
    fontSize: 12,
    color: "#1e1e1e",
    cursor: "pointer",
  },
  primaryBtn: {
    width: "100%",
    padding: "7px 0",
    fontSize: 13,
    fontWeight: "600",
    background: "#0073CF",
    color: "#fff",
    border: "none",
    cursor: "pointer",
  },
  secondaryBtn: {
    width: "100%",
    padding: "6px 0",
    fontSize: 12,
    background: "transparent",
    color: "#0073CF",
    border: "1px solid #0073CF",
    cursor: "pointer",
  },
  destructiveBtn: {
    width: "100%",
    padding: "6px 0",
    fontSize: 12,
    background: "transparent",
    color: "#c00",
    border: "1px solid #c00",
    cursor: "pointer",
  },
  resultText: {
    fontSize: 11,
    color: "#3c3c3c",
    textAlign: "center",
  },
};
