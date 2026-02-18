/**
 * LinkedItemsBar — always-visible bottom bar.
 * Shows total linked item count and the Highlight Links toggle.
 */
import React from "react";

export default function LinkedItemsBar({ count, highlightActive, onToggleHighlight }) {
  return (
    <div style={styles.bar}>
      <span style={styles.count}>
        {count} linked {count === 1 ? "item" : "items"}
      </span>
      <button
        style={{
          ...styles.highlightBtn,
          background: highlightActive ? "#0073CF" : "transparent",
          color: highlightActive ? "#fff" : "#0073CF",
        }}
        onClick={onToggleHighlight}
        aria-pressed={highlightActive}
      >
        {highlightActive ? "Hide Highlights" : "Highlight Links"}
      </button>
    </div>
  );
}

const styles = {
  bar: {
    position: "fixed",
    bottom: 0,
    left: 0,
    right: 0,
    height: 36,
    background: "#e8e8e8",
    borderTop: "1px solid #c0c0c0",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "0 12px",
    zIndex: 100,
  },
  count: {
    fontSize: 11,
    color: "#3c3c3c",
  },
  highlightBtn: {
    padding: "4px 10px",
    fontSize: 11,
    border: "1px solid #0073CF",
    cursor: "pointer",
    fontWeight: "500",
  },
};
