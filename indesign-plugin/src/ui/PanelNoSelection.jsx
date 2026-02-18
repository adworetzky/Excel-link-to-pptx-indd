/**
 * PanelNoSelection — shown when nothing is selected in InDesign.
 */
import React from "react";

export default function PanelNoSelection() {
  return (
    <div style={styles.container}>
      <p style={styles.heading}>DataLink</p>
      <p style={styles.message}>
        Select a text frame in InDesign to link it to an Excel cell.
      </p>
      <p style={styles.hint}>
        After selecting a frame, you can browse for an Excel file, choose a
        sheet and cell, then click <strong>Link</strong>.
      </p>
    </div>
  );
}

const styles = {
  container: {
    padding: 16,
    display: "flex",
    flexDirection: "column",
    gap: 12,
  },
  heading: {
    fontSize: 13,
    fontWeight: "600",
    color: "#1e1e1e",
  },
  message: {
    fontSize: 12,
    color: "#3c3c3c",
    lineHeight: 1.5,
  },
  hint: {
    fontSize: 11,
    color: "#767676",
    lineHeight: 1.5,
  },
};
