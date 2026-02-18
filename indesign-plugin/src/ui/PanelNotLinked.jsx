/**
 * PanelNotLinked — shown when a text frame is selected but has no DataLink.
 * Provides file browse, sheet picker, cell input, and Link button.
 */
import React, { useState, useEffect } from "react";

const uxp = typeof require !== "undefined" ? require("uxp") : null;

export default function PanelNotLinked({ onLink }) {
  const [filePath, setFilePath] = useState("");
  const [sheets, setSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [cell, setCell] = useState("");
  const [loading, setLoading] = useState(false);
  const [sheetLoading, setSheetLoading] = useState(false);
  const [errors, setErrors] = useState({});

  // When filePath changes, fetch sheet names
  useEffect(() => {
    if (!filePath) {
      setSheets([]);
      setSelectedSheet("");
      return;
    }
    (async () => {
      setSheetLoading(true);
      setErrors((e) => ({ ...e, file: undefined }));
      try {
        // Import getSheetNames dynamically to avoid bundling issues
        const { getSheetNames } = await import("@datalink/shared");
        const names = await getSheetNames(filePath);
        setSheets(names);
        setSelectedSheet(names[0] || "");
      } catch (err) {
        setErrors((e) => ({ ...e, file: `Cannot read file: ${err.message}` }));
        setSheets([]);
        setSelectedSheet("");
      } finally {
        setSheetLoading(false);
      }
    })();
  }, [filePath]);

  async function handleBrowse() {
    if (!uxp) return;
    try {
      const result = await uxp.storage.localFileSystem.getFileForOpening({
        types: ["xlsx"],
      });
      if (result) {
        // UXP returns a File object; get the native path
        const nativePath = result.nativePath || result.name;
        setFilePath(nativePath);
      }
    } catch {
      // user cancelled
    }
  }

  function validate() {
    const errs = {};
    if (!filePath) errs.file = "Choose an Excel file.";
    if (!selectedSheet) errs.sheet = "Choose a sheet.";
    if (!cell.trim()) errs.cell = "Enter a cell address (e.g. B4).";
    else if (!/^[A-Za-z]+\d+$/.test(cell.trim()))
      errs.cell = "Invalid cell address. Use A1 notation (e.g. B4).";
    return errs;
  }

  async function handleLink() {
    const errs = validate();
    if (Object.keys(errs).length > 0) {
      setErrors(errs);
      return;
    }
    setErrors({});
    setLoading(true);
    try {
      await onLink({ filePath, sheet: selectedSheet, cell: cell.trim().toUpperCase() });
    } catch (err) {
      setErrors({ general: err.message });
    } finally {
      setLoading(false);
    }
  }

  return (
    <div style={styles.container}>
      <p style={styles.sectionLabel}>Excel File</p>
      <div style={styles.fileRow}>
        <input
          style={{ ...styles.input, flex: 1, cursor: "default" }}
          value={filePath}
          readOnly
          placeholder="No file selected"
          title={filePath}
        />
        <button style={styles.browseBtn} onClick={handleBrowse}>
          Browse…
        </button>
      </div>
      {errors.file && <p style={styles.error}>{errors.file}</p>}

      <p style={styles.sectionLabel}>Sheet</p>
      {sheetLoading ? (
        <p style={styles.hint}>Loading sheets…</p>
      ) : (
        <select
          style={styles.select}
          value={selectedSheet}
          onChange={(e) => setSelectedSheet(e.target.value)}
          disabled={sheets.length === 0}
        >
          {sheets.length === 0 && (
            <option value="">— select a file first —</option>
          )}
          {sheets.map((s) => (
            <option key={s} value={s}>
              {s}
            </option>
          ))}
        </select>
      )}
      {errors.sheet && <p style={styles.error}>{errors.sheet}</p>}

      <p style={styles.sectionLabel}>Cell</p>
      <input
        style={styles.input}
        value={cell}
        onChange={(e) => setCell(e.target.value)}
        placeholder="e.g. B4"
        maxLength={10}
      />
      {errors.cell && <p style={styles.error}>{errors.cell}</p>}

      {errors.general && <p style={styles.error}>{errors.general}</p>}

      <button
        style={{ ...styles.primaryBtn, marginTop: 8 }}
        onClick={handleLink}
        disabled={loading}
      >
        {loading ? "Linking…" : "Link"}
      </button>
    </div>
  );
}

const styles = {
  container: {
    padding: 16,
    display: "flex",
    flexDirection: "column",
    gap: 6,
  },
  sectionLabel: {
    fontSize: 11,
    fontWeight: "600",
    color: "#3c3c3c",
    marginTop: 8,
    marginBottom: 2,
  },
  fileRow: {
    display: "flex",
    gap: 6,
    alignItems: "center",
  },
  input: {
    width: "100%",
    padding: "4px 6px",
    fontSize: 12,
    border: "1px solid #b0b0b0",
    background: "#fff",
    color: "#1e1e1e",
    outline: "none",
  },
  select: {
    width: "100%",
    padding: "4px 6px",
    fontSize: 12,
    border: "1px solid #b0b0b0",
    background: "#fff",
    color: "#1e1e1e",
  },
  browseBtn: {
    padding: "4px 10px",
    fontSize: 12,
    border: "1px solid #b0b0b0",
    background: "#f4f4f4",
    color: "#1e1e1e",
    cursor: "pointer",
    whiteSpace: "nowrap",
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
  error: {
    fontSize: 11,
    color: "#c00",
    marginTop: 2,
  },
  hint: {
    fontSize: 11,
    color: "#767676",
  },
};
