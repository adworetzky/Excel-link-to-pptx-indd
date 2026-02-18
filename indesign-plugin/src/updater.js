/**
 * indesign-plugin/src/updater.js
 *
 * Core update logic: reads a value from Excel and writes it to an InDesign
 * text frame — without touching any formatting.
 */

import { readCell } from "@datalink/shared";
import { writeLinkToFrame, findFrameByLinkId, getAllLinkedFrames } from "./storage.js";

/**
 * Update a single linked frame from its Excel source.
 *
 * @param {object}  app          - InDesign app global
 * @param {DataLink} linkMeta
 * @param {boolean}  forceUpdate - if true, write even when value hasn't changed
 * @returns {Promise<{ updated: boolean, newValue: string, error?: string }>}
 */
export async function updateLinkedItem(app, linkMeta, forceUpdate = false) {
  let newValue;
  try {
    const result = await readCell(linkMeta.excelFile, linkMeta.sheet, linkMeta.cell);
    newValue = result.value;
  } catch (err) {
    const isNotFound =
      err.message.includes("File not found") ||
      err.message.includes("ENOENT");
    return {
      updated: false,
      newValue: linkMeta.lastValue,
      error: isNotFound
        ? `FILE_NOT_FOUND:${linkMeta.excelFile}`
        : err.message,
    };
  }

  if (newValue === linkMeta.lastValue && !forceUpdate) {
    return { updated: false, newValue };
  }

  // Find the frame in the document
  const entry = findFrameByLinkId(app, linkMeta.linkId);
  if (!entry) {
    return {
      updated: false,
      newValue,
      error: "Frame not found in document",
    };
  }

  // Write text content only — never touch formatting
  try {
    entry.frame.parentStory.contents = newValue;
  } catch (err) {
    return { updated: false, newValue, error: `Write error: ${err.message}` };
  }

  // Update metadata
  const updated = {
    ...linkMeta,
    lastValue: newValue,
    lastUpdated: new Date().toISOString(),
  };
  writeLinkToFrame(entry.frame, updated);

  return { updated: true, newValue };
}

/**
 * Update all linked frames that have autoUpdate === true.
 *
 * @param {object} app
 * @returns {Promise<Array<{ linkId: string, result: object }>>}
 */
export async function updateAllAutoLinks(app) {
  const linked = getAllLinkedFrames(app);
  const results = [];
  for (const { link } of linked) {
    if (link.autoUpdate) {
      const result = await updateLinkedItem(app, link);
      results.push({ linkId: link.linkId, result });
    }
  }
  return results;
}

/**
 * Update every linked frame in the document, regardless of autoUpdate flag.
 *
 * @param {object} app
 * @returns {Promise<Array<{ linkId: string, result: object }>>}
 */
export async function updateAllLinks(app) {
  const linked = getAllLinkedFrames(app);
  const results = [];
  for (const { link } of linked) {
    const result = await updateLinkedItem(app, link, true);
    results.push({ linkId: link.linkId, result });
  }
  return results;
}

/**
 * Repoint all links from one Excel file path to a new path.
 * Called when the user selects a replacement file for a broken link.
 *
 * @param {object} app
 * @param {string} oldPath
 * @param {string} newPath
 */
export function repointExcelFile(app, oldPath, newPath) {
  const linked = getAllLinkedFrames(app);
  for (const { frame, link } of linked) {
    if (link.excelFile === oldPath) {
      const updated = { ...link, excelFile: newPath };
      writeLinkToFrame(frame, updated);
    }
  }
}
