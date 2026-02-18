/**
 * powerpoint-addin/src/updater.js
 *
 * Core update logic for the PowerPoint add-in.
 * Reads a formatted value from Excel and writes it to the shape's text range,
 * without touching any text formatting properties.
 */

import { readCell } from "@datalink/shared";
import { findShapeByLinkId, getAllLinkedShapes, writeLinkToShape } from "./storage.js";

/**
 * Update a single linked shape from its Excel source.
 * Must be called inside PowerPoint.run() with a valid context.
 *
 * @param {Office.RequestContext} context
 * @param {DataLink} linkMeta
 * @param {boolean} forceUpdate
 * @returns {Promise<{ updated: boolean, newValue: string, error?: string }>}
 */
export async function updateLinkedItem(context, linkMeta, forceUpdate = false) {
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

  const entry = await findShapeByLinkId(context, linkMeta.linkId);
  if (!entry) {
    return { updated: false, newValue, error: "Shape not found in presentation" };
  }

  // Write text content only — never touch formatting
  entry.shape.textFrame.textRange.text = newValue;

  const updated = {
    ...linkMeta,
    lastValue: newValue,
    lastUpdated: new Date().toISOString(),
  };
  await writeLinkToShape(entry.shape, updated, context);
  await context.sync();

  return { updated: true, newValue };
}

/**
 * Update all auto-linked shapes.
 *
 * @param {Office.RequestContext} context
 * @returns {Promise<Array<{ linkId: string, result: object }>>}
 */
export async function updateAllAutoLinks(context) {
  const linked = await getAllLinkedShapes(context);
  const results = [];
  for (const { link } of linked) {
    if (link.autoUpdate) {
      const result = await updateLinkedItem(context, link);
      results.push({ linkId: link.linkId, result });
    }
  }
  return results;
}

/**
 * Update every linked shape in the presentation.
 *
 * @param {Office.RequestContext} context
 * @returns {Promise<Array<{ linkId: string, result: object }>>}
 */
export async function updateAllLinks(context) {
  const linked = await getAllLinkedShapes(context);
  const results = [];
  for (const { link } of linked) {
    const result = await updateLinkedItem(context, link, true);
    results.push({ linkId: link.linkId, result });
  }
  return results;
}

/**
 * Repoint all links from an old Excel file path to a new one.
 *
 * @param {Office.RequestContext} context
 * @param {string} oldPath
 * @param {string} newPath
 */
export async function repointExcelFile(context, oldPath, newPath) {
  const linked = await getAllLinkedShapes(context);
  for (const { shape, link } of linked) {
    if (link.excelFile === oldPath) {
      const updated = { ...link, excelFile: newPath };
      await writeLinkToShape(shape, updated, context);
    }
  }
  await context.sync();
}
