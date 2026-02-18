/**
 * powerpoint-addin/src/linker.js
 *
 * Creates, reads, and removes DataLink associations on PowerPoint shapes.
 * All operations are async and must be called from within Office JS context.
 */

import { v4 as uuidv4 } from "uuid";
import { writeLinkToShape, readLinkFromShape, removeLinkFromShape } from "./storage.js";

/**
 * Create a new DataLink on a shape and write the initial Excel value.
 * Must be called inside PowerPoint.run().
 *
 * @param {Office.Shape} shape
 * @param {string} excelFile
 * @param {string} sheet
 * @param {string} cell
 * @param {string} initialValue  - formatted display string
 * @param {boolean} autoUpdate
 * @param {Office.RequestContext} context
 * @returns {Promise<DataLink>}
 */
export async function createLink(shape, excelFile, sheet, cell, initialValue, autoUpdate, context) {
  // Write text content only — no formatting change
  shape.textFrame.textRange.text = initialValue;

  const dataLink = {
    linkId: uuidv4(),
    excelFile,
    sheet,
    cell,
    lastValue: initialValue,
    lastUpdated: new Date().toISOString(),
    autoUpdate: !!autoUpdate,
  };

  await writeLinkToShape(shape, dataLink, context);
  return dataLink;
}

/**
 * Read the DataLink from a shape.
 *
 * @param {Office.Shape} shape
 * @param {Office.RequestContext} context
 * @returns {Promise<DataLink|null>}
 */
export async function readLink(shape, context) {
  return readLinkFromShape(shape, context);
}

/**
 * Remove the DataLink from a shape. Does not change text content.
 *
 * @param {Office.Shape} shape
 * @param {Office.RequestContext} context
 */
export async function removeLink(shape, context) {
  return removeLinkFromShape(shape, context);
}

/**
 * Update the autoUpdate flag on a linked shape.
 *
 * @param {Office.Shape} shape
 * @param {boolean} enabled
 * @param {Office.RequestContext} context
 * @returns {Promise<DataLink|null>}
 */
export async function setAutoUpdate(shape, enabled, context) {
  const link = await readLinkFromShape(shape, context);
  if (!link) return null;
  link.autoUpdate = enabled;
  await writeLinkToShape(shape, link, context);
  return link;
}
