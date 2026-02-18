/**
 * indesign-plugin/src/linker.js
 *
 * Creates, reads, and removes DataLink associations on InDesign frames.
 */

import { v4 as uuidv4 } from "uuid";
import {
  readLinkFromFrame,
  writeLinkToFrame,
  removeLinkFromFrame,
} from "./storage.js";

/**
 * Create a new DataLink on a frame and write the initial Excel value.
 *
 * @param {object} frame       - InDesign text frame
 * @param {string} excelFile   - Absolute path to the .xlsx file
 * @param {string} sheet       - Sheet name
 * @param {string} cell        - Cell address (A1 notation)
 * @param {string} initialValue - Formatted string value to write immediately
 * @param {boolean} autoUpdate
 * @returns {DataLink}
 */
export function createLink(frame, excelFile, sheet, cell, initialValue, autoUpdate = false) {
  const dataLink = {
    linkId: uuidv4(),
    excelFile,
    sheet,
    cell,
    lastValue: initialValue,
    lastUpdated: new Date().toISOString(),
    autoUpdate,
  };

  // Write formatted value to frame text — content only, never formatting
  frame.parentStory.contents = initialValue;

  writeLinkToFrame(frame, dataLink);
  return dataLink;
}

/**
 * Read the DataLink from a frame.
 *
 * @param {object} frame
 * @returns {DataLink|null}
 */
export function readLink(frame) {
  return readLinkFromFrame(frame);
}

/**
 * Remove the DataLink from a frame. Does not change the text content.
 *
 * @param {object} frame
 */
export function removeLink(frame) {
  removeLinkFromFrame(frame);
}

/**
 * Toggle the autoUpdate flag on a linked frame.
 *
 * @param {object} frame
 * @param {boolean} enabled
 * @returns {DataLink|null} updated link or null if not linked
 */
export function setAutoUpdate(frame, enabled) {
  const link = readLinkFromFrame(frame);
  if (!link) return null;
  link.autoUpdate = enabled;
  writeLinkToFrame(frame, link);
  return link;
}
