/**
 * indesign-plugin/src/storage.js
 *
 * Encodes and decodes DataLink metadata stored in InDesign frame labels.
 * The label is set to a JSON string of the DataLink object.
 * Non-DataLink frames have labels that are either empty or not valid JSON
 * containing a linkId field.
 */

/**
 * Decode a DataLink from a frame's label string.
 * Returns null if the label doesn't contain a valid DataLink.
 *
 * @param {string} label
 * @returns {DataLink|null}
 */
export function decodeLabel(label) {
  if (!label || label.trim() === "") return null;
  try {
    const obj = JSON.parse(label);
    if (obj && typeof obj.linkId === "string") {
      return obj;
    }
  } catch {
    // not JSON or not a DataLink
  }
  return null;
}

/**
 * Encode a DataLink into the JSON string that goes into frame.label.
 *
 * @param {DataLink} dataLink
 * @returns {string}
 */
export function encodeLabel(dataLink) {
  return JSON.stringify(dataLink);
}

/**
 * Read the DataLink from a frame (or any InDesign page item).
 *
 * @param {object} frame - InDesign page item with a .label property
 * @returns {DataLink|null}
 */
export function readLinkFromFrame(frame) {
  try {
    return decodeLabel(frame.label);
  } catch {
    return null;
  }
}

/**
 * Write a DataLink to a frame's label.
 *
 * @param {object} frame
 * @param {DataLink} dataLink
 */
export function writeLinkToFrame(frame, dataLink) {
  frame.label = encodeLabel(dataLink);
}

/**
 * Remove the DataLink from a frame by clearing its label.
 *
 * @param {object} frame
 */
export function removeLinkFromFrame(frame) {
  frame.label = "";
}

/**
 * Scan all page items in the active document and return those that
 * have a DataLink stored in their label.
 *
 * @param {object} app - InDesign app global
 * @returns {Array<{ frame: object, link: DataLink }>}
 */
export function getAllLinkedFrames(app) {
  const results = [];
  if (!app.activeDocument) return results;

  const pages = app.activeDocument.pages.everyItem().getElements();
  for (const page of pages) {
    for (const item of page.allPageItems) {
      const link = readLinkFromFrame(item);
      if (link) {
        results.push({ frame: item, link });
      }
    }
  }
  return results;
}

/**
 * Find the single frame that contains a DataLink with the given linkId.
 *
 * @param {object} app
 * @param {string} linkId
 * @returns {{ frame: object, link: DataLink }|null}
 */
export function findFrameByLinkId(app, linkId) {
  const all = getAllLinkedFrames(app);
  return all.find((entry) => entry.link.linkId === linkId) || null;
}
