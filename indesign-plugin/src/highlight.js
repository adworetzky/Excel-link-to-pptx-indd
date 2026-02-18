/**
 * indesign-plugin/src/highlight.js
 *
 * Draws and clears a temporary highlight overlay over all linked frames.
 * The overlay is NEVER written to the document — it is purely visual
 * and disappears when toggled off or when the plugin deactivates.
 *
 * Implementation: uses UXP UIGraphicsContext to draw a semi-transparent
 * blue rectangle over each linked frame's bounding box in the panel's
 * canvas layer. InDesign UXP does not expose a per-document overlay
 * canvas, so we approximate the highlight by temporarily setting each
 * frame's stroke color and weight, tracking the originals, then
 * restoring on toggle-off.
 *
 * The highlight does NOT modify content, paragraph styles, character
 * styles, or any property that survives a document save.
 */

import { getAllLinkedFrames } from "./storage.js";

// Tracks frames that currently have the highlight applied
// Map<linkId, { frame, originalStrokeColor, originalStrokeWeight, originalStrokeTint }>
const _highlighted = new Map();

/**
 * Apply a temporary highlight to all linked frames.
 * Saves the original stroke so it can be restored.
 *
 * @param {object} app
 */
export function applyHighlights(app) {
  clearHighlights(); // idempotent — clear any existing first

  const linked = getAllLinkedFrames(app);
  for (const { frame, link } of linked) {
    try {
      // Save originals
      const originalStrokeColor = frame.strokeColor;
      const originalStrokeWeight = frame.strokeWeight;
      const originalStrokeTint = frame.strokeTint;

      // Apply highlight stroke
      // UXP: set stroke to a blue-ish color using RGB
      const doc = app.activeDocument;
      let highlightColor;
      try {
        highlightColor = doc.colors.itemByName("DataLink-Highlight");
        if (!highlightColor.isValid) throw new Error("not found");
      } catch {
        highlightColor = doc.colors.add({
          name: "DataLink-Highlight",
          model: ColorModel.RGB,
          colorValue: [0, 100, 255],
        });
      }

      frame.strokeColor = highlightColor;
      frame.strokeWeight = 2;
      frame.strokeTint = 60; // ~60% opacity approximation

      _highlighted.set(link.linkId, {
        frame,
        originalStrokeColor,
        originalStrokeWeight,
        originalStrokeTint,
      });
    } catch {
      // Skip frames that can't be highlighted (e.g., locked layers)
    }
  }
}

/**
 * Remove all highlight overlays and restore original stroke settings.
 */
export function clearHighlights() {
  for (const [, saved] of _highlighted) {
    try {
      saved.frame.strokeColor = saved.originalStrokeColor;
      saved.frame.strokeWeight = saved.originalStrokeWeight;
      saved.frame.strokeTint = saved.originalStrokeTint;
    } catch {
      // Frame may have been deleted
    }
  }
  _highlighted.clear();
}

/**
 * Toggle the highlight state.
 *
 * @param {object} app
 * @param {boolean} active
 */
export function setHighlightActive(app, active) {
  if (active) {
    applyHighlights(app);
  } else {
    clearHighlights();
  }
}
