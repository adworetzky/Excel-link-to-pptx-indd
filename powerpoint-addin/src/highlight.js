/**
 * powerpoint-addin/src/highlight.js
 *
 * Applies and removes a temporary highlight on all linked shapes.
 * Uses shape fill color as a visual indicator.
 * The original fill is saved and restored when highlight is toggled off.
 * This does NOT persist to the document — the highlight state is purely
 * in-memory and will be cleared when the add-in is reloaded.
 *
 * Highlight color: rgba(0, 100, 255, 0.15) approximated as a light blue
 * solid fill at low opacity. We can't use true alpha in Office JS shape
 * fill, so we use a light blue (#CCE0FF) as the highlight fill.
 */

import { getAllLinkedShapes } from "./storage.js";

const HIGHLIGHT_FILL = "#CCE0FF"; // light blue ~rgba(0,100,255,0.15)

// Map<linkId, { shapeName, originalFillType, originalFillColor }>
// Used to restore fill when highlight is cleared.
const _savedFills = new Map();

/**
 * Apply highlight fill to all linked shapes.
 *
 * @param {Office.RequestContext} context
 */
export async function applyHighlights(context) {
  await clearHighlights(context);

  const linked = await getAllLinkedShapes(context);
  for (const { shape, link } of linked) {
    try {
      // Save original fill
      shape.fill.load("type,foregroundColor");
      await context.sync();

      const originalFillType = shape.fill.type;
      let originalFillColor = null;
      try {
        originalFillColor = shape.fill.foregroundColor;
      } catch {
        // fill type may not have foregroundColor
      }

      _savedFills.set(link.linkId, {
        shape,
        originalFillType,
        originalFillColor,
      });

      // Apply highlight
      shape.fill.setSolidColor(HIGHLIGHT_FILL);
    } catch {
      // Skip shapes that can't be highlighted
    }
  }
  await context.sync();
}

/**
 * Remove all highlights and restore original fills.
 *
 * @param {Office.RequestContext} context
 */
export async function clearHighlights(context) {
  for (const [, saved] of _savedFills) {
    try {
      const { shape, originalFillType, originalFillColor } = saved;
      if (originalFillType === "NoFill" || originalFillType === 0) {
        shape.fill.clear();
      } else if (originalFillColor) {
        shape.fill.setSolidColor(originalFillColor);
      } else {
        shape.fill.clear();
      }
    } catch {
      // Shape may have been deleted
    }
  }
  if (_savedFills.size > 0) {
    try {
      await context.sync();
    } catch {}
  }
  _savedFills.clear();
}

/**
 * Toggle highlight state.
 *
 * @param {Office.RequestContext} context
 * @param {boolean} active
 */
export async function setHighlightActive(context, active) {
  if (active) {
    await applyHighlights(context);
  } else {
    await clearHighlights(context);
  }
}
