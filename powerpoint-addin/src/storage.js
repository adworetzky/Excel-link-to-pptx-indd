/**
 * powerpoint-addin/src/storage.js
 *
 * Encodes and decodes DataLink metadata stored in PowerPoint shape tags.
 * Uses shape.tags.add("DATALINK", jsonString) to store invisibly.
 *
 * All functions that touch Office JS APIs return Promises and must be
 * called inside a PowerPoint.run() context.
 */

const TAG_KEY = "DATALINK";

/**
 * Write a DataLink to a shape's tags.
 * Must be called inside PowerPoint.run() with a loaded shape.
 *
 * @param {Office.Shape} shape
 * @param {DataLink} dataLink
 * @param {Office.RequestContext} context
 */
export async function writeLinkToShape(shape, dataLink, context) {
  shape.tags.add(TAG_KEY, JSON.stringify(dataLink));
  await context.sync();
}

/**
 * Read the DataLink from a shape's tags.
 * Returns null if the shape has no DataLink tag.
 *
 * @param {Office.Shape} shape - shape with tags already loaded
 * @param {Office.RequestContext} context
 * @returns {Promise<DataLink|null>}
 */
export async function readLinkFromShape(shape, context) {
  const tag = shape.tags.getItemOrNullObject(TAG_KEY);
  tag.load("key,value,isNullObject");
  await context.sync();

  if (tag.isNullObject) return null;
  try {
    const obj = JSON.parse(tag.value);
    if (obj && typeof obj.linkId === "string") return obj;
  } catch {
    // corrupted tag
  }
  return null;
}

/**
 * Remove the DataLink tag from a shape.
 *
 * @param {Office.Shape} shape
 * @param {Office.RequestContext} context
 */
export async function removeLinkFromShape(shape, context) {
  const tag = shape.tags.getItemOrNullObject(TAG_KEY);
  await context.sync();
  if (!tag.isNullObject) {
    tag.delete();
    await context.sync();
  }
}

/**
 * Scan all shapes on all slides and return those with a DataLink tag.
 *
 * @param {Office.RequestContext} context
 * @returns {Promise<Array<{ shape: Office.Shape, slideIndex: number, link: DataLink }>>}
 */
export async function getAllLinkedShapes(context) {
  const results = [];
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();

  for (let si = 0; si < slides.items.length; si++) {
    const slide = slides.items[si];
    slide.shapes.load("items");
    await context.sync();

    for (const shape of slide.shapes.items) {
      shape.tags.load("items");
      await context.sync();

      const tag = shape.tags.getItemOrNullObject(TAG_KEY);
      tag.load("key,value,isNullObject");
      await context.sync();

      if (!tag.isNullObject) {
        try {
          const obj = JSON.parse(tag.value);
          if (obj && typeof obj.linkId === "string") {
            results.push({ shape, slideIndex: si, link: obj });
          }
        } catch {}
      }
    }
  }
  return results;
}

/**
 * Find the shape with the given linkId.
 *
 * @param {Office.RequestContext} context
 * @param {string} linkId
 * @returns {Promise<{ shape: Office.Shape, slideIndex: number, link: DataLink }|null>}
 */
export async function findShapeByLinkId(context, linkId) {
  const all = await getAllLinkedShapes(context);
  return all.find((e) => e.link.linkId === linkId) || null;
}
