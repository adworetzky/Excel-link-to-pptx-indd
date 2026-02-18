/**
 * shared/index.js
 *
 * Main export for @datalink/shared.
 * Re-exports all public API from the excel-reader module.
 */

export { readCell, getSheetNames, validateFile } from "./excel-reader.js";
