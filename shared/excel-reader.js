/**
 * shared/excel-reader.js
 *
 * Reads cell values from a local .xlsx file using SheetJS.
 * Number formatting from the cell's format string is applied so
 * the returned value is the display string (e.g. "$1,234.56").
 */

"use strict";

const XLSX = require("xlsx");
const { readFile, utils } = XLSX;
const { existsSync } = require("fs");

/**
 * Read a single cell value from a local xlsx file, formatted as a string.
 *
 * @param {string} filePath    - Absolute path to the .xlsx file
 * @param {string} sheetName   - Sheet name
 * @param {string} cellAddress - A1 notation e.g. "B4"
 * @returns {Promise<{ value: string, rawValue: any, formatString: string }>}
 * @throws {Error} if file not found, sheet not found, or cell empty
 */
async function readCell(filePath, sheetName, cellAddress) {
  if (!existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  let workbook;
  try {
    workbook = readFile(filePath, { cellStyles: true, cellNF: true, cellDates: true });
  } catch (err) {
    throw new Error(`Failed to open workbook: ${err.message}`);
  }

  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    const available = workbook.SheetNames.join(", ");
    throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${available}`);
  }

  // Resolve merged cell: if the address is in a merge range,
  // redirect to the top-left cell of that range.
  const resolvedAddress = resolveMerge(sheet, cellAddress);
  const cell = sheet[resolvedAddress];

  if (cell === undefined || cell === null) {
    throw new Error(`Cell ${cellAddress} is empty`);
  }

  const formatString = cell.z || "General";
  const value = utils.format_cell(cell);
  const rawValue = cell.v;

  return { value, rawValue, formatString };
}

/**
 * Get all sheet names from a workbook.
 *
 * @param {string} filePath
 * @returns {Promise<string[]>}
 */
async function getSheetNames(filePath) {
  if (!existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  let workbook;
  try {
    workbook = readFile(filePath, { bookSheets: true });
  } catch (err) {
    throw new Error(`Failed to open workbook: ${err.message}`);
  }

  return workbook.SheetNames;
}

/**
 * Validate that a file path points to a readable xlsx file.
 *
 * @param {string} filePath
 * @returns {Promise<boolean>}
 */
async function validateFile(filePath) {
  if (!existsSync(filePath)) {
    return false;
  }
  try {
    readFile(filePath, { bookSheets: true });
    return true;
  } catch {
    return false;
  }
}

/**
 * If the given address falls within a merged cell range, return the
 * top-left address of that range. Otherwise return the original address.
 *
 * @param {object} sheet  - SheetJS worksheet object
 * @param {string} addr   - A1 address string
 * @returns {string}
 */
function resolveMerge(sheet, addr) {
  if (!sheet["!merges"]) return addr;

  const target = utils.decode_cell(addr);
  for (const merge of sheet["!merges"]) {
    const { s, e } = merge; // s = start, e = end (both {r, c})
    if (
      target.r >= s.r &&
      target.r <= e.r &&
      target.c >= s.c &&
      target.c <= e.c
    ) {
      return utils.encode_cell(s);
    }
  }
  return addr;
}

module.exports = { readCell, getSheetNames, validateFile };
