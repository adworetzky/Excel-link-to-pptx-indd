/**
 * shared/test/excel-reader.test.js
 *
 * CLI test harness for the Excel reader module.
 * Creates a temporary .xlsx file, then validates:
 *   - currency formatting
 *   - percentage formatting
 *   - date formatting
 *   - plain number formatting
 *   - merged cell resolution
 *   - error cases (file not found, bad sheet, empty cell)
 *
 * Usage: node test/excel-reader.test.js
 */

import assert from "assert";
import { writeFileSync, unlinkSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { readCell, getSheetNames, validateFile } from "../index.js";

// We use SheetJS to create the test workbook inline.
import { utils, write } from "xlsx";

// ─── Helpers ────────────────────────────────────────────────────────────────

let passed = 0;
let failed = 0;

async function test(name, fn) {
  try {
    await fn();
    console.log(`  ✓ ${name}`);
    passed++;
  } catch (err) {
    console.error(`  ✗ ${name}`);
    console.error(`    ${err.message}`);
    failed++;
  }
}

function assertIncludes(str, substr, msg) {
  assert(
    String(str).includes(substr),
    msg || `Expected "${str}" to include "${substr}"`
  );
}

// ─── Build test workbook ─────────────────────────────────────────────────────

function buildTestWorkbook() {
  const wb = utils.book_new();

  // --- Sheet1: various number formats ---
  const ws1 = utils.aoa_to_sheet([
    ["Label", "Value"],
    ["Currency", 1234567.89],
    ["Percentage", 0.4256],
    ["Date", new Date(2024, 0, 15)], // Jan 15 2024
    ["Plain", 42],
    ["String", "hello world"],
  ]);

  // Apply number formats
  ws1["B2"].z = '$#,##0.00';
  ws1["B3"].z = '0.00%';
  ws1["B4"].z = 'mm/dd/yyyy';
  ws1["B5"].z = '0';

  // Merged cell: A7:B8 → value in A7, rest are empty
  ws1["A7"] = { v: "MergedCell", t: "s" };
  ws1["!merges"] = [{ s: { r: 6, c: 0 }, e: { r: 7, c: 1 } }];
  ws1["!ref"] = "A1:B8";

  utils.book_append_sheet(wb, ws1, "Sheet1");

  // --- Sheet2: simple values ---
  const ws2 = utils.aoa_to_sheet([["only one cell"]]);
  utils.book_append_sheet(wb, ws2, "Data");

  return wb;
}

// Write workbook to temp file
const tmpPath = join(tmpdir(), `datalink-test-${Date.now()}.xlsx`);
const wb = buildTestWorkbook();
writeFileSync(tmpPath, write(wb, { type: "buffer", bookType: "xlsx" }));

// ─── Run tests ───────────────────────────────────────────────────────────────

console.log("\nDataLink shared/excel-reader tests\n");

// --- validateFile ---
console.log("validateFile:");
await test("returns true for valid xlsx", async () => {
  assert.strictEqual(await validateFile(tmpPath), true);
});
await test("returns false for missing file", async () => {
  assert.strictEqual(await validateFile("/nonexistent/path/file.xlsx"), false);
});

// --- getSheetNames ---
console.log("\ngetSheetNames:");
await test("returns all sheet names", async () => {
  const names = await getSheetNames(tmpPath);
  assert.deepStrictEqual(names, ["Sheet1", "Data"]);
});
await test("throws for missing file", async () => {
  await assert.rejects(
    () => getSheetNames("/nonexistent/file.xlsx"),
    /File not found/
  );
});

// --- readCell: number formatting ---
console.log("\nreadCell — number formatting:");
await test("currency: $#,##0.00", async () => {
  const { value, formatString } = await readCell(tmpPath, "Sheet1", "B2");
  assert.strictEqual(formatString, "$#,##0.00");
  assertIncludes(value, "$", "should contain dollar sign");
  assertIncludes(value, "1,234,567", "should contain formatted number");
});
await test("percentage: 0.00%", async () => {
  const { value } = await readCell(tmpPath, "Sheet1", "B3");
  assertIncludes(value, "%", "should contain percent sign");
});
await test("date: mm/dd/yyyy", async () => {
  const { value } = await readCell(tmpPath, "Sheet1", "B4");
  // SheetJS formats date cells; just verify it has slashes
  assertIncludes(value, "/", "should contain date separator");
});
await test("plain number: 0", async () => {
  const { value } = await readCell(tmpPath, "Sheet1", "B5");
  assert.strictEqual(value, "42");
});
await test("string cell", async () => {
  const { value } = await readCell(tmpPath, "Sheet1", "B6");
  assert.strictEqual(value, "hello world");
});

// --- readCell: merged cells ---
console.log("\nreadCell — merged cells:");
await test("reads top-left of merged range when targeting A7", async () => {
  const { value } = await readCell(tmpPath, "Sheet1", "A7");
  assert.strictEqual(value, "MergedCell");
});
await test("reads top-left of merged range when targeting B8 (interior)", async () => {
  const { value } = await readCell(tmpPath, "Sheet1", "B8");
  assert.strictEqual(value, "MergedCell");
});

// --- readCell: error cases ---
console.log("\nreadCell — error cases:");
await test("throws for missing file", async () => {
  await assert.rejects(
    () => readCell("/nonexistent/file.xlsx", "Sheet1", "A1"),
    /File not found/
  );
});
await test("throws for missing sheet", async () => {
  await assert.rejects(
    () => readCell(tmpPath, "NoSuchSheet", "A1"),
    /Sheet "NoSuchSheet" not found/
  );
});
await test("throws for empty cell", async () => {
  await assert.rejects(
    () => readCell(tmpPath, "Sheet1", "Z99"),
    /empty/i
  );
});

// ─── Cleanup & summary ───────────────────────────────────────────────────────

unlinkSync(tmpPath);

console.log(`\n${passed + failed} tests: ${passed} passed, ${failed} failed\n`);
if (failed > 0) {
  process.exit(1);
}
