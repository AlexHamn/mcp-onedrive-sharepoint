import test from "node:test";
import assert from "node:assert/strict";

import { sanitizeFileName, analyzePath } from "../tools/utils/path-helper.js";

test("sanitizeFileName preserves spaces", () => {
  assert.equal(sanitizeFileName("My File.pdf"), "My File.pdf");
  assert.equal(sanitizeFileName("S01 DE 14.pdf"), "S01 DE 14.pdf");
  assert.equal(sanitizeFileName("Soportes Balbuena.dwg"), "Soportes Balbuena.dwg");
});

test("sanitizeFileName preserves spaces alongside hyphens and digits", () => {
  assert.equal(
    sanitizeFileName("00-PH-01-JBA PERFIL HIDRÁULICO-Model.pdf"),
    "00-PH-01-JBA PERFIL HIDRÁULICO-Model.pdf",
  );
});

test("sanitizeFileName collapses runs of spaces to a single space", () => {
  assert.equal(sanitizeFileName("A    B   C.pdf"), "A B C.pdf");
});

test("sanitizeFileName replaces control characters (incl. tab) with underscores", () => {
  // Tab is in the control-char range and is replaced before whitespace collapse.
  assert.equal(sanitizeFileName("A\tB.pdf"), "A_B.pdf");
});

test("sanitizeFileName replaces SharePoint-forbidden chars with underscores", () => {
  assert.equal(sanitizeFileName('weird<>:"|?*name.pdf'), "weird_______name.pdf");
  assert.equal(sanitizeFileName("with/slash.pdf"), "with_slash.pdf");
  assert.equal(sanitizeFileName("with\\backslash.pdf"), "with_backslash.pdf");
});

test("sanitizeFileName trims leading/trailing whitespace and dots", () => {
  assert.equal(sanitizeFileName("  spaced.pdf  "), "spaced.pdf");
  assert.equal(sanitizeFileName("...leading-dots.pdf"), "leading-dots.pdf");
  assert.equal(sanitizeFileName("trailing-dots.pdf..."), "trailing-dots.pdf");
});

test("sanitizeFileName prefixes Windows reserved names", () => {
  assert.equal(sanitizeFileName("CON.txt"), "file_CON.txt");
  assert.equal(sanitizeFileName("nul.log"), "file_nul.log");
  assert.equal(sanitizeFileName("Lpt1.dat"), "file_Lpt1.dat");
});

test("sanitizeFileName falls back to untitled_file for empty input", () => {
  assert.equal(sanitizeFileName(""), "untitled_file");
  assert.equal(sanitizeFileName("   "), "untitled_file");
  assert.equal(sanitizeFileName("..."), "untitled_file");
});

test("sanitizeFileName truncates names longer than 200 chars while keeping the extension", () => {
  const longBase = "a".repeat(300);
  const result = sanitizeFileName(`${longBase}.pdf`);
  assert.equal(result.length, 200);
  assert.ok(result.endsWith(".pdf"));
});

test("analyzePath preserves spaces in both folder segments and filename", () => {
  const info = analyzePath("/My Folder/Sub Folder/My File.pdf");
  assert.equal(info.folderPath, "My Folder/Sub Folder");
  assert.equal(info.fileName, "My File.pdf");
  assert.equal(info.sanitizedPath, "My Folder/Sub Folder/My File.pdf");
  assert.equal(info.needsFolderCreation, true);
});

test("analyzePath normalizes redundant slashes without disturbing names", () => {
  const info = analyzePath("///A/B//C.pdf///");
  assert.equal(info.sanitizedPath, "A/B/C.pdf");
});

test("analyzePath handles root-only file (no folder creation)", () => {
  const info = analyzePath("report.xlsx");
  assert.equal(info.folderPath, "");
  assert.equal(info.fileName, "report.xlsx");
  assert.equal(info.needsFolderCreation, false);
});
