import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";
import JSZip from "jszip";


import {
  inspectWorkbookPackage,
  readWorksheetTable,
} from "../scripts/lib/ooxml-reader.mjs";
import { PauseError } from "../scripts/lib/pause.mjs";
import {
  addOrphanWorksheetRelationship,
  createWorkbookIoFixtures,
  removeReferencedWorksheetPart,
} from "./helpers/fixture-workbooks.mjs";

const testDirectory = fileURLToPath(new URL("./tmp/ooxml-reader/", import.meta.url));

test.before(async () => {
  await fs.mkdir(testDirectory, { recursive: true });
});

test("unreferenced missing worksheet relationship is reported but tolerated", async () => {
  const { normalPath } = await createWorkbookIoFixtures(testDirectory);
  const orphanPath = path.join(testDirectory, "workbook-orphan-relationship.xlsx");
  await addOrphanWorksheetRelationship(normalPath, orphanPath);

  const info = await inspectWorkbookPackage(orphanPath);

  assert.deepEqual(info.sheets.map((sheet) => sheet.name), ["Data", "Notes"]);
  assert.deepEqual(
    info.orphanRelationships.map((relationship) => relationship.id),
    ["rIdOrphan"],
  );
});

test("referenced missing worksheet part pauses package inspection", async () => {
  const { normalPath } = await createWorkbookIoFixtures(testDirectory);
  const missingPath = path.join(testDirectory, "workbook-referenced-missing.xlsx");
  await removeReferencedWorksheetPart(normalPath, missingPath);

  await assert.rejects(
    inspectWorkbookPackage(missingPath),
    (error) =>
      error instanceof PauseError &&
      error.code === "INVALID_WORKBOOK_PACKAGE" &&
      error.evidence.partPath === "xl/worksheets/sheet1.xml",
  );
});

async function createTypedWorkbook(
  outputPath,
  { formula = "=D2*2", mergeInside = false } = {},
) {
  await fs.rm(outputPath, { force: true });
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Typed");
  sheet.getRange("A1:G2").values = [
    ["城市", "年份", "文本", "数值", "布尔", "日期", "公式"],
    [
      "厦门市",
      2021,
      "甲&乙",
      12.5,
      true,
      new Date("2021-01-02T00:00:00Z"),
      null,
    ],
  ];
  sheet.getRange("G2").formulas = [[formula]];
  sheet.getRange("F2").format.numberFormat = "yyyy-mm-dd";
  if (mergeInside) sheet.getRange("A2:B2").merge();
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outputPath);
  return outputPath;
}

async function removeFormulaCachedValue(sourcePath, outputPath, cellAddress = "G2") {
  const zip = await JSZip.loadAsync(await fs.readFile(sourcePath));
  const partPath = "xl/worksheets/sheet1.xml";
  const part = zip.file(partPath);
  if (!part) throw new Error(`Missing OOXML part: ${partPath}`);
  let xml = await part.async("string");
  const cellExpression = new RegExp(
    `(<(?:[\\w.-]+:)?c\\b[^>]*\\br="${cellAddress}"[^>]*>)([\\s\\S]*?)(<\\/(?:[\\w.-]+:)?c>)`,
  );
  const match = xml.match(cellExpression);
  if (!match) throw new Error(`Missing formula cell: ${cellAddress}`);
  const bodyWithoutValue = match[2].replace(
    /<(?:[\w.-]+:)?v\b[^>]*>[\s\S]*?<\/(?:[\w.-]+:)?v>/i,
    "",
  );
  xml = xml.replace(
    cellExpression,
    `${match[1]}${bodyWithoutValue}${match[3]}`,
  );
  zip.file(partPath, xml);
  await fs.writeFile(
    outputPath,
    await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }),
  );
  return outputPath;
}

async function prependSelfClosingBlankCell(sourcePath, outputPath) {
  const zip = await JSZip.loadAsync(await fs.readFile(sourcePath));
  const partPath = "xl/worksheets/sheet1.xml";
  const part = zip.file(partPath);
  if (!part) throw new Error(`Missing worksheet part: ${partPath}`);
  const xml = await part.async("string");
  const rowStart = /(<(?:[\w.-]+:)?row\b[^>]*\br="1"[^>]*>)/i;
  if (!rowStart.test(xml)) throw new Error("Missing first worksheet row");
  zip.file(partPath, xml.replace(rowStart, '$1<c r="A1" s="0"/>'));
  await fs.writeFile(
    outputPath,
    await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }),
  );
  return outputPath;
}

test("typed worksheet values and cached formula results are decoded", async () => {
  const workbookPath = path.join(testDirectory, "typed-values.xlsx");
  await createTypedWorkbook(workbookPath);

  const sheet = await readWorksheetTable(workbookPath, "Typed");

  assert.deepEqual(sheet.matrix[0], [
    "城市",
    "年份",
    "文本",
    "数值",
    "布尔",
    "日期",
    "公式",
  ]);
  assert.equal(sheet.matrix[1][0], "厦门市");
  assert.equal(sheet.matrix[1][1], 2021);
  assert.equal(sheet.matrix[1][2], "甲&乙");
  assert.equal(sheet.matrix[1][3], 12.5);
  assert.equal(sheet.matrix[1][4], true);
  assert.ok(sheet.matrix[1][5] instanceof Date);
  assert.equal(sheet.matrix[1][5].toISOString(), "2021-01-02T00:00:00.000Z");
  assert.equal(sheet.matrix[1][6], 25);
  assert.equal(sheet.formulaEvents[0].cell, "G2");
});

test("self-closing blank cell does not steal the following shared-string cell", async () => {
  const basePath = path.join(testDirectory, "sparse-header-base.xlsx");
  const sparsePath = path.join(testDirectory, "sparse-header.xlsx");
  await fs.rm(basePath, { force: true });
  await fs.rm(sparsePath, { force: true });
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("SparseHeader");
  sheet.getRange("B1:C2").values = [
    ["城市", "值"],
    ["厦门市", 1],
  ];
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(basePath);
  await prependSelfClosingBlankCell(basePath, sparsePath);

  const parsed = await readWorksheetTable(sparsePath, "SparseHeader");

  assert.deepEqual(parsed.matrix[0], [null, "城市", "值"]);
  assert.deepEqual(
    parsed.headerMetadata.map((item) => item.address),
    ["A1", "B1", "C1"],
  );
});

test("formula errors and missing cached values pause reading", async () => {
  const errorPath = path.join(testDirectory, "typed-formula-error.xlsx");
  await createTypedWorkbook(errorPath, { formula: "=1/0" });
  await assert.rejects(
    readWorksheetTable(errorPath, "Typed"),
    (error) => error instanceof PauseError && error.code === "FORMULA_ERROR",
  );

  const normalPath = path.join(testDirectory, "typed-formula-normal.xlsx");
  const missingPath = path.join(testDirectory, "typed-formula-no-cache.xlsx");
  await createTypedWorkbook(normalPath);
  await removeFormulaCachedValue(normalPath, missingPath);
  await assert.rejects(
    readWorksheetTable(missingPath, "Typed"),
    (error) =>
      error instanceof PauseError &&
      error.code === "FORMULA_VALUE_UNAVAILABLE",
  );
});

test("unsafe merges pause and projected columns retain source cells", async () => {
  const mergedPath = path.join(testDirectory, "typed-merged.xlsx");
  await createTypedWorkbook(mergedPath, { mergeInside: true });
  await assert.rejects(
    readWorksheetTable(mergedPath, "Typed"),
    (error) =>
      error instanceof PauseError && error.code === "UNSAFE_MERGED_CELLS",
  );

  const normalPath = path.join(testDirectory, "typed-projection.xlsx");
  await createTypedWorkbook(normalPath);
  const sheet = await readWorksheetTable(normalPath, "Typed", {
    projectHeaders: ["城市", "年份", "公式"],
  });
  assert.deepEqual(sheet.matrix[0], ["城市", "年份", "公式"]);
  assert.deepEqual(sheet.rowRecords[0].sourceCells, ["A2", "B2", "G2"]);
});
