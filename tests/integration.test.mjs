import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

import { readMappingRows } from "../scripts/lib/mapping-store.mjs";
import { runAlignment } from "../scripts/run-align.mjs";

const tempDirectory = fileURLToPath(new URL("./tmp/", import.meta.url));
const fixedNow = new Date(2026, 6, 30, 12, 0, 0);

async function writeWorkbook(filePath, sheets) {
  await fs.rm(filePath, { force: true });
  const workbook = Workbook.create();
  for (const [name, matrix] of Object.entries(sheets)) {
    const sheet = workbook.worksheets.add(name);
    sheet.getRangeByIndexes(0, 0, matrix.length, matrix[0].length).values = matrix;
  }
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(filePath);
}

async function makeFixture(name, sheets = null) {
  await fs.mkdir(tempDirectory, { recursive: true });
  const inputPath = path.join(tempDirectory, `${name}-source.xlsx`);
  const cityOrderPath = path.join(tempDirectory, `${name}-城市顺序.xlsx`);
  const mappingPath = path.join(tempDirectory, `${name}-城市名称映射表.xlsx`);
  const outputPath = path.join(tempDirectory, `${name}-source_城市面板对齐_20260730-120000.xlsx`);
  await fs.rm(mappingPath, { force: true });
  await fs.rm(outputPath, { force: true });
  await writeWorkbook(cityOrderPath, {
    Sheet1: [
      ["序号", "城市名"],
      [1, "北京市"],
      [2, "厦门市"],
    ],
  });
  await writeWorkbook(inputPath, sheets ?? {
    Data: [
      ["城市", "年份", "值"],
      ["厦冂市", 2021, 10],
    ],
  });
  return { inputPath, cityOrderPath, mappingPath };
}

function baseConfig(paths, selectedSheets = ["Data"]) {
  return {
    ...paths,
    selectedSheets,
    startYear: 2021,
    endYear: 2021,
    outputDir: tempDirectory,
    approvedSheetNames: {},
    approvedExcludedYears: [],
    confirmedMappings: [],
  };
}

const runtime = { now: () => fixedNow };

test("does not create formal output when a typo needs confirmation", async () => {
  const paths = await makeFixture("paused-typo");
  const result = await runAlignment(baseConfig(paths), runtime);
  assert.equal(result.status, "paused");
  assert.equal(result.code, "CITY_CONFIRMATION_REQUIRED");
  assert.equal(await fs.access(result.expectedOutputPath).then(() => true, () => false), false);
  assert.equal(await fs.access(paths.mappingPath).then(() => true, () => false), false);
});

test("resumes after explicit confirmation and appends mapping", async () => {
  const paths = await makeFixture("confirmed-typo");
  const result = await runAlignment({
    ...baseConfig(paths),
    confirmedMappings: [{
      sourceName: "厦冂市",
      standardName: "厦门市",
      matchType: "错字确认",
      confirmedAt: "2026-07-30T12:00:00+08:00",
      sourceFile: path.basename(paths.inputPath),
      sourceSheet: "Data",
      note: "用户确认",
    }],
  }, runtime);
  assert.equal(result.status, "passed");
  assert.equal(await fs.access(result.outputPath).then(() => true, () => false), true);
  const mappings = await readMappingRows(paths.mappingPath);
  assert.equal(mappings.length, 1);
  assert.equal(mappings[0].标准城市名, "厦门市");
});

test("uses the source base name and timestamp without overwriting", async () => {
  const paths = await makeFixture("occupied-output", {
    Data: [
      ["城市", "年份", "值"],
      ["厦门市", 2021, 10],
    ],
  });
  const expected = path.join(tempDirectory, "occupied-output-source_城市面板对齐_20260730-120000.xlsx");
  await fs.writeFile(expected, "occupied");
  const before = await fs.readFile(expected);
  const result = await runAlignment(baseConfig(paths), runtime);
  assert.equal(result.status, "paused");
  assert.equal(result.code, "OUTPUT_FILE_OCCUPIED");
  assert.equal(result.expectedOutputPath, expected);
  assert.deepEqual(await fs.readFile(expected), before);
});

test("invalid derived sheet names pause until approved names are supplied", async () => {
  const longName = "长".repeat(29);
  const paths = await makeFixture("long-sheet-name", {
    [longName]: [
      ["城市", "年份", "值"],
      ["厦门市", 2021, 10],
    ],
    短表: [
      ["城市", "年份", "值"],
      ["北京市", 2021, 20],
    ],
  });
  const result = await runAlignment(
    baseConfig(paths, [longName, "短表"]),
    runtime,
  );
  assert.equal(result.status, "paused");
  assert.equal(result.code, "INVALID_SHEET_NAME");
  assert.equal(await fs.access(result.expectedOutputPath).then(() => true, () => false), false);
});

test("multiple selected sheets use approved names plus one audit sheet", async () => {
  const longName = "长".repeat(29);
  const paths = await makeFixture("multi-sheet", {
    [longName]: [
      ["城市", "年份", "值"],
      ["厦门市", 2021, 10],
    ],
    短表: [
      ["城市", "年份", "值"],
      ["北京市", 2021, 20],
    ],
  });
  const result = await runAlignment({
    ...baseConfig(paths, [longName, "短表"]),
    approvedSheetNames: { [longName]: "结果一", 短表: "结果二" },
  }, runtime);
  assert.equal(result.status, "passed");
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(result.outputPath));
  assert.deepEqual(workbook.worksheets.items.map((sheet) => sheet.name), [
    "结果一",
    "结果二",
    "核验结果",
    "缺失键",
  ]);
});