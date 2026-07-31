import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";
import JSZip from "jszip";

import {
  estimateAlignmentModes,
  runAlignmentV2,
} from "../scripts/run-align-v2.mjs";

const tempDirectory = fileURLToPath(new URL("./tmp/v2/", import.meta.url));
const fixedNow = new Date(2026, 6, 31, 15, 0, 0);

async function writeWorkbook(filePath, sheets) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.rm(filePath, { force: true });
  const workbook = Workbook.create();
  for (const [name, matrix] of Object.entries(sheets)) {
    const sheet = workbook.worksheets.add(name);
    sheet.getRangeByIndexes(0, 0, matrix.length, matrix[0].length).values = matrix;
  }
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(filePath);
}

async function fixture(name) {
  const inputPath = path.join(tempDirectory, `${name}-source.xlsx`);
  const cityOrderPath = path.join(tempDirectory, `${name}-城市顺序.xlsx`);
  const outputDir = path.join(tempDirectory, `${name}-output`);
  await fs.rm(outputDir, { recursive: true, force: true });
  await fs.mkdir(outputDir, { recursive: true });
  await writeWorkbook(cityOrderPath, {
    Sheet1: [
      ["序号", "城市名"],
      [1, "北京市"],
      [2, "厦门市"],
    ],
  });
  await writeWorkbook(inputPath, {
    Data: [
      ["城市", "年份", "可支配收入", "消费支出"],
      ["儋州市", 2022, 31, 10],
      ["北京市", 2022, 11, 5],
      ["三沙市", 2022, 21, 8],
      ["厦门市", 2022, 15, 7],
      ["北京市", 2020, 9, 4],
      ["厦门市", 2020, 13, 6],
      ["儋州市", 2020, 29, 9],
      ["三沙市", 2020, 19, 7],
    ],
  });
  return { inputPath, cityOrderPath, outputDir };
}

test("preflight recommendation reports both modes with time and quota", async () => {
  const paths = await fixture("estimate");
  const recommendations = await estimateAlignmentModes({
    ...paths,
    selectedSheets: ["Data"],
    selectedIndicatorCount: 1,
  });
  assert.equal(recommendations.length, 1);
  assert.equal(recommendations[0].sheetName, "Data");
  assert.match(recommendations[0].fullRow.timeEstimate, /分钟/);
  assert.match(recommendations[0].fullRow.fiveHourQuotaEstimate, /%/);
  assert.match(recommendations[0].selectedIndicators.timeEstimate, /分钟/);
});

test("selected-indicator mode uses all source years and appends outside cities", async () => {
  const paths = await fixture("selected");
  const result = await runAlignmentV2({
    ...paths,
    selectedSheets: ["Data"],
    outputMode: "selected_indicators",
    selectedIndicators: ["可支配收入"],
  }, { now: () => fixedNow });

  assert.equal(result.status, "passed");
  assert.deepEqual(result.yearsBySheet.Data, [2022, 2020]);
  assert.deepEqual(result.outsideOrderCities, ["儋州市", "三沙市"]);
  assert.equal(result.cityOrderDerivedFileCreated, false);
  assert.deepEqual(result.additionalOutputFiles, []);

  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(result.outputPath));
  assert.deepEqual(workbook.worksheets.items.map((sheet) => sheet.name), ["对齐结果"]);
  const values = workbook.worksheets.getItem("对齐结果").getUsedRange().values;
  assert.deepEqual(values[0], ["城市", "年份", "可支配收入"]);
  assert.deepEqual(values.slice(1).map((row) => row[0]), [
    "北京市", "北京市",
    "厦门市", "厦门市",
    "儋州市", "儋州市",
    "三沙市", "三沙市",
  ]);
  assert.equal(values.slice(1).some((row) => row[1] === 2021), false);
});

test("clean writer creates no solid black fill and no audit worksheets", async () => {
  const paths = await fixture("clean");
  const result = await runAlignmentV2({
    ...paths,
    selectedSheets: ["Data"],
    outputMode: "preserve_rows",
  }, { now: () => fixedNow });
  assert.equal(result.status, "passed");

  const zip = await JSZip.loadAsync(await fs.readFile(result.outputPath));
  const styles = await zip.file("xl/styles.xml").async("string");
  const fills = styles.match(/<fills\b[^>]*>[\s\S]*?<\/fills>/u)?.[0] ?? "";
  assert.equal(/patternType="solid"[\s\S]*?fgColor rgb="FF000000"/u.test(fills), false);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(result.outputPath));
  assert.deepEqual(workbook.worksheets.items.map((sheet) => sheet.name), ["对齐结果"]);
});
