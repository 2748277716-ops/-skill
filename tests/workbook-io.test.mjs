import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { PauseError } from "../scripts/lib/pause.mjs";
import {
  readWorkbookModel,
  writeResultWorkbook,
} from "../scripts/lib/workbook-io.mjs";
import { createWorkbookIoFixtures } from "./helpers/fixture-workbooks.mjs";

const tempDirectory = fileURLToPath(new URL("./tmp/", import.meta.url));

async function fixtures() {
  return createWorkbookIoFixtures(tempDirectory);
}

function assertPause(code) {
  return (error) => error instanceof PauseError && error.code === code;
}

test("reads formula text and current value separately", async () => {
  const paths = await fixtures();
  const model = await readWorkbookModel(paths.normalPath, ["Data"]);
  assert.deepEqual(model.sheetNames, ["Data", "Notes"]);
  assert.equal(model.sheets.length, 1);
  assert.equal(model.sheets[0].cells.D2.formula, "=C2*2");
  assert.equal(model.sheets[0].cells.D2.value, 20);
  assert.equal(model.formulaEvents[0].currentValue, 20);
  assert.equal(model.formulaEvents[0].formula, "=C2*2");
  assert.match(model.sha256, /^[0-9a-f]{64}$/u);
  assert.match(model.sheets[0].sheetSha256, /^[0-9a-f]{64}$/u);
});

test("captures style, comment, hyperlink and source address", async () => {
  const paths = await fixtures();
  const model = await readWorkbookModel(paths.normalPath, ["Data"]);
  const valueCell = model.sheets[0].cells.C2;
  assert.equal(valueCell.address, "C2");
  assert.equal(valueCell.style.numberFormat, "0.00");
  assert.equal(valueCell.style.fill, "#FFFF00");
  assert.equal(valueCell.style.borders.top.style, "thin");
  assert.equal(valueCell.comment, "fixture-comment");
  assert.equal(model.sheets[0].cells.E2.hyperlink.target, "https://example.com/");
});

test("allows a merged note outside the contiguous data region", async () => {
  const paths = await fixtures();
  const model = await readWorkbookModel(paths.normalPath, ["Data"]);
  assert.equal(model.sheets[0].dataRange, "A1:E3");
  assert.deepEqual(model.sheets[0].mergedRanges, ["H1:I1"]);
});

test("pauses on formula errors", async () => {
  const paths = await fixtures();
  await assert.rejects(
    readWorkbookModel(paths.errorPath, ["Data"]),
    assertPause("FORMULA_ERROR"),
  );
});

test("pauses on a merge intersecting the data region", async () => {
  const paths = await fixtures();
  await assert.rejects(
    readWorkbookModel(paths.mergedPath, ["Data"]),
    assertPause("UNSAFE_MERGED_CELLS"),
  );
});

test("writes values only and preserves transferable style, comment and hyperlink", async () => {
  const paths = await fixtures();
  const source = await readWorkbookModel(paths.normalPath, ["Data"]);
  const sheet = source.sheets[0];
  const outputPath = path.join(tempDirectory, "workbook-io-output.xlsx");
  await fs.rm(outputPath, { force: true });

  const records = sheet.rowRecords.map((record) => ({
    ...record,
    values: [...record.values],
  }));
  const result = {
    outputs: [{
      sheetName: "Data",
      outputSheetName: "对齐结果",
      headers: [...sheet.matrix[0]],
      rows: records.map((record) => record.values),
      records,
      cityColumn: 0,
      yearColumn: 1,
      targetYears: [2021, 2020],
      missingKeys: [],
      mappingEvents: [],
    }],
    audit: {
      passed: true,
      auditRows: [{ section: "结论", item: "最终结论", status: "passed", value: "通过", details: null }],
    },
    formulaEvents: source.formulaEvents,
  };

  const written = await writeResultWorkbook(result, outputPath);
  assert.equal(written.outputPath, outputPath);
  const reread = await readWorkbookModel(outputPath, ["对齐结果"]);
  assert.equal(reread.sheets[0].cells.D2.formula, null);
  assert.equal(reread.sheets[0].cells.D2.value, 20);
  assert.equal(reread.sheets[0].cells.C2.style.fill, "#FFFF00");
  assert.equal(reread.sheets[0].cells.C2.comment, "fixture-comment");
  assert.equal(reread.sheets[0].cells.E2.hyperlink.target, "https://example.com/");
});

test("never overwrites an existing output", async () => {
  const paths = await fixtures();
  const outputPath = path.join(tempDirectory, "occupied-output.xlsx");
  await fs.copyFile(paths.normalPath, outputPath);
  await assert.rejects(
    writeResultWorkbook({ outputs: [], audit: { passed: true, auditRows: [] } }, outputPath),
    assertPause("OUTPUT_FILE_OCCUPIED"),
  );
});