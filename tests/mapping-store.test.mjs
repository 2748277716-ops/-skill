import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

import { validateCityOrder } from "../scripts/lib/cities.mjs";
import {
  appendConfirmedMappings,
  readMappingRows,
} from "../scripts/lib/mapping-store.mjs";
import { PauseError } from "../scripts/lib/pause.mjs";

const tempDirectory = fileURLToPath(new URL("./tmp/", import.meta.url));
const cityOrder = validateCityOrder([
  { 序号: 1, 城市名: "北京市" },
  { 序号: 2, 城市名: "厦门市" },
]);
const mapping = {
  sourceName: "厦冂市",
  standardName: "厦门市",
  matchType: "错字确认",
  confirmedAt: "2026-07-30T12:00:00+08:00",
  sourceFile: "来源.xlsx",
  sourceSheet: "数据",
  note: "用户确认",
};

function testPath(name) {
  return path.join(tempDirectory, `mapping-${name}.xlsx`);
}

async function freshPath(name) {
  await fs.mkdir(tempDirectory, { recursive: true });
  const target = testPath(name);
  await fs.rm(target, { force: true });
  return target;
}

function assertPause(code) {
  return (error) => error instanceof PauseError && error.code === code;
}

test("creates the mapping workbook on first confirmed mapping", async () => {
  const target = await freshPath("create");
  const result = await appendConfirmedMappings(target, [mapping], cityOrder);
  const rows = await readMappingRows(target);
  assert.equal(result.appendedCount, 1);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].原始城市名, "厦冂市");
  assert.equal(rows[0].标准城市名, "厦门市");
  assert.equal(rows[0].来源文件, "来源.xlsx");
});

test("does not duplicate an identical mapping", async () => {
  const target = await freshPath("dedupe");
  await appendConfirmedMappings(target, [mapping], cityOrder);
  const before = await fs.readFile(target);
  const result = await appendConfirmedMappings(target, [mapping], cityOrder);
  const after = await fs.readFile(target);
  assert.equal(result.appendedCount, 0);
  assert.equal((await readMappingRows(target)).length, 1);
  assert.deepEqual(after, before);
});

test("rejects a conflicting confirmed mapping before writing", async () => {
  const target = await freshPath("conflict");
  await appendConfirmedMappings(target, [mapping], cityOrder);
  const before = await fs.readFile(target);
  await assert.rejects(
    appendConfirmedMappings(
      target,
      [{ ...mapping, standardName: "北京市" }],
      cityOrder,
    ),
    assertPause("MAPPING_CONFLICT"),
  );
  assert.deepEqual(await fs.readFile(target), before);
});

test("occupied destination leaves original bytes unchanged", async () => {
  const target = await freshPath("occupied");
  await appendConfirmedMappings(target, [mapping], cityOrder);
  const before = await fs.readFile(target);
  const fileAdapter = {
    ...fs,
    rename: async (from, to) => {
      if (path.resolve(from) === path.resolve(target) && String(to).includes(".backup-")) {
        const error = new Error("occupied");
        error.code = "EBUSY";
        throw error;
      }
      return fs.rename(from, to);
    },
  };
  await assert.rejects(
    appendConfirmedMappings(
      target,
      [{ ...mapping, sourceName: "北亰市", standardName: "北京市" }],
      cityOrder,
      { fileAdapter },
    ),
    assertPause("MAPPING_FILE_OCCUPIED"),
  );
  assert.deepEqual(await fs.readFile(target), before);
});

test("temporary export failure leaves original bytes unchanged", async () => {
  const target = await freshPath("export-failure");
  await appendConfirmedMappings(target, [mapping], cityOrder);
  const before = await fs.readFile(target);
  await assert.rejects(
    appendConfirmedMappings(
      target,
      [{ ...mapping, sourceName: "北亰市", standardName: "北京市" }],
      cityOrder,
      { exportWorkbook: async () => { throw new Error("injected export failure"); } },
    ),
    assertPause("MAPPING_WRITE_FAILED"),
  );
  assert.deepEqual(await fs.readFile(target), before);
});

test("final verification failure restores original bytes", async () => {
  const target = await freshPath("verify-failure");
  await appendConfirmedMappings(target, [mapping], cityOrder);
  const before = await fs.readFile(target);
  await assert.rejects(
    appendConfirmedMappings(
      target,
      [{ ...mapping, sourceName: "北亰市", standardName: "北京市" }],
      cityOrder,
      { verifyWorkbook: async () => { throw new Error("injected verification failure"); } },
    ),
    assertPause("MAPPING_WRITE_FAILED"),
  );
  assert.deepEqual(await fs.readFile(target), before);
});
test("preserves existing custom columns and values", async () => {
  const target = await freshPath("custom-column");
  await appendConfirmedMappings(target, [mapping], cityOrder);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(target));
  const sheet = workbook.worksheets.getItemAt(0);
  sheet.getRange("H1:H2").values = [["自定义列"], ["必须保留"]];
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(target);

  await appendConfirmedMappings(
    target,
    [{ ...mapping, sourceName: "北亰市", standardName: "北京市" }],
    cityOrder,
  );
  const rows = await readMappingRows(target);
  assert.equal(rows[0].自定义列, "必须保留");
  assert.equal(rows.length, 2);
});