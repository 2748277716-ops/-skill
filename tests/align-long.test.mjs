import assert from "node:assert/strict";
import test from "node:test";

import { alignLongTable } from "../scripts/lib/align-long.mjs";
import { validateCityOrder } from "../scripts/lib/cities.mjs";
import { PauseError } from "../scripts/lib/pause.mjs";

const cityContext = validateCityOrder([
  { 序号: 1, 城市名: "北京市" },
  { 序号: 2, 城市名: "厦门市" },
]);

function row(values, sourceRow, standardName, method = "exact", metadata = null) {
  return {
    values,
    sourceRow,
    year: values[2],
    cityMatch: { status: "matched", standardName, method },
    cellMetadata: metadata,
  };
}

const model = {
  kind: "long",
  sheetName: "收入",
  headers: ["省份", "城市", "年份", "收入", "代码"],
  cityColumn: 1,
  yearColumn: 2,
  rows: [
    row(["福建", "厦门", 2020, 100, 3502], 2, "厦门市", "unique_suffix", ["x0", "x1", "x2", "x3", "x4"]),
    row(["北京", "北京市", 2021, 50, 11], 3, "北京市"),
    row(["福建", "厦门", 2021, 120, 3502], 4, "厦门市", "unique_suffix"),
    row(["北京", "北京市", 2020, 40, 11], 5, "北京市"),
  ],
};

function assertPause(fn, code) {
  assert.throws(fn, (error) => error instanceof PauseError && error.code === code);
}

test("sorts complete rows and preserves all auxiliary columns", () => {
  const result = alignLongTable(model, cityContext, { startYear: 2020, endYear: 2021 });
  assert.deepEqual(result.headers, ["省份", "城市", "年份", "收入", "代码"]);
  assert.deepEqual(result.rows, [
    ["北京", "北京市", 2021, 50, 11],
    ["北京", "北京市", 2020, 40, 11],
    ["福建", "厦门市", 2021, 120, 3502],
    ["福建", "厦门市", 2020, 100, 3502],
  ]);
});

test("creates missing rows with only city and year populated", () => {
  const missingYearModel = { ...model, rows: model.rows.filter((item) => item.sourceRow !== 2) };
  const result = alignLongTable(missingYearModel, cityContext, {
    startYear: 2020,
    endYear: 2021,
  });
  const inserted = result.rows.find(
    (values) => values[1] === "厦门市" && values[2] === 2020,
  );
  assert.deepEqual(inserted, [null, "厦门市", 2020, null, null]);
  assert.deepEqual(result.missingKeys, [{ city: "厦门市", year: 2020 }]);
});

test("multiple indicators and auxiliary values move together", () => {
  const multiModel = {
    ...model,
    headers: ["省份", "城市", "年份", "收入", "支出", "代码", "英文名"],
    rows: model.rows.map((item) => ({
      ...item,
      values: [
        item.values[0],
        item.values[1],
        item.values[2],
        item.values[3],
        item.values[3] / 2,
        item.values[4],
        item.cityMatch.standardName === "北京市" ? "Beijing" : "Xiamen",
      ],
    })),
  };
  const result = alignLongTable(multiModel, cityContext, {
    startYear: 2020,
    endYear: 2021,
  });
  assert.deepEqual(result.rows[0], ["北京", "北京市", 2021, 50, 25, 11, "Beijing"]);
  assert.deepEqual(result.rows[2], ["福建", "厦门市", 2021, 120, 60, 3502, "Xiamen"]);
  assert.equal(result.sourceRecordCount, 4);
});

test("an entirely absent city receives every target-year blank row", () => {
  const beijingOnly = {
    ...model,
    rows: model.rows.filter((item) => item.cityMatch.standardName === "北京市"),
  };
  const result = alignLongTable(beijingOnly, cityContext, {
    startYear: 2020,
    endYear: 2021,
  });
  assert.deepEqual(result.rows.slice(2), [
    [null, "厦门市", 2021, null, null],
    [null, "厦门市", 2020, null, null],
  ]);
});

test("existing source provenance and cell metadata stay on the full row", () => {
  const result = alignLongTable(model, cityContext, { startYear: 2020, endYear: 2021 });
  const record = result.records.find(
    (item) => item.values[1] === "厦门市" && item.values[2] === 2020,
  );
  assert.equal(record.sourceRow, 2);
  assert.equal(record.missing, false);
  assert.deepEqual(record.cellMetadata, ["x0", "x1", "x2", "x3", "x4"]);
});

test("duplicate city-year pauses even when rows are identical", () => {
  const duplicate = { ...model, rows: [...model.rows, { ...model.rows[0], sourceRow: 8 }] };
  assertPause(
    () => alignLongTable(duplicate, cityContext, { startYear: 2020, endYear: 2021 }),
    "DUPLICATE_CITY_YEAR",
  );
});

test("source rows outside an approved target range are retained separately", () => {
  const outOfRange = {
    ...model,
    rows: [
      ...model.rows,
      row(["福建", "厦门", 2019, 90, 3502], 9, "厦门市", "unique_suffix"),
    ],
  };
  const result = alignLongTable(outOfRange, cityContext, {
    startYear: 2020,
    endYear: 2021,
    allowOutOfRangeYears: true,
  });
  assert.equal(result.outOfRangeRows.length, 1);
  assert.deepEqual(result.outOfRangeRows[0].values, ["福建", "厦门", 2019, 90, 3502]);
  assert.equal(result.rows.length, 4);
});