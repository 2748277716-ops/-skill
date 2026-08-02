import assert from "node:assert/strict";
import test from "node:test";

import { validateCityOrder } from "../scripts/lib/cities.mjs";
import { detectTableModel } from "../scripts/lib/detect.mjs";
import { PauseError } from "../scripts/lib/pause.mjs";

const cityContext = validateCityOrder([
  { 序号: 1, 城市名: "北京市" },
  { 序号: 2, 城市名: "厦门市" },
]);
const baseConfig = {
  sheetName: "数据",
  startYear: 2020,
  endYear: 2021,
  workbookSheets: ["数据"],
  selectedSheets: ["数据"],
  mergedRanges: [],
};

function assertPause(fn, code) {
  assert.throws(fn, (error) => error instanceof PauseError && error.code === code);
}

test("detects a long table from city and year values", () => {
  const matrix = [
    ["省份", "城市", "年份", "收入"],
    ["福建", "厦门市", 2021, 120],
    ["福建", "厦门市", 2020, 100],
  ];
  const model = detectTableModel(matrix, baseConfig, cityContext);
  assert.equal(model.kind, "long");
  assert.equal(model.cityColumn, 1);
  assert.equal(model.yearColumn, 2);
  assert.deepEqual(model.rows.map((row) => row.sourceRow), [2, 3]);
});

test("detects a wide table from numeric and trimmed-text year headers", () => {
  const matrix = [
    ["省份", "城市", "代码", 2020, " 2021 "],
    ["福建", "厦门市", 3502, 100, 120],
  ];
  const model = detectTableModel(matrix, baseConfig, cityContext);
  assert.equal(model.kind, "wide");
  assert.deepEqual(model.yearColumns, [
    { year: 2020, column: 3 },
    { year: 2021, column: 4 },
  ]);
  assert.deepEqual(model.auxiliaryColumns, [0, 1, 2]);
});

test("detects a wide table from date-formatted year header values", () => {
  const matrix = [
    ["城市", new Date("2020-12-01T00:00:00Z"), new Date("2021-12-01T00:00:00Z")],
    ["厦门市", 100, 120],
  ];
  const model = detectTableModel(matrix, baseConfig, cityContext);
  assert.equal(model.kind, "wide");
  assert.deepEqual(model.yearColumns, [
    { year: 2020, column: 1 },
    { year: 2021, column: 2 },
  ]);
});

test("two city-like columns pause instead of guessing", () => {
  const matrix = [
    ["城市", "城市名称", "年份", "值"],
    ["厦门市", "厦门市", 2021, 1],
  ];
  assertPause(
    () => detectTableModel(matrix, baseConfig, cityContext),
    "AMBIGUOUS_CITY_COLUMN",
  );
});

test("two long year-like columns pause", () => {
  const matrix = [
    ["城市", "年份", "统计年份", "值"],
    ["厦门市", 2021, 2021, 1],
  ];
  assertPause(
    () => detectTableModel(matrix, baseConfig, cityContext),
    "AMBIGUOUS_YEAR_COLUMN",
  );
});

test("source year outside target range pauses", () => {
  const matrix = [
    ["城市", "年份", "值"],
    ["厦门市", 1999, 1],
  ];
  assertPause(
    () => detectTableModel(matrix, baseConfig, cityContext),
    "OUT_OF_RANGE_YEARS",
  );
});

test("wide source year header outside target range pauses", () => {
  const matrix = [
    ["城市", 1999, 2020, 2021],
    ["厦门市", 1, 2, 3],
  ];
  assertPause(
    () => detectTableModel(matrix, baseConfig, cityContext),
    "OUT_OF_RANGE_YEARS",
  );
});

test("merged data range pauses", () => {
  const matrix = [
    ["城市", "年份", "值"],
    ["厦门市", 2021, 1],
  ];
  assertPause(
    () =>
      detectTableModel(
        matrix,
        { ...baseConfig, mergedRanges: ["A2:B2"] },
        cityContext,
      ),
    "UNSAFE_MERGED_CELLS",
  );
});

test("duplicate wide city rows pause", () => {
  const matrix = [
    ["城市", 2020, 2021],
    ["厦门市", 1, 2],
    ["厦门", 1, 2],
  ];
  assertPause(
    () => detectTableModel(matrix, baseConfig, cityContext),
    "DUPLICATE_CITY_YEAR",
  );
});

test("unselected multi-sheet workbook pauses", () => {
  const matrix = [
    ["城市", "年份", "值"],
    ["厦门市", 2021, 1],
  ];
  assertPause(
    () =>
      detectTableModel(
        matrix,
        {
          ...baseConfig,
          workbookSheets: ["数据", "说明"],
          selectedSheets: [],
        },
        cityContext,
      ),
    "MULTIPLE_SHEETS",
  );
});

test("unmatched city pauses with source evidence", () => {
  const matrix = [
    ["城市", "年份", "值"],
    ["不存在市", 2021, 1],
  ];
  assertPause(
    () => detectTableModel(matrix, baseConfig, cityContext),
    "UNMATCHED_CITY",
  );
});

test("unrecognized structure pauses", () => {
  const matrix = [
    ["名称", "备注"],
    ["A", "B"],
  ];
  assertPause(
    () => detectTableModel(matrix, baseConfig, cityContext),
    "UNRECOGNIZED_TABLE",
  );
});