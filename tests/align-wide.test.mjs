import assert from "node:assert/strict";
import test from "node:test";

import { alignWideTable } from "../scripts/lib/align-wide.mjs";
import { validateCityOrder } from "../scripts/lib/cities.mjs";
import { PauseError } from "../scripts/lib/pause.mjs";

const xiamenOnly = validateCityOrder([{ 序号: 1, 城市名: "厦门市" }]);
const twoCities = validateCityOrder([
  { 序号: 1, 城市名: "北京市" },
  { 序号: 2, 城市名: "厦门市" },
]);

function wideRow(values, sourceRow, standardName = "厦门市", metadata = null) {
  return {
    values,
    sourceRow,
    cityMatch: { status: "matched", standardName, method: "exact" },
    cellMetadata: metadata,
  };
}

const model = {
  kind: "wide",
  sheetName: "指标",
  headers: ["省份", "城市", "代码", 2020, 2021],
  rows: [
    wideRow(
      ["福建", "厦门市", 3502, 100, 120],
      2,
      "厦门市",
      ["province-style", "city-style", "code-style", "v2020-style", "v2021-style"],
    ),
  ],
  cityColumn: 1,
  yearColumns: [
    { year: 2020, column: 3 },
    { year: 2021, column: 4 },
  ],
  auxiliaryColumns: [0, 1, 2],
  indicatorLabel: "指标值",
  outOfRangeColumns: [],
};

function assertPause(fn, code) {
  assert.throws(fn, (error) => error instanceof PauseError && error.code === code);
}

test("un-pivots one indicator while retaining auxiliary columns", () => {
  const result = alignWideTable(model, xiamenOnly, { startYear: 2020, endYear: 2021 });
  assert.deepEqual(result.headers, ["省份", "城市", "代码", "年份", "指标值"]);
  assert.deepEqual(result.rows, [
    ["福建", "厦门市", 3502, 2021, 120],
    ["福建", "厦门市", 3502, 2020, 100],
  ]);
});

test("keeps a source year cell with no value truly blank", () => {
  const missingValueModel = {
    ...model,
    rows: [wideRow(["福建", "厦门市", 3502, null, 120], 2)],
  };
  const result = alignWideTable(missingValueModel, xiamenOnly, {
    startYear: 2020,
    endYear: 2021,
  });
  assert.equal(result.rows.find((row) => row[3] === 2020)[4], null);
  assert.equal(result.records.find((record) => record.year === 2020).missing, false);
});

test("value-cell metadata follows the unpivoted indicator cell", () => {
  const result = alignWideTable(model, xiamenOnly, { startYear: 2020, endYear: 2021 });
  const record2020 = result.records.find((record) => record.year === 2020);
  assert.equal(record2020.sourceRow, 2);
  assert.equal(record2020.sourceColumn, 3);
  assert.deepEqual(record2020.cellMetadata, [
    "province-style",
    "city-style",
    "code-style",
    null,
    "v2020-style",
  ]);
});

test("a target year absent from source headers creates a city-year-only row", () => {
  const oneYearModel = {
    ...model,
    headers: ["省份", "城市", "代码", 2021],
    rows: [wideRow(["福建", "厦门市", 3502, 120], 2)],
    yearColumns: [{ year: 2021, column: 3 }],
    auxiliaryColumns: [0, 1, 2],
  };
  const result = alignWideTable(oneYearModel, xiamenOnly, {
    startYear: 2020,
    endYear: 2021,
  });
  assert.deepEqual(result.rows[1], [null, "厦门市", null, 2020, null]);
  assert.deepEqual(result.missingKeys, [{ city: "厦门市", year: 2020 }]);
});

test("an entirely absent city gets city-year-only rows", () => {
  const result = alignWideTable(model, twoCities, { startYear: 2020, endYear: 2021 });
  assert.deepEqual(result.rows.slice(0, 2), [
    [null, "北京市", null, 2021, null],
    [null, "北京市", null, 2020, null],
  ]);
});

test("duplicate source city rows pause", () => {
  const duplicate = {
    ...model,
    rows: [...model.rows, wideRow(["福建", "厦门", 3502, 100, 120], 3)],
  };
  assertPause(
    () => alignWideTable(duplicate, xiamenOnly, { startYear: 2020, endYear: 2021 }),
    "DUPLICATE_CITY_YEAR",
  );
});

test("guarded multi-indicator mode emits shared city-year rows", () => {
  const multi = {
    ...model,
    headers: ["省份", "城市", "代码", "收入2020", "收入2021", "支出2020", "支出2021"],
    rows: [wideRow(["福建", "厦门市", 3502, 100, 120, 60, 70], 2)],
    yearColumns: undefined,
    indicatorBlocks: [
      {
        label: "收入",
        yearColumns: [
          { year: 2020, column: 3 },
          { year: 2021, column: 4 },
        ],
      },
      {
        label: "支出",
        yearColumns: [
          { year: 2020, column: 5 },
          { year: 2021, column: 6 },
        ],
      },
    ],
  };
  const result = alignWideTable(multi, xiamenOnly, {
    startYear: 2020,
    endYear: 2021,
  });
  assert.deepEqual(result.headers, ["省份", "城市", "代码", "年份", "收入", "支出"]);
  assert.deepEqual(result.rows, [
    ["福建", "厦门市", 3502, 2021, 120, 70],
    ["福建", "厦门市", 3502, 2020, 100, 60],
  ]);
});

test("multi-indicator mode pauses on duplicate labels", () => {
  const ambiguous = {
    ...model,
    yearColumns: undefined,
    indicatorBlocks: [
      { label: "收入", yearColumns: model.yearColumns },
      { label: "收入", yearColumns: model.yearColumns },
    ],
  };
  assertPause(
    () => alignWideTable(ambiguous, xiamenOnly, { startYear: 2020, endYear: 2021 }),
    "AMBIGUOUS_MULTI_INDICATOR",
  );
});

test("multi-indicator mode pauses when target-year sets differ", () => {
  const ambiguous = {
    ...model,
    yearColumns: undefined,
    indicatorBlocks: [
      { label: "收入", yearColumns: model.yearColumns },
      { label: "支出", yearColumns: [{ year: 2021, column: 4 }] },
    ],
  };
  assertPause(
    () => alignWideTable(ambiguous, xiamenOnly, { startYear: 2020, endYear: 2021 }),
    "AMBIGUOUS_MULTI_INDICATOR",
  );
});