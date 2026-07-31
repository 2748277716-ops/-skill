import assert from "node:assert/strict";
import test from "node:test";

import { alignLongTable } from "../scripts/lib/align-long.mjs";
import { validateCityOrder } from "../scripts/lib/cities.mjs";
import { recommendProcessingModes } from "../scripts/lib/recommendation.mjs";

const cityContext = validateCityOrder([
  { 序号: 1, 城市名: "北京市" },
  { 序号: 2, 城市名: "厦门市" },
]);

function row(city, year, income, expenditure, sourceRow, match = null) {
  return {
    values: [city, year, income, expenditure],
    sourceRow,
    year,
    cityMatch: match ?? {
      status: cityContext.rankByName.has(city) ? "matched" : "outside_order",
      standardName: city,
      method: cityContext.rankByName.has(city) ? "exact" : "outside_order",
    },
    cellMetadata: [
      { address: `A${sourceRow}` },
      { address: `B${sourceRow}` },
      { address: `C${sourceRow}` },
      { address: `D${sourceRow}` },
    ],
  };
}

function model(rows) {
  return {
    kind: "long",
    sheetName: "数据",
    headers: ["城市", "年份", "可支配收入", "消费支出"],
    cityColumn: 0,
    yearColumn: 1,
    rows,
  };
}

test("omitted year range uses exactly all source years in descending order", () => {
  const result = alignLongTable(model([
    row("北京市", 2022, 10, 5, 2),
    row("北京市", 2020, 8, 4, 3),
    row("厦门市", 2022, 20, 9, 4),
    row("厦门市", 2020, 18, 8, 5),
  ]), cityContext);

  assert.deepEqual(result.targetYears, [2022, 2020]);
  assert.equal(result.rows.some((values) => values[1] === 2021), false);
});

test("cities outside the fixed order append by first source appearance", () => {
  const result = alignLongTable(model([
    row("儋州市", 2021, 31, 10, 2),
    row("北京市", 2021, 11, 5, 3),
    row("三沙市", 2021, 21, 8, 4),
    row("厦门市", 2021, 15, 7, 5),
  ]), cityContext);

  assert.deepEqual(result.records.map((record) => record.city), [
    "北京市",
    "厦门市",
    "儋州市",
    "三沙市",
  ]);
  assert.deepEqual(result.outsideOrderCities, ["儋州市", "三沙市"]);
});

test("selected indicators emit a clean city-year-indicator table", () => {
  const result = alignLongTable(model([
    row("北京市", 2021, 11, 5, 2),
    row("厦门市", 2021, 15, 7, 3),
  ]), cityContext, {
    outputMode: "selected_indicators",
    selectedIndicators: ["可支配收入"],
  });

  assert.deepEqual(result.headers, ["城市", "年份", "可支配收入"]);
  assert.deepEqual(result.rows, [
    ["北京市", 2021, 11],
    ["厦门市", 2021, 15],
  ]);
  assert.deepEqual(
    result.records[0].cellMetadata.map((item) => item.address),
    ["A2", "B2", "C2"],
  );
});

test("mode recommendation compares time, quota and cell workload", () => {
  const recommendation = recommendProcessingModes({
    rowCount: 50_000,
    columnCount: 200,
    selectedIndicatorCount: 1,
  });

  assert.equal(recommendation.fullRow.outputCellCount, 10_000_000);
  assert.equal(recommendation.selectedIndicators.outputCellCount, 150_000);
  assert.equal(recommendation.selectedIndicators.workloadReductionPercent, 98.5);
  assert.match(recommendation.fullRow.timeEstimate, /分钟/);
  assert.match(recommendation.selectedIndicators.fiveHourQuotaEstimate, /%/);
  assert.equal(recommendation.recommendedMode, "selected_indicators");
});
