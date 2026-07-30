import assert from "node:assert/strict";
import test from "node:test";

import { PauseError } from "../scripts/lib/pause.mjs";
import {
  normalizeSafeCityName,
  resolveCityName,
  validateCityOrder,
  validateMappingRows,
} from "../scripts/lib/cities.mjs";

const order = validateCityOrder([
  { 序号: 1, 城市名: "北京市" },
  { 序号: 2, 城市名: "厦门市" },
]);

function assertPause(fn, code) {
  assert.throws(fn, (error) => error instanceof PauseError && error.code === code);
}

test("safe normalization removes Chinese-city whitespace and fullwidth spacing", () => {
  assert.equal(normalizeSafeCityName(" 厦　 门市 "), "厦门市");
});

test("exact match has the highest precedence", () => {
  assert.deepEqual(resolveCityName("厦门市", order, []), {
    status: "matched",
    standardName: "厦门市",
    method: "exact",
  });
});

test("unique suffix difference becomes an automatic standard match", () => {
  assert.deepEqual(resolveCityName("厦门", order, []), {
    status: "matched",
    standardName: "厦门市",
    method: "unique_suffix",
  });
});

test("confirmed mapping is reused after safe matching rules", () => {
  const mappings = validateMappingRows(
    [{ 原始城市名: "厦冂市", 标准城市名: "厦门市" }],
    order,
  );
  assert.deepEqual(resolveCityName("厦冂市", order, mappings), {
    status: "matched",
    standardName: "厦门市",
    method: "confirmed_mapping",
  });
});

test("typo is only a candidate", () => {
  const result = resolveCityName("厦冂市", order, []);
  assert.equal(result.status, "needs_confirmation");
  assert.equal(result.candidates[0].standardName, "厦门市");
  assert.equal(result.candidates[0].distance, 1);
});

test("non-contiguous city sequence pauses", () => {
  assertPause(
    () =>
      validateCityOrder([
        { 序号: 1, 城市名: "北京市" },
        { 序号: 3, 城市名: "厦门市" },
      ]),
    "INVALID_CITY_ORDER",
  );
});

test("duplicate standard city pauses", () => {
  assertPause(
    () =>
      validateCityOrder([
        { 序号: 1, 城市名: "厦门市" },
        { 序号: 2, 城市名: "厦门市" },
      ]),
    "INVALID_CITY_ORDER",
  );
});

test("missing required city-order headers pauses", () => {
  assertPause(() => validateCityOrder([{ 排名: 1, 名称: "厦门市" }]), "INVALID_CITY_ORDER");
});

test("mapping conflict pauses", () => {
  assertPause(
    () =>
      validateMappingRows(
        [
          { 原始城市名: "厦冂市", 标准城市名: "厦门市" },
          { 原始城市名: "厦冂市", 标准城市名: "北京市" },
        ],
        order,
      ),
    "MAPPING_CONFLICT",
  );
});

test("mapping to a city absent from the order pauses", () => {
  assertPause(
    () =>
      validateMappingRows(
        [{ 原始城市名: "福州市", 标准城市名: "福州市" }],
        order,
      ),
    "MAPPING_CONFLICT",
  );
});

test("safe normalization collision is never automatic", () => {
  const collisionOrder = validateCityOrder([
    { 序号: 1, 城市名: "厦门市" },
    { 序号: 2, 城市名: "厦 门市" },
  ]);
  const result = resolveCityName("厦　门市", collisionOrder, []);
  assert.equal(result.status, "needs_confirmation");
  assert.deepEqual(
    result.candidates.map((candidate) => candidate.standardName),
    ["厦门市", "厦 门市"],
  );
});

test("PauseError serializes the stable paused contract", () => {
  const error = new PauseError(
    "DUPLICATE_CITY_YEAR",
    "存在重复的城市＋年份记录",
    { sheetName: "Sheet1", keys: ["厦门市|2021"], sourceRows: [2, 3] },
  );
  assert.deepEqual(error.toJSON(), {
    status: "paused",
    code: "DUPLICATE_CITY_YEAR",
    message: "存在重复的城市＋年份记录",
    evidence: {
      sheetName: "Sheet1",
      keys: ["厦门市|2021"],
      sourceRows: [2, 3],
    },
  });
});