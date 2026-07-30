import assert from "node:assert/strict";
import test from "node:test";

import { alignLongTable } from "../scripts/lib/align-long.mjs";
import { alignWideTable } from "../scripts/lib/align-wide.mjs";
import { buildAudit } from "../scripts/lib/audit.mjs";
import { validateCityOrder } from "../scripts/lib/cities.mjs";
import { PauseError } from "../scripts/lib/pause.mjs";
import {
  verifyLongAlignment,
  verifyWideAlignment,
} from "../scripts/lib/verify.mjs";

const cityContext = validateCityOrder([
  { 序号: 1, 城市名: "北京市" },
  { 序号: 2, 城市名: "厦门市" },
]);

const longSource = {
  kind: "long",
  sheetName: "长表",
  headers: ["省份", "城市", "年份", "收入", "代码"],
  cityColumn: 1,
  yearColumn: 2,
  rows: [
    {
      values: ["福建", "厦门", 2020, 20, 3502],
      sourceRow: 2,
      year: 2020,
      cityMatch: { status: "matched", standardName: "厦门市", method: "unique_suffix" },
    },
    {
      values: ["北京", "北京市", 2020, 10, 11],
      sourceRow: 3,
      year: 2020,
      cityMatch: { status: "matched", standardName: "北京市", method: "exact" },
    },
  ],
};

const wideSource = {
  kind: "wide",
  sheetName: "宽表",
  headers: ["省份", "城市", 2020],
  cityColumn: 1,
  auxiliaryColumns: [0, 1],
  yearColumns: [{ year: 2020, column: 2 }],
  indicatorLabel: "指标值",
  rows: [
    {
      values: ["北京", "北京市", 10],
      sourceRow: 2,
      cityMatch: { status: "matched", standardName: "北京市", method: "exact" },
    },
    {
      values: ["福建", "厦门市", 20],
      sourceRow: 3,
      cityMatch: { status: "matched", standardName: "厦门市", method: "exact" },
    },
  ],
};

function assertPause(fn, code) {
  assert.throws(fn, (error) => error instanceof PauseError && error.code === code);
}

test("accepts only authorized long-table differences", () => {
  const aligned = alignLongTable(longSource, cityContext, {
    startYear: 2020,
    endYear: 2020,
  });
  const result = verifyLongAlignment(longSource, aligned, { cityContext });
  assert.equal(result.passed, true);
  assert.equal(result.checks.sourceRowsAccountedFor, true);
  assert.equal(result.checks.onlyAuthorizedChanges, true);
  assert.equal(result.checks.reverseKeysVerified, true);
});

test("matches formula events by absolute worksheet column outside A1", () => {
  const source = {
    ...structuredClone(longSource),
    dataStartColumn: 2,
    rows: [structuredClone(longSource.rows[0])],
  };
  source.rows[0].values[3] = 999;
  const aligned = alignLongTable(source, cityContext, {
    startYear: 2020,
    endYear: 2020,
  });
  aligned.records.find((record) => !record.missing).values[3] = 20;
  const result = verifyLongAlignment(source, aligned, {
    cityContext,
    formulaEvents: [{
      sheetName: "长表",
      sourceRow: 2,
      sourceColumn: 5,
      currentValue: 20,
    }],
  });
  assert.equal(result.passed, true);
});

test("detects a long value moved to the wrong city despite equal counts", () => {
  const aligned = alignLongTable(longSource, cityContext, {
    startYear: 2020,
    endYear: 2020,
  });
  const swapped = structuredClone(aligned);
  [swapped.records[0].values[3], swapped.records[1].values[3]] = [
    swapped.records[1].values[3],
    swapped.records[0].values[3],
  ];
  swapped.rows = swapped.records.map((record) => record.values);
  assertPause(
    () => verifyLongAlignment(longSource, swapped, { cityContext }),
    "REVERSE_VERIFICATION_FAILED",
  );
});

test("detects a wide value moved to the wrong city despite equal counts", () => {
  const aligned = alignWideTable(wideSource, cityContext, {
    startYear: 2020,
    endYear: 2020,
  });
  const swapped = structuredClone(aligned);
  [swapped.records[0].values[3], swapped.records[1].values[3]] = [
    swapped.records[1].values[3],
    swapped.records[0].values[3],
  ];
  swapped.rows = swapped.records.map((record) => record.values);
  assertPause(
    () => verifyWideAlignment(wideSource, swapped, { cityContext }),
    "REVERSE_VERIFICATION_FAILED",
  );
});

test("accepts a correct wide conversion", () => {
  const aligned = alignWideTable(wideSource, cityContext, {
    startYear: 2020,
    endYear: 2020,
  });
  const result = verifyWideAlignment(wideSource, aligned, { cityContext });
  assert.equal(result.passed, true);
  assert.equal(result.checks.sourceCellsAccountedFor, true);
  assert.equal(result.checks.reverseKeysVerified, true);
});

test("rejects non-empty data in an inserted missing row", () => {
  const source = { ...longSource, rows: [longSource.rows[0]] };
  const aligned = alignLongTable(source, cityContext, {
    startYear: 2020,
    endYear: 2020,
  });
  const contaminated = structuredClone(aligned);
  const missing = contaminated.records.find((record) => record.missing);
  missing.values[3] = 999;
  assertPause(
    () => verifyLongAlignment(source, contaminated, { cityContext }),
    "REVERSE_VERIFICATION_FAILED",
  );
});

test("rejects duplicate output keys", () => {
  const aligned = alignWideTable(wideSource, cityContext, {
    startYear: 2020,
    endYear: 2020,
  });
  const duplicate = structuredClone(aligned);
  duplicate.records.push(structuredClone(duplicate.records[0]));
  assertPause(
    () => verifyWideAlignment(wideSource, duplicate, { cityContext }),
    "REVERSE_VERIFICATION_FAILED",
  );
});

test("audit includes hashes, dimensions, matches, missing keys, formulas and verdict", () => {
  const aligned = alignLongTable(longSource, cityContext, {
    startYear: 2020,
    endYear: 2021,
  });
  const verification = verifyLongAlignment(longSource, aligned, { cityContext });
  const audit = buildAudit({
    fileMetadata: {
      inputPath: "C:/data/source.xlsx",
      inputSha256: "abc",
      cityOrderPath: "C:/data/城市顺序.xlsx",
      cityOrderSha256: "def",
    },
    config: { selectedSheets: ["长表"], startYear: 2020, endYear: 2021 },
    sourceModels: [longSource],
    outputModels: [aligned],
    verifications: [verification],
    mappingEvents: aligned.mappingEvents,
    formulaEvents: [{ sheetName: "长表", cell: "D2", formula: "=1+1", currentValue: 2 }],
  });
  assert.equal(audit.passed, true);
  assert.equal(audit.verdict, "通过");
  for (const section of ["文件", "范围", "结构", "城市匹配", "缺失键", "公式处理", "核验", "结论"]) {
    assert.ok(audit.auditRows.some((row) => row.section === section), section);
  }
});