import { PauseError } from "./pause.mjs";

function fail(message, evidence) {
  throw new PauseError("REVERSE_VERIFICATION_FAILED", message, evidence);
}

function key(city, year) {
  return `${city}\u0000${year}`;
}

function valuesEqual(left, right) {
  if (Object.is(left, right)) return true;
  if (left instanceof Date && right instanceof Date) {
    return left.getTime() === right.getTime();
  }
  if (Array.isArray(left) && Array.isArray(right)) {
    return left.length === right.length && left.every((value, index) => valuesEqual(value, right[index]));
  }
  return false;
}

function formulaValue(formulaEvents, sourceModel, row, column, fallback) {
  const event = formulaEvents.find(
    (item) =>
      (item.sheetName ?? item.sourceSheet) === sourceModel.sheetName &&
      item.sourceRow === row.sourceRow &&
      item.sourceColumn === column,
  );
  return event ? event.currentValue : fallback;
}

function targetYearSet(outputModel) {
  return new Set(outputModel.targetYears ?? outputModel.records.map((record) => record.year));
}

function indexOutputRecords(outputModel) {
  const byKey = new Map();
  const duplicates = [];
  for (const record of outputModel.records ?? []) {
    const recordKey = key(record.city, record.year);
    if (byKey.has(recordKey)) duplicates.push({ city: record.city, year: record.year });
    byKey.set(recordKey, record);
  }
  if (duplicates.length) {
    fail("输出存在重复城市＋年份键", { duplicates });
  }
  return byKey;
}

function verifyCompleteGrid(outputModel, byKey, cityContext) {
  if (!cityContext?.entries) return true;
  const expected = [];
  for (const city of cityContext.entries) {
    for (const year of outputModel.targetYears) expected.push(key(city.standardName, year));
  }
  const missing = expected.filter((item) => !byKey.has(item));
  const extra = [...byKey.keys()].filter((item) => !expected.includes(item));
  if (missing.length || extra.length) {
    fail("输出城市×年份网格不完整", {
      missingKeys: missing.slice(0, 50),
      extraKeys: extra.slice(0, 50),
    });
  }
  return true;
}

function verifyMissingRows(outputModel, expectedKeys) {
  const contaminated = [];
  for (const record of outputModel.records) {
    const recordKey = key(record.city, record.year);
    if (!record.missing) continue;
    if (expectedKeys.has(recordKey)) {
      contaminated.push({ city: record.city, year: record.year, reason: "source key marked missing" });
      continue;
    }
    const extras = record.values.flatMap((value, column) =>
      column !== outputModel.cityColumn &&
      column !== outputModel.yearColumn &&
      value !== null &&
      value !== undefined &&
      value !== ""
        ? [{ column, value }]
        : [],
    );
    if (extras.length) contaminated.push({ city: record.city, year: record.year, extras });
  }
  if (contaminated.length) {
    fail("缺失键行包含未经授权的数据", { contaminated: contaminated.slice(0, 50) });
  }
  return true;
}

export function verifyLongAlignment(sourceModel, outputModel, config = {}) {
  if (sourceModel?.kind !== "long" || outputModel?.kind !== "long-aligned") {
    fail("长表核验模型类型不匹配", {
      sourceKind: sourceModel?.kind,
      outputKind: outputModel?.kind,
    });
  }
  const formulaEvents = config.formulaEvents ?? [];
  const years = targetYearSet(outputModel);
  const outputByKey = indexOutputRecords(outputModel);
  const expectedKeys = new Set();
  const mismatches = [];

  for (const sourceRow of sourceModel.rows) {
    const city = sourceRow.cityMatch?.standardName;
    const year = Number(sourceRow.year ?? sourceRow.values[sourceModel.yearColumn]);
    if (!years.has(year)) continue;
    const sourceKey = key(city, year);
    if (expectedKeys.has(sourceKey)) {
      fail("源长表存在重复城市＋年份键", {
        city,
        year,
        sourceRow: sourceRow.sourceRow,
      });
    }
    expectedKeys.add(sourceKey);
    const output = outputByKey.get(sourceKey);
    if (!output || output.missing) {
      mismatches.push({ city, year, sourceRow: sourceRow.sourceRow, reason: "missing output record" });
      continue;
    }
    const expectedValues = sourceRow.values.map((value, column) =>
      formulaValue(formulaEvents, sourceModel, sourceRow, column, value),
    );
    expectedValues[sourceModel.cityColumn] = city;
    expectedValues[sourceModel.yearColumn] = year;
    if (!valuesEqual(expectedValues, output.values) || output.sourceRow !== sourceRow.sourceRow) {
      mismatches.push({
        city,
        year,
        sourceRow: sourceRow.sourceRow,
        outputSourceRow: output.sourceRow,
        expectedValues,
        actualValues: output.values,
      });
    }
  }

  for (const record of outputModel.records) {
    const recordKey = key(record.city, record.year);
    if (!record.missing && !expectedKeys.has(recordKey)) {
      mismatches.push({ city: record.city, year: record.year, reason: "extra non-missing output" });
    }
  }
  if (mismatches.length) {
    fail("长表逐键对应关系不一致", { mismatches: mismatches.slice(0, 50) });
  }

  verifyMissingRows(outputModel, expectedKeys);
  verifyCompleteGrid(outputModel, outputByKey, config.cityContext);
  return {
    passed: true,
    kind: "long",
    checks: {
      uniqueOutputKeys: true,
      completeGrid: true,
      sourceRowsAccountedFor: true,
      onlyAuthorizedChanges: true,
      missingRowsBlank: true,
      reverseKeysVerified: true,
    },
    sourceKeyCount: expectedKeys.size,
    outputKeyCount: outputByKey.size,
  };
}

function wideBlocks(sourceModel) {
  if (Array.isArray(sourceModel.indicatorBlocks)) {
    return sourceModel.indicatorBlocks.map((block) => ({
      label: block.label,
      years: new Map(block.yearColumns.map((item) => [Number(item.year), Number(item.column)])),
    }));
  }
  return [{
    label: sourceModel.indicatorLabel ?? "指标值",
    years: new Map((sourceModel.yearColumns ?? []).map((item) => [Number(item.year), Number(item.column)])),
  }];
}

export function verifyWideAlignment(sourceModel, outputModel, config = {}) {
  if (sourceModel?.kind !== "wide" || !String(outputModel?.kind).startsWith("wide-")) {
    fail("宽表核验模型类型不匹配", {
      sourceKind: sourceModel?.kind,
      outputKind: outputModel?.kind,
    });
  }
  const formulaEvents = config.formulaEvents ?? [];
  const outputByKey = indexOutputRecords(outputModel);
  const blocks = wideBlocks(sourceModel);
  const expectedKeys = new Set();
  const mismatches = [];
  let sourceCellCount = 0;

  for (const sourceRow of sourceModel.rows) {
    const city = sourceRow.cityMatch?.standardName;
    for (const year of outputModel.targetYears) {
      const columns = blocks.map((block) => block.years.get(year));
      if (!columns.every((column) => Number.isInteger(column))) continue;
      const sourceKey = key(city, year);
      if (expectedKeys.has(sourceKey)) {
        fail("源宽表存在重复城市＋年份键", {
          city,
          year,
          sourceRow: sourceRow.sourceRow,
        });
      }
      expectedKeys.add(sourceKey);
      sourceCellCount += columns.length;
      const output = outputByKey.get(sourceKey);
      if (!output || output.missing) {
        mismatches.push({ city, year, sourceRow: sourceRow.sourceRow, reason: "missing output record" });
        continue;
      }
      const expectedValues = [
        ...sourceModel.auxiliaryColumns.map((column) => sourceRow.values[column]),
        year,
        ...columns.map((column) =>
          formulaValue(formulaEvents, sourceModel, sourceRow, column, sourceRow.values[column] ?? null),
        ),
      ];
      expectedValues[outputModel.cityColumn] = city;
      if (!valuesEqual(expectedValues, output.values) || output.sourceRow !== sourceRow.sourceRow) {
        mismatches.push({
          city,
          year,
          sourceRow: sourceRow.sourceRow,
          outputSourceRow: output.sourceRow,
          expectedValues,
          actualValues: output.values,
        });
      }
    }
  }

  for (const record of outputModel.records) {
    const recordKey = key(record.city, record.year);
    if (!record.missing && !expectedKeys.has(recordKey)) {
      mismatches.push({ city: record.city, year: record.year, reason: "extra non-missing output" });
    }
  }
  if (mismatches.length) {
    fail("宽表逐键对应关系不一致", { mismatches: mismatches.slice(0, 50) });
  }

  verifyMissingRows(outputModel, expectedKeys);
  verifyCompleteGrid(outputModel, outputByKey, config.cityContext);
  return {
    passed: true,
    kind: "wide",
    checks: {
      uniqueOutputKeys: true,
      completeGrid: true,
      sourceCellsAccountedFor: true,
      onlyAuthorizedChanges: true,
      missingRowsBlank: true,
      reverseKeysVerified: true,
    },
    sourceKeyCount: expectedKeys.size,
    sourceCellCount,
    outputKeyCount: outputByKey.size,
  };
}