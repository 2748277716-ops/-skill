import { PauseError } from "./pause.mjs";

function fail(message, evidence = {}) {
  throw new PauseError("REVERSE_VERIFICATION_FAILED", message, evidence);
}

function key(city, year) {
  return `${city}\u0000${year}`;
}

function valuesEqual(left, right) {
  if (Object.is(left, right)) return true;
  if (left instanceof Date && right instanceof Date) return left.getTime() === right.getTime();
  if (Array.isArray(left) && Array.isArray(right)) {
    return left.length === right.length && left.every(
      (value, index) => valuesEqual(value, right[index]),
    );
  }
  return false;
}

export function verifyLongAlignmentFast(sourceModel, outputModel, cityContext) {
  if (sourceModel?.kind !== "long" || outputModel?.kind !== "long-aligned") {
    fail("长表核验模型类型不匹配", {
      sourceKind: sourceModel?.kind,
      outputKind: outputModel?.kind,
    });
  }
  const years = new Set(outputModel.targetYears);
  const outputByKey = new Map();
  for (const record of outputModel.records) {
    const recordKey = key(record.city, record.year);
    if (outputByKey.has(recordKey)) {
      fail("输出存在重复城市＋年份键", {
        city: record.city,
        year: record.year,
      });
    }
    outputByKey.set(recordKey, record);
  }

  const expectedGrid = new Set();
  for (const city of cityContext.entries) {
    for (const year of outputModel.targetYears) {
      expectedGrid.add(key(city.standardName, year));
    }
  }
  if (expectedGrid.size !== outputByKey.size) {
    fail("输出城市×年份网格规模不一致", {
      expected: expectedGrid.size,
      actual: outputByKey.size,
    });
  }
  for (const expectedKey of expectedGrid) {
    if (!outputByKey.has(expectedKey)) {
      fail("输出城市×年份网格缺少键", { key: expectedKey });
    }
  }

  const sourceKeys = new Set();
  for (const sourceRow of sourceModel.rows) {
    const city = sourceRow.cityMatch?.standardName;
    const year = Number(sourceRow.year ?? sourceRow.values[sourceModel.yearColumn]);
    if (!years.has(year)) continue;
    const sourceKey = key(city, year);
    if (sourceKeys.has(sourceKey)) {
      fail("源长表存在重复城市＋年份键", {
        city,
        year,
        sourceRow: sourceRow.sourceRow,
      });
    }
    sourceKeys.add(sourceKey);
    const output = outputByKey.get(sourceKey);
    if (!output || output.missing) {
      fail("源记录在输出中缺失", {
        city,
        year,
        sourceRow: sourceRow.sourceRow,
      });
    }
    const expectedValues = [...sourceRow.values];
    expectedValues[sourceModel.cityColumn] = city;
    expectedValues[sourceModel.yearColumn] = year;
    if (!valuesEqual(expectedValues, output.values) || output.sourceRow !== sourceRow.sourceRow) {
      fail("长表逐键对应关系不一致", {
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
    if (!record.missing && !sourceKeys.has(recordKey)) {
      fail("输出存在无法追溯的非缺失记录", {
        city: record.city,
        year: record.year,
      });
    }
    if (record.missing) {
      const extras = record.values.flatMap((value, column) =>
        column !== outputModel.cityColumn &&
        column !== outputModel.yearColumn &&
        value !== null &&
        value !== undefined &&
        value !== ""
          ? [{ column, value }]
          : [],
      );
      if (extras.length) {
        fail("缺失键行包含未经授权的数据", {
          city: record.city,
          year: record.year,
          extras,
        });
      }
    }
  }

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
      linearKeyIndexUsed: true,
    },
    sourceKeyCount: sourceKeys.size,
    outputKeyCount: outputByKey.size,
  };
}
