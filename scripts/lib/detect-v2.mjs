import { resolveCityName } from "./cities.mjs";
import { detectTableModel as detectTableModelLegacy } from "./detect.mjs";
import { PauseError } from "./pause.mjs";

function pause(code, message, evidence = {}) {
  throw new PauseError(code, message, evidence);
}

function normalizedHeader(value) {
  return String(value ?? "").normalize("NFKC").replace(/[\s\u00a0\u3000]+/gu, "").toLowerCase();
}

function parseYear(value) {
  if (typeof value === "number" && Number.isInteger(value)) {
    return value >= 1000 && value <= 9999 ? value : null;
  }
  const text = String(value ?? "").trim();
  return /^\d{4}$/u.test(text) ? Number(text) : null;
}

function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function nonBlankRows(matrix) {
  return matrix.slice(1).flatMap((values, index) =>
    Array.isArray(values) && values.some((value) => !isBlank(value))
      ? [{ values: [...values], sourceRow: index + 2 }]
      : [],
  );
}

function targetYearContext(config) {
  const first = Number(config?.startYear);
  const second = Number(config?.endYear);
  if (!Number.isInteger(first) || !Number.isInteger(second)) {
    pause("INVALID_CONFIG", "快速检测需要可靠的年份边界", {
      startYear: config?.startYear,
      endYear: config?.endYear,
    });
  }
  const minimum = Math.min(first, second);
  const maximum = Math.max(first, second);
  const targetYears = [];
  for (let year = maximum; year >= minimum; year -= 1) targetYears.push(year);
  return { minimum, maximum, targetYears };
}

function resolveRows(rows, cityColumn, cityContext, mappingRows) {
  return rows.map((row) => {
    const rawCity = row.values[cityColumn];
    const cityMatch = resolveCityName(rawCity, cityContext, mappingRows);
    if (cityMatch.status !== "matched") {
      pause("UNMATCHED_CITY", "城市名称无法可靠匹配", {
        sourceRow: row.sourceRow,
        rawCity,
        reason: cityMatch.reason,
        candidates: cityMatch.candidates,
      });
    }
    return { ...row, cityMatch };
  });
}

export function detectTableModel(sheetMatrix, config = {}, cityContext) {
  if ((config.mergedRanges ?? []).length > 0) {
    pause("UNSAFE_MERGED_CELLS", "数据区域存在无法安全展开的合并单元格", {
      sheetName: config.sheetName,
      mergedRanges: config.mergedRanges,
    });
  }
  if (!Array.isArray(sheetMatrix) || !Array.isArray(sheetMatrix[0])) {
    pause("UNRECOGNIZED_TABLE", "工作表没有可识别的二维数据区域", {
      sheetName: config.sheetName,
    });
  }
  const headers = [...sheetMatrix[0]];
  const rows = nonBlankRows(sheetMatrix);
  const cityColumns = headers.flatMap((header, column) =>
    ["城市", "城市名", "城市名称", "地级市", "city", "cityname"].includes(normalizedHeader(header))
      ? [column]
      : [],
  );
  if (cityColumns.length !== 1) {
    return detectTableModelLegacy(sheetMatrix, config, cityContext);
  }
  const cityColumn = cityColumns[0];
  const longYearColumns = headers.flatMap((header, column) =>
    column !== cityColumn &&
    ["年份", "年度", "统计年份", "year"].includes(normalizedHeader(header))
      ? [column]
      : [],
  );
  const wideYearColumns = headers.flatMap((header, column) => {
    const year = parseYear(header);
    return year === null ? [] : [{ year, column }];
  });
  if (longYearColumns.length > 0 && wideYearColumns.length > 0) {
    pause("AMBIGUOUS_YEAR_COLUMN", "同时检测到长表年份列和宽表年份表头", {
      longYearColumns,
      wideYearColumns,
    });
  }
  if (longYearColumns.length > 1) {
    pause("AMBIGUOUS_YEAR_COLUMN", "存在多个可能的长表年份列", {
      longYearColumns,
    });
  }
  if (longYearColumns.length === 0 && wideYearColumns.length === 0) {
    return detectTableModelLegacy(sheetMatrix, config, cityContext);
  }

  const yearContext = targetYearContext(config);
  const mappingRows = config.mappingRows ?? [];
  const resolvedRows = resolveRows(rows, cityColumn, cityContext, mappingRows);
  if (longYearColumns.length === 1) {
    const yearColumn = longYearColumns[0];
    const invalidRows = [];
    const rowYears = resolvedRows.map((row) => {
      const year = parseYear(row.values[yearColumn]);
      if (year === null) invalidRows.push({ sourceRow: row.sourceRow, value: row.values[yearColumn] });
      return year;
    });
    if (invalidRows.length) {
      pause("AMBIGUOUS_YEAR_COLUMN", "年份列包含无法解析的值", {
        column: yearColumn,
        invalidRows,
      });
    }
    const outOfRangeYears = [...new Set(rowYears.filter(
      (year) => year < yearContext.minimum || year > yearContext.maximum,
    ))].sort((left, right) => left - right);
    if (outOfRangeYears.length && !config.allowOutOfRangeYears) {
      pause("OUT_OF_RANGE_YEARS", "长表包含目标范围外年份", {
        sheetName: config.sheetName,
        years: outOfRangeYears,
      });
    }
    return {
      kind: "long",
      sheetName: config.sheetName,
      headers,
      rows: resolvedRows.map((row, index) => ({ ...row, year: rowYears[index] })),
      cityColumn,
      yearColumn,
      targetYears: yearContext.targetYears,
      outOfRangeYears,
      metadata: config.metadata ?? {},
    };
  }

  const yearColumns = wideYearColumns.filter(
    (item) => item.year >= yearContext.minimum && item.year <= yearContext.maximum,
  );
  const outOfRangeYears = [...new Set(wideYearColumns
    .filter((item) => item.year < yearContext.minimum || item.year > yearContext.maximum)
    .map((item) => item.year))].sort((left, right) => left - right);
  if (outOfRangeYears.length && !config.allowOutOfRangeYears) {
    pause("OUT_OF_RANGE_YEARS", "宽表包含目标范围外年份", {
      sheetName: config.sheetName,
      years: outOfRangeYears,
    });
  }
  const seen = new Map();
  for (const row of resolvedRows) {
    const city = row.cityMatch.standardName;
    if (seen.has(city)) {
      pause("DUPLICATE_CITY_YEAR", "宽表中同一标准城市出现多行", {
        city,
        sourceRows: [seen.get(city), row.sourceRow],
      });
    }
    seen.set(city, row.sourceRow);
  }
  return {
    kind: "wide",
    sheetName: config.sheetName,
    headers,
    rows: resolvedRows,
    cityColumn,
    yearColumns,
    auxiliaryColumns: headers
      .map((_, column) => column)
      .filter((column) => !wideYearColumns.some((item) => item.column === column)),
    targetYears: yearContext.targetYears,
    outOfRangeYears,
    outOfRangeColumns: wideYearColumns.filter((item) => outOfRangeYears.includes(item.year)),
    indicatorLabel: config.indicatorLabel ?? "指标值",
    metadata: config.metadata ?? {},
  };
}
