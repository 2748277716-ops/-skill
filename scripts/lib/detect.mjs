import { resolveCityName } from "./cities.mjs";
import { PauseError } from "./pause.mjs";

function pause(code, message, evidence = {}) {
  throw new PauseError(code, message, evidence);
}

function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function parseYear(value) {
  if (typeof value === "number" && Number.isInteger(value)) {
    return value >= 1000 && value <= 9999 ? value : null;
  }
  const text = String(value ?? "").trim();
  return /^\d{4}$/u.test(text) ? Number(text) : null;
}

function targetYearContext(config) {
  const first = Number(config?.startYear);
  const second = Number(config?.endYear);
  if (!Number.isInteger(first) || !Number.isInteger(second)) {
    pause("UNRECOGNIZED_TABLE", "目标年份范围必须显式提供为整数", {
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

function validateSheetSelection(config) {
  const workbookSheets = config?.workbookSheets ?? [config?.sheetName].filter(Boolean);
  const selectedSheets = config?.selectedSheets ?? [];
  if (
    workbookSheets.length > 1 &&
    (selectedSheets.length === 0 || !selectedSheets.includes(config?.sheetName))
  ) {
    pause("MULTIPLE_SHEETS", "工作簿包含多个工作表，需要先明确选择", {
      workbookSheets,
      selectedSheets,
    });
  }
}

function cityHeaderScore(value) {
  const text = String(value ?? "").normalize("NFKC").trim().toLowerCase();
  if (["城市", "城市名", "城市名称", "地级市", "city", "cityname", "city name"].includes(text)) {
    return 3;
  }
  if ((text.includes("城市") || text.includes("city")) && !text.includes("代码")) {
    return 2;
  }
  return 0;
}

function yearHeaderScore(value) {
  const text = String(value ?? "").normalize("NFKC").trim().toLowerCase();
  if (["年份", "年度", "统计年份", "year"].includes(text)) return 3;
  if (text.includes("年份") || text.endsWith("年度")) return 2;
  return 0;
}

function nonBlankRows(matrix) {
  return matrix.slice(1).flatMap((values, index) =>
    Array.isArray(values) && values.some((value) => !isBlank(value))
      ? [{ values: [...values], sourceRow: index + 2 }]
      : [],
  );
}

function chooseCityColumn(headers, rows, cityContext, mappingRows) {
  const candidates = headers.flatMap((header, column) => {
    const samples = rows.map((row) => row.values[column]).filter((value) => !isBlank(value));
    const matched = samples.filter(
      (value) => resolveCityName(value, cityContext, mappingRows).status === "matched",
    ).length;
    const matchRate = samples.length ? matched / samples.length : 0;
    const headerScore = cityHeaderScore(header);
    return headerScore > 0 || matchRate >= 0.6
      ? [{ column, header, headerScore, matchRate, score: headerScore + matchRate * 2 }]
      : [];
  });

  if (candidates.length === 0) {
    pause("UNRECOGNIZED_TABLE", "无法可靠识别城市列", {
      headers,
    });
  }
  if (candidates.length !== 1) {
    pause("AMBIGUOUS_CITY_COLUMN", "存在多个可能的城市列", {
      candidates,
    });
  }
  return candidates[0].column;
}

function chooseLongYearColumns(headers, rows, cityColumn) {
  return headers.flatMap((header, column) => {
    if (column === cityColumn) return [];
    const samples = rows.map((row) => row.values[column]).filter((value) => !isBlank(value));
    const parsed = samples.filter((value) => parseYear(value) !== null).length;
    const parseRate = samples.length ? parsed / samples.length : 0;
    const headerScore = yearHeaderScore(header);
    return headerScore > 0 && parseRate >= 0.6
      ? [{ column, header, headerScore, parseRate }]
      : [];
  });
}

function resolveRows(rows, cityColumn, cityContext, mappingRows) {
  return rows.map((row) => {
    const rawCity = row.values[cityColumn];
    const cityMatch = resolveCityName(rawCity, cityContext, mappingRows);
    if (cityMatch.status !== "matched") {
      pause("UNMATCHED_CITY", "城市名称无法可靠匹配", {
        sourceRow: row.sourceRow,
        rawCity,
        normalizedCity: String(rawCity ?? "").normalize("NFKC"),
        candidates: cityMatch.candidates,
        reason: cityMatch.reason,
      });
    }
    return { ...row, cityMatch };
  });
}

function findOutOfRange(years, yearContext) {
  return [...new Set(years.filter(
    (year) => year < yearContext.minimum || year > yearContext.maximum,
  ))].sort((left, right) => left - right);
}

export function detectTableModel(sheetMatrix, config = {}, cityContext) {
  validateSheetSelection(config);
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
  if (headers.length === 0 || rows.length === 0) {
    pause("UNRECOGNIZED_TABLE", "工作表缺少表头或数据行", {
      sheetName: config.sheetName,
      headerCount: headers.length,
      rowCount: rows.length,
    });
  }

  const yearContext = targetYearContext(config);
  const mappingRows = config.mappingRows ?? [];
  const cityColumn = chooseCityColumn(headers, rows, cityContext, mappingRows);
  const resolvedRows = resolveRows(rows, cityColumn, cityContext, mappingRows);
  const allYearHeaders = headers.flatMap((header, column) => {
    const year = parseYear(header);
    return year === null ? [] : [{ year, column }];
  });
  const outOfRangeHeaderYears = findOutOfRange(
    allYearHeaders.map((item) => item.year),
    yearContext,
  );
  if (outOfRangeHeaderYears.length && !config.allowOutOfRangeYears) {
    pause("OUT_OF_RANGE_YEARS", "宽表包含目标范围外年份", {
      sheetName: config.sheetName,
      years: outOfRangeHeaderYears,
      columns: allYearHeaders.filter((item) => outOfRangeHeaderYears.includes(item.year)),
    });
  }

  const yearColumns = allYearHeaders.filter(
    (item) => item.year >= yearContext.minimum && item.year <= yearContext.maximum,
  );
  const longYearCandidates = chooseLongYearColumns(headers, rows, cityColumn);

  if (yearColumns.length > 0 && longYearCandidates.length > 0) {
    pause("AMBIGUOUS_YEAR_COLUMN", "同时检测到长表年份列和宽表年份表头", {
      longYearCandidates,
      yearColumns,
    });
  }
  if (longYearCandidates.length > 1) {
    pause("AMBIGUOUS_YEAR_COLUMN", "存在多个可能的长表年份列", {
      candidates: longYearCandidates,
    });
  }

  if (longYearCandidates.length === 1) {
    const yearColumn = longYearCandidates[0].column;
    const invalidRows = [];
    const rowYears = resolvedRows.map((row) => {
      const year = parseYear(row.values[yearColumn]);
      if (year === null) {
        invalidRows.push({ sourceRow: row.sourceRow, value: row.values[yearColumn] });
      }
      return year;
    });
    if (invalidRows.length) {
      pause("AMBIGUOUS_YEAR_COLUMN", "年份列包含无法解析的值", {
        column: yearColumn,
        invalidRows,
      });
    }
    const outOfRangeYears = findOutOfRange(rowYears, yearContext);
    if (outOfRangeYears.length && !config.allowOutOfRangeYears) {
      pause("OUT_OF_RANGE_YEARS", "长表包含目标范围外年份", {
        sheetName: config.sheetName,
        years: outOfRangeYears,
        sourceRows: resolvedRows
          .filter((row, index) => outOfRangeYears.includes(rowYears[index]))
          .map((row) => row.sourceRow),
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

  if (yearColumns.length > 0) {
    const seen = new Map();
    for (const row of resolvedRows) {
      const standardName = row.cityMatch.standardName;
      const previous = seen.get(standardName);
      if (previous) {
        pause("DUPLICATE_CITY_YEAR", "宽表中同一标准城市出现多行", {
          sheetName: config.sheetName,
          city: standardName,
          sourceRows: [previous.sourceRow, row.sourceRow],
          years: yearColumns.map((item) => item.year),
        });
      }
      seen.set(standardName, row);
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
        .filter((column) => !allYearHeaders.some((item) => item.column === column)),
      targetYears: yearContext.targetYears,
      outOfRangeYears: outOfRangeHeaderYears,
      outOfRangeColumns: allYearHeaders.filter((item) => outOfRangeHeaderYears.includes(item.year)),
      indicatorLabel: config.indicatorLabel ?? "指标值",
      metadata: config.metadata ?? {},
    };
  }

  pause("UNRECOGNIZED_TABLE", "无法唯一识别长表或宽表结构", {
    sheetName: config.sheetName,
    headers,
    cityColumn,
  });
}