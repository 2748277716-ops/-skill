import { PauseError } from "./pause.mjs";

function pause(code, message, evidence = {}) {
  throw new PauseError(code, message, evidence);
}

function targetYears(config) {
  const first = Number(config?.startYear);
  const second = Number(config?.endYear);
  if (!Number.isInteger(first) || !Number.isInteger(second)) {
    pause("UNSAFE_STRUCTURE", "目标年份范围无效", {
      startYear: config?.startYear,
      endYear: config?.endYear,
    });
  }
  const minimum = Math.min(first, second);
  const maximum = Math.max(first, second);
  const years = [];
  for (let year = maximum; year >= minimum; year -= 1) years.push(year);
  return { minimum, maximum, years };
}

function cityYearKey(city, year) {
  return `${city}\u0000${year}`;
}

function cloneMetadata(metadata) {
  if (Array.isArray(metadata)) return [...metadata];
  if (metadata && typeof metadata === "object") return { ...metadata };
  return metadata ?? null;
}

function validateModel(model, cityContext) {
  if (model?.kind !== "long") {
    pause("UNSAFE_STRUCTURE", "长表对齐器收到非长表模型", {
      kind: model?.kind,
    });
  }
  if (!Array.isArray(model.headers) || !Array.isArray(model.rows)) {
    pause("UNSAFE_STRUCTURE", "长表模型缺少表头或数据行", {});
  }
  if (
    !Number.isInteger(model.cityColumn) ||
    !Number.isInteger(model.yearColumn) ||
    model.cityColumn < 0 ||
    model.yearColumn < 0 ||
    model.cityColumn >= model.headers.length ||
    model.yearColumn >= model.headers.length
  ) {
    pause("UNSAFE_STRUCTURE", "长表模型的城市列或年份列无效", {
      cityColumn: model.cityColumn,
      yearColumn: model.yearColumn,
      columnCount: model.headers.length,
    });
  }
  if (!cityContext?.entries?.length) {
    pause("INVALID_CITY_ORDER", "城市顺序上下文为空", {});
  }
}

export function alignLongTable(model, cityContext, config = {}) {
  validateModel(model, cityContext);
  const yearContext = targetYears(config);
  const sourceByKey = new Map();
  const outOfRangeRows = [];
  const mappingEvents = [];

  for (const row of model.rows) {
    if (!Array.isArray(row.values) || row.values.length !== model.headers.length) {
      pause("UNSAFE_STRUCTURE", "源行列数与表头不一致", {
        sheetName: model.sheetName,
        sourceRow: row.sourceRow,
        expectedColumns: model.headers.length,
        actualColumns: row.values?.length,
      });
    }
    const standardName = row.cityMatch?.standardName;
    if (!standardName || !cityContext.rankByName.has(standardName)) {
      pause("UNMATCHED_CITY", "源行缺少已确认的标准城市名", {
        sheetName: model.sheetName,
        sourceRow: row.sourceRow,
        rawCity: row.values[model.cityColumn],
      });
    }
    const year = Number(row.year ?? row.values[model.yearColumn]);
    if (!Number.isInteger(year)) {
      pause("AMBIGUOUS_YEAR_COLUMN", "源行年份不是整数", {
        sheetName: model.sheetName,
        sourceRow: row.sourceRow,
        value: row.year ?? row.values[model.yearColumn],
      });
    }
    const key = cityYearKey(standardName, year);
    const previous = sourceByKey.get(key);
    if (previous) {
      pause("DUPLICATE_CITY_YEAR", "存在重复的城市＋年份记录", {
        sheetName: model.sheetName,
        keys: [{ city: standardName, year }],
        sourceRows: [previous.sourceRow, row.sourceRow],
      });
    }

    const sourceRecord = { ...row, standardName, year };
    sourceByKey.set(key, sourceRecord);
    if (year < yearContext.minimum || year > yearContext.maximum) {
      outOfRangeRows.push({
        ...sourceRecord,
        values: [...row.values],
        cellMetadata: cloneMetadata(row.cellMetadata),
      });
    }
    if (row.cityMatch.method !== "exact") {
      mappingEvents.push({
        sourceRow: row.sourceRow,
        rawCity: row.values[model.cityColumn],
        standardName,
        method: row.cityMatch.method,
      });
    }
  }

  if (outOfRangeRows.length && !config.allowOutOfRangeYears) {
    pause("OUT_OF_RANGE_YEARS", "长表包含目标范围外年份", {
      sheetName: model.sheetName,
      years: [...new Set(outOfRangeRows.map((row) => row.year))].sort(),
      sourceRows: outOfRangeRows.map((row) => row.sourceRow),
    });
  }

  const records = [];
  const missingKeys = [];
  for (const city of cityContext.entries) {
    for (const year of yearContext.years) {
      const existing = sourceByKey.get(cityYearKey(city.standardName, year));
      if (existing) {
        const values = [...existing.values];
        values[model.cityColumn] = city.standardName;
        values[model.yearColumn] = year;
        records.push({
          values,
          city: city.standardName,
          year,
          missing: false,
          sourceSheet: model.sheetName,
          sourceRow: existing.sourceRow,
          sourceCells: existing.sourceCells ?? null,
          cellMetadata: cloneMetadata(existing.cellMetadata),
          rowMetadata: cloneMetadata(existing.rowMetadata),
          cityMatch: { ...existing.cityMatch },
        });
      } else {
        const values = Array(model.headers.length).fill(null);
        values[model.cityColumn] = city.standardName;
        values[model.yearColumn] = year;
        records.push({
          values,
          city: city.standardName,
          year,
          missing: true,
          sourceSheet: null,
          sourceRow: null,
          sourceCells: null,
          cellMetadata: null,
          rowMetadata: null,
          cityMatch: null,
        });
        missingKeys.push({ city: city.standardName, year });
      }
    }
  }

  const inRangeSourceCount = model.rows.length - outOfRangeRows.length;
  const emittedSourceCount = records.filter((record) => !record.missing).length;
  if (emittedSourceCount !== inRangeSourceCount) {
    pause("REVERSE_VERIFICATION_FAILED", "长表对齐遗漏或重复了源记录", {
      sheetName: model.sheetName,
      inRangeSourceCount,
      emittedSourceCount,
    });
  }

  return {
    kind: "long-aligned",
    sheetName: model.sheetName,
    headers: [...model.headers],
    rows: records.map((record) => record.values),
    records,
    cityColumn: model.cityColumn,
    yearColumn: model.yearColumn,
    targetYears: yearContext.years,
    missingKeys,
    outOfRangeRows,
    mappingEvents,
    sourceRecordCount: model.rows.length,
    inRangeSourceRecordCount: inRangeSourceCount,
  };
}