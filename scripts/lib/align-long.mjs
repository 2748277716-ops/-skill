import { PauseError } from "./pause.mjs";

function pause(code, message, evidence = {}) {
  throw new PauseError(code, message, evidence);
}

function sourceYears(model) {
  const years = [...new Set(model.rows.map(
    (row) => Number(row.year ?? row.values?.[model.yearColumn]),
  ))];
  if (!years.length || years.some((year) => !Number.isInteger(year))) {
    pause("AMBIGUOUS_YEAR_COLUMN", "无法从源表可靠提取全部年份", {
      sheetName: model.sheetName,
      years,
    });
  }
  return years.sort((left, right) => right - left);
}

function targetYears(config, model) {
  if (Array.isArray(config?.targetYears) && config.targetYears.length) {
    const years = [...new Set(config.targetYears.map(Number))];
    if (years.some((year) => !Number.isInteger(year))) {
      pause("UNSAFE_STRUCTURE", "目标年份集合包含非整数", {
        targetYears: config.targetYears,
      });
    }
    years.sort((left, right) => right - left);
    return { years, yearSet: new Set(years) };
  }

  const hasStart = config?.startYear !== null && config?.startYear !== undefined;
  const hasEnd = config?.endYear !== null && config?.endYear !== undefined;
  if (!hasStart && !hasEnd) {
    const years = sourceYears(model);
    return { years, yearSet: new Set(years) };
  }
  const first = Number(config?.startYear);
  const second = Number(config?.endYear);
  if (!Number.isInteger(first) || !Number.isInteger(second)) {
    pause("UNSAFE_STRUCTURE", "年份范围必须同时提供两个整数；均未提供时自动使用源文件全部年份", {
      startYear: config?.startYear,
      endYear: config?.endYear,
    });
  }
  const minimum = Math.min(first, second);
  const maximum = Math.max(first, second);
  const years = [];
  for (let year = maximum; year >= minimum; year -= 1) years.push(year);
  return { years, yearSet: new Set(years) };
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

function normalizedHeader(value) {
  return String(value ?? "")
    .normalize("NFKC")
    .replace(/[\s\u00a0\u3000]+/gu, "")
    .toLowerCase();
}

function selectedSourceColumns(model, config) {
  if (config?.outputMode !== "selected_indicators") {
    return model.headers.map((_, column) => column);
  }
  const requested = Array.isArray(config.selectedIndicators)
    ? config.selectedIndicators.map((item) => String(item ?? "").trim()).filter(Boolean)
    : [];
  if (!requested.length) {
    pause("INDICATOR_SELECTION_REQUIRED", "精简指标模式必须明确指定至少一个指标", {
      headers: model.headers,
    });
  }
  const columns = [model.cityColumn, model.yearColumn];
  for (const indicator of requested) {
    const key = normalizedHeader(indicator);
    const matches = model.headers.flatMap((header, column) =>
      normalizedHeader(header) === key ? [{ header, column }] : [],
    );
    if (matches.length === 0) {
      pause("INDICATOR_NOT_FOUND", "指定指标在源表中不存在", {
        indicator,
        headers: model.headers,
      });
    }
    if (matches.length > 1) {
      pause("AMBIGUOUS_INDICATOR", "指定指标对应多个源列，不能自动选择", {
        indicator,
        matches,
      });
    }
    if (!columns.includes(matches[0].column)) columns.push(matches[0].column);
  }
  return columns;
}

export function alignLongTable(model, cityContext, config = {}) {
  validateModel(model, cityContext);
  const yearContext = targetYears(config, model);
  const sourceByKey = new Map();
  const outOfRangeRows = [];
  const mappingEvents = [];
  const outsideOrderEntries = [];
  const outsideOrderNames = new Set();

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
    if (!standardName) {
      pause("UNMATCHED_CITY", "源行缺少可用城市名", {
        sheetName: model.sheetName,
        sourceRow: row.sourceRow,
        rawCity: row.values[model.cityColumn],
      });
    }
    if (!cityContext.rankByName.has(standardName) && !outsideOrderNames.has(standardName)) {
      if (config.allowOutsideOrderCities === false) {
        pause("UNMATCHED_CITY", "源城市不在城市顺序表中，且配置禁止自动追加", {
          sheetName: model.sheetName,
          sourceRow: row.sourceRow,
          rawCity: row.values[model.cityColumn],
        });
      }
      outsideOrderNames.add(standardName);
      outsideOrderEntries.push({
        sequence: cityContext.entries.length + outsideOrderEntries.length + 1,
        standardName,
        sourceRow: row.sourceRow,
        rank: cityContext.entries.length + outsideOrderEntries.length,
        outsideOrder: true,
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
    if (!yearContext.yearSet.has(year)) {
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
  const effectiveCityEntries = [...cityContext.entries, ...outsideOrderEntries];
  for (const city of effectiveCityEntries) {
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

  const sourceColumns = selectedSourceColumns(model, config);
  const projectedRecords = records.map((record) => ({
    ...record,
    values: sourceColumns.map((column) => record.values[column]),
    cellMetadata: Array.isArray(record.cellMetadata)
      ? sourceColumns.map((column) => cloneMetadata(record.cellMetadata[column]))
      : record.cellMetadata,
    sourceCells: Array.isArray(record.sourceCells)
      ? sourceColumns.map((column) => record.sourceCells[column])
      : record.sourceCells,
  }));

  return {
    kind: "long-aligned",
    sheetName: model.sheetName,
    headers: sourceColumns.map((column) => model.headers[column]),
    rows: projectedRecords.map((record) => record.values),
    records: projectedRecords,
    cityColumn: sourceColumns.indexOf(model.cityColumn),
    yearColumn: sourceColumns.indexOf(model.yearColumn),
    sourceColumns,
    outputMode: config.outputMode ?? "preserve_rows",
    targetYears: yearContext.years,
    missingKeys,
    outOfRangeRows,
    mappingEvents,
    outsideOrderCities: outsideOrderEntries.map((entry) => entry.standardName),
    cityEntries: effectiveCityEntries,
    sourceRecordCount: model.rows.length,
    inRangeSourceRecordCount: inRangeSourceCount,
  };
}
