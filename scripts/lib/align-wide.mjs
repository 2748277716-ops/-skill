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

function cloneMetadata(value) {
  if (Array.isArray(value)) return [...value];
  if (value && typeof value === "object") return { ...value };
  return value ?? null;
}

function validateBaseModel(model, cityContext) {
  if (model?.kind !== "wide") {
    pause("UNSAFE_STRUCTURE", "宽表转换器收到非宽表模型", { kind: model?.kind });
  }
  if (!Array.isArray(model.headers) || !Array.isArray(model.rows)) {
    pause("UNSAFE_STRUCTURE", "宽表模型缺少表头或数据行", {});
  }
  if (!Array.isArray(model.auxiliaryColumns) || !model.auxiliaryColumns.includes(model.cityColumn)) {
    pause("UNSAFE_STRUCTURE", "宽表辅助列必须包含城市列", {
      cityColumn: model.cityColumn,
      auxiliaryColumns: model.auxiliaryColumns,
    });
  }
  if (!cityContext?.entries?.length) {
    pause("INVALID_CITY_ORDER", "城市顺序上下文为空", {});
  }
}

function normalizeBlock(block, model, target, multiMode) {
  const label = String(block?.label ?? "").trim();
  if (!label || !Array.isArray(block.yearColumns)) {
    pause("AMBIGUOUS_MULTI_INDICATOR", "指标标签或年份映射缺失", {
      label,
      yearColumns: block?.yearColumns,
    });
  }
  const years = new Map();
  const usedColumns = new Set();
  for (const item of block.yearColumns) {
    const year = Number(item?.year);
    const column = Number(item?.column);
    if (
      !Number.isInteger(year) ||
      !Number.isInteger(column) ||
      column < 0 ||
      column >= model.headers.length ||
      years.has(year) ||
      usedColumns.has(column)
    ) {
      pause("AMBIGUOUS_MULTI_INDICATOR", "指标年份到列的映射不唯一", {
        label,
        yearColumns: block.yearColumns,
      });
    }
    years.set(year, column);
    usedColumns.add(column);
  }
  if (multiMode) {
    const actual = [...years.keys()].sort((left, right) => right - left);
    if (
      actual.length !== target.years.length ||
      actual.some((year, index) => year !== target.years[index])
    ) {
      pause("AMBIGUOUS_MULTI_INDICATOR", "多个指标没有共享相同目标年份集合", {
        label,
        expectedYears: target.years,
        actualYears: actual,
      });
    }
  }
  return { label, years, usedColumns };
}

function indicatorBlocks(model, target) {
  const multiMode = Array.isArray(model.indicatorBlocks);
  const sourceBlocks = multiMode
    ? model.indicatorBlocks
    : [{ label: model.indicatorLabel ?? "指标值", yearColumns: model.yearColumns ?? [] }];
  if (multiMode && sourceBlocks.length < 2) {
    pause("AMBIGUOUS_MULTI_INDICATOR", "多指标模式至少需要两个可靠指标块", {
      blockCount: sourceBlocks.length,
    });
  }
  const blocks = sourceBlocks.map((block) => normalizeBlock(block, model, target, multiMode));
  const labels = blocks.map((block) => block.label);
  if (new Set(labels).size !== labels.length) {
    pause("AMBIGUOUS_MULTI_INDICATOR", "多个指标标签不唯一", { labels });
  }
  if (multiMode) {
    const allColumns = blocks.flatMap((block) => [...block.usedColumns]);
    if (new Set(allColumns).size !== allColumns.length) {
      pause("AMBIGUOUS_MULTI_INDICATOR", "多个指标块复用了同一源数据列", {
        columns: allColumns,
      });
    }
  }
  return { blocks, multiMode };
}

function rowByCity(model, cityContext) {
  const result = new Map();
  for (const row of model.rows) {
    if (!Array.isArray(row.values) || row.values.length !== model.headers.length) {
      pause("UNSAFE_STRUCTURE", "宽表源行列数与表头不一致", {
        sheetName: model.sheetName,
        sourceRow: row.sourceRow,
        expectedColumns: model.headers.length,
        actualColumns: row.values?.length,
      });
    }
    const city = row.cityMatch?.standardName;
    if (!city || !cityContext.rankByName.has(city)) {
      pause("UNMATCHED_CITY", "宽表源行缺少已确认的标准城市名", {
        sheetName: model.sheetName,
        sourceRow: row.sourceRow,
        rawCity: row.values[model.cityColumn],
      });
    }
    if (result.has(city)) {
      pause("DUPLICATE_CITY_YEAR", "宽表中同一标准城市出现多行", {
        sheetName: model.sheetName,
        city,
        sourceRows: [result.get(city).sourceRow, row.sourceRow],
      });
    }
    result.set(city, row);
  }
  return result;
}

function metadataForOutput(row, auxiliaryColumns, indicatorColumns) {
  const source = Array.isArray(row.cellMetadata) ? row.cellMetadata : [];
  return [
    ...auxiliaryColumns.map((column) => cloneMetadata(source[column])),
    null,
    ...indicatorColumns.map((column) => cloneMetadata(source[column])),
  ];
}

export function alignWideTable(model, cityContext, config = {}) {
  validateBaseModel(model, cityContext);
  const target = targetYears(config);
  const { blocks, multiMode } = indicatorBlocks(model, target);
  const sourceRows = rowByCity(model, cityContext);
  const outputCityColumn = model.auxiliaryColumns.indexOf(model.cityColumn);
  const headers = [
    ...model.auxiliaryColumns.map((column) => model.headers[column]),
    "年份",
    ...blocks.map((block) => block.label),
  ];

  const outOfRangeColumns = model.outOfRangeColumns ?? [];
  if (outOfRangeColumns.length && !config.allowOutOfRangeYears) {
    pause("OUT_OF_RANGE_YEARS", "宽表包含目标范围外年份", {
      sheetName: model.sheetName,
      years: [...new Set(outOfRangeColumns.map((item) => item.year))],
      columns: outOfRangeColumns,
    });
  }
  const outOfRangeRecords = [];
  if (config.allowOutOfRangeYears) {
    for (const row of model.rows) {
      for (const item of outOfRangeColumns) {
        outOfRangeRecords.push({
          sourceSheet: model.sheetName,
          sourceRow: row.sourceRow,
          sourceColumn: item.column,
          year: item.year,
          values: [...row.values],
          value: row.values[item.column],
        });
      }
    }
  }

  const records = [];
  const missingKeys = [];
  const mappingEvents = [];
  for (const row of model.rows) {
    if (row.cityMatch?.method !== "exact") {
      mappingEvents.push({
        sourceRow: row.sourceRow,
        rawCity: row.values[model.cityColumn],
        standardName: row.cityMatch.standardName,
        method: row.cityMatch.method,
      });
    }
  }

  for (const city of cityContext.entries) {
    const sourceRow = sourceRows.get(city.standardName);
    for (const year of target.years) {
      const indicatorColumns = blocks.map((block) => block.years.get(year));
      const hasCompleteSource =
        Boolean(sourceRow) && indicatorColumns.every((column) => Number.isInteger(column));
      if (!hasCompleteSource) {
        const values = Array(headers.length).fill(null);
        values[outputCityColumn] = city.standardName;
        values[model.auxiliaryColumns.length] = year;
        records.push({
          values,
          city: city.standardName,
          year,
          missing: true,
          sourceSheet: null,
          sourceRow: null,
          sourceColumn: null,
          sourceColumns: null,
          cellMetadata: null,
          cityMatch: null,
        });
        missingKeys.push({ city: city.standardName, year });
        continue;
      }

      const values = [
        ...model.auxiliaryColumns.map((column) => sourceRow.values[column]),
        year,
        ...indicatorColumns.map((column) => sourceRow.values[column] ?? null),
      ];
      values[outputCityColumn] = city.standardName;
      records.push({
        values,
        city: city.standardName,
        year,
        missing: false,
        sourceSheet: model.sheetName,
        sourceRow: sourceRow.sourceRow,
        sourceColumn: multiMode ? null : indicatorColumns[0],
        sourceColumns: Object.fromEntries(
          blocks.map((block, index) => [block.label, indicatorColumns[index]]),
        ),
        cellMetadata: metadataForOutput(
          sourceRow,
          model.auxiliaryColumns,
          indicatorColumns,
        ),
        cityMatch: { ...sourceRow.cityMatch },
      });
    }
  }

  return {
    kind: multiMode ? "wide-multi-aligned" : "wide-aligned",
    sheetName: model.sheetName,
    headers,
    rows: records.map((record) => record.values),
    records,
    cityColumn: outputCityColumn,
    yearColumn: model.auxiliaryColumns.length,
    indicatorColumns: blocks.map((block, index) => ({
      label: block.label,
      column: model.auxiliaryColumns.length + 1 + index,
    })),
    targetYears: target.years,
    missingKeys,
    outOfRangeRecords,
    mappingEvents,
    sourceCityCount: model.rows.length,
    multiIndicator: multiMode,
  };
}