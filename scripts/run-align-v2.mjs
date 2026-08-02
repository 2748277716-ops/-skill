import fs from "node:fs/promises";
import path from "node:path";

import { alignLongTable } from "./lib/align-long.mjs";
import { alignWideTable } from "./lib/align-wide.mjs";
import {
  normalizeSafeCityName,
  resolveCityName,
  validateCityOrder,
  validateMappingRows,
} from "./lib/cities.mjs";
import { detectTableModel } from "./lib/detect-v2.mjs";
import { PauseError } from "./lib/pause.mjs";
import { recommendProcessingModes } from "./lib/recommendation.mjs";
import { verifyLongAlignment, verifyWideAlignment } from "./lib/verify-v2.mjs";
import {
  computeFileSha256,
  listWorkbookStructure,
  readWorkbookFast,
  writeCleanResultWorkbook,
} from "./lib/workbook-fast.mjs";

function pause(code, message, evidence = {}) {
  throw new PauseError(code, message, evidence);
}

function requiredPath(config, key) {
  const value = config?.[key];
  if (!value || typeof value !== "string") {
    pause("INVALID_CONFIG", `缺少必需路径：${key}`, { key, value });
  }
  return path.resolve(value);
}

function cityOrderRows(sheetModel) {
  const matrix = sheetModel.matrix;
  if (!matrix.length) pause("INVALID_CITY_ORDER", "城市顺序工作表为空", {});
  const headers = matrix[0].map((value) => String(value ?? "").trim());
  const sequenceColumn = headers.indexOf("序号");
  const cityColumn = headers.indexOf("城市名");
  if (sequenceColumn < 0 || cityColumn < 0) {
    pause("INVALID_CITY_ORDER", "城市顺序工作表缺少序号或城市名表头", {
      headers,
    });
  }
  return matrix.slice(1).flatMap((row) => {
    const sequence = row[sequenceColumn];
    const cityName = row[cityColumn];
    return sequence === null && cityName === null
      ? []
      : [{ 序号: sequence, 城市名: cityName }];
  });
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

function inferSourceYears(matrix) {
  const headers = matrix[0] ?? [];
  const yearColumns = headers.flatMap((header, column) =>
    ["年份", "年度", "统计年份", "year"].includes(normalizedHeader(header))
      ? [column]
      : [],
  );
  if (yearColumns.length === 1) {
    const years = matrix.slice(1).flatMap((row) => {
      const year = parseYear(row[yearColumns[0]]);
      return year === null ? [] : [year];
    });
    return [...new Set(years)].sort((left, right) => right - left);
  }
  const years = headers.flatMap((header) => {
    const year = parseYear(header);
    return year === null ? [] : [year];
  });
  return [...new Set(years)].sort((left, right) => right - left);
}

function configuredYears(config, matrix) {
  const hasStart = config?.startYear !== null && config?.startYear !== undefined;
  const hasEnd = config?.endYear !== null && config?.endYear !== undefined;
  if (hasStart || hasEnd) {
    const startYear = Number(config.startYear);
    const endYear = Number(config.endYear);
    if (!Number.isInteger(startYear) || !Number.isInteger(endYear)) {
      pause("INVALID_CONFIG", "年份范围必须同时提供两个整数", {
        startYear: config.startYear,
        endYear: config.endYear,
      });
    }
    const years = [];
    for (
      let year = Math.max(startYear, endYear);
      year >= Math.min(startYear, endYear);
      year -= 1
    ) years.push(year);
    return { years, explicitlyConfigured: true };
  }
  const years = inferSourceYears(matrix);
  if (!years.length) {
    pause("AMBIGUOUS_YEAR_COLUMN", "未指定年份范围，且无法从源表提取年份", {});
  }
  return { years, explicitlyConfigured: false };
}

function findCityColumn(matrix) {
  const headers = matrix[0] ?? [];
  const candidates = headers.flatMap((header, column) =>
    ["城市", "城市名", "城市名称", "地级市", "city", "cityname"].includes(normalizedHeader(header))
      ? [column]
      : [],
  );
  return candidates.length === 1 ? candidates[0] : null;
}

function extendCityContext(matrix, baseCityContext, mappingRows) {
  const cityColumn = findCityColumn(matrix);
  if (cityColumn === null) {
    return { cityContext: baseCityContext, outsideOrderCities: [] };
  }
  const outsideOrderCities = [];
  const seen = new Set(baseCityContext.names);
  for (const row of matrix.slice(1)) {
    const rawCity = row[cityColumn];
    const normalized = normalizeSafeCityName(rawCity);
    if (!normalized) continue;
    const match = resolveCityName(rawCity, baseCityContext, mappingRows);
    if (match.status === "matched") continue;
    if (!seen.has(normalized)) {
      seen.add(normalized);
      outsideOrderCities.push(normalized);
    }
  }
  if (!outsideOrderCities.length) {
    return { cityContext: baseCityContext, outsideOrderCities };
  }
  const rows = [
    ...baseCityContext.entries.map((entry) => ({
      序号: entry.sequence,
      城市名: entry.standardName,
    })),
    ...outsideOrderCities.map((city, index) => ({
      序号: baseCityContext.entries.length + index + 1,
      城市名: city,
    })),
  ];
  return {
    cityContext: validateCityOrder(rows),
    outsideOrderCities,
  };
}

function attachWorkbookMetadata(detected, sheetModel) {
  const rows = detected.rows.map((row) => {
    const sourceRecord = sheetModel.rowRecords[row.sourceRow - 2];
    if (!sourceRecord) {
      pause("UNSAFE_STRUCTURE", "无法把检测行追溯到源工作表行", {
        sheetName: sheetModel.name,
        detectedSourceRow: row.sourceRow,
      });
    }
    return {
      ...row,
      sourceRow: sourceRecord.sourceRow,
      sourceCells: sourceRecord.sourceCells,
      cellMetadata: sourceRecord.cellMetadata,
    };
  });
  return {
    ...detected,
    rows,
    headerMetadata: sheetModel.headerMetadata,
    dataStartRow: sheetModel.dataStartRow,
    dataStartColumn: sheetModel.dataStartColumn,
    sourceSheetSha256: sheetModel.sheetSha256,
  };
}

function projectedSourceModel(sourceModel, outputModel) {
  if (outputModel.outputMode !== "selected_indicators") return sourceModel;
  const columns = outputModel.sourceColumns;
  return {
    ...sourceModel,
    headers: columns.map((column) => sourceModel.headers[column]),
    cityColumn: outputModel.cityColumn,
    yearColumn: outputModel.yearColumn,
    dataStartColumn: 0,
    rows: sourceModel.rows.map((row) => ({
      ...row,
      values: columns.map((column) => row.values[column]),
      sourceCells: Array.isArray(row.sourceCells)
        ? columns.map((column) => row.sourceCells[column])
        : row.sourceCells,
      cellMetadata: Array.isArray(row.cellMetadata)
        ? columns.map((column) => row.cellMetadata[column])
        : row.cellMetadata,
    })),
  };
}

function filterWideToExactYears(output, years) {
  const allowed = new Set(years);
  output.records = output.records.filter((record) => allowed.has(record.year));
  output.rows = output.records.map((record) => record.values);
  output.missingKeys = output.missingKeys.filter((item) => allowed.has(item.year));
  output.targetYears = [...years];
  return output;
}

async function selectSheets(inputPath, configuredSheets) {
  const structure = await listWorkbookStructure(inputPath);
  let selectedSheets = Array.isArray(configuredSheets) ? [...configuredSheets] : [];
  if (selectedSheets.length === 0) {
    if (structure.sheets.length !== 1) {
      pause("MULTIPLE_SHEETS", "源工作簿包含多个工作表，需要先明确选择", {
        sheets: structure.sheets,
      });
    }
    selectedSheets = [structure.sheets[0].name];
  }
  const unknownSheets = selectedSheets.filter(
    (name) => !structure.sheets.some((sheet) => sheet.name === name),
  );
  if (unknownSheets.length) {
    pause("MULTIPLE_SHEETS", "选定工作表不存在", {
      selectedSheets,
      unknownSheets,
      sheets: structure.sheets,
    });
  }
  return { selectedSheets, structure };
}

export async function estimateAlignmentModes(config) {
  const inputPath = requiredPath(config, "inputPath");
  const { selectedSheets, structure } = await selectSheets(
    inputPath,
    config.selectedSheets,
  );
  return selectedSheets.map((sheetName) => {
    const sheet = structure.sheets.find((item) => item.name === sheetName);
    return {
      sheetName,
      ...recommendProcessingModes({
        rowCount: Math.max(1, sheet.rowCount),
        columnCount: Math.max(1, sheet.columnCount),
        selectedIndicatorCount: Array.isArray(config.selectedIndicators)
          ? config.selectedIndicators.length
          : Number(config.selectedIndicatorCount ?? 1),
      }),
    };
  });
}

export async function runAlignmentV2(config, runtime = {}) {
  let createdOutputPath = null;
  try {
    const inputPath = requiredPath(config, "inputPath");
    const cityOrderPath = requiredPath(config, "cityOrderPath");
    const outputDir = path.resolve(config.outputDir ?? path.join(path.dirname(inputPath), "处理后数据"));
    const { selectedSheets } = await selectSheets(inputPath, config.selectedSheets);
    const outputMode = config.outputMode ?? (
      Array.isArray(config.selectedIndicators) && config.selectedIndicators.length
        ? "selected_indicators"
        : "preserve_rows"
    );
    if (!["preserve_rows", "selected_indicators"].includes(outputMode)) {
      pause("INVALID_CONFIG", "outputMode 只能是 preserve_rows 或 selected_indicators", {
        outputMode,
      });
    }

    const orderStructure = await listWorkbookStructure(cityOrderPath);
    const cityOrderSheet = config.cityOrderSheet ?? (
      orderStructure.sheets.length === 1 ? orderStructure.sheets[0].name : null
    );
    if (!cityOrderSheet) {
      pause("MULTIPLE_SHEETS", "城市顺序工作簿包含多个工作表，需要指定 cityOrderSheet", {
        sheets: orderStructure.sheets,
      });
    }
    const cityOrderWorkbook = await readWorkbookFast(cityOrderPath, [cityOrderSheet]);
    const baseCityContext = validateCityOrder(cityOrderRows(cityOrderWorkbook.sheets[0]));
    const mappingPath = path.resolve(
      config.mappingPath ?? path.join(path.dirname(cityOrderPath), "城市名称映射表.xlsx"),
    );
    const { readMappingRows } = await import("./lib/mapping-store.mjs");
    const existingMappings = await readMappingRows(mappingPath);
    const confirmedMappings = Array.isArray(config.confirmedMappings)
      ? config.confirmedMappings.map((item) => ({
          原始城市名: item.sourceName,
          标准城市名: item.standardName,
        }))
      : [];
    const validatedMappings = validateMappingRows(
      [...existingMappings, ...confirmedMappings],
      baseCityContext,
    );

    const sourceReadOptions = outputMode === "selected_indicators"
      ? {
          projectHeaders: ["城市", "年份", ...(config.selectedIndicators ?? [])],
          projectionRoles: {
            cityHeaders: ["城市", "城市名", "城市名称", "地级市", "city", "cityname"],
            yearHeaders: ["年份", "年度", "统计年份", "year"],
          },
        }
      : {};
    const sourceWorkbook = await readWorkbookFast(
      inputPath,
      selectedSheets,
      sourceReadOptions,
    );
    const sourceModels = [];
    const outputModels = [];
    const verifications = [];
    const outsideOrderCities = [];
    const yearsBySheet = {};
    for (const sheetModel of sourceWorkbook.sheets) {
      const yearSelection = configuredYears(config, sheetModel.matrix);
      yearsBySheet[sheetModel.name] = yearSelection.years;
      const extended = extendCityContext(
        sheetModel.matrix,
        baseCityContext,
        validatedMappings,
      );
      for (const city of extended.outsideOrderCities) {
        if (!outsideOrderCities.includes(city)) outsideOrderCities.push(city);
      }
      const detected = detectTableModel(
        sheetModel.matrix,
        {
          startYear: Math.max(...yearSelection.years),
          endYear: Math.min(...yearSelection.years),
          sheetName: sheetModel.name,
          workbookSheets: selectedSheets,
          selectedSheets,
          mergedRanges: sheetModel.unsafeMerges,
          mappingRows: validatedMappings,
          allowOutOfRangeYears: true,
        },
        extended.cityContext,
      );
      const sourceModel = attachWorkbookMetadata(detected, sheetModel);
      sourceModels.push(sourceModel);
      if (sourceModel.kind === "wide" && outputMode === "selected_indicators") {
        pause(
          "INDICATOR_SELECTION_UNSUPPORTED",
          "宽表的指标通常由工作表或指标区块定义；请先选择对应工作表，再使用完整行模式",
          { sheetName: sheetModel.name, selectedIndicators: config.selectedIndicators },
        );
      }
      const transformConfig = {
        startYear: Math.max(...yearSelection.years),
        endYear: Math.min(...yearSelection.years),
        targetYears: yearSelection.years,
        allowOutOfRangeYears: false,
        allowOutsideOrderCities: true,
        outputMode,
        selectedIndicators: config.selectedIndicators ?? [],
      };
      const output = sourceModel.kind === "long"
        ? alignLongTable(sourceModel, extended.cityContext, transformConfig)
        : filterWideToExactYears(
            alignWideTable(sourceModel, extended.cityContext, transformConfig),
            yearSelection.years,
          );
      output.outputSheetName = selectedSheets.length === 1
        ? "对齐结果"
        : (config.approvedSheetNames?.[sheetModel.name] ?? `对齐结果${outputModels.length + 1}`);
      output.outsideOrderCities = extended.outsideOrderCities;
      outputModels.push(output);
      if (sourceModel.kind === "long") {
        const verificationSource = projectedSourceModel(sourceModel, output);
        verifications.push(verifyLongAlignment(verificationSource, output, {
          cityContext: extended.cityContext,
          formulaEvents: [],
        }));
      } else {
        verifications.push(verifyWideAlignment(sourceModel, output, {
          cityContext: extended.cityContext,
          formulaEvents: sourceWorkbook.formulaEvents,
        }));
      }
    }

    const auditPassed = verifications.length > 0 && verifications.every(
      (verification) => verification.passed === true,
    );
    if (!auditPassed) {
      pause("REVERSE_VERIFICATION_FAILED", "结构核验未通过", {
        verifications,
      });
    }
    const now = runtime.now ? runtime.now() : new Date();
    const { buildOutputPath } = await import("./run-align.mjs");
    const outputPath = buildOutputPath(inputPath, outputDir, now);
    const written = await writeCleanResultWorkbook(outputModels, outputPath);
    createdOutputPath = written.outputPath;

    const inputHashAfter = await computeFileSha256(inputPath);
    const cityOrderHashAfter = await computeFileSha256(cityOrderPath);
    if (
      inputHashAfter !== sourceWorkbook.sha256 ||
      cityOrderHashAfter !== cityOrderWorkbook.sha256
    ) {
      await fs.rm(createdOutputPath, { force: true });
      createdOutputPath = null;
      pause("SOURCE_HASH_CHANGED", "输出期间源文件或城市顺序文件发生变化", {
        inputBefore: sourceWorkbook.sha256,
        inputAfter: inputHashAfter,
        cityOrderBefore: cityOrderWorkbook.sha256,
        cityOrderAfter: cityOrderHashAfter,
      });
    }

    return {
      status: "passed",
      outputPath: written.outputPath,
      outputSha256: written.outputSha256,
      outputMode,
      selectedIndicators: config.selectedIndicators ?? [],
      selectedSheets,
      yearsBySheet,
      outsideOrderCities,
      cityOrderDerivedFileCreated: false,
      additionalOutputFiles: [],
      audit: {
        passed: true,
        sourceColumnCountsBySheet: Object.fromEntries(
          sourceWorkbook.sheets.map((sheet) =>
            [sheet.name, sheet.projectedColumnCount ?? sheet.matrix[0]?.length ?? 0]),
        ),
        verifications,
        inputSha256: sourceWorkbook.sha256,
        inputSha256After: inputHashAfter,
        cityOrderSha256: cityOrderWorkbook.sha256,
        cityOrderSha256After: cityOrderHashAfter,
        formulaValueCount: sourceWorkbook.formulaEvents.length,
        visualRendering: "not performed",
      },
    };
  } catch (error) {
    if (createdOutputPath) {
      await fs.rm(createdOutputPath, { force: true }).catch(() => {});
    }
    if (error instanceof PauseError) return error.toJSON();
    return {
      status: "failed",
      code: "UNEXPECTED_ERROR",
      message: error?.message ?? String(error),
      evidence: { stack: error?.stack ?? null },
    };
  }
}
