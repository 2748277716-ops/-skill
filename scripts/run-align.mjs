import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { alignLongTable } from "./lib/align-long.mjs";
import { alignWideTable } from "./lib/align-wide.mjs";
import { buildAudit } from "./lib/audit.mjs";
import {
  validateCityOrder,
  validateMappingRows,
} from "./lib/cities.mjs";
import { detectTableModel } from "./lib/detect.mjs";
import {
  appendConfirmedMappings,
  readMappingRows,
} from "./lib/mapping-store.mjs";
import { PauseError } from "./lib/pause.mjs";
import {
  computeFileSha256,
  listWorkbookStructure,
  readWorkbookModel,
  writeResultWorkbook,
} from "./lib/workbook-io.mjs";
import {
  verifyLongAlignment,
  verifyWideAlignment,
} from "./lib/verify.mjs";

function pause(code, message, evidence = {}) {
  throw new PauseError(code, message, evidence);
}

function requiredPath(config, key) {
  if (!config?.[key] || typeof config[key] !== "string") {
    pause("INVALID_CONFIG", `缺少必需路径：${key}`, { key, value: config?.[key] });
  }
  return path.resolve(config[key]);
}

function pad(value) {
  return String(value).padStart(2, "0");
}

export function formatTimestamp(date) {
  return (
    `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-` +
    `${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`
  );
}

export function buildOutputPath(inputPath, outputDir, date) {
  const input = path.parse(inputPath);
  return path.join(
    outputDir,
    `${input.name}_城市面板对齐_${formatTimestamp(date)}.xlsx`,
  );
}

function validateYears(config) {
  const startYear = Number(config?.startYear);
  const endYear = Number(config?.endYear);
  if (!Number.isInteger(startYear) || !Number.isInteger(endYear)) {
    pause("INVALID_CONFIG", "必须显式提供整数目标年份范围", {
      startYear: config?.startYear,
      endYear: config?.endYear,
    });
  }
  return { startYear, endYear };
}

function validSheetName(name) {
  return Boolean(name) && name.length <= 31 && !/[\[\]:*?/\\]/u.test(name);
}

function outputSheetNames(selectedSheets, approvedSheetNames = {}) {
  const names = new Map();
  for (const sheetName of selectedSheets) {
    const derived = selectedSheets.length === 1 ? "对齐结果" : `对齐结果_${sheetName}`;
    const outputName = approvedSheetNames[sheetName] ?? derived;
    if (!validSheetName(outputName)) {
      pause("INVALID_SHEET_NAME", "派生输出工作表名称无效，需要明确批准安全名称", {
        sourceSheet: sheetName,
        derivedName: derived,
        approvedName: approvedSheetNames[sheetName] ?? null,
      });
    }
    names.set(sheetName, outputName);
  }
  const duplicates = [...names.values()].filter(
    (name, index, all) => all.indexOf(name) !== index,
  );
  if (duplicates.length) {
    pause("INVALID_SHEET_NAME", "批准的输出工作表名称存在重复", {
      duplicates: [...new Set(duplicates)],
    });
  }
  return names;
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

function outputHeaderMetadata(sourceModel) {
  if (sourceModel.kind === "long") return sourceModel.headerMetadata ?? [];
  const auxiliary = sourceModel.auxiliaryColumns.map(
    (column) => sourceModel.headerMetadata?.[column] ?? null,
  );
  const indicatorCount = Array.isArray(sourceModel.indicatorBlocks)
    ? sourceModel.indicatorBlocks.length
    : 1;
  return [...auxiliary, null, ...Array(indicatorCount).fill(null)];
}

function serializePause(error, expectedOutputPath) {
  if (error.code === "UNMATCHED_CITY" || error.code === "AMBIGUOUS_CITY_MATCH") {
    return {
      status: "paused",
      code: "CITY_CONFIRMATION_REQUIRED",
      message: "存在只能由用户确认的城市名称映射",
      evidence: error.evidence,
      expectedOutputPath,
    };
  }
  return { ...error.toJSON(), expectedOutputPath };
}

export async function runAlignment(config, runtime = {}) {
  let expectedOutputPath = null;
  let createdOutputPath = null;
  try {
    const inputPath = requiredPath(config, "inputPath");
    const cityOrderPath = requiredPath(config, "cityOrderPath");
    const years = validateYears(config);
    const outputDir = path.resolve(config.outputDir ?? path.dirname(inputPath));
    const mappingPath = path.resolve(
      config.mappingPath ?? path.join(path.dirname(cityOrderPath), "城市名称映射表.xlsx"),
    );
    const now = runtime.now ? runtime.now() : new Date();
    expectedOutputPath = buildOutputPath(inputPath, outputDir, now);
    const protectedPaths = [inputPath, cityOrderPath, mappingPath].map((item) => path.resolve(item));
    if (protectedPaths.includes(path.resolve(expectedOutputPath))) {
      pause("INVALID_CONFIG", "输出路径不得等于任何输入或映射路径", {
        expectedOutputPath,
        protectedPaths,
      });
    }

    const inputStructure = await listWorkbookStructure(inputPath);
    let selectedSheets = Array.isArray(config.selectedSheets) ? [...config.selectedSheets] : [];
    if (selectedSheets.length === 0) {
      if (inputStructure.sheets.length === 1) {
        selectedSheets = [inputStructure.sheets[0].name];
      } else {
        pause("MULTIPLE_SHEETS", "源工作簿包含多个工作表，需要先明确选择", {
          sheets: inputStructure.sheets,
        });
      }
    }
    const unknownSheets = selectedSheets.filter(
      (name) => !inputStructure.sheets.some((sheet) => sheet.name === name),
    );
    if (unknownSheets.length) {
      pause("MULTIPLE_SHEETS", "选定工作表不存在", {
        selectedSheets,
        unknownSheets,
        sheets: inputStructure.sheets,
      });
    }
    const nameMap = outputSheetNames(selectedSheets, config.approvedSheetNames ?? {});

    const orderStructure = await listWorkbookStructure(cityOrderPath);
    let cityOrderSheet = config.cityOrderSheet ?? null;
    if (!cityOrderSheet) {
      if (orderStructure.sheets.length !== 1) {
        pause("MULTIPLE_SHEETS", "城市顺序工作簿包含多个工作表，需要指定 cityOrderSheet", {
          sheets: orderStructure.sheets,
        });
      }
      cityOrderSheet = orderStructure.sheets[0].name;
    }
    const cityOrderWorkbook = await readWorkbookModel(cityOrderPath, [cityOrderSheet]);
    const cityContext = validateCityOrder(cityOrderRows(cityOrderWorkbook.sheets[0]));

    const existingMappings = await readMappingRows(mappingPath);
    const confirmedMappings = Array.isArray(config.confirmedMappings)
      ? config.confirmedMappings
      : [];
    const effectiveMappingRows = [
      ...existingMappings,
      ...confirmedMappings.map((item) => ({
        原始城市名: item.sourceName,
        标准城市名: item.standardName,
      })),
    ];
    const validatedMappings = validateMappingRows(effectiveMappingRows, cityContext);

    const sourceWorkbook = await readWorkbookModel(inputPath, selectedSheets);
    const sourceModels = [];
    const outputModels = [];
    const verifications = [];
    const approvedExcludedYears = new Set(
      (config.approvedExcludedYears ?? []).map((year) => Number(year)),
    );

    for (const sheetModel of sourceWorkbook.sheets) {
      const detected = detectTableModel(
        sheetModel.matrix,
        {
          ...years,
          sheetName: sheetModel.name,
          workbookSheets: inputStructure.sheets.map((sheet) => sheet.name),
          selectedSheets,
          mergedRanges: sheetModel.unsafeMerges,
          mappingRows: validatedMappings,
          allowOutOfRangeYears: true,
        },
        cityContext,
      );
      const sourceModel = attachWorkbookMetadata(detected, sheetModel);
      const unapprovedYears = (sourceModel.outOfRangeYears ?? []).filter(
        (year) => !approvedExcludedYears.has(Number(year)),
      );
      if (unapprovedYears.length) {
        pause("OUT_OF_RANGE_YEARS", "存在尚未批准排除的范围外年份", {
          sheetName: sheetModel.name,
          years: unapprovedYears,
          approvedExcludedYears: [...approvedExcludedYears],
        });
      }
      sourceModels.push(sourceModel);
      const transformConfig = {
        ...years,
        allowOutOfRangeYears: (sourceModel.outOfRangeYears ?? []).length > 0,
      };
      const output = sourceModel.kind === "long"
        ? alignLongTable(sourceModel, cityContext, transformConfig)
        : alignWideTable(sourceModel, cityContext, transformConfig);
      output.outputSheetName = nameMap.get(sheetModel.name);
      output.headerMetadata = outputHeaderMetadata(sourceModel);
      output.mappingEvents = (output.mappingEvents ?? []).map((event) => ({
        ...event,
        sheetName: sheetModel.name,
      }));
      outputModels.push(output);
      verifications.push(
        sourceModel.kind === "long"
          ? verifyLongAlignment(sourceModel, output, {
              cityContext,
              formulaEvents: sourceWorkbook.formulaEvents,
            })
          : verifyWideAlignment(sourceModel, output, {
              cityContext,
              formulaEvents: sourceWorkbook.formulaEvents,
            }),
      );
    }

    const mappingResult = confirmedMappings.length
      ? await appendConfirmedMappings(
          mappingPath,
          confirmedMappings,
          cityContext,
        )
      : {
          mappingPath,
          appendedCount: 0,
          beforeSha256: await fs.access(mappingPath).then(
            () => computeFileSha256(mappingPath),
            () => null,
          ),
          afterSha256: await fs.access(mappingPath).then(
            () => computeFileSha256(mappingPath),
            () => null,
          ),
        };

    const inputHashBeforeWrite = await computeFileSha256(inputPath);
    const cityOrderHashBeforeWrite = await computeFileSha256(cityOrderPath);
    if (
      inputHashBeforeWrite !== sourceWorkbook.sha256 ||
      cityOrderHashBeforeWrite !== cityOrderWorkbook.sha256
    ) {
      pause("SOURCE_HASH_CHANGED", "执行期间源文件或城市顺序文件发生变化", {
        inputBefore: sourceWorkbook.sha256,
        inputNow: inputHashBeforeWrite,
        cityOrderBefore: cityOrderWorkbook.sha256,
        cityOrderNow: cityOrderHashBeforeWrite,
      });
    }

    const audit = buildAudit({
      fileMetadata: {
        inputPath,
        inputSha256: sourceWorkbook.sha256,
        cityOrderPath,
        cityOrderSha256: cityOrderWorkbook.sha256,
        mappingPath,
        mappingBeforeSha256: mappingResult.beforeSha256,
        mappingAfterSha256: mappingResult.afterSha256,
      },
      config: { ...years, selectedSheets },
      sourceModels,
      outputModels,
      verifications,
      formulaEvents: sourceWorkbook.formulaEvents,
    });
    if (!audit.passed) {
      pause("REVERSE_VERIFICATION_FAILED", "审计结论未通过", { audit });
    }

    const written = await writeResultWorkbook(
      {
        outputs: outputModels,
        audit,
        formulaEvents: sourceWorkbook.formulaEvents,
      },
      expectedOutputPath,
    );
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
      audit: {
        ...audit,
        inputSha256After: inputHashAfter,
        cityOrderSha256After: cityOrderHashAfter,
        outputSha256: written.outputSha256,
        mappingAppend: mappingResult,
        visual_rendering: "not performed - user approval not requested",
      },
    };
  } catch (error) {
    if (createdOutputPath) {
      await fs.rm(createdOutputPath, { force: true }).catch(() => {});
    }
    if (error instanceof PauseError) {
      return serializePause(error, expectedOutputPath);
    }
    return {
      status: "failed",
      code: "UNEXPECTED_ERROR",
      message: error?.message ?? String(error),
      evidence: { stack: error?.stack ?? null },
      expectedOutputPath,
    };
  }
}

async function main() {
  const argument = process.argv[2];
  if (!argument) {
    throw new Error("Usage: node scripts/run-align.mjs <config-json-or-path>");
  }
  let configText = argument;
  try {
    configText = await fs.readFile(path.resolve(argument), "utf8");
  } catch {
    // Treat the argument itself as JSON.
  }
  const result = await runAlignment(JSON.parse(configText));
  process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
  if (result.status === "failed") process.exitCode = 1;
}

const runtimeProcess = globalThis.process;
if (
  runtimeProcess?.argv?.[1] &&
  import.meta.url === pathToFileURL(runtimeProcess.argv[1]).href
) {
  await main();
}