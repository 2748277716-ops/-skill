import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

import {
  normalizeSafeCityName,
  validateMappingRows,
} from "./cities.mjs";
import { PauseError } from "./pause.mjs";

const CANONICAL_HEADERS = [
  "原始城市名",
  "标准城市名",
  "匹配类型",
  "确认时间",
  "来源文件",
  "来源工作表",
  "备注",
];

function pause(code, message, evidence = {}) {
  throw new PauseError(code, message, evidence);
}

function rowIsBlank(row) {
  return row.every((value) => value === null || value === undefined || String(value).trim() === "");
}

function rowsToObjects(headers, rows) {
  return rows
    .filter((row) => !rowIsBlank(row))
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? null])));
}

async function exists(fileAdapter, filePath) {
  try {
    await fileAdapter.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function hashFile(fileAdapter, filePath) {
  return crypto.createHash("sha256").update(await fileAdapter.readFile(filePath)).digest("hex");
}

async function loadMappingDocument(mappingPath) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(mappingPath));
  const sheet = workbook.worksheets.getItemAt(0);
  const usedRange = sheet.getUsedRange();
  if (!usedRange) {
    pause("MAPPING_CONFLICT", "城市名称映射表没有表头", { path: mappingPath });
  }
  const values = usedRange.values;
  const formulas = usedRange.formulas;
  if (formulas.flat().some((formula) => Boolean(formula))) {
    pause("MAPPING_CONFLICT", "城市名称映射表不得使用公式", { path: mappingPath });
  }
  const headers = values[0].map((value) => String(value ?? "").trim());
  if (!headers.includes("原始城市名") || !headers.includes("标准城市名")) {
    pause("MAPPING_CONFLICT", "城市名称映射表缺少必需表头", {
      path: mappingPath,
      headers,
      requiredHeaders: ["原始城市名", "标准城市名"],
    });
  }
  const rowArrays = values.slice(1).filter((row) => !rowIsBlank(row));
  return {
    workbook,
    sheet,
    headers,
    rowArrays,
    rows: rowsToObjects(headers, rowArrays),
  };
}

export async function readMappingRows(mappingPath) {
  try {
    await fs.access(mappingPath);
  } catch {
    return [];
  }
  return (await loadMappingDocument(mappingPath)).rows;
}

function mappingObject(confirmed, now) {
  return {
    原始城市名: String(confirmed.sourceName ?? ""),
    标准城市名: String(confirmed.standardName ?? ""),
    匹配类型: confirmed.matchType ?? "用户确认",
    确认时间: confirmed.confirmedAt ?? now,
    来源文件: confirmed.sourceFile ?? null,
    来源工作表: confirmed.sourceSheet ?? null,
    备注: confirmed.note ?? "用户确认",
  };
}

function sameRequiredMapping(left, right) {
  return (
    normalizeSafeCityName(left.原始城市名) === normalizeSafeCityName(right.原始城市名) &&
    left.标准城市名 === right.标准城市名
  );
}

function verifyRowsPreserved(beforeRows, afterRows, additions) {
  if (afterRows.length !== beforeRows.length + additions.length) return false;
  for (let index = 0; index < beforeRows.length; index += 1) {
    for (const header of Object.keys(beforeRows[index])) {
      if (!Object.is(afterRows[index][header] ?? null, beforeRows[index][header] ?? null)) return false;
    }
  }
  return additions.every((addition) => afterRows.some((row) => sameRequiredMapping(row, addition)));
}

async function defaultExportWorkbook(workbook, targetPath) {
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(targetPath);
}

export async function appendConfirmedMappings(
  mappingPath,
  confirmedMappings,
  cityOrder,
  options = {},
) {
  if (!Array.isArray(confirmedMappings)) {
    pause("MAPPING_CONFLICT", "confirmedMappings 必须是明确确认记录数组", {});
  }
  const absolutePath = path.resolve(mappingPath);
  const fileAdapter = { ...fs, ...(options.fileAdapter ?? {}) };
  const exportWorkbook = options.exportWorkbook ?? defaultExportWorkbook;
  const verifyWorkbook = options.verifyWorkbook ?? readMappingRows;
  const existed = await exists(fileAdapter, absolutePath);
  const beforeHash = existed ? await hashFile(fileAdapter, absolutePath) : null;

  let document;
  if (existed) {
    document = await loadMappingDocument(absolutePath);
  } else {
    const workbook = Workbook.create();
    const sheet = workbook.worksheets.add("城市名称映射表");
    sheet.getRangeByIndexes(0, 0, 1, CANONICAL_HEADERS.length).values = [CANONICAL_HEADERS];
    document = {
      workbook,
      sheet,
      headers: [...CANONICAL_HEADERS],
      rowArrays: [],
      rows: [],
    };
  }

  validateMappingRows(document.rows, cityOrder);
  const additions = [];
  const known = new Map(
    document.rows.map((row) => [normalizeSafeCityName(row.原始城市名), row]),
  );
  const now = new Date().toISOString();
  for (const confirmed of confirmedMappings) {
    const addition = mappingObject(confirmed, now);
    const sourceKey = normalizeSafeCityName(addition.原始城市名);
    if (!sourceKey || !cityOrder.byExact.has(addition.标准城市名)) {
      pause("MAPPING_CONFLICT", "确认映射包含空原始名或未知标准城市", {
        sourceName: addition.原始城市名,
        standardName: addition.标准城市名,
      });
    }
    const existing = known.get(sourceKey);
    if (existing) {
      if (existing.标准城市名 !== addition.标准城市名) {
        pause("MAPPING_CONFLICT", "确认映射与长期映射表冲突", {
          sourceName: addition.原始城市名,
          existingStandardName: existing.标准城市名,
          proposedStandardName: addition.标准城市名,
        });
      }
      continue;
    }
    known.set(sourceKey, addition);
    additions.push(addition);
  }
  validateMappingRows([...document.rows, ...additions], cityOrder);

  if (additions.length === 0) {
    return {
      mappingPath: absolutePath,
      appendedCount: 0,
      skippedCount: confirmedMappings.length,
      beforeSha256: beforeHash,
      afterSha256: beforeHash,
    };
  }

  const headers = [...document.headers];
  for (const header of CANONICAL_HEADERS) {
    if (!headers.includes(header)) headers.push(header);
  }
  document.sheet.getRangeByIndexes(0, 0, 1, headers.length).values = [headers];
  const startRow = document.rowArrays.length + 1;
  const additionMatrix = additions.map((addition) =>
    headers.map((header) => addition[header] ?? null),
  );
  document.sheet
    .getRangeByIndexes(startRow, 0, additionMatrix.length, headers.length)
    .values = additionMatrix;

  await fileAdapter.mkdir(path.dirname(absolutePath), { recursive: true });
  const parsed = path.parse(absolutePath);
  const token = `${process.pid}-${Date.now()}`;
  const tempPath = path.join(parsed.dir, `${parsed.name}.tmp-${token}${parsed.ext || ".xlsx"}`);
  const backupPath = path.join(parsed.dir, `${parsed.name}.backup-${token}${parsed.ext || ".xlsx"}`);
  let originalMoved = false;
  let finalInstalled = false;
  let restoreFailure = null;

  try {
    await exportWorkbook(document.workbook, tempPath);
    const temporaryRows = await readMappingRows(tempPath);
    if (!verifyRowsPreserved(document.rows, temporaryRows, additions)) {
      throw new Error("temporary mapping verification failed");
    }

    if (existed) {
      await fileAdapter.rename(absolutePath, backupPath);
      originalMoved = true;
    }
    await fileAdapter.rename(tempPath, absolutePath);
    finalInstalled = true;

    const finalRows = await verifyWorkbook(absolutePath);
    if (!verifyRowsPreserved(document.rows, finalRows, additions)) {
      throw new Error("final mapping verification failed");
    }
    const afterHash = await hashFile(fileAdapter, absolutePath);
    if (originalMoved) await fileAdapter.rm(backupPath, { force: true });
    return {
      mappingPath: absolutePath,
      appendedCount: additions.length,
      skippedCount: confirmedMappings.length - additions.length,
      beforeSha256: beforeHash,
      afterSha256: afterHash,
    };
  } catch (error) {
    try {
      if (finalInstalled) await fileAdapter.rm(absolutePath, { force: true });
      if (originalMoved) await fileAdapter.rename(backupPath, absolutePath);
    } catch (restoreError) {
      restoreFailure = restoreError;
    }
    await fileAdapter.rm(tempPath, { force: true }).catch(() => {});
    if (!originalMoved) await fileAdapter.rm(backupPath, { force: true }).catch(() => {});
    if (error instanceof PauseError) throw error;
    const code = ["EBUSY", "EPERM", "EACCES"].includes(error?.code)
      ? "MAPPING_FILE_OCCUPIED"
      : "MAPPING_WRITE_FAILED";
    pause(
      code,
      code === "MAPPING_FILE_OCCUPIED"
        ? "城市名称映射表被占用，无法安全替换"
        : "城市名称映射表写入或复核失败",
      {
        path: absolutePath,
        error: error?.message ?? String(error),
        systemCode: error?.code ?? null,
        originalRestored: originalMoved ? restoreFailure === null : true,
        backupPath: restoreFailure ? backupPath : null,
        restoreError: restoreFailure?.message ?? null,
      },
    );
  }
}