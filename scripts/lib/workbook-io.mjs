import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

import { applyHyperlinksToXlsx, readHyperlinksFromXlsx } from "./ooxml-links.mjs";
import { PauseError } from "./pause.mjs";

const FORMULA_ERRORS = new Set([
  "#DIV/0!",
  "#N/A",
  "#NAME?",
  "#NULL!",
  "#NUM!",
  "#REF!",
  "#SPILL!",
  "#VALUE!",
]);

function pause(code, message, evidence = {}) {
  throw new PauseError(code, message, evidence);
}

async function sha256(filePath) {
  return crypto.createHash("sha256").update(await fs.readFile(filePath)).digest("hex");
}

function columnNumber(label) {
  return [...label.toUpperCase()].reduce(
    (total, character) => total * 26 + character.charCodeAt(0) - 64,
    0,
  ) - 1;
}

function columnLabel(index) {
  let value = index + 1;
  let label = "";
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return label;
}

function cellAddress(rowIndex, columnIndex) {
  return `${columnLabel(columnIndex)}${rowIndex + 1}`;
}

function parseCellAddress(address) {
  const match = String(address).replace(/^.*!/u, "").match(/^\$?([A-Z]+)\$?(\d+)$/iu);
  if (!match) throw new Error(`Invalid cell address: ${address}`);
  return { row: Number(match[2]) - 1, column: columnNumber(match[1]) };
}

function parseRange(address) {
  const clean = String(address).replace(/^.*!/u, "");
  const [startText, endText = startText] = clean.split(":");
  const start = parseCellAddress(startText);
  const end = parseCellAddress(endText);
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endColumn: Math.max(start.column, end.column),
  };
}

function rangesIntersect(leftAddress, rightAddress) {
  const left = parseRange(leftAddress);
  const right = parseRange(rightAddress);
  return !(
    left.endRow < right.startRow ||
    right.endRow < left.startRow ||
    left.endColumn < right.startColumn ||
    right.endColumn < left.startColumn
  );
}

function colorHex(color) {
  try {
    return color?.hex ?? null;
  } catch {
    return null;
  }
}

function borderSnapshot(border) {
  if (!border?.style) return null;
  return {
    style: border.style,
    color: colorHex(border.color),
    weight: border.weight ?? null,
  };
}

function styleSnapshot(range) {
  const format = range.format;
  const borders = {
    top: borderSnapshot(format.borders.top),
    bottom: borderSnapshot(format.borders.bottom),
    left: borderSnapshot(format.borders.left),
    right: borderSnapshot(format.borders.right),
    insideHorizontal: borderSnapshot(format.borders.insideHorizontal),
    insideVertical: borderSnapshot(format.borders.insideVertical),
    diagonalUp: borderSnapshot(format.borders.diagonalUp),
    diagonalDown: borderSnapshot(format.borders.diagonalDown),
  };
  return {
    numberFormat: format.numberFormat ?? null,
    fill: colorHex(format.fill?.color),
    font: {
      bold: format.font?.bold ?? null,
      italic: format.font?.italic ?? null,
      size: format.font?.size ?? null,
      name: format.font?.name ?? null,
      color: colorHex(format.font?.color),
    },
    borders,
    wrapText: format.wrapText ?? null,
    horizontalAlignment: format.horizontalAlignment ?? null,
    verticalAlignment: format.verticalAlignment ?? null,
  };
}

function parseThreadInspection(inspection) {
  return String(inspection?.ndjson ?? "")
    .split(/\r?\n/u)
    .filter(Boolean)
    .map((line) => JSON.parse(line))
    .filter((item) => item.kind === "thread");
}

function isFormulaError(value, rawValue) {
  return rawValue?.kind === "Error" || FORMULA_ERRORS.has(String(value ?? ""));
}

function valuesEqual(left, right) {
  if (Object.is(left, right)) return true;
  if (left instanceof Date && right instanceof Date) return left.getTime() === right.getTime();
  if (Array.isArray(left) && Array.isArray(right)) {
    return left.length === right.length && left.every((value, index) => valuesEqual(value, right[index]));
  }
  return false;
}

export async function readWorkbookModel(workbookPath, selectedSheets = []) {
  const absolutePath = path.resolve(workbookPath);
  try {
    await fs.access(absolutePath);
  } catch {
    pause("SOURCE_NOT_FOUND", "源工作簿不存在", { path: absolutePath });
  }
  const inputSha256 = await sha256(absolutePath);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(absolutePath));
  const sheetItems = workbook.worksheets.items;
  const sheetNames = sheetItems.map((sheet) => sheet.name);
  const selected = selectedSheets.length ? selectedSheets : sheetNames;
  const unknownSheets = selected.filter((name) => !sheetNames.includes(name));
  if (unknownSheets.length) {
    pause("MULTIPLE_SHEETS", "选择了不存在的工作表", {
      sheetNames,
      selectedSheets,
      unknownSheets,
    });
  }

  const hyperlinks = await readHyperlinksFromXlsx(absolutePath, selected);
  const hyperlinkMap = new Map(
    hyperlinks.map((item) => [`${item.sheetName}\u0000${item.address}`, item]),
  );
  const threadMap = new Map();
  for (const sheetName of selected) {
    const threadInspection = await workbook.inspect({
      kind: "thread",
      sheetId: sheetName,
      maxChars: 1000000,
      options: { maxResults: 10000 },
    });
    for (const item of parseThreadInspection(threadInspection)) {
      threadMap.set(`${item.sheet}\u0000${item.target}`, item.text);
    }
  }

  const summaries = sheetItems.map((sheet) => ({
    name: sheet.name,
    usedRange: sheet.getUsedRange()?.address ?? null,
  }));
  const sheets = [];
  const formulaEvents = [];

  for (const sheetName of selected) {
    const sheet = workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRange();
    if (!usedRange) {
      sheets.push({
        name: sheetName,
        usedRange: null,
        dataRange: null,
        matrix: [],
        formulas: [],
        cells: {},
        rowRecords: [],
        headerMetadata: [],
        mergedRanges: [],
        sheetSha256: crypto.createHash("sha256").update("[]").digest("hex"),
      });
      continue;
    }
    const dataRange = usedRange.getCell(0, 0).getCurrentRegion();
    const values = dataRange.values;
    const formulas = dataRange.formulas;
    const rawValues = dataRange.rawValues;
    const formulaInfos = dataRange.formulaInfos;
    const mergedRanges = typeof sheet.__getMergedCells === "function"
      ? sheet.__getMergedCells().map((merge) => `${merge.startAddress}:${merge.endAddress}`)
      : [];
    const unsafeMerges = mergedRanges.filter((merge) => rangesIntersect(merge, dataRange.address));
    if (unsafeMerges.length) {
      pause("UNSAFE_MERGED_CELLS", "合并单元格与数据区域相交", {
        sheetName,
        dataRange: dataRange.address,
        mergedRanges: unsafeMerges,
      });
    }

    const cells = {};
    const metadataMatrix = [];
    for (let rowOffset = 0; rowOffset < dataRange.rowCount; rowOffset += 1) {
      const metadataRow = [];
      for (let columnOffset = 0; columnOffset < dataRange.columnCount; columnOffset += 1) {
        const absoluteRow = dataRange.rowIndex + rowOffset;
        const absoluteColumn = dataRange.columnIndex + columnOffset;
        const address = cellAddress(absoluteRow, absoluteColumn);
        const formula = formulas[rowOffset]?.[columnOffset] || null;
        const value = values[rowOffset]?.[columnOffset] ?? null;
        const rawValue = rawValues[rowOffset]?.[columnOffset] ?? null;
        if (formula && isFormulaError(value, rawValue)) {
          pause("FORMULA_ERROR", "公式单元格包含错误值", {
            sheetName,
            cell: address,
            formula,
            value,
          });
        }
        if (formula && (value === null || value === undefined)) {
          pause("FORMULA_VALUE_UNAVAILABLE", "公式缺少可用的当前计算值", {
            sheetName,
            cell: address,
            formula,
          });
        }
        if (formula) {
          formulaEvents.push({
            sheetName,
            cell: address,
            sourceRow: absoluteRow + 1,
            sourceColumn: absoluteColumn,
            formula,
            currentValue: value,
          });
        }
        const metadata = {
          address,
          sourceRow: absoluteRow + 1,
          sourceColumn: absoluteColumn,
          style: styleSnapshot(sheet.getCell(absoluteRow, absoluteColumn)),
          comment: threadMap.get(`${sheetName}\u0000${address}`) ?? null,
          hyperlink: hyperlinkMap.get(`${sheetName}\u0000${address}`) ?? null,
          formula,
          formulaInfo: formulaInfos[rowOffset]?.[columnOffset] ?? null,
        };
        metadataRow.push(metadata);
        cells[address] = { value, ...metadata };
      }
      metadataMatrix.push(metadataRow);
    }

    const rowRecords = values.slice(1).map((rowValues, index) => ({
      values: [...rowValues],
      sourceRow: dataRange.rowIndex + index + 2,
      cellMetadata: metadataMatrix[index + 1].map((item) => ({ ...item })),
      sourceCells: metadataMatrix[index + 1].map((item) => item.address),
    }));
    const sheetSha256 = crypto
      .createHash("sha256")
      .update(JSON.stringify({ values, formulas }))
      .digest("hex");
    sheets.push({
      name: sheetName,
      usedRange: usedRange.address,
      dataRange: dataRange.address,
      dataStartRow: dataRange.rowIndex,
      dataStartColumn: dataRange.columnIndex,
      matrix: values,
      formulas,
      cells,
      rowRecords,
      headerMetadata: metadataMatrix[0]?.map((item) => ({ ...item })) ?? [],
      mergedRanges,
      unsafeMerges,
      hyperlinks: hyperlinks.filter((item) => item.sheetName === sheetName),
      sheetSha256,
    });
  }

  return {
    path: absolutePath,
    sha256: inputSha256,
    sheetNames,
    selectedSheets: selected,
    summaries,
    sheets,
    formulaEvents,
  };
}

function formatConfig(style) {
  if (!style) return null;
  const config = {};
  if (style.numberFormat) config.numberFormat = style.numberFormat;
  if (style.fill) config.fill = style.fill;
  const font = Object.fromEntries(
    Object.entries(style.font ?? {}).filter(([, value]) => value !== null && value !== undefined),
  );
  if (Object.keys(font).length) config.font = font;
  const borders = {};
  for (const [edge, value] of Object.entries(style.borders ?? {})) {
    if (value?.style) {
      borders[edge] = Object.fromEntries(
        Object.entries(value).filter(([, item]) => item !== null && item !== undefined),
      );
    }
  }
  if (Object.keys(borders).length) config.borders = borders;
  if (style.wrapText !== null && style.wrapText !== undefined) config.wrapText = style.wrapText;
  if (style.horizontalAlignment) config.horizontalAlignment = style.horizontalAlignment;
  if (style.verticalAlignment) config.verticalAlignment = style.verticalAlignment;
  return config;
}

function validateSheetName(name) {
  if (!name || name.length > 31 || /[\[\]:*?/\\]/u.test(name)) {
    pause("INVALID_SHEET_NAME", "输出工作表名称无效或超过 31 个字符", { name });
  }
}

function optionalSheetName(workbook, requested) {
  let name = requested;
  let counter = 2;
  while (workbook.worksheets.items.some((sheet) => sheet.name === name)) {
    name = `${requested.slice(0, 28)}_${counter}`;
    counter += 1;
  }
  validateSheetName(name);
  return name;
}

function addMatrixSheet(workbook, name, matrix) {
  validateSheetName(name);
  const sheet = workbook.worksheets.add(name);
  if (matrix.length && matrix[0]?.length) {
    sheet.getRangeByIndexes(0, 0, matrix.length, matrix[0].length).values = matrix;
  }
  return sheet;
}

async function outputExists(outputPath) {
  try {
    await fs.access(outputPath);
    return true;
  } catch {
    return false;
  }
}

export async function writeResultWorkbook(result, outputPath) {
  const absoluteOutputPath = path.resolve(outputPath);
  if (await outputExists(absoluteOutputPath)) {
    pause("OUTPUT_FILE_OCCUPIED", "输出文件已存在或被占用", { path: absoluteOutputPath });
  }
  if (result?.audit?.passed !== true) {
    pause("REVERSE_VERIFICATION_FAILED", "核验未通过，禁止创建正式输出", {
      audit: result?.audit ?? null,
    });
  }
  const outputs = result.outputs ?? [];
  if (!outputs.length) {
    pause("REVERSE_VERIFICATION_FAILED", "没有可写入且已核验的对齐结果", {});
  }

  await fs.mkdir(path.dirname(absoluteOutputPath), { recursive: true });
  const parsed = path.parse(absoluteOutputPath);
  const tempPath = path.join(
    parsed.dir,
    `${parsed.name}.tmp-${process.pid}-${Date.now()}${parsed.ext || ".xlsx"}`,
  );
  const workbook = Workbook.create();
  const hyperlinks = [];
  const outputSheetNames = [];
  let commentsInitialized = false;

  try {
    for (const [outputIndex, output] of outputs.entries()) {
      const requestedName = output.outputSheetName ?? (
        outputs.length === 1 ? "对齐结果" : `对齐结果_${output.sheetName}`
      );
      validateSheetName(requestedName);
      if (workbook.worksheets.items.some((sheet) => sheet.name === requestedName)) {
        pause("INVALID_SHEET_NAME", "多个输出工作表名称重复", { name: requestedName });
      }
      const sheet = workbook.worksheets.add(requestedName);
      outputSheetNames.push(requestedName);
      const matrix = [[...output.headers], ...output.rows.map((row) => [...row])];
      if (matrix.length && matrix[0].length) {
        sheet.getRangeByIndexes(0, 0, matrix.length, matrix[0].length).values = matrix;
      }

      const metadataRows = [output.headerMetadata ?? [], ...(output.records ?? []).map((record) => record.cellMetadata ?? [])];
      for (let rowIndex = 0; rowIndex < metadataRows.length; rowIndex += 1) {
        for (let columnIndex = 0; columnIndex < metadataRows[rowIndex].length; columnIndex += 1) {
          const metadata = metadataRows[rowIndex][columnIndex];
          if (!metadata) continue;
          const cell = sheet.getCell(rowIndex, columnIndex);
          const format = formatConfig(metadata.style);
          if (format && Object.keys(format).length) cell.format = format;
          if (metadata.comment) {
            if (!commentsInitialized) {
              workbook.comments.setSelf({ displayName: "Codex" });
              commentsInitialized = true;
            }
            workbook.comments.addThread({ cell }, metadata.comment);
          }
          if (metadata.hyperlink?.target || metadata.hyperlink?.location) {
            hyperlinks.push({
              sheetName: requestedName,
              address: cellAddress(rowIndex, columnIndex),
              target: metadata.hyperlink.target,
              location: metadata.hyperlink.location,
              tooltip: metadata.hyperlink.tooltip,
              display: metadata.hyperlink.display,
            });
          }
        }
      }
      output.outputSheetName = requestedName;
      output.outputIndex = outputIndex;
    }

    const auditRows = result.audit.auditRows ?? [];
    addMatrixSheet(workbook, optionalSheetName(workbook, "核验结果"), [
      ["分类", "项目", "状态", "值", "明细"],
      ...auditRows.map((item) => [item.section, item.item, item.status, item.value, item.details]),
    ]);

    const mappingEvents = outputs.flatMap((output) => output.mappingEvents ?? []);
    if (mappingEvents.length) {
      addMatrixSheet(workbook, optionalSheetName(workbook, "城市匹配明细"), [
        ["源工作表", "源行", "原始城市名", "标准城市名", "匹配方法"],
        ...mappingEvents.map((item) => [item.sheetName ?? null, item.sourceRow, item.rawCity, item.standardName, item.method]),
      ]);
    }
    const missingKeys = outputs.flatMap((output) => output.missingKeys ?? []);
    if (missingKeys.length) {
      addMatrixSheet(workbook, optionalSheetName(workbook, "缺失键"), [
        ["城市", "年份"],
        ...missingKeys.map((item) => [item.city, item.year]),
      ]);
    }
    const outOfRange = outputs.flatMap((output) => output.outOfRangeRows ?? output.outOfRangeRecords ?? []);
    if (outOfRange.length) {
      addMatrixSheet(workbook, optionalSheetName(workbook, "范围外数据"), [
        ["源工作表", "源行", "年份", "原始内容"],
        ...outOfRange.map((item) => [item.sourceSheet ?? null, item.sourceRow, item.year, JSON.stringify(item.values ?? item.value)]),
      ]);
    }
    const formulaEvents = result.formulaEvents ?? [];
    if (formulaEvents.length) {
      addMatrixSheet(workbook, optionalSheetName(workbook, "公式处理"), [
        ["工作表", "单元格", "公式", "当前计算值"],
        ...formulaEvents.map((item) => [item.sheetName, item.cell, item.formula, item.currentValue]),
      ]);
    }

    const exported = await SpreadsheetFile.exportXlsx(workbook);
    await exported.save(tempPath);
    await applyHyperlinksToXlsx(tempPath, hyperlinks);

    const reimported = await SpreadsheetFile.importXlsx(await FileBlob.load(tempPath));
    for (const output of outputs) {
      const sheet = reimported.worksheets.getItem(output.outputSheetName);
      const expected = [[...output.headers], ...output.rows.map((row) => [...row])];
      const actualRange = sheet.getRangeByIndexes(0, 0, expected.length, expected[0].length);
      if (!valuesEqual(actualRange.values, expected)) {
        throw new Error(`Temporary output value verification failed: ${output.outputSheetName}`);
      }
      if (actualRange.formulas.flat().some((formula) => Boolean(formula))) {
        throw new Error(`Temporary output still contains formulas: ${output.outputSheetName}`);
      }
    }
    const writtenLinks = await readHyperlinksFromXlsx(tempPath, outputSheetNames);
    if (writtenLinks.length !== hyperlinks.length) {
      throw new Error("Temporary output hyperlink verification failed");
    }

    await fs.copyFile(tempPath, absoluteOutputPath, fs.constants.COPYFILE_EXCL);
    await fs.rm(tempPath, { force: true });
    return {
      outputPath: absoluteOutputPath,
      outputSha256: await sha256(absoluteOutputPath),
      outputSheetNames,
      hyperlinkCount: hyperlinks.length,
    };
  } catch (error) {
    await fs.rm(tempPath, { force: true }).catch(() => {});
    if (error instanceof PauseError) throw error;
    if (["EEXIST", "EBUSY", "EPERM", "EACCES"].includes(error?.code)) {
      pause("OUTPUT_FILE_OCCUPIED", "输出文件已存在或被占用", {
        path: absoluteOutputPath,
        systemCode: error.code,
      });
    }
    pause("OUTPUT_WRITE_FAILED", "输出工作簿写入或复核失败", {
      path: absoluteOutputPath,
      error: error?.message ?? String(error),
    });
  }
}