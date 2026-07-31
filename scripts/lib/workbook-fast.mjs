import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

import { PauseError } from "./pause.mjs";

function pause(code, message, evidence = {}) {
  throw new PauseError(code, message, evidence);
}

function cellAddress(rowIndex, columnIndex) {
  let number = columnIndex + 1;
  let label = "";
  while (number > 0) {
    const remainder = (number - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    number = Math.floor((number - 1) / 26);
  }
  return `${label}${rowIndex + 1}`;
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

function formulaError(value, rawValue) {
  const text = String(value ?? rawValue ?? "").trim();
  return /^#(?:DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!|ERROR!)$/iu.test(text);
}

async function sha256(filePath) {
  const bytes = await fs.readFile(filePath);
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

export async function readWorkbookFast(workbookPath, selectedSheets = []) {
  const absolutePath = path.resolve(workbookPath);
  const inputSha256 = await sha256(absolutePath);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(absolutePath));
  const sheetNames = workbook.worksheets.items.map((sheet) => sheet.name);
  const selected = selectedSheets.length ? selectedSheets : sheetNames;
  const unknownSheets = selected.filter((name) => !sheetNames.includes(name));
  if (unknownSheets.length) {
    pause("MULTIPLE_SHEETS", "选择了不存在的工作表", {
      sheetNames,
      selectedSheets,
      unknownSheets,
    });
  }

  const sheets = [];
  const formulaEvents = [];
  for (const sheetName of selected) {
    const sheet = workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRange();
    if (!usedRange) {
      sheets.push({
        name: sheetName,
        matrix: [],
        rowRecords: [],
        headerMetadata: [],
        unsafeMerges: [],
        dataStartRow: 0,
        dataStartColumn: 0,
      });
      continue;
    }
    const dataRange = usedRange.getCell(0, 0).getCurrentRegion();
    const values = dataRange.values;
    const formulas = dataRange.formulas;
    const rawValues = dataRange.rawValues;
    const mergedRanges = typeof sheet.__getMergedCells === "function"
      ? sheet.__getMergedCells().map((merge) => `${merge.startAddress}:${merge.endAddress}`)
      : [];
    if (mergedRanges.length) {
      pause("UNSAFE_MERGED_CELLS", "数据区域存在合并单元格，快速模式不能安全处理", {
        sheetName,
        dataRange: dataRange.address,
        mergedRanges,
      });
    }

    const metadataMatrix = values.map((row, rowOffset) => row.map((_, columnOffset) => {
      const absoluteRow = dataRange.rowIndex + rowOffset;
      const absoluteColumn = dataRange.columnIndex + columnOffset;
      const formula = formulas[rowOffset]?.[columnOffset] || null;
      const value = values[rowOffset]?.[columnOffset] ?? null;
      const rawValue = rawValues[rowOffset]?.[columnOffset] ?? null;
      const address = cellAddress(absoluteRow, absoluteColumn);
      if (formula && formulaError(value, rawValue)) {
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
      return {
        address,
        sourceRow: absoluteRow + 1,
        sourceColumn: absoluteColumn,
        formula,
      };
    }));

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
      rowRecords,
      headerMetadata: metadataMatrix[0]?.map((item) => ({ ...item })) ?? [],
      mergedRanges,
      unsafeMerges: [],
      sheetSha256,
    });
  }

  return {
    path: absolutePath,
    sha256: inputSha256,
    sheetNames,
    selectedSheets: selected,
    sheets,
    formulaEvents,
  };
}

export async function writeCleanResultWorkbook(outputs, outputPath) {
  const absoluteOutputPath = path.resolve(outputPath);
  if (!Array.isArray(outputs) || outputs.length === 0) {
    pause("REVERSE_VERIFICATION_FAILED", "没有可写入且已核验的对齐结果", {});
  }
  if (await fs.access(absoluteOutputPath).then(() => true, () => false)) {
    pause("OUTPUT_FILE_OCCUPIED", "输出文件已存在或被占用", {
      path: absoluteOutputPath,
    });
  }

  await fs.mkdir(path.dirname(absoluteOutputPath), { recursive: true });
  const parsed = path.parse(absoluteOutputPath);
  const tempPath = path.join(
    parsed.dir,
    `${parsed.name}.tmp-${crypto.randomUUID()}-${Date.now()}${parsed.ext || ".xlsx"}`,
  );
  const workbook = Workbook.create();
  try {
    for (const [index, output] of outputs.entries()) {
      const requestedName = output.outputSheetName ?? (
        outputs.length === 1 ? "对齐结果" : `对齐结果${index + 1}`
      );
      const sheet = workbook.worksheets.add(requestedName);
      const matrix = [[...output.headers], ...output.rows.map((row) => [...row])];
      sheet.getRangeByIndexes(0, 0, matrix.length, matrix[0].length).values = matrix;
      sheet.getRangeByIndexes(0, 0, 1, matrix[0].length).format.font = { bold: true };
    }

    const exported = await SpreadsheetFile.exportXlsx(workbook);
    await exported.save(tempPath);
    const reimported = await SpreadsheetFile.importXlsx(await FileBlob.load(tempPath));
    for (const output of outputs) {
      const sheet = reimported.worksheets.getItem(output.outputSheetName);
      const expected = [[...output.headers], ...output.rows.map((row) => [...row])];
      const actualRange = sheet.getRangeByIndexes(0, 0, expected.length, expected[0].length);
      if (!valuesEqual(actualRange.values, expected)) {
        throw new Error(`Temporary output value verification failed: ${output.outputSheetName}`);
      }
      if (actualRange.formulas.flat().some(Boolean)) {
        throw new Error(`Temporary output still contains formulas: ${output.outputSheetName}`);
      }
    }
    await fs.copyFile(tempPath, absoluteOutputPath, fs.constants.COPYFILE_EXCL);
    await fs.rm(tempPath, { force: true });
    return {
      outputPath: absoluteOutputPath,
      outputSha256: await sha256(absoluteOutputPath),
      outputSheetNames: outputs.map((output) => output.outputSheetName),
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
