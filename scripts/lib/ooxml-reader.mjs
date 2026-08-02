import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import { PauseError } from "./pause.mjs";

const REQUIRED_PACKAGE_PARTS = [
  "[Content_Types].xml",
  "xl/workbook.xml",
  "xl/_rels/workbook.xml.rels",
];

export function decodeXmlText(text = "") {
  return String(text)
    .replace(/&#x([0-9a-f]+);/gi, (_, value) =>
      String.fromCodePoint(Number.parseInt(value, 16)),
    )
    .replace(/&#([0-9]+);/g, (_, value) =>
      String.fromCodePoint(Number.parseInt(value, 10)),
    )
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

export function parseXmlAttributes(tag) {
  const attributes = {};
  const expression = /(?:^|\s)([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  for (const match of String(tag).matchAll(expression)) {
    attributes[match[1]] = decodeXmlText(match[2] ?? match[3] ?? "");
  }
  return attributes;
}

export function resolvePartPath(sourcePart, target) {
  const normalizedTarget = String(target).replace(/\\/g, "/");
  const combined = normalizedTarget.startsWith("/")
    ? normalizedTarget.slice(1)
    : path.posix.join(path.posix.dirname(sourcePart), normalizedTarget);
  const normalized = path.posix.normalize(combined);
  if (
    normalized === ".." ||
    normalized.startsWith("../") ||
    path.posix.isAbsolute(normalized)
  ) {
    throw new PauseError(
      "INVALID_WORKBOOK_PACKAGE",
      "工作簿关系目标超出 OOXML 包范围。",
      { sourcePart, target, partPath: normalized },
    );
  }
  return normalized;
}

export function parseWorkbookSheets(workbookXml) {
  return [...String(workbookXml).matchAll(/<(?:[\w.-]+:)?sheet\b[^>]*\/?\s*>/gi)]
    .map((match) => parseXmlAttributes(match[0]))
    .map((attributes) => ({
      name: attributes.name,
      state: attributes.state ?? "visible",
      relationshipId: attributes["r:id"] ?? attributes.id,
    }))
    .filter((sheet) => sheet.name);
}

export function parseRelationships(relationshipsXml) {
  return [
    ...String(relationshipsXml).matchAll(
      /<(?:[\w.-]+:)?Relationship\b[^>]*\/?\s*>/gi,
    ),
  ]
    .map((match) => parseXmlAttributes(match[0]))
    .map((attributes) => ({
      id: attributes.Id ?? attributes.id,
      type: attributes.Type ?? attributes.type,
      target: attributes.Target ?? attributes.target,
      targetMode: attributes.TargetMode ?? attributes.targetMode ?? null,
    }))
    .filter((relationship) => relationship.id);
}

function columnIndexFromLetters(letters) {
  let index = 0;
  for (const character of String(letters).toUpperCase()) {
    index = index * 26 + character.charCodeAt(0) - 64;
  }
  return index;
}

export function parseDimension(sheetXml) {
  const match = String(sheetXml).match(/<(?:[\w.-]+:)?dimension\b[^>]*\/?\s*>/i);
  if (!match) {
    return {
      dimension: null,
      usedRange: null,
      totalRowCount: 0,
      rowCount: 0,
      columnCount: 0,
    };
  }
  const reference = parseXmlAttributes(match[0]).ref ?? null;
  const endReference = reference?.split(":").at(-1)?.replace(/\$/g, "") ?? "";
  const cellMatch = endReference.match(/^([A-Za-z]+)(\d+)$/);
  const totalRowCount = cellMatch ? Number.parseInt(cellMatch[2], 10) : 0;
  const columnCount = cellMatch ? columnIndexFromLetters(cellMatch[1]) : 0;
  return {
    dimension: reference,
    usedRange: reference,
    totalRowCount,
    rowCount: Math.max(0, totalRowCount - 1),
    columnCount,
  };
}

function invalidPackage(message, evidence = {}) {
  throw new PauseError("INVALID_WORKBOOK_PACKAGE", message, evidence);
}

async function requireZipText(zip, partPath) {
  const part = zip.file(partPath);
  if (!part) invalidPackage("工作簿缺少必需的 OOXML 部件。", { partPath });
  return part.async("string");
}

export async function inspectWorkbookPackage(workbookPath) {
  const bytes = await fs.readFile(workbookPath);
  let zip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch (error) {
    invalidPackage("文件不是可读取的 XLSX/OOXML 包。", {
      path: workbookPath,
      cause: error.message,
    });
  }
  for (const partPath of REQUIRED_PACKAGE_PARTS) {
    if (!zip.file(partPath)) {
      invalidPackage("工作簿缺少必需的 OOXML 部件。", { partPath });
    }
  }

  const workbookXml = await requireZipText(zip, "xl/workbook.xml");
  const relationshipsXml = await requireZipText(
    zip,
    "xl/_rels/workbook.xml.rels",
  );
  const sheetEntries = parseWorkbookSheets(workbookXml);
  const relationships = parseRelationships(relationshipsXml);
  const relationshipsById = new Map(
    relationships.map((relationship) => [relationship.id, relationship]),
  );
  const usedRelationshipIds = new Set();
  const sheets = [];

  for (const sheet of sheetEntries) {
    const relationship = relationshipsById.get(sheet.relationshipId);
    if (!relationship?.target) {
      invalidPackage("工作表缺少实际引用的关系定义。", {
        sheetName: sheet.name,
        relationshipId: sheet.relationshipId,
        target: relationship?.target ?? null,
        partPath: null,
      });
    }
    usedRelationshipIds.add(relationship.id);
    const partPath = resolvePartPath("xl/workbook.xml", relationship.target);
    const part = zip.file(partPath);
    if (!part) {
      invalidPackage("工作表实际引用的 OOXML 部件缺失。", {
        sheetName: sheet.name,
        relationshipId: relationship.id,
        target: relationship.target,
        partPath,
      });
    }
    const dimension = parseDimension(await part.async("string"));
    sheets.push({ ...sheet, partPath, ...dimension });
  }

  const orphanRelationships = relationships
    .filter(
      (relationship) =>
        relationship.type?.endsWith("/worksheet") &&
        !usedRelationshipIds.has(relationship.id),
    )
    .map((relationship) => ({
      ...relationship,
      partPath: relationship.target
        ? resolvePartPath("xl/workbook.xml", relationship.target)
        : null,
    }));

  return {
    path: workbookPath,
    sha256: crypto.createHash("sha256").update(bytes).digest("hex"),
    sheets,
    orphanRelationships,
  };
}

const FORMULA_ERROR_VALUES = new Set([
  "#DIV/0!",
  "#N/A",
  "#NAME?",
  "#NULL!",
  "#NUM!",
  "#REF!",
  "#SPILL!",
  "#VALUE!",
  "#CALC!",
  "#ERROR!",
]);

const BUILTIN_DATE_FORMAT_IDS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22,
  27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
  45, 46, 47,
  50, 51, 52, 53, 54, 55, 56, 57, 58,
]);

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

function parseCellReference(reference) {
  const match = String(reference).replace(/\$/g, "").match(/^([A-Za-z]+)(\d+)$/);
  if (!match) return null;
  return {
    column: columnIndexFromLetters(match[1]) - 1,
    row: Number.parseInt(match[2], 10) - 1,
  };
}

function parseRangeReference(reference) {
  const [startText, endText = startText] = String(reference).split(":");
  const start = parseCellReference(startText);
  const end = parseCellReference(endText);
  if (!start || !end) return null;
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endColumn: Math.max(start.column, end.column),
  };
}

function normalizedHeader(value) {
  return String(value ?? "").normalize("NFKC").replace(/\s+/gu, "");
}

function xmlTextDescendants(xml) {
  return [
    ...String(xml).matchAll(
      /<(?:[\w.-]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?t>/gi,
    ),
  ].map((match) => decodeXmlText(match[1])).join("");
}

export function parseSharedStrings(xml) {
  if (!xml) return [];
  return [
    ...String(xml).matchAll(
      /<(?:[\w.-]+:)?si\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?si>/gi,
    ),
  ].map((match) => xmlTextDescendants(match[1]));
}

export function parseWorkbookDateSystem(workbookXml) {
  const match = String(workbookXml).match(
    /<(?:[\w.-]+:)?workbookPr\b[^>]*\/?\s*>/i,
  );
  if (!match) return false;
  const value = parseXmlAttributes(match[0]).date1904;
  return value === "1" || String(value).toLowerCase() === "true";
}

function looksLikeDateFormat(formatCode) {
  const cleaned = String(formatCode)
    .replace(/"[^"]*"/g, "")
    .replace(/\\./g, "")
    .replace(/\[(?!h\]|m\]|s\])[^\]]*\]/gi, "")
    .replace(/_.|\*./g, "");
  return /(?:^|[^a-z])[ymdhis]+(?:[^a-z]|$)/i.test(cleaned) ||
    /\[(?:h|m|s)\]/i.test(cleaned);
}

export function parseDateStyleIndexes(stylesXml) {
  if (!stylesXml) return new Set();
  const customDateFormatIds = new Set();
  for (const match of String(stylesXml).matchAll(
    /<(?:[\w.-]+:)?numFmt\b[^>]*\/?\s*>/gi,
  )) {
    const attributes = parseXmlAttributes(match[0]);
    const formatId = Number.parseInt(attributes.numFmtId, 10);
    if (Number.isInteger(formatId) && looksLikeDateFormat(attributes.formatCode)) {
      customDateFormatIds.add(formatId);
    }
  }
  const cellXfsMatch = String(stylesXml).match(
    /<(?:[\w.-]+:)?cellXfs\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?cellXfs>/i,
  );
  if (!cellXfsMatch) return new Set();
  const dateStyleIndexes = new Set();
  const xfs = [
    ...cellXfsMatch[1].matchAll(/<(?:[\w.-]+:)?xf\b[^>]*\/?\s*>/gi),
  ];
  xfs.forEach((match, index) => {
    const formatId = Number.parseInt(
      parseXmlAttributes(match[0]).numFmtId ?? "0",
      10,
    );
    if (
      BUILTIN_DATE_FORMAT_IDS.has(formatId) ||
      customDateFormatIds.has(formatId)
    ) {
      dateStyleIndexes.add(index);
    }
  });
  return dateStyleIndexes;
}

function excelSerialToDate(value, date1904) {
  const epoch = date1904
    ? Date.UTC(1904, 0, 1)
    : Date.UTC(1899, 11, 30);
  return new Date(epoch + Number(value) * 86400000);
}

function extractElementText(xml, localName) {
  const expression = new RegExp(
    "<(?:[\\w.-]+:)?" + localName +
      "\\b[^>]*>([\\s\\S]*?)<\\/(?:[\\w.-]+:)?" +
      localName + ">",
    "i",
  );
  return String(xml).match(expression)?.[1] ?? null;
}

function hasElement(xml, localName) {
  const expression = new RegExp(
    "<(?:[\\w.-]+:)?" + localName + "\\b",
    "i",
  );
  return expression.test(String(xml));
}

function decodeCellValue(cell, context) {
  const { attributes, body, address, column, row } = cell;
  const type = attributes.t ?? null;
  const styleIndex = Number.parseInt(attributes.s ?? "0", 10);
  const formulaPresent = hasElement(body, "f");
  const formulaText = formulaPresent
    ? decodeXmlText(extractElementText(body, "f") ?? "")
    : null;
  const rawValue = extractElementText(body, "v");

  if (formulaPresent && rawValue === null) {
    throw new PauseError(
      "FORMULA_VALUE_UNAVAILABLE",
      "公式缺少可用的当前计算值。",
      {
        sheetName: context.sheetName,
        cell: address,
        formula: formulaText,
      },
    );
  }

  let value = null;
  if (type === "inlineStr") {
    value = xmlTextDescendants(extractElementText(body, "is") ?? body);
  } else if (rawValue !== null) {
    const decodedRawValue = decodeXmlText(rawValue);
    if (type === "s") {
      value = context.sharedStrings[Number.parseInt(decodedRawValue, 10)] ?? "";
    } else if (type === "str") {
      value = decodedRawValue;
    } else if (type === "b") {
      value = decodedRawValue === "1" || decodedRawValue.toLowerCase() === "true";
    } else if (type === "e") {
      value = decodedRawValue;
    } else {
      const numericValue = Number(decodedRawValue);
      value = Number.isNaN(numericValue) ? decodedRawValue : numericValue;
      if (
        typeof value === "number" &&
        context.dateStyleIndexes.has(styleIndex)
      ) {
        value = excelSerialToDate(value, context.date1904);
      }
    }
  }

  if (
    formulaPresent &&
    (type === "e" || FORMULA_ERROR_VALUES.has(String(value).trim()))
  ) {
    throw new PauseError("FORMULA_ERROR", "公式单元格包含错误值。", {
      sheetName: context.sheetName,
      cell: address,
      formula: formulaText,
      value,
    });
  }

  const formula = formulaPresent ? formulaText : null;
  return {
    value,
    formula,
    metadata: {
      address,
      sourceRow: row + 1,
      sourceColumn: column,
      formula,
    },
  };
}

function parseWorksheetCells(sheetXml, context) {
  const rows = new Map();
  const rowExpression =
    /<(?:[\w.-]+:)?row\b([^>]*)>([\s\S]*?)<\/(?:[\w.-]+:)?row>/gi;
  for (const rowMatch of String(sheetXml).matchAll(rowExpression)) {
    const rowAttributes = parseXmlAttributes(rowMatch[1]);
    const rowNumber = Number.parseInt(rowAttributes.r, 10);
    const rowIndex = Number.isInteger(rowNumber) ? rowNumber - 1 : rows.size;
    const cells = new Map();
    const cellExpression =
      /<(?:[\w.-]+:)?c\b([^>]*?)(?:\/\s*>|>([\s\S]*?)<\/(?:[\w.-]+:)?c>)/gi;
    for (const cellMatch of rowMatch[2].matchAll(cellExpression)) {
      const attributes = parseXmlAttributes(cellMatch[1]);
      const parsedReference = parseCellReference(attributes.r);
      const column = parsedReference?.column ?? cells.size;
      const row = parsedReference?.row ?? rowIndex;
      const address = attributes.r ?? (columnLabel(column) + String(row + 1));
      cells.set(column, decodeCellValue({
        attributes,
        body: cellMatch[2] ?? "",
        address,
        column,
        row,
      }, context));
    }
    rows.set(rowIndex, cells);
  }
  return rows;
}

function parseMergedRanges(sheetXml) {
  return [
    ...String(sheetXml).matchAll(
      /<(?:[\w.-]+:)?mergeCell\b[^>]*\/?\s*>/gi,
    ),
  ]
    .map((match) => parseXmlAttributes(match[0]).ref)
    .filter(Boolean);
}

function selectProjectedColumns(headerCells, options) {
  const headerEntries = [...headerCells.entries()].map(([column, cell]) => ({
    column,
    value: cell.value,
    normalized: normalizedHeader(cell.value),
  }));
  if (!options.projectHeaders?.length && !options.projectionRoles) {
    if (!headerEntries.length) return [];
    const start = Math.min(...headerEntries.map((entry) => entry.column));
    const end = Math.max(...headerEntries.map((entry) => entry.column));
    return Array.from({ length: end - start + 1 }, (_, index) => start + index);
  }

  function uniqueExactHeader(requestedHeader) {
    const normalized = normalizedHeader(requestedHeader);
    const matches = headerEntries.filter((entry) => entry.normalized === normalized);
    if (!matches.length) {
      throw new PauseError("INDICATOR_NOT_FOUND", "指定表头在源表中不存在。", {
        requestedHeader,
        availableHeaders: headerEntries.map((entry) => entry.value),
      });
    }
    if (matches.length > 1) {
      throw new PauseError("AMBIGUOUS_INDICATOR", "指定表头对应多个源列。", {
        requestedHeader,
        columns: matches.map((entry) => columnLabel(entry.column)),
      });
    }
    return matches[0].column;
  }

  if (!options.projectionRoles) {
    return options.projectHeaders.map(uniqueExactHeader);
  }

  function uniqueRole(role, aliases, missingCode, ambiguousCode) {
    const normalizedAliases = new Set(aliases.map(normalizedHeader));
    const matches = headerEntries.filter((entry) =>
      normalizedAliases.has(entry.normalized),
    );
    if (!matches.length) {
      throw new PauseError(missingCode, "未找到必需的键列。", {
        role,
        aliases,
        availableHeaders: headerEntries.map((entry) => entry.value),
      });
    }
    if (matches.length > 1) {
      throw new PauseError(ambiguousCode, "必需键列不能唯一识别。", {
        role,
        aliases,
        columns: matches.map((entry) => columnLabel(entry.column)),
      });
    }
    return matches[0];
  }

  const cityAliases = options.projectionRoles.cityHeaders ?? [];
  const yearAliases = options.projectionRoles.yearHeaders ?? [];
  const city = uniqueRole(
    "city",
    cityAliases,
    "AMBIGUOUS_CITY_COLUMN",
    "AMBIGUOUS_CITY_COLUMN",
  );
  const year = uniqueRole(
    "year",
    yearAliases,
    "AMBIGUOUS_YEAR_COLUMN",
    "AMBIGUOUS_YEAR_COLUMN",
  );
  const roleAliases = new Set(
    [...cityAliases, ...yearAliases].map(normalizedHeader),
  );
  const indicators = (options.projectHeaders ?? [])
    .filter((header) => !roleAliases.has(normalizedHeader(header)))
    .map(uniqueExactHeader);
  return [city.column, year.column, ...indicators];
}

function mergeIntersectsColumns(reference, startRow, endRow, columns) {
  const range = parseRangeReference(reference);
  if (!range) return false;
  if (range.endRow < startRow || range.startRow > endRow) return false;
  return columns.some(
    (column) => column >= range.startColumn && column <= range.endColumn,
  );
}

export async function readWorksheetTable(workbookPath, sheetName, options = {}) {
  const absolutePath = path.resolve(workbookPath);
  const packageInfo = await inspectWorkbookPackage(absolutePath);
  const sheetInfo = packageInfo.sheets.find((sheet) => sheet.name === sheetName);
  if (!sheetInfo) {
    throw new PauseError("MULTIPLE_SHEETS", "选择了不存在的工作表。", {
      sheetName,
      sheetNames: packageInfo.sheets.map((sheet) => sheet.name),
    });
  }

  const bytes = await fs.readFile(absolutePath);
  const zip = await JSZip.loadAsync(bytes);
  const workbookXml = await requireZipText(zip, "xl/workbook.xml");
  const sheetXml = await requireZipText(zip, sheetInfo.partPath);
  const sharedStringsXml = zip.file("xl/sharedStrings.xml")
    ? await zip.file("xl/sharedStrings.xml").async("string")
    : null;
  const stylesXml = zip.file("xl/styles.xml")
    ? await zip.file("xl/styles.xml").async("string")
    : null;
  const context = {
    sheetName,
    sharedStrings: parseSharedStrings(sharedStringsXml),
    date1904: parseWorkbookDateSystem(workbookXml),
    dateStyleIndexes: parseDateStyleIndexes(stylesXml),
  };
  const parsedRows = parseWorksheetCells(sheetXml, context);
  const populatedRows = [...parsedRows.keys()].sort((left, right) => left - right);
  if (!populatedRows.length) {
    return {
      name: sheetName,
      matrix: [],
      formulas: [],
      rowRecords: [],
      headerMetadata: [],
      mergedRanges: parseMergedRanges(sheetXml),
      unsafeMerges: [],
      usedRange: sheetInfo.usedRange,
      dataRange: null,
      dataStartRow: 0,
      dataStartColumn: 0,
      formulaEvents: [],
      sheetSha256: crypto.createHash("sha256").update("[]").digest("hex"),
    };
  }

  const headerRowIndex = populatedRows[0];
  const lastRowIndex = populatedRows.at(-1);
  const headerCells = parsedRows.get(headerRowIndex);
  const selectedColumns = selectProjectedColumns(headerCells, options);
  const mergedRanges = parseMergedRanges(sheetXml);
  const unsafeMerges = mergedRanges.filter((reference) =>
    mergeIntersectsColumns(
      reference,
      headerRowIndex,
      lastRowIndex,
      selectedColumns,
    ),
  );
  if (unsafeMerges.length) {
    throw new PauseError(
      "UNSAFE_MERGED_CELLS",
      "合并单元格与物化数据区域相交。",
      {
        sheetName,
        mergedRanges: unsafeMerges,
      },
    );
  }

  const matrix = [];
  const formulas = [];
  const metadataMatrix = [];
  for (let rowIndex = headerRowIndex; rowIndex <= lastRowIndex; rowIndex += 1) {
    const rowCells = parsedRows.get(rowIndex) ?? new Map();
    matrix.push(selectedColumns.map((column) => rowCells.get(column)?.value ?? null));
    formulas.push(selectedColumns.map((column) => rowCells.get(column)?.formula ?? null));
    metadataMatrix.push(selectedColumns.map((column) => {
      const cell = rowCells.get(column);
      return cell?.metadata ?? {
        address: columnLabel(column) + String(rowIndex + 1),
        sourceRow: rowIndex + 1,
        sourceColumn: column,
        formula: null,
      };
    }));
  }

  const rowRecords = matrix.slice(1).map((values, index) => ({
    values: [...values],
    sourceRow: headerRowIndex + index + 2,
    cellMetadata: metadataMatrix[index + 1].map((item) => ({ ...item })),
    sourceCells: metadataMatrix[index + 1].map((item) => item.address),
  }));
  const formulaEvents = metadataMatrix.flat()
    .filter((metadata) => metadata.formula !== null)
    .map((metadata) => ({
      sheetName,
      cell: metadata.address,
      sourceRow: metadata.sourceRow,
      sourceColumn: metadata.sourceColumn,
      formula: metadata.formula,
      currentValue: parsedRows
        .get(metadata.sourceRow - 1)
        ?.get(metadata.sourceColumn)?.value ?? null,
    }));
  const startColumn = selectedColumns.length ? Math.min(...selectedColumns) : 0;
  const endColumn = selectedColumns.length ? Math.max(...selectedColumns) : 0;
  const dataRange = selectedColumns.length
    ? columnLabel(startColumn) + String(headerRowIndex + 1) + ":" +
      columnLabel(endColumn) + String(lastRowIndex + 1)
    : null;
  const sheetSha256 = crypto
    .createHash("sha256")
    .update(JSON.stringify({ matrix, formulas }))
    .digest("hex");

  return {
    name: sheetName,
    usedRange: sheetInfo.usedRange,
    dataRange,
    dataStartRow: headerRowIndex,
    dataStartColumn: startColumn,
    matrix,
    formulas,
    rowRecords,
    headerMetadata: metadataMatrix[0]?.map((item) => ({ ...item })) ?? [],
    mergedRanges,
    unsafeMerges,
    formulaEvents,
    sheetSha256,
    projectedColumnCount: selectedColumns.length,
  };
}
