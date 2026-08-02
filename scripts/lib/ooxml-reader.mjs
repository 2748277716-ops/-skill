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
