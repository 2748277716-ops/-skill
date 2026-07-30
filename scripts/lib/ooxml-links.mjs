import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

const HYPERLINK_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_NS =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_REL_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function xmlUnescape(value) {
  return String(value ?? "")
    .replace(/&quot;/gu, '"')
    .replace(/&apos;/gu, "'")
    .replace(/&lt;/gu, "<")
    .replace(/&gt;/gu, ">")
    .replace(/&amp;/gu, "&");
}

function xmlEscape(value) {
  return String(value ?? "")
    .replace(/&/gu, "&amp;")
    .replace(/"/gu, "&quot;")
    .replace(/</gu, "&lt;")
    .replace(/>/gu, "&gt;");
}

function attribute(fragment, localName) {
  const pattern = new RegExp(`(?:\\w+:)?${localName}="([^"]*)"`, "u");
  return fragment.match(pattern)?.[1] ?? null;
}

function normalizePartTarget(baseDirectory, target) {
  const decoded = xmlUnescape(target).replace(/\\/gu, "/");
  if (decoded.startsWith("/")) return decoded.slice(1);
  return path.posix.normalize(path.posix.join(baseDirectory, decoded));
}

async function relationships(zip, relationshipPath, baseDirectory) {
  const file = zip.file(relationshipPath);
  if (!file) return new Map();
  const xml = await file.async("string");
  const result = new Map();
  for (const match of xml.matchAll(/<Relationship\b([^>]*)\/?\s*>/gu)) {
    const id = attribute(match[1], "Id");
    const target = attribute(match[1], "Target");
    if (id && target) {
      result.set(id, {
        id,
        type: attribute(match[1], "Type"),
        target: normalizePartTarget(baseDirectory, target),
        rawTarget: xmlUnescape(target),
        targetMode: attribute(match[1], "TargetMode"),
      });
    }
  }
  return result;
}

async function workbookSheetParts(zip) {
  const workbookFile = zip.file("xl/workbook.xml");
  if (!workbookFile) throw new Error("Missing xl/workbook.xml");
  const workbookXml = await workbookFile.async("string");
  const workbookRels = await relationships(
    zip,
    "xl/_rels/workbook.xml.rels",
    "xl",
  );
  const result = new Map();
  for (const match of workbookXml.matchAll(/<(?:\w+:)?sheet\b([^>]*)\/?\s*>/gu)) {
    const name = xmlUnescape(attribute(match[1], "name"));
    const relationshipId = attribute(match[1], "id");
    const relationship = workbookRels.get(relationshipId);
    if (name && relationship) result.set(name, relationship.target);
  }
  return result;
}

function sheetRelationshipPath(sheetPart) {
  return path.posix.join(
    path.posix.dirname(sheetPart),
    "_rels",
    `${path.posix.basename(sheetPart)}.rels`,
  );
}

export async function readHyperlinksFromXlsx(workbookPath, selectedSheets = null) {
  const zip = await JSZip.loadAsync(await fs.readFile(workbookPath));
  const sheetParts = await workbookSheetParts(zip);
  const selected = selectedSheets ? new Set(selectedSheets) : null;
  const hyperlinks = [];
  for (const [sheetName, sheetPart] of sheetParts) {
    if (selected && !selected.has(sheetName)) continue;
    const sheetFile = zip.file(sheetPart);
    if (!sheetFile) continue;
    const sheetXml = await sheetFile.async("string");
    const rels = await relationships(
      zip,
      sheetRelationshipPath(sheetPart),
      path.posix.dirname(sheetPart),
    );
    for (const match of sheetXml.matchAll(/<(?:\w+:)?hyperlink\b([^>]*)\/?\s*>/gu)) {
      const address = attribute(match[1], "ref");
      const relationshipId = attribute(match[1], "id");
      const relationship = relationshipId ? rels.get(relationshipId) : null;
      if (!address) continue;
      hyperlinks.push({
        sheetName,
        address,
        target: relationship?.rawTarget ?? null,
        location: xmlUnescape(attribute(match[1], "location")),
        tooltip: xmlUnescape(attribute(match[1], "tooltip")),
        display: xmlUnescape(attribute(match[1], "display")),
      });
    }
  }
  return hyperlinks;
}

export async function applyHyperlinksToXlsx(workbookPath, hyperlinkEntries = []) {
  if (!hyperlinkEntries.length) return;
  const zip = await JSZip.loadAsync(await fs.readFile(workbookPath));
  const sheetParts = await workbookSheetParts(zip);
  const grouped = new Map();
  for (const entry of hyperlinkEntries) {
    const items = grouped.get(entry.sheetName) ?? [];
    items.push(entry);
    grouped.set(entry.sheetName, items);
  }

  for (const [sheetName, entries] of grouped) {
    const sheetPart = sheetParts.get(sheetName);
    if (!sheetPart) throw new Error(`Unknown output sheet for hyperlink: ${sheetName}`);
    const sheetFile = zip.file(sheetPart);
    if (!sheetFile) throw new Error(`Missing worksheet part: ${sheetPart}`);
    let sheetXml = await sheetFile.async("string");
    const relPath = sheetRelationshipPath(sheetPart);
    const relFile = zip.file(relPath);
    let relXml = relFile
      ? await relFile.async("string")
      : `<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="${REL_NS}"></Relationships>`;
    const existingIds = new Set(
      [...relXml.matchAll(/\bId="([^"]+)"/gu)].map((match) => match[1]),
    );
    const hyperlinkFragments = [];

    for (const [index, entry] of entries.entries()) {
      if (!entry.address || (!entry.target && !entry.location)) {
        throw new Error("Hyperlink entry needs an address and target or location");
      }
      const existingAddressPattern = new RegExp(
        `<(?:\\w+:)?hyperlink\\b[^>]*\\bref="${String(entry.address).replace(/[.*+?^${}()|[\]\\]/gu, "\\$&")}"`,
        "u",
      );
      if (existingAddressPattern.test(sheetXml)) {
        throw new Error(`Output hyperlink address already exists: ${sheetName}!${entry.address}`);
      }

      let relationshipAttribute = "";
      if (entry.target) {
        let relationshipId = `RpanelHyperlink${index + 1}`;
        while (existingIds.has(relationshipId)) {
          relationshipId = `${relationshipId}x`;
        }
        existingIds.add(relationshipId);
        const relationshipXml =
          `<Relationship Type="${HYPERLINK_REL}" Target="${xmlEscape(entry.target)}" ` +
          `TargetMode="External" Id="${relationshipId}" />`;
        relXml = relXml.replace("</Relationships>", `${relationshipXml}</Relationships>`);
        relationshipAttribute = ` r:id="${relationshipId}" xmlns:r="${OFFICE_REL_NS}"`;
      }
      const optional = [
        entry.location ? ` location="${xmlEscape(entry.location)}"` : "",
        entry.tooltip ? ` tooltip="${xmlEscape(entry.tooltip)}"` : "",
        entry.display ? ` display="${xmlEscape(entry.display)}"` : "",
      ].join("");
      hyperlinkFragments.push(
        `<x:hyperlink ref="${xmlEscape(entry.address)}"${relationshipAttribute}${optional} />`,
      );
    }

    const fragments = hyperlinkFragments.join("");
    if (/<(?:\w+:)?hyperlinks\b[^>]*>/u.test(sheetXml)) {
      sheetXml = sheetXml.replace(
        /<\/(?:\w+:)?hyperlinks>/u,
        `${fragments}</x:hyperlinks>`,
      );
    } else {
      const block = `<x:hyperlinks>${fragments}</x:hyperlinks>`;
      if (/<(?:\w+:)?pageMargins\b/u.test(sheetXml)) {
        sheetXml = sheetXml.replace(/<(?:\w+:)?pageMargins\b/u, `${block}<x:pageMargins`);
      } else {
        sheetXml = sheetXml.replace(/<\/(?:\w+:)?worksheet>/u, `${block}</x:worksheet>`);
      }
    }
    zip.file(sheetPart, sheetXml);
    zip.file(relPath, relXml);
  }

  await fs.writeFile(
    workbookPath,
    await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }),
  );
}