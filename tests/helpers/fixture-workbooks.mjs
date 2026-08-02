import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";
import JSZip from "jszip";

const HYPERLINK_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_NS =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_REL_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const tempDirectory = fileURLToPath(new URL("../tmp/", import.meta.url));
const capabilityPath = path.join(tempDirectory, "capability.xlsx");

async function addNormalHyperlink(workbookPath, target, cellAddress = "C2") {
  const zip = await JSZip.loadAsync(await fs.readFile(workbookPath));
  const sheetPath = "xl/worksheets/sheet1.xml";
  const relsPath = "xl/worksheets/_rels/sheet1.xml.rels";
  const relationshipId = "RcapHyperlink";

  const sheetFile = zip.file(sheetPath);
  if (!sheetFile) throw new Error(`Missing OOXML part: ${sheetPath}`);
  let sheetXml = await sheetFile.async("string");
  if (sheetXml.includes(`r:id="${relationshipId}"`)) {
    throw new Error("Capability hyperlink relationship already exists");
  }
  const hyperlinkXml =
    `<x:hyperlinks><x:hyperlink ref="${cellAddress}" r:id="${relationshipId}" ` +
    `xmlns:r="${OFFICE_REL_NS}" /></x:hyperlinks>`;
  if (!sheetXml.includes("</x:sheetData>")) {
    throw new Error("Unsupported worksheet XML: missing sheetData boundary");
  }
  sheetXml = sheetXml.replace("</x:sheetData>", `</x:sheetData>${hyperlinkXml}`);
  zip.file(sheetPath, sheetXml);

  const relsFile = zip.file(relsPath);
  let relsXml = relsFile
    ? await relsFile.async("string")
    : `<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="${REL_NS}"></Relationships>`;
  if (!relsXml.includes("</Relationships>")) {
    throw new Error("Unsupported relationship XML");
  }
  const relationshipXml =
    `<Relationship Type="${HYPERLINK_REL}" Target="${target}" ` +
    `TargetMode="External" Id="${relationshipId}" />`;
  relsXml = relsXml.replace("</Relationships>", `${relationshipXml}</Relationships>`);
  zip.file(relsPath, relsXml);

  await fs.writeFile(
    workbookPath,
    await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }),
  );
}

async function inspectNormalHyperlink(workbookPath) {
  const zip = await JSZip.loadAsync(await fs.readFile(workbookPath));
  const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
  const relsXml = await zip
    .file("xl/worksheets/_rels/sheet1.xml.rels")
    ?.async("string");
  if (!sheetXml || !relsXml) return null;

  const linkMatch = sheetXml.match(
    /<x:hyperlink\b(?=[^>]*\bref="C2")(?=[^>]*\br:id="([^"]+)")[^>]*\/>/,
  );
  if (!linkMatch) return null;
  const escapedId = linkMatch[1].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const relMatch = relsXml.match(
    new RegExp(
      `<Relationship\\b(?=[^>]*\\bType="${HYPERLINK_REL}")` +
        `(?=[^>]*\\bId="${escapedId}")(?=[^>]*\\bTarget="([^"]+)")[^>]*/>`,
    ),
  );
  return relMatch?.[1] ?? null;
}

export async function createCapabilityWorkbook(outputPath = capabilityPath) {
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Sheet1");
  sheet.getRange("A2").values = [[10]];
  sheet.getRange("A2").format.numberFormat = "0.00";
  sheet.getRange("B2").formulas = [["=A2*2"]];
  sheet.getRange("C2").values = [["Example"]];
  workbook.comments.setSelf({ displayName: "fixture" });
  workbook.comments.addThread(
    { cell: sheet.getRange("A2") },
    "fixture-comment",
  );

  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(outputPath);
  await addNormalHyperlink(outputPath, "https://example.com/");
  return outputPath;
}

export async function inspectCapabilityRoundTrip(workbookPath) {
  const imported = await SpreadsheetFile.importXlsx(
    await FileBlob.load(workbookPath),
  );
  const sheet = imported.worksheets.getItemAt(0);
  const threadInspection = await imported.inspect({
    kind: "thread",
    maxChars: 3000,
  });
  const threads = String(threadInspection.ndjson ?? "")
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => JSON.parse(line));

  return {
    value: sheet.getRange("A2").values[0][0],
    formula: sheet.getRange("B2").formulas[0][0],
    calculated: sheet.getRange("B2").values[0][0],
    numberFormat: sheet.getRange("A2").format.numberFormat,
    comment: threads.find((thread) => thread.target === "A2")?.text ?? null,
    hyperlinkTarget: await inspectNormalHyperlink(workbookPath),
  };
}

const workbookIoFixturePromises = new Map();

async function createWorkbookIoFixture(workbookPath, { formulaError = false, mergeInside = false } = {}) {
  await fs.rm(workbookPath, { force: true });
  const workbook = Workbook.create();
  const data = workbook.worksheets.add("Data");
  const notes = workbook.worksheets.add("Notes");
  data.getRange("A1:E3").values = [
    ["城市", "年份", "值", "计算值", "链接"],
    ["厦门市", 2021, 10, null, "Example"],
    ["北京市", 2020, 5, null, null],
  ];
  data.getRange("D2").formulas = [[formulaError ? "=1/0" : "=C2*2"]];
  data.getRange("C2").format = {
    numberFormat: "0.00",
    fill: "#FFFF00",
    font: { bold: true, color: "#FF0000" },
    borders: { preset: "all", style: "thin", color: "#00FF00" },
  };
  workbook.comments.setSelf({ displayName: "fixture" });
  workbook.comments.addThread(
    { cell: data.getRange("C2") },
    "fixture-comment",
  );
  data.getRange("H1:I1").merge();
  data.getRange("H1").values = [["outside-note"]];
  if (mergeInside) data.getRange("A2:B2").merge();
  notes.getRange("A1:B2").values = [
    ["说明", "不处理"],
    ["仅用于工作表列表", true],
  ];

  const exported = await SpreadsheetFile.exportXlsx(workbook);
  await exported.save(workbookPath);
  await addNormalHyperlink(workbookPath, "https://example.com/", "E2");
}

export function createWorkbookIoFixtures(outputDirectory) {
  const cacheKey = path.resolve(outputDirectory);
  if (!workbookIoFixturePromises.has(cacheKey)) {
    workbookIoFixturePromises.set(cacheKey, (async () => {
      await fs.mkdir(outputDirectory, { recursive: true });
      const paths = {
        normalPath: path.join(outputDirectory, "workbook-io-normal.xlsx"),
        errorPath: path.join(outputDirectory, "workbook-io-error.xlsx"),
        mergedPath: path.join(outputDirectory, "workbook-io-merged.xlsx"),
      };
      await createWorkbookIoFixture(paths.normalPath);
      await createWorkbookIoFixture(paths.errorPath, { formulaError: true });
      await createWorkbookIoFixture(paths.mergedPath, { mergeInside: true });
      return paths;
    })());
  }
  return workbookIoFixturePromises.get(cacheKey);
}
export async function addOrphanWorksheetRelationship(
  sourcePath,
  outputPath,
  {
    id = "rIdOrphan",
    target = "worksheets/missing-orphan.xml",
  } = {},
) {
  const zip = await JSZip.loadAsync(await fs.readFile(sourcePath));
  const relsPath = "xl/_rels/workbook.xml.rels";
  const relsFile = zip.file(relsPath);
  if (!relsFile) throw new Error(`Missing OOXML part: ${relsPath}`);
  let relsXml = await relsFile.async("string");
  if (relsXml.includes(`Id="${id}"`)) {
    throw new Error(`Relationship already exists: ${id}`);
  }
  const relationshipXml =
    `<Relationship Id="${id}" ` +
    `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ` +
    `Target="${target}"/>`;
  if (!relsXml.includes("</Relationships>")) {
    throw new Error("Unsupported workbook relationship XML");
  }
  relsXml = relsXml.replace(
    "</Relationships>",
    `${relationshipXml}</Relationships>`,
  );
  zip.file(relsPath, relsXml);
  await fs.writeFile(
    outputPath,
    await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }),
  );
  return outputPath;
}

export async function removeReferencedWorksheetPart(
  sourcePath,
  outputPath,
  partPath = "xl/worksheets/sheet1.xml",
) {
  const zip = await JSZip.loadAsync(await fs.readFile(sourcePath));
  if (!zip.file(partPath)) throw new Error(`Missing OOXML part: ${partPath}`);
  zip.remove(partPath);
  await fs.writeFile(
    outputPath,
    await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }),
  );
  return outputPath;
}
