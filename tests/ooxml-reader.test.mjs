import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { inspectWorkbookPackage } from "../scripts/lib/ooxml-reader.mjs";
import { PauseError } from "../scripts/lib/pause.mjs";
import {
  addOrphanWorksheetRelationship,
  createWorkbookIoFixtures,
  removeReferencedWorksheetPart,
} from "./helpers/fixture-workbooks.mjs";

const testDirectory = fileURLToPath(new URL("./tmp/ooxml-reader/", import.meta.url));

test.before(async () => {
  await fs.mkdir(testDirectory, { recursive: true });
});

test("unreferenced missing worksheet relationship is reported but tolerated", async () => {
  const { normalPath } = await createWorkbookIoFixtures(testDirectory);
  const orphanPath = path.join(testDirectory, "workbook-orphan-relationship.xlsx");
  await addOrphanWorksheetRelationship(normalPath, orphanPath);

  const info = await inspectWorkbookPackage(orphanPath);

  assert.deepEqual(info.sheets.map((sheet) => sheet.name), ["Data", "Notes"]);
  assert.deepEqual(
    info.orphanRelationships.map((relationship) => relationship.id),
    ["rIdOrphan"],
  );
});

test("referenced missing worksheet part pauses package inspection", async () => {
  const { normalPath } = await createWorkbookIoFixtures(testDirectory);
  const missingPath = path.join(testDirectory, "workbook-referenced-missing.xlsx");
  await removeReferencedWorksheetPart(normalPath, missingPath);

  await assert.rejects(
    inspectWorkbookPackage(missingPath),
    (error) =>
      error instanceof PauseError &&
      error.code === "INVALID_WORKBOOK_PACKAGE" &&
      error.evidence.partPath === "xl/worksheets/sheet1.xml",
  );
});
