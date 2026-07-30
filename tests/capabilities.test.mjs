import assert from "node:assert/strict";
import test from "node:test";

import { assertRuntimeAvailable } from "./helpers/artifact-runtime.mjs";
import {
  createCapabilityWorkbook,
  inspectCapabilityRoundTrip,
} from "./helpers/fixture-workbooks.mjs";

test("artifact runtime round-trips required spreadsheet features", async () => {
  assertRuntimeAvailable();
  const workbookPath = await createCapabilityWorkbook();
  const result = await inspectCapabilityRoundTrip(workbookPath);
  assert.deepEqual(result, {
    value: 10,
    formula: "=A2*2",
    calculated: 20,
    numberFormat: "0.00",
    comment: "fixture-comment",
    hyperlinkTarget: "https://example.com/",
  });
});
