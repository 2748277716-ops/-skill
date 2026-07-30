# City Panel Data Alignment Skill Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an installed Codex Skill that converts long or wide city datasets into a complete, descending-year panel order while preserving source relationships, pausing on ambiguity, and producing an auditable independent workbook.

**Architecture:** Keep spreadsheet I/O behind an `@oai/artifact-tool` adapter, keep matching/detection/transformation/verification as pure JavaScript modules, and expose one deterministic `runAlignment(config)` orchestration entry point. `SKILL.md` performs dialogue, approval, worksheet selection, and pause/resume control; the scripts never guess through an ambiguous case.

**Tech Stack:** Codex Skill format, JavaScript ESM, Node.js built-in test runner, `@oai/artifact-tool`, bundled `jszip` for OOXML hyperlink preservation, PowerShell on Windows, Git.

## Global Constraints

- Work from an isolated Git worktree created with `superpowers:using-git-worktrees` when implementation begins.
- Do not install dependencies. Use only the Node.js and package paths returned by `codex_app.load_workspace_dependencies`.
- Use `@oai/artifact-tool` for workbook import, creation, editing, inspection, and export.
- Use bundled `jszip` only for targeted OOXML hyperlink extraction/remapping after an artifact-tool export; do not use it as a second spreadsheet authoring engine.
- Never modify or overwrite the source data workbook or `城市顺序.xlsx`.
- Only append a city mapping after the user explicitly confirms it; never overwrite a conflicting mapping.
- Complete the goal with the fewest necessary operations. Do not delete auxiliary columns, rename unrelated columns, change units, fill missing values, beautify, autofit, or restyle.
- Every target year range is explicitly supplied by the user. Output years descend from the highest year to the lowest year.
- Long tables move complete rows. Wide tables retain every non-year column and append `年份` plus indicator column(s).
- Duplicate city-year keys, ambiguous columns, unmatched city names, mapping conflicts, out-of-range years, invalid formulas, unsafe merged cells, occupied writable files, or failed reverse verification must pause before formal output.
- Normal formulas become their current calculated values. Formula errors or missing calculated values pause.
- Existing normal hyperlinks should be preserved when safely remappable. `HYPERLINK(...)` formula cells follow the formula-to-value rule and are recorded in the audit.
- Missing city-year rows contain only standard city name and year; all other cells are truly blank.
- Formal output is a new timestamped workbook in the source directory. It is never an overwrite.
- Do not run Excel, LibreOffice, PDF export, PNG rendering, browser preview, or any external GUI/CLI converter without separate user approval.
- Do not use parallel tasks or subagents unless the user separately approves the time, efficiency, and quota trade-off.
- Do not commit user workbooks, generated test workbooks, `node_modules`, temporary output, or mapping data.

---

## Planned File Structure

```text
align-city-panel-data/
├─ SKILL.md
├─ agents/
│  └─ openai.yaml
├─ references/
│  ├─ workflow.md
│  ├─ pause-codes.md
│  └─ output-contract.md
├─ scripts/
│  ├─ run-align.mjs
│  └─ lib/
│     ├─ pause.mjs
│     ├─ cities.mjs
│     ├─ detect.mjs
│     ├─ align-long.mjs
│     ├─ align-wide.mjs
│     ├─ verify.mjs
│     ├─ audit.mjs
│     ├─ workbook-io.mjs
│     ├─ mapping-store.mjs
│     └─ ooxml-links.mjs
├─ tests/
│  ├─ helpers/
│  │  ├─ artifact-runtime.mjs
│  │  └─ fixture-workbooks.mjs
│  ├─ capabilities.test.mjs
│  ├─ skill-contract.test.mjs
│  ├─ cities.test.mjs
│  ├─ detect.test.mjs
│  ├─ align-long.test.mjs
│  ├─ align-wide.test.mjs
│  ├─ verify.test.mjs
│  ├─ workbook-io.test.mjs
│  ├─ mapping-store.test.mjs
│  └─ integration.test.mjs
├─ docs/
│  └─ superpowers/
│     ├─ specs/
│     │  └─ 2026-07-30-align-city-panel-data-design.md
│     └─ plans/
│        └─ 2026-07-30-align-city-panel-data-implementation.md
└─ .gitignore
```

### Module Interfaces

```js
// scripts/lib/pause.mjs
export class PauseError extends Error {
  constructor(code, message, evidence = {}) { /* store code and evidence */ }
}

// scripts/lib/cities.mjs
export function validateCityOrder(rows);
export function validateMappingRows(rows, cityOrder);
export function normalizeSafeCityName(value);
export function resolveCityName(rawName, cityOrder, mappingRows);

// scripts/lib/detect.mjs
export function detectTableModel(sheetMatrix, config, cityContext);

// scripts/lib/align-long.mjs
export function alignLongTable(model, cityContext, config);

// scripts/lib/align-wide.mjs
export function alignWideTable(model, cityContext, config);

// scripts/lib/verify.mjs
export function verifyLongAlignment(sourceModel, outputModel, config);
export function verifyWideAlignment(sourceModel, outputModel, config);

// scripts/lib/workbook-io.mjs
export async function readWorkbookModel(path, selectedSheets);
export async function writeResultWorkbook(result, outputPath);

// scripts/lib/mapping-store.mjs
export async function appendConfirmedMappings(mappingPath, confirmedMappings, cityOrder);

// scripts/run-align.mjs
export async function runAlignment(config);
```

`runAlignment(config)` returns one of:

```js
{ status: "passed", outputPath, audit }
{ status: "paused", code, message, evidence }
{ status: "failed", code, message, evidence }
```

It must never return `passed` unless formal output exists and reverse verification has passed.

---

### Task 1: Runtime capability gate and repository hygiene

**Files:**
- Create: `.gitignore`
- Create: `tests/helpers/artifact-runtime.mjs`
- Create: `tests/helpers/fixture-workbooks.mjs`
- Create: `tests/capabilities.test.mjs`

**Interfaces:**
- Consumes: bundled Node executable and bundled Node package directory.
- Produces: `createCapabilityWorkbook(path)` and a passing capability gate for values, formulas, formats, comments, and OOXML hyperlink access.

- [ ] **Step 1: Create `.gitignore` before creating runtime links or test output**

```gitignore
node_modules/
work/
tests/tmp/
*.xlsx
*.xls
*.csv
*.tsv
*.tmp
*.log
```

- [ ] **Step 2: Create the test runtime helper**

```js
// tests/helpers/artifact-runtime.mjs
import fs from "node:fs";
import path from "node:path";

export function assertRuntimeAvailable() {
  const nodeModules = String.raw`C:\Users\86182\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules`;
  for (const name of ["@oai/artifact-tool", "jszip"]) {
    const target = path.join(nodeModules, ...name.split("/"));
    if (!fs.existsSync(target)) throw new Error(`Missing bundled module: ${name}`);
  }
  return nodeModules;
}
```

- [ ] **Step 3: Create a failing capability test**

```js
// tests/capabilities.test.mjs
import assert from "node:assert/strict";
import test from "node:test";
import { assertRuntimeAvailable } from "./helpers/artifact-runtime.mjs";
import { createCapabilityWorkbook, inspectCapabilityRoundTrip } from "./helpers/fixture-workbooks.mjs";

test("artifact runtime round-trips required spreadsheet features", async () => {
  assertRuntimeAvailable();
  const path = await createCapabilityWorkbook();
  const result = await inspectCapabilityRoundTrip(path);
  assert.deepEqual(result, {
    value: 10,
    formula: "=A2*2",
    calculated: 20,
    numberFormat: "0.00",
    comment: "fixture-comment",
    hyperlinkTarget: "https://example.com/",
  });
});
```

- [ ] **Step 4: Run the test and verify RED**

Run:

```powershell
$alignNodeExe = 'C:\Users\86182\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe'
& $alignNodeExe --test tests/capabilities.test.mjs
```

Expected: FAIL because `tests/helpers/fixture-workbooks.mjs` does not exist.

- [ ] **Step 5: Implement the fixture helper with artifact-tool and jszip**

The helper must:

- create `tests/tmp/capability.xlsx`;
- write `A2=10`, `B2="=A2*2"`, number format `0.00`, and a threaded comment;
- export with artifact-tool;
- add one normal hyperlink relationship to `C2` using jszip;
- re-import with artifact-tool;
- inspect formula/value/style/comment;
- inspect the exported OOXML hyperlink relationship with jszip.

Return exactly the object asserted in Step 3. If any required feature cannot be round-tripped, stop implementation and report the unsupported feature to the user before choosing another library or weakening the requirement.

- [ ] **Step 6: Run the capability test and verify GREEN**

Run:

```powershell
& $alignNodeExe --test tests/capabilities.test.mjs
```

Expected: PASS, one test.

- [ ] **Step 7: Verify ignored runtime artifacts**

Run:

```powershell
git status --short
```

Expected: only `.gitignore` and test source files appear; `node_modules` and `tests/tmp/*.xlsx` do not appear.

- [ ] **Step 8: Commit the capability gate**

```powershell
git add .gitignore tests/helpers/artifact-runtime.mjs tests/helpers/fixture-workbooks.mjs tests/capabilities.test.mjs
git commit -m "test: add spreadsheet runtime capability gate"
```

---

### Task 2: Skill contract, scaffold, and instruction references

**Files:**
- Create: `tests/skill-contract.test.mjs`
- Create: `SKILL.md`
- Create: `agents/openai.yaml`
- Create: `references/workflow.md`
- Create: `references/pause-codes.md`
- Create: `references/output-contract.md`

**Interfaces:**
- Consumes: approved design document.
- Produces: installed skill metadata and concise instructions that route agents to the deterministic runner.

- [ ] **Step 1: Write the failing skill contract test**

```js
import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";

test("SKILL.md exposes required trigger and safety contract", () => {
  const text = fs.readFileSync(new URL("../SKILL.md", import.meta.url), "utf8");
  assert.match(text, /^---[\s\S]*name: align-city-panel-data/m);
  assert.match(text, /description: Use when/i);
  for (const phrase of [
    "城市顺序.xlsx",
    "城市名称映射表.xlsx",
    "长表",
    "宽表",
    "暂停正式输出",
    "最少操作",
    "逐键反向核验",
  ]) assert.ok(text.includes(phrase), `Missing phrase: ${phrase}`);
});
```

- [ ] **Step 2: Run the contract test and verify RED**

Run:

```powershell
& $alignNodeExe --test tests/skill-contract.test.mjs
```

Expected: FAIL because `SKILL.md` does not exist.

- [ ] **Step 3: Run the required Skill Creator scaffold in a temporary directory**

Run:

```powershell
$alignPythonExe = 'C:\Users\86182\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$alignInitScript = 'C:\Users\86182\.codex\skills\.system\skill-creator\scripts\init_skill.py'
$alignScaffoldRoot = 'C:\tmp\align-city-panel-data-scaffold-20260730'
New-Item -ItemType Directory -Force -Path $alignScaffoldRoot | Out-Null
& $alignPythonExe $alignInitScript align-city-panel-data --path $alignScaffoldRoot --resources scripts,references --interface 'display_name=城市面板数据对齐' --interface 'short_description=安全对齐城市与年份顺序并生成核验结果' --interface 'default_prompt=将城市长表或宽表按城市顺序和目标年份安全对齐，遇到歧义时暂停。'
```

Copy only the generated `SKILL.md`, `agents/openai.yaml`, and empty requested resource directories into the worktree. Do not copy example placeholders.

- [ ] **Step 4: Author the minimal Skill workflow**

`SKILL.md` must:

- start its description with `Use when`;
- trigger on city order alignment, panel data, long/wide conversion, descending years, `城市顺序.xlsx`, or city-name correction;
- require `spreadsheets:Spreadsheets`;
- require `codex_app.load_workspace_dependencies`;
- require an explicit time/quota estimate before each run;
- gather source path, selected sheets, city-order path, target year range, and output location;
- call the deterministic runner;
- surface `paused` evidence without creating formal output;
- obtain confirmation before adding a fuzzy city mapping;
- treat that confirmation as mapping-write permission;
- require source and result hashes in the audit;
- prohibit rendering until separately approved.

Keep the body under 500 lines. Put the detailed state machine, pause code catalog, and output schema in the three reference files.

- [ ] **Step 5: Generate `agents/openai.yaml` deterministically**

Run:

```powershell
$alignYamlScript = 'C:\Users\86182\.codex\skills\.system\skill-creator\scripts\generate_openai_yaml.py'
& $alignPythonExe $alignYamlScript . --interface 'display_name=城市面板数据对齐' --interface 'short_description=安全对齐城市与年份顺序并生成核验结果' --interface 'default_prompt=将城市长表或宽表按城市顺序和目标年份安全对齐，遇到歧义时暂停。'
```

- [ ] **Step 6: Run the contract test and Skill validator**

Run:

```powershell
& $alignNodeExe --test tests/skill-contract.test.mjs
$alignValidateScript = 'C:\Users\86182\.codex\skills\.system\skill-creator\scripts\quick_validate.py'
& $alignPythonExe $alignValidateScript .
```

Expected: contract test PASS; validator exits 0.

- [ ] **Step 7: Commit the scaffold and contract**

```powershell
git add SKILL.md agents/openai.yaml references tests/skill-contract.test.mjs
git commit -m "feat: scaffold city panel alignment skill"
```

---

### Task 3: Pause model, city order validation, and name matching

**Files:**
- Create: `scripts/lib/pause.mjs`
- Create: `scripts/lib/cities.mjs`
- Create: `tests/cities.test.mjs`
- Modify: `references/pause-codes.md`

**Interfaces:**
- Consumes: raw city-order rows and optional mapping rows.
- Produces: validated city context and deterministic match results.

- [ ] **Step 1: Write failing city tests**

```js
import assert from "node:assert/strict";
import test from "node:test";
import {
  normalizeSafeCityName,
  resolveCityName,
  validateCityOrder,
} from "../scripts/lib/cities.mjs";

const order = validateCityOrder([
  { 序号: 1, 城市名: "北京市" },
  { 序号: 2, 城市名: "厦门市" },
]);

test("safe normalization removes Chinese-city whitespace", () => {
  assert.equal(normalizeSafeCityName(" 厦　 门市 "), "厦门市");
});

test("unique suffix difference becomes an automatic standard match", () => {
  assert.deepEqual(resolveCityName("厦门", order, []), {
    status: "matched",
    standardName: "厦门市",
    method: "unique_suffix",
  });
});

test("typo is only a candidate", () => {
  const result = resolveCityName("厦冂市", order, []);
  assert.equal(result.status, "needs_confirmation");
  assert.equal(result.candidates[0].standardName, "厦门市");
});
```

- [ ] **Step 2: Run the tests and verify RED**

Run:

```powershell
& $alignNodeExe --test tests/cities.test.mjs
```

Expected: FAIL because city modules do not exist.

- [ ] **Step 3: Implement pause codes and city functions**

`PauseError` must serialize as:

```js
{
  status: "paused",
  code: "DUPLICATE_CITY_YEAR",
  message: "存在重复的城市＋年份记录",
  evidence: { sheetName, keys, sourceRows }
}
```

`validateCityOrder()` must reject missing headers, blanks, duplicate sequence values, non-contiguous sequence, and duplicate city names.

`validateMappingRows()` must reject a standard city not in the order file and a source name mapped to two standards.

`resolveCityName()` must use this exact precedence:

```text
exact → safe normalization → unique suffix → confirmed mapping → candidate only
```

Candidate scoring must never change `status` to `matched`.

- [ ] **Step 4: Add conflict and invalid-order tests**

Add assertions for:

- sequence `1, 3` → `INVALID_CITY_ORDER`;
- duplicate `厦门市` → `INVALID_CITY_ORDER`;
- conflicting mapping → `MAPPING_CONFLICT`;
- safe normalization collision → `needs_confirmation`, not automatic.

- [ ] **Step 5: Run tests and commit**

```powershell
& $alignNodeExe --test tests/cities.test.mjs
git add scripts/lib/pause.mjs scripts/lib/cities.mjs tests/cities.test.mjs references/pause-codes.md
git commit -m "feat: add deterministic city matching"
```

Expected: all city tests PASS.

---

### Task 4: Table detection and preflight

**Files:**
- Create: `scripts/lib/detect.mjs`
- Create: `tests/detect.test.mjs`
- Modify: `references/workflow.md`

**Interfaces:**
- Consumes: a sheet matrix, selected sheet name, target years, and city context.
- Produces: a unique `LongTableModel` or `WideTableModel`, or throws `PauseError`.

- [ ] **Step 1: Write failing detection tests**

```js
test("detects a long table from city and year values", () => {
  const matrix = [
    ["省份", "城市", "年份", "收入"],
    ["福建", "厦门市", 2021, 120],
    ["福建", "厦门市", 2020, 100],
  ];
  const model = detectTableModel(matrix, config, cityContext);
  assert.equal(model.kind, "long");
  assert.equal(model.cityColumn, 1);
  assert.equal(model.yearColumn, 2);
});

test("detects a wide table from year headers", () => {
  const matrix = [
    ["省份", "城市", "代码", 2020, 2021],
    ["福建", "厦门市", 3502, 100, 120],
  ];
  const model = detectTableModel(matrix, config, cityContext);
  assert.deepEqual(model.yearColumns, [
    { year: 2020, column: 3 },
    { year: 2021, column: 4 },
  ]);
});
```

- [ ] **Step 2: Run and verify RED**

```powershell
& $alignNodeExe --test tests/detect.test.mjs
```

Expected: FAIL because `detect.mjs` does not exist.

- [ ] **Step 3: Implement unique detection**

The detector must:

- score city candidates using header text plus match rate against city context;
- recognize four-digit years as numeric or trimmed text;
- constrain years to the user target range;
- detect source years outside the target and pause with `OUT_OF_RANGE_YEAR`;
- require exactly one city column;
- require exactly one long-table year column or at least one wide year header;
- return `AMBIGUOUS_CITY_COLUMN`, `AMBIGUOUS_YEAR_COLUMN`, or `UNRECOGNIZED_TABLE` instead of guessing;
- reject data-region merged cells passed in preflight metadata;
- require selected sheets before processing a multi-sheet workbook.

- [ ] **Step 4: Add negative tests**

Cover:

- two city-like columns;
- two year-like long columns;
- source year 1999 when target is 2000–2024;
- merged data range;
- duplicate wide city rows;
- unselected multi-sheet workbook.

- [ ] **Step 5: Run tests and commit**

```powershell
& $alignNodeExe --test tests/detect.test.mjs
git add scripts/lib/detect.mjs tests/detect.test.mjs references/workflow.md
git commit -m "feat: add table preflight detection"
```

---

### Task 5: Long-table alignment

**Files:**
- Create: `scripts/lib/align-long.mjs`
- Create: `tests/align-long.test.mjs`

**Interfaces:**
- Consumes: `LongTableModel`, validated city context, and `{ startYear, endYear }`.
- Produces: aligned headers, existing rows with source provenance, inserted blank rows, and audit events.

- [ ] **Step 1: Write failing long-table tests**

```js
test("sorts complete rows and preserves all auxiliary columns", () => {
  const result = alignLongTable(model, cityContext, { startYear: 2020, endYear: 2021 });
  assert.deepEqual(result.headers, ["省份", "城市", "年份", "收入", "代码"]);
  assert.deepEqual(result.rows.map(r => [r[1], r[2], r[3]]), [
    ["北京市", 2021, 50],
    ["北京市", 2020, 40],
    ["厦门市", 2021, 120],
    ["厦门市", 2020, 100],
  ]);
});

test("creates missing rows with only city and year populated", () => {
  const result = alignLongTable(missingYearModel, cityContext, { startYear: 2020, endYear: 2021 });
  const inserted = result.rows.find(r => r[1] === "厦门市" && r[2] === 2020);
  assert.deepEqual(inserted, [null, "厦门市", 2020, null, null]);
});
```

- [ ] **Step 2: Run and verify RED**

```powershell
& $alignNodeExe --test tests/align-long.test.mjs
```

- [ ] **Step 3: Implement long alignment**

Implementation requirements:

- normalize only the city output cell to the standard name;
- preserve every other existing cell value and source provenance;
- sort by city rank ascending, then year descending;
- generate exactly one row per city per target year;
- populate only city and year for a missing key;
- throw `DUPLICATE_CITY_YEAR` even when duplicate rows are identical;
- keep style/comment/hyperlink references attached to the full existing row for the workbook adapter.

- [ ] **Step 4: Add tests for multiple indicators and complete missing cities**

Assert that:

- multiple value columns move together;
- an entirely absent city receives all target-year blank rows;
- no source row is omitted;
- no existing auxiliary value changes.

- [ ] **Step 5: Run tests and commit**

```powershell
& $alignNodeExe --test tests/align-long.test.mjs
git add scripts/lib/align-long.mjs tests/align-long.test.mjs
git commit -m "feat: align long city panel tables"
```

---

### Task 6: Wide-to-long alignment

**Files:**
- Create: `scripts/lib/align-wide.mjs`
- Create: `tests/align-wide.test.mjs`

**Interfaces:**
- Consumes: `WideTableModel`, validated city context, and target years.
- Produces: retained auxiliary columns followed by `年份` and one or more indicator columns.

- [ ] **Step 1: Write failing wide-table tests**

```js
test("un-pivots one indicator while retaining auxiliary columns", () => {
  const result = alignWideTable(model, cityContext, { startYear: 2020, endYear: 2021 });
  assert.deepEqual(result.headers, ["省份", "城市", "代码", "年份", "指标值"]);
  assert.deepEqual(result.rows, [
    ["福建", "厦门市", 3502, 2021, 120],
    ["福建", "厦门市", 3502, 2020, 100],
  ]);
});

test("keeps a missing year value truly blank", () => {
  const result = alignWideTable(missingYearModel, cityContext, { startYear: 2020, endYear: 2021 });
  assert.equal(result.rows.find(r => r[3] === 2020)[4], null);
});
```

- [ ] **Step 2: Run and verify RED**

```powershell
& $alignNodeExe --test tests/align-wide.test.mjs
```

- [ ] **Step 3: Implement single-indicator conversion**

The implementation must:

- preserve all non-year columns in their original order;
- append `年份` and `指标值` for a single-indicator sheet;
- use the uniquely identified indicator labels only in the guarded multi-indicator mode;
- repeat auxiliary values for every target year;
- standardize only the city cell;
- inherit value-cell style metadata from the source year cell;
- emit rows in city-rank then descending-year order;
- reject duplicate source city rows.

- [ ] **Step 4: Implement guarded multi-indicator conversion**

Only combine multiple indicators when each indicator block has:

- a unique indicator label;
- the same target-year set;
- the same unique city set;
- an unambiguous year-to-column mapping.

Otherwise return `AMBIGUOUS_MULTI_INDICATOR` with block and header evidence.

- [ ] **Step 5: Run tests and commit**

```powershell
& $alignNodeExe --test tests/align-wide.test.mjs
git add scripts/lib/align-wide.mjs tests/align-wide.test.mjs
git commit -m "feat: convert wide city tables to panel rows"
```

---

### Task 7: Forward/reverse verification and audit model

**Files:**
- Create: `scripts/lib/verify.mjs`
- Create: `scripts/lib/audit.mjs`
- Create: `tests/verify.test.mjs`
- Modify: `references/output-contract.md`

**Interfaces:**
- Consumes: source model, aligned output model, mapping events, formula events, and file metadata.
- Produces: `{ passed, checks, auditRows }`; throws on a relationship mismatch.

- [ ] **Step 1: Write failing reverse-verification tests**

```js
test("detects a value moved to the wrong city despite equal counts", () => {
  assert.throws(
    () => verifyWideAlignment(sourceModel, swappedOutputModel, config),
    error => error.code === "REVERSE_VERIFICATION_FAILED",
  );
});

test("accepts only authorized long-table differences", () => {
  const result = verifyLongAlignment(sourceModel, alignedModel, config);
  assert.equal(result.passed, true);
  assert.equal(result.checks.sourceRowsAccountedFor, true);
  assert.equal(result.checks.onlyAuthorizedChanges, true);
});
```

- [ ] **Step 2: Run and verify RED**

```powershell
& $alignNodeExe --test tests/verify.test.mjs
```

- [ ] **Step 3: Implement key-level verification**

Long verification must compare every existing source row to its output key and allow only:

- city value replaced by the resolved standard name;
- formula cell replaced by the recorded current value;
- new blank grid rows.

Wide verification must compare every source year cell to exactly one output key:

```text
standard city + year + indicator
```

Both verifiers must detect omissions, duplicates, extra non-empty values, and swapped values even when aggregate counts agree.

- [ ] **Step 4: Implement the audit model**

Audit sections must include:

- file paths and SHA-256 hashes;
- selected sheets;
- target years;
- detected table kind and column indices;
- source/output dimensions;
- match-method counts and original-to-standard mappings;
- missing city-year keys;
- formula-to-value events;
- out-of-range, duplicate, ambiguity, and pause checks;
- forward/reverse verification results;
- final `通过` or `未通过`.

- [ ] **Step 5: Run tests and commit**

```powershell
& $alignNodeExe --test tests/verify.test.mjs
git add scripts/lib/verify.mjs scripts/lib/audit.mjs tests/verify.test.mjs references/output-contract.md
git commit -m "feat: add auditable relationship verification"
```

---

### Task 8: Workbook adapter, formula conversion, comments, styles, and hyperlinks

**Files:**
- Create: `scripts/lib/workbook-io.mjs`
- Create: `scripts/lib/ooxml-links.mjs`
- Create: `tests/workbook-io.test.mjs`
- Modify: `tests/helpers/fixture-workbooks.mjs`

**Interfaces:**
- Consumes: workbook paths and aligned in-memory models.
- Produces: source models with provenance and exported result workbooks.

- [ ] **Step 1: Write failing workbook round-trip tests**

Create an input fixture with:

- two sheets;
- values and blanks;
- one normal formula with cached result;
- one formula error;
- number format, fill, border, comment, and normal hyperlink;
- one merged cell outside the data region and one fixture with a merged cell inside it.

Assert:

```js
test("reads formula text and current value separately", async () => {
  const model = await readWorkbookModel(inputPath, ["Data"]);
  assert.equal(model.sheets[0].cells.B2.formula, "=A2*2");
  assert.equal(model.sheets[0].cells.B2.value, 20);
});

test("pauses on formula errors and data-region merges", async () => {
  await assert.rejects(
    readWorkbookModel(errorPath, ["Data"]),
    error => ["FORMULA_ERROR", "UNSAFE_MERGE"].includes(error.code),
  );
});
```

- [ ] **Step 2: Run and verify RED**

```powershell
& $alignNodeExe --test tests/workbook-io.test.mjs
```

- [ ] **Step 3: Implement artifact-tool import and provenance capture**

Capture:

- values, formulas, formula information, computed styles, comments/threads, merges, and source cell addresses;
- source workbook and selected-sheet SHA-256 hashes;
- formula events with source cell, formula text, and current value;
- hyperlink objects from OOXML with relationship targets.

Do not load or process unselected sheets beyond listing their names and structural summaries.

- [ ] **Step 4: Implement workbook output**

Output requirements:

- one `对齐结果` sheet for one selected source sheet;
- `对齐结果_<工作表名>` for multiple selected sheets after user-approved safe names;
- one `核验结果` sheet;
- optional `范围外数据`;
- values written as typed values;
- formula cells written as current values;
- transferable styles, comments, and normal hyperlinks attached to remapped output cells;
- no autofit, restyling, charts, or rendering.

Export first to a temporary path in the destination directory, re-import, verify values/formulas, then rename to the final timestamped path only after verification passes.

- [ ] **Step 5: Run tests and commit**

```powershell
& $alignNodeExe --test tests/workbook-io.test.mjs
git add scripts/lib/workbook-io.mjs scripts/lib/ooxml-links.mjs tests/workbook-io.test.mjs tests/helpers/fixture-workbooks.mjs
git commit -m "feat: add safe workbook input and output"
```

---

### Task 9: Persistent confirmed-mapping store

**Files:**
- Create: `scripts/lib/mapping-store.mjs`
- Create: `tests/mapping-store.test.mjs`

**Interfaces:**
- Consumes: mapping path, explicitly confirmed mappings, and validated city order.
- Produces: safely appended and re-verified `城市名称映射表.xlsx`.

- [ ] **Step 1: Write failing mapping tests**

```js
test("creates the mapping workbook on first confirmed mapping", async () => {
  await appendConfirmedMappings(path, [{
    sourceName: "厦冂市",
    standardName: "厦门市",
    matchType: "错字确认",
    confirmedAt: "2026-07-30",
    note: "用户确认",
  }], cityOrder);
  const rows = await readMappingRows(path);
  assert.equal(rows.length, 1);
});

test("does not duplicate an identical mapping", async () => {
  await appendConfirmedMappings(path, [mapping], cityOrder);
  await appendConfirmedMappings(path, [mapping], cityOrder);
  assert.equal((await readMappingRows(path)).length, 1);
});
```

- [ ] **Step 2: Run and verify RED**

```powershell
& $alignNodeExe --test tests/mapping-store.test.mjs
```

- [ ] **Step 3: Implement append-only update**

The function must:

- require every proposed row to be explicitly present in `confirmedMappings`;
- verify each standard city exists;
- reject conflicts before writing;
- write a verified temporary workbook in the same directory;
- preserve all existing rows and columns;
- replace the original only after re-reading the temporary workbook;
- restore the original from a task-created temporary backup if replacement fails;
- re-read the final path and verify the appended rows;
- return `MAPPING_FILE_OCCUPIED` or `MAPPING_WRITE_FAILED` rather than choosing another path.

- [ ] **Step 4: Add injected-failure tests**

Use an injectable file adapter to simulate:

- destination occupied;
- temporary export failure;
- final verification failure;
- mapping conflict.

Assert that the original mapping bytes remain unchanged in every failure case.

- [ ] **Step 5: Run tests and commit**

```powershell
& $alignNodeExe --test tests/mapping-store.test.mjs
git add scripts/lib/mapping-store.mjs tests/mapping-store.test.mjs
git commit -m "feat: persist confirmed city mappings safely"
```

---

### Task 10: Deterministic orchestration and pause/resume contract

**Files:**
- Create: `scripts/run-align.mjs`
- Create: `tests/integration.test.mjs`
- Modify: `references/workflow.md`
- Modify: `SKILL.md`

**Interfaces:**
- Consumes: one JSON-serializable alignment config.
- Produces: `passed`, `paused`, or `failed` result objects with no ambiguous side effects.

- [ ] **Step 1: Write the failing orchestration test**

```js
test("does not create formal output when a typo needs confirmation", async () => {
  const result = await runAlignment(configWithUnconfirmedTypo);
  assert.equal(result.status, "paused");
  assert.equal(result.code, "CITY_CONFIRMATION_REQUIRED");
  assert.equal(fs.existsSync(result.expectedOutputPath), false);
});

test("resumes after explicit confirmation and appends mapping", async () => {
  const result = await runAlignment({
    ...configWithUnconfirmedTypo,
    confirmedMappings: [{ sourceName: "厦冂市", standardName: "厦门市" }],
  });
  assert.equal(result.status, "passed");
  assert.equal(fs.existsSync(result.outputPath), true);
});
```

- [ ] **Step 2: Run and verify RED**

```powershell
& $alignNodeExe --test tests/integration.test.mjs
```

- [ ] **Step 3: Implement `runAlignment(config)`**

Config shape:

```js
{
  inputPath: "C:\\data\\source.xlsx",
  cityOrderPath: "C:\\data\\城市顺序.xlsx",
  mappingPath: "C:\\data\\城市名称映射表.xlsx",
  selectedSheets: ["Sheet1"],
  startYear: 2000,
  endYear: 2024,
  outputDir: "C:\\data",
  approvedSheetNames: {},
  approvedExcludedYears: [],
  confirmedMappings: []
}
```

Execution order:

```text
validate config
→ hash source/reference/mapping
→ read structural summaries
→ preflight formulas/merges/years
→ pause unless every range-excluded year appears in approvedExcludedYears
→ resolve cities
→ pause if confirmation required
→ append confirmed mappings
→ transform selected sheets
→ forward/reverse verify
→ export temporary workbook
→ re-import and verify
→ rename to timestamped formal output
→ return passed
```

Catch `PauseError` and return its serialized form. Any unexpected error returns `failed` and must not leave a formal output file.

- [ ] **Step 4: Add timestamp and sheet-name tests**

Assert:

- output name uses the source base name without its extension: `原文件名_城市面板对齐_YYYYMMDD-HHMMSS.xlsx`;
- an existing output is never overwritten;
- invalid or overlong sheet names pause until `approvedSheetNames` is supplied;
- multiple selected sheets create one output sheet per source plus one audit sheet.

- [ ] **Step 5: Update `SKILL.md` with the exact runner contract**

Document:

- how to build config after user answers;
- how to run via `node_repl` after adding the loader-provided module directory;
- how to present pause evidence;
- how to accept a mapping confirmation;
- how to rerun with `confirmedMappings`;
- how to report the output and audit verdict;
- how to avoid claiming completion on a `paused` or `failed` result.

- [ ] **Step 6: Run tests and commit**

```powershell
& $alignNodeExe --test tests/integration.test.mjs
git add scripts/run-align.mjs tests/integration.test.mjs SKILL.md references/workflow.md
git commit -m "feat: orchestrate safe panel alignment"
```

---

### Task 11: Full regression, source-integrity checks, and deployment validation

**Files:**
- Modify: tests as required only to fix proven gaps.
- Modify: `SKILL.md` or references only when a failing scenario proves a gap.
- Modify: `agents/openai.yaml` only if metadata no longer matches `SKILL.md`.

**Interfaces:**
- Consumes: all prior modules and generated fixtures.
- Produces: validated skill ready for user review.

- [ ] **Step 1: Record pre-test source hashes**

For every generated fixture and the user-provided `城市顺序.xlsx` used only for read-only validation, compute SHA-256 before tests. Do not copy the user workbook into the repository.

- [ ] **Step 2: Run the complete automated suite**

```powershell
& $alignNodeExe --test tests/*.test.mjs
```

Expected: all tests PASS.

- [ ] **Step 3: Run the required scenario matrix**

Verify all 18 design scenarios:

1. long table with multiple indicators and auxiliary columns;
2. single-indicator wide table;
3. reliable multi-indicator wide table;
4. ambiguous multi-indicator pause;
5. whitespace variants;
6. unique suffix variant;
7. typo candidate pause;
8. mapping reuse, append, and conflict;
9. missing cities and years;
10. duplicate city-year pause;
11. out-of-range pause;
12. formula-to-value success;
13. formula-error pause;
14. multiple-sheet selection;
15. style/comment/hyperlink transfer;
16. forward/reverse mismatch detection;
17. output or mapping file occupied;
18. source/reference hashes unchanged.

- [ ] **Step 4: Run Skill validation**

```powershell
$alignPythonExe = 'C:\Users\86182\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$alignValidateScript = 'C:\Users\86182\.codex\skills\.system\skill-creator\scripts\quick_validate.py'
& $alignPythonExe $alignValidateScript .
```

Expected: exit 0.

- [ ] **Step 5: Run static hygiene checks**

```powershell
rg -n "TO[D]O|TB[D]|PLACEH[O]LDER|implement[ ]later|fill[ ]in" SKILL.md agents references scripts tests
git status --short
git ls-files | rg "\.(xlsx|xls|csv|tsv|tmp|log)$"
```

Expected:

- no placeholder hits;
- no generated workbook or user data tracked;
- only intended implementation changes remain.

- [ ] **Step 6: Recompute source hashes**

Assert the source workbook and `城市顺序.xlsx` hashes match Step 1 exactly. Treat a mismatch as a failed release.

- [ ] **Step 7: Do not render without approval**

Record `visual_rendering: not performed - user approval not requested` in the handoff. Structural verification and workbook re-import checks are not visual verification.

- [ ] **Step 8: Commit verified implementation**

```powershell
git add SKILL.md agents references scripts tests .gitignore
git commit -m "test: verify city panel alignment skill"
```

- [ ] **Step 9: Run final verification after commit**

```powershell
& $alignNodeExe --test tests/*.test.mjs
& $alignPythonExe $alignValidateScript .
git status --short --branch
git log --oneline --decorate -8
```

Expected:

- all tests PASS;
- Skill validation PASS;
- worktree clean;
- commits are local only and nothing is pushed.

---

## Plan Self-Review Checklist

- [x] Every design requirement maps to at least one task.
- [x] No task authorizes source or city-order modification.
- [x] Ambiguity and validation failures pause before formal output.
- [x] Formula conversion, comments, styles, and hyperlinks have explicit tests.
- [x] Long and wide transformations have forward and reverse relationship checks.
- [x] Mapping writes are append-only, user-confirmed, verified, and recoverable.
- [x] No external dependency install, rendering, parallel execution, or push is included.
- [x] No placeholders remain.
- [x] Function names and return shapes match across tasks.

## Execution Handoff

After the user approves this plan, choose one execution mode:

1. **Subagent-Driven** — requires separate user approval before dispatching any agents. Expected implementation time 1.5–2.5 hours, total completion time 1.5–2.5 hours, and estimated five-hour quota 25%–40%; parallel review may save approximately 25%–35% elapsed time but adds approximately 20%–30% quota.
2. **Inline Execution** — execute in the current session with `superpowers:executing-plans`, using checkpoints after Tasks 2, 6, 9, and 11. Expected implementation time 2–3 hours and estimated five-hour quota 15%–25%, with no extra parallel-agent quota.

Do not begin either mode until the user explicitly chooses it.
