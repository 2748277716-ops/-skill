# OOXML 定向读取 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让城市面板对齐技能无需完整导入源工作簿即可预检和读取目标数据，并安全容忍未被工作簿引用的孤立 OOXML 关系。

**Architecture:** 新增一个只负责 ZIP/OOXML 输入读取的模块，使用已捆绑的 `JSZip` 解析工作簿关系、共享字符串、样式和工作表单元格。现有排序、城市匹配、年份规则和核验逻辑继续复用；`@oai/artifact-tool` 只负责最终输出创建和重新导入核验。

**Tech Stack:** Node.js ESM、`jszip`、`node:test`、`@oai/artifact-tool`、Excel OOXML。

## Global Constraints

- 源工作簿和 `城市顺序.xlsx` 始终只读。
- 仅忽略未被 `workbook.xml` 引用的孤立关系；实际引用部件缺失必须暂停。
- 不安装新依赖，只使用加载入口返回的 Node.js、`jszip` 和 `@oai/artifact-tool`。
- 精简模式只物化城市、年份和用户指定指标列。
- 完整行模式保留全部源列值。
- 公式只使用已有缓存值；错误值或无缓存值必须暂停。
- 数据区域存在合并单元格时必须暂停。
- 未指定年份时使用源文件中实际出现的全部唯一年份并降序排列。
- 固定城市顺序不修改；范围外城市按源文件首次出现顺序放在结果末尾。
- 不生成派生城市顺序表、审计工作表或额外正式文件。
- 不执行视觉渲染，除非用户另行授权。

---

## File Structure

- Create: `scripts/lib/ooxml-reader.mjs`
  - OOXML 实体解码、关系解析、工作簿结构检查、共享字符串、样式、表头和工作表值读取。
- Create: `tests/ooxml-reader.test.mjs`
  - 包结构、孤立关系、实际缺失部件、类型读取、公式、合并单元格和列投影测试。
- Modify: `scripts/lib/workbook-io.mjs`
  - `listWorkbookStructure` 和 `computeFileSha256` 接入 OOXML 包检查器。
- Modify: `scripts/lib/workbook-fast.mjs`
  - `readWorkbookFast` 改用 OOXML 工作表读取器；写出逻辑保持 `artifact-tool`。
- Modify: `scripts/run-align-v2.mjs`
  - 预检使用结构统计；精简模式把指标选择下推到读取器。
- Modify: `tests/v2-integration.test.mjs`
  - 加入孤立关系预检和精简列投影端到端回归。
- Modify: `tests/helpers/fixture-workbooks.mjs`
  - 提供孤立关系与实际缺失部件夹具修改器。
- Modify: `tests/helpers/artifact-runtime.mjs`
  - 明确 `jszip` 是输入读取运行时依赖。
- Modify: `SKILL.md`
  - 说明预检与输入读取使用 OOXML 定向读取，输出使用 `artifact-tool`。
- Modify: `references/workflow.md`
  - 增加包结构检查、孤立关系处理和按列读取步骤。
- Modify: `references/pause-codes.md`
  - 增加 `INVALID_WORKBOOK_PACKAGE` 暂停证据。

---

### Task 1: OOXML 包结构检查器

**Files:**
- Create: `scripts/lib/ooxml-reader.mjs`
- Create: `tests/ooxml-reader.test.mjs`
- Modify: `tests/helpers/fixture-workbooks.mjs`

**Interfaces:**
- Produces:
  - `inspectWorkbookPackage(workbookPath: string): Promise<WorkbookPackageInfo>`
  - `WorkbookPackageInfo = { path, sha256, sheets, orphanRelationships }`
  - `sheets[] = { name, state, relationshipId, partPath, dimension, rowCount, columnCount }`
- Consumes: `PauseError` from `scripts/lib/pause.mjs` and bundled `JSZip`.

- [ ] **Step 1: Add a failing orphan-relationship fixture and test**

Add a helper that loads a valid fixture with `JSZip`, inserts:

```xml
<Relationship
  Id="rIdOrphan"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
  Target="worksheets/missing-orphan.xml"/>
```

without adding a matching `<sheet r:id="rIdOrphan">` to `xl/workbook.xml`.

Test:

```js
test("unreferenced missing worksheet relationship is reported but tolerated", async () => {
  const info = await inspectWorkbookPackage(orphanPath);
  assert.deepEqual(info.sheets.map((sheet) => sheet.name), ["Data", "Notes"]);
  assert.deepEqual(info.orphanRelationships.map((item) => item.id), ["rIdOrphan"]);
});
```

- [ ] **Step 2: Add a failing referenced-missing-part test**

Delete `xl/worksheets/sheet1.xml` from a valid fixture while keeping the
`workbook.xml` sheet entry and relationship.

```js
await assert.rejects(
  inspectWorkbookPackage(referencedMissingPath),
  (error) =>
    error instanceof PauseError &&
    error.code === "INVALID_WORKBOOK_PACKAGE" &&
    error.evidence.partPath === "xl/worksheets/sheet1.xml",
);
```

- [ ] **Step 3: Run the focused tests and confirm RED**

Run:

```powershell
& 'C:\Users\86182\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' --test tests/ooxml-reader.test.mjs
```

Expected: FAIL because `scripts/lib/ooxml-reader.mjs` does not exist.

- [ ] **Step 4: Implement minimal package parsing**

Implement these focused helpers in `ooxml-reader.mjs`:

```js
export async function inspectWorkbookPackage(workbookPath) {}
function parseXmlAttributes(tag) {}
function decodeXmlText(text) {}
function resolvePartPath(sourcePart, target) {}
function parseWorkbookSheets(workbookXml) {}
function parseRelationships(relationshipsXml) {}
function parseDimension(sheetXml) {}
```

Required behavior:

- load the ZIP once with `JSZip.loadAsync`;
- require `[Content_Types].xml`, `xl/workbook.xml` and
  `xl/_rels/workbook.xml.rels`;
- map only `<sheet>` entries from `workbook.xml` to relationships;
- validate each mapped relationship and target part;
- list, but do not fail on, unused relationships;
- parse `A1:HF10201` into 10,201 total rows and 214 columns;
- return data row count as `max(0, totalRows - 1)`;
- throw `PauseError("INVALID_WORKBOOK_PACKAGE", ...)` with exact relationship ID,
  target and resolved part path for an actual missing part.

- [ ] **Step 5: Run focused tests and confirm GREEN**

Run the same focused command.

Expected: both package-structure tests pass.

- [ ] **Step 6: Commit Task 1**

```powershell
git add scripts/lib/ooxml-reader.mjs tests/ooxml-reader.test.mjs tests/helpers/fixture-workbooks.mjs
git commit -m "feat: inspect xlsx package without full import"
```

---

### Task 2: Typed OOXML worksheet reader

**Files:**
- Modify: `scripts/lib/ooxml-reader.mjs`
- Modify: `tests/ooxml-reader.test.mjs`

**Interfaces:**
- Consumes: `inspectWorkbookPackage`.
- Produces:
  - `readWorksheetTable(workbookPath, sheetName, options?): Promise<SheetModel>`
  - `options.projectHeaders?: string[]`
  - `options.requiredKeyHeaders?: string[]`
  - `SheetModel = { name, matrix, rowRecords, headerMetadata, mergedRanges, unsafeMerges, formulaEvents, usedRange, dataRange, dataStartRow, dataStartColumn, sheetSha256 }`

- [ ] **Step 1: Add failing typed-value tests**

Create a fixture containing:

```js
[
  ["城市", "年份", "文本", "数值", "布尔", "日期", "公式"],
  ["厦门市", 2021, "甲&乙", 12.5, true, new Date("2021-01-02T00:00:00Z"), null],
]
```

with `G2` set to `=D2*2`.

Assert that `readWorksheetTable` returns:

- decoded shared or inline strings;
- number `12.5`;
- boolean `true`;
- a `Date` representing 2021-01-02;
- formula cached value `25`;
- formula event with source cell `G2`.

- [ ] **Step 2: Add failing formula safety tests**

Create one fixture with `=1/0` and one fixture whose formula `<v>` element is
removed with `JSZip`.

Assert:

```js
error.code === "FORMULA_ERROR"
error.code === "FORMULA_VALUE_UNAVAILABLE"
```

- [ ] **Step 3: Add failing merge and projection tests**

Assert that a merge intersecting `A1:G2` produces
`UNSAFE_MERGED_CELLS`.

Then read the normal fixture with:

```js
{
  projectHeaders: ["城市", "年份", "公式"],
}
```

and assert:

```js
sheet.matrix[0] === ["城市", "年份", "公式"]
sheet.rowRecords[0].sourceCells === ["A2", "B2", "G2"]
```

- [ ] **Step 4: Run focused tests and confirm RED**

Run:

```powershell
& 'C:\Users\86182\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' --test tests/ooxml-reader.test.mjs
```

Expected: new typed-reader tests fail.

- [ ] **Step 5: Implement shared strings, styles and cell decoding**

Add:

```js
function parseSharedStrings(xml) {}
function parseWorkbookDateSystem(workbookXml) {}
function parseDateStyleIndexes(stylesXml) {}
function parseWorksheetRows(sheetXml, context) {}
function cellValue(cellXml, context) {}
export async function readWorksheetTable(workbookPath, sheetName, options = {}) {}
```

Rules:

- concatenate every `<t>` descendant inside each `<si>` rich-text item;
- support cell types `s`, `inlineStr`, `str`, `b`, `e` and numeric/default;
- recognize built-in Excel date number formats and custom formats containing
  date/time tokens after quoted and escaped text is removed;
- use 1899-12-30 or the workbook `date1904` epoch for serial conversion;
- preserve original source row and source column in metadata;
- materialize only projected columns when `projectHeaders` is present;
- projection matching uses NFKC plus whitespace removal and requires one match;
- collect formulas and require a non-error cached value;
- parse merge ranges and pause when any merge intersects the materialized data
  region.

- [ ] **Step 6: Run focused tests and confirm GREEN**

Expected: all OOXML reader tests pass.

- [ ] **Step 7: Commit Task 2**

```powershell
git add scripts/lib/ooxml-reader.mjs tests/ooxml-reader.test.mjs
git commit -m "feat: read projected worksheet values from ooxml"
```

---

### Task 3: Integrate preflight and V2 processing

**Files:**
- Modify: `scripts/lib/workbook-io.mjs`
- Modify: `scripts/lib/workbook-fast.mjs`
- Modify: `scripts/run-align-v2.mjs`
- Modify: `tests/v2-integration.test.mjs`

**Interfaces:**
- Consumes: `inspectWorkbookPackage`, `readWorksheetTable`.
- Preserves public exports:
  - `listWorkbookStructure`
  - `readWorkbookFast`
  - `estimateAlignmentModes`
  - `runAlignmentV2`

- [ ] **Step 1: Add a failing V2 orphan preflight test**

Modify a normal V2 source fixture to include the missing `rIdOrphan`
relationship, then call:

```js
const result = await estimateAlignmentModes({
  inputPath,
  selectedSheets: ["Data"],
  selectedIndicators: ["可支配收入"],
});
```

Assert the preflight returns both modes and does not import the missing orphan.

- [ ] **Step 2: Add a failing selected-projection integration test**

Create a 40-column long table containing city, year, two requested indicators
and 36 unrequested columns. Run selected mode and assert:

```js
output headers === ["城市", "年份", "指标A", "指标B"]
source reader projected column count === 4
reverse verification passed === true
```

Expose `sourceColumnCountsBySheet` in the returned audit solely as a numeric
diagnostic so the test can prove early projection occurred.

- [ ] **Step 3: Run V2 integration tests and confirm RED**

Run:

```powershell
& 'C:\Users\86182\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' --test tests/v2-integration.test.mjs
```

- [ ] **Step 4: Replace structure enumeration**

Change `listWorkbookStructure` to return the `inspectWorkbookPackage` sheet
summaries:

```js
{
  path,
  sha256,
  sheets: sheets.map(({ name, state, usedRange }) => ({ name, state, usedRange })),
  orphanRelationships,
}
```

Do not call `SpreadsheetFile.importXlsx`.

- [ ] **Step 5: Replace input reading**

Change:

```js
readWorkbookFast(workbookPath, selectedSheets = [], options = {})
```

to call `readWorksheetTable` for selected sheets and aggregate formula events.
Keep `writeCleanResultWorkbook` unchanged except for imports no longer needed by
the read path.

- [ ] **Step 6: Make preflight metadata-only**

Change `estimateAlignmentModes` to use `listWorkbookStructure` row and column
counts directly:

```js
recommendProcessingModes({
  rowCount: sheet.rowCount,
  columnCount: sheet.columnCount,
  selectedIndicatorCount,
})
```

Preflight must not call `readWorkbookFast`.

- [ ] **Step 7: Push selected indicators into the source reader**

Before reading source data, construct:

```js
const readOptions = outputMode === "selected_indicators"
  ? {
      projectHeaders: ["城市", "年份", ...config.selectedIndicators],
      projectionRoles: {
        cityHeaders: ["城市", "城市名", "城市名称", "地级市", "city", "cityname"],
        yearHeaders: ["年份", "年度", "统计年份", "year"],
      },
    }
  : {};
```

The reader must identify one city and one year header from the accepted aliases,
then append the exact selected indicators. The downstream model receives the
projected header order `城市列、年份列、指标列` without loading unrelated columns.

- [ ] **Step 8: Run V2 tests and confirm GREEN**

Run the focused V2 test command, followed by:

```powershell
& 'C:\Users\86182\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' --test tests/v2-rules.test.mjs tests/v2-integration.test.mjs
```

- [ ] **Step 9: Commit Task 3**

```powershell
git add scripts/lib/workbook-io.mjs scripts/lib/workbook-fast.mjs scripts/run-align-v2.mjs tests/v2-integration.test.mjs
git commit -m "feat: use projected ooxml reads in alignment v2"
```

---

### Task 4: Skill contract and pause documentation

**Files:**
- Modify: `SKILL.md`
- Modify: `references/workflow.md`
- Modify: `references/pause-codes.md`
- Modify: `tests/skill-contract.test.mjs`
- Modify: `tests/helpers/artifact-runtime.mjs`

**Interfaces:**
- Documents the behavior implemented by Tasks 1–3.

- [ ] **Step 1: Add failing contract assertions**

Assert the skill documentation contains:

- `OOXML 定向读取`;
- `孤立关系`;
- `实际引用部件缺失`;
- `INVALID_WORKBOOK_PACKAGE`;
- `精简模式只物化`;
- `artifact-tool 仅用于结果写出和复核`.

- [ ] **Step 2: Run contract tests and confirm RED**

Run:

```powershell
& 'C:\Users\86182\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' --test tests/skill-contract.test.mjs
```

- [ ] **Step 3: Update skill and references**

Document:

- loader-provided Node.js and bundled `jszip`;
- preflight package inspection;
- orphan relationship tolerance;
- actual referenced-part failure boundary;
- selected-column materialization;
- formula and merge pause rules;
- no source repair or temporary normalized source requirement.

- [ ] **Step 4: Run contract tests and confirm GREEN**

- [ ] **Step 5: Commit Task 4**

```powershell
git add SKILL.md references/workflow.md references/pause-codes.md tests/skill-contract.test.mjs tests/helpers/artifact-runtime.mjs
git commit -m "docs: define ooxml input compatibility contract"
```

---

### Task 5: Full regression and real-workbook verification

**Files:**
- Test only; no source file modifications.

**Interfaces:**
- Consumes the public V2 API implemented above.

- [ ] **Step 1: Run all formal tests**

Run:

```powershell
& 'C:\Users\86182\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' --test tests
```

Expected: all tests pass with zero failures.

- [ ] **Step 2: Run formatting and sensitive-path checks**

Run:

```powershell
git diff --check
rg -n 'C:\\Users\\86182|中国城市数据库6\.0版|暂未整理数据' SKILL.md references scripts tests
```

Expected: no diff errors and no hard-coded real data paths in formal skill files.

- [ ] **Step 3: Measure real preflight**

Call `estimateAlignmentModes` against:

- `中国城市数据库6.0版.xlsx`;
- selected sheet `原始数据`;
- seven requested GDP indicators.

Record elapsed wall time and assert it completes within 120 seconds without a
normalized temporary copy.

- [ ] **Step 4: Run real selected-mode output in a unique `C:\tmp` directory**

Call `runAlignmentV2` with:

```js
{
  outputMode: "selected_indicators",
  selectedSheets: ["原始数据"],
  selectedIndicators: [
    "地区生产总值(万元)",
    "第一产业增加值(万元)",
    "第二产业增加值(万元)",
    "第三产业增加值(万元)",
    "第一产业增加值占GDP比重(%)",
    "第二产业增加值占GDP比重(%)",
    "第三产业增加值占GDP比重(%)",
  ],
}
```

Do not pass `startYear` or `endYear`.

- [ ] **Step 5: Verify the real output**

Require:

- `status === "passed"`;
- one result sheet only;
- exactly nine output columns;
- all actual source years in descending order;
- fixed 295 cities first;
- source-only cities after the fixed list in first-appearance order;
- no formulas;
- forward and reverse verification passed;
- input and city-order hashes unchanged;
- no derived city-order file;
- no extra formal output files.

- [ ] **Step 6: Remove only the verified real-test temporary directory**

Resolve the absolute path, confirm it starts with
`C:\tmp\align-city-panel-real-test-`, then remove that directory and verify it no
longer exists.

- [ ] **Step 7: Review final diff and repository state**

Run:

```powershell
git status --short --branch
git diff --stat HEAD
```

List every modified file and confirm no source data or test output is staged.

- [ ] **Step 8: Final implementation commit**

If Tasks 1–4 were not committed separately, stage only the planned skill files
and commit:

```powershell
git commit -m "feat: add tolerant ooxml city panel reader"
```

Do not push unless the user explicitly requests it.
