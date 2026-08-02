---
name: align-city-panel-data
description: Use when a city-level Excel/CSV table must be aligned to 城市顺序.xlsx, converted from 长表 or 宽表 into a descending-year panel, or reduced to selected indicators while preserving city-year-value correspondence.
---

# 城市面板数据对齐

以最少操作输出一个干净、独立、可核验的结果工作簿。源文件与固定的 `城市顺序.xlsx` 只读，不覆盖、不生成派生城市顺序表，不执行 Git 操作，不默认生成审计工作表或其他附加文件。

## 必须使用的能力

- 使用 `spreadsheets:Spreadsheets` 处理独立表格文件。
- 开始前调用 `codex_app.load_workspace_dependencies`，只使用返回的 Node.js、`@oai/artifact-tool` 和已捆绑依赖；缺少依赖时先询问用户，不擅自安装。
- 输入预检与读取使用已捆绑 jszip 完成 OOXML 定向读取；不要求先生成修复版或规范化临时源文件。
- `@oai/artifact-tool` 不得由估算、预检或源工作簿读取路径静态导入；仅在正式执行需要映射读取或结果写出时动态导入。
- 加载依赖后只做一次阶段化解析预检：估算路径验证 `jszip`，正式写出前再验证 `@oai/artifact-tool`。加载器报告可用但模块无法解析时，按捆绑运行时损坏报告，不逐包安装。
- 调用确定性运行器 [run-align-v2.mjs](scripts/run-align-v2.mjs)，不要在对话中临时重写排序逻辑。
- 未单独取得许可，不启动 Excel、LibreOffice、PDF/PNG 渲染、浏览器预览或外部转换器。

## 处理前必须给出两种模式建议

收到源文件和城市顺序文件后，先只读调用：

```js
const { estimateAlignmentModes } = await import("<skill>/scripts/run-align-v2.mjs");
const recommendations = await estimateAlignmentModes(config);
```

向用户同时报告：

1. 完整行模式预计时间、五小时额度、周额度和输出单元格量；
2. 指标精简模式预计时间、五小时额度、周额度、预计减少的处理量；
3. 推荐模式和原因；
4. 两种模式的信息保留差异。

用户已经明确指定模式时仍简短报告估算，然后按其选择执行。没有明确选择时，给出建议并等待用户决定，不写文件。

## 两种输出模式

### `preserve_rows`

- 长表移动完整数据行，保留全部源列的值。
- 宽表保留辅助列并展开为城市—年份记录。
- 适合归档、未知后续用途或用户要求“每行内容不变”的任务。

### `selected_indicators`

- 仅用于列式长表；始终保留城市列、年份列和用户明确指定的指标列。
- `selectedIndicators` 必须使用源表实际表头。只允许安全空格和全半角归一后的唯一精确匹配，不用模糊匹配猜指标。
- 指标不存在或对应多个同名列时暂停，列出候选让用户确认。
- 宽表的指标通常由工作表或指标区块定义；先选择对应工作表，再使用完整行模式。
- 不生成“提取后未对齐”的中间文件，直接生成最终精简对齐结果。

## 固定默认规则

- 未声明年份时，使用源文件中实际出现的全部唯一年份，按降序排列；不擅自在两个年份之间补出源文件不存在的年份。
- 用户明确给出年份范围时，按该范围生成降序年份集合。
- 先按 `城市顺序.xlsx` 的固定顺序排列。
- 源文件中不在城市顺序表内的非空城市，按其在源文件首次出现顺序追加在最后。
## OOXML 输入兼容边界

- 预检只读取包结构、工作表关系、使用区域和行列统计，不导入全部单元格。
- 只校验 workbook.xml 中工作表实际使用的关系；未被引用且目标缺失的孤立关系只记录，不阻断处理。
- 实际引用部件缺失时以 INVALID_WORKBOOK_PACKAGE 暂停，并报告关系 ID、目标与解析后的部件路径。
- 精简模式只物化城市列、年份列和用户指定指标列；完整行模式才物化目标数据区的全部源列。
- 共享字符串、内联字符串、数值、布尔值和带样式日期保持类型；公式只使用 OOXML 中已有缓存值。
- 公式错误、公式无缓存值或物化数据区域存在合并单元格时暂停。
- 不修改或修复源文件，不删除孤立关系，不输出临时“修复版源文件”。

- 不修改、扩展或派生 `城市顺序.xlsx`；不得生成派生顺序表。
- 空城市名、重复城市—年份键、城市列或年份列歧义、指标列歧义、公式错误、数据区域合并单元格等情况暂停正式输出。
- 不填充、不插值、不复制相邻值。缺失城市—年份只保留城市和年份，指标保持真正空白。
- 普通公式输出当前计算值；公式错误或缺少可用当前值时暂停。

## 执行

先加载工作区 Node 依赖，再动态导入：

```js
const { runAlignmentV2 } = await import("<skill>/scripts/run-align-v2.mjs");
const result = await runAlignmentV2({
  inputPath,
  cityOrderPath,
  cityOrderSheet,
  selectedSheets,
  outputDir,
  outputMode: "preserve_rows", // 或 selected_indicators
  selectedIndicators: [],      // 精简模式必填
  // startYear/endYear 均省略时自动使用源文件全部实际年份
});
```

`outputDir` 默认使用项目文件夹中的 `处理后数据`。每次只生成一个带时间戳的 `.xlsx` 结果文件；多个被选工作表可在同一个结果工作簿中形成多个结果表。不得覆盖任何输入或既有输出。

只有 `status === "passed"`、正式文件存在、正向核验与逐键反向核验通过、输入哈希未变时才能报告完成。`paused` 只报告最小证据和需要用户决定的一件事；`failed` 报告失败阶段。

## 输出与格式

- 默认工作簿只包含对齐结果表，不包含“核验结果”“公式处理”“城市匹配明细”等附加表。
- 快速模式只写数据值和简洁表头，不逐单元格复制源填充色、批注或超链接，避免黑色填充和不必要的样式开销。
- 源文件、`城市顺序.xlsx` 和既有 `城市名称映射表.xlsx` 保持只读；不会为范围外城市创建映射或派生文件。
- 不做 Git 状态检查、提交或推送。

## 交付说明

最终回复只需列出：

- 输出文件；
- 采用的模式及指定指标；
- 处理工作表、年份集合、固定顺序城市数和追加城市；
- 输出行列数与缺失键数量；
- 正向核验、逐键反向核验和输入哈希结果；
- 是否执行视觉渲染；未获授权时写明未执行。
