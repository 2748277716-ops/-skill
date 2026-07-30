---
name: align-city-panel-data
description: Use when a city-level Excel/CSV table must be aligned to 城市顺序.xlsx, converted from 长表 or 宽表 into a descending-year panel, city names need safe correction, or an auditable independent result is required without altering source data.
---

# 城市面板数据对齐

将城市数据整理为独立、可核验的标准长表：每行保持“城市＋年份＋数值/源表辅助列”的对应关系，城市按指定基准排序，年份从高到低。始终坚持最少操作，不覆盖源文件，不把歧义当作事实。

## 必须使用的能力

- 使用 `spreadsheets:Spreadsheets` 处理独立 `.xlsx/.xls/.csv/.tsv` 文件。
- 开始前调用 `codex_app.load_workspace_dependencies`，只使用返回的 Node.js、`@oai/artifact-tool` 和已捆绑依赖；不得擅自安装或改用其他表格库。
- 运行确定性脚本 `scripts/run-align.mjs`；不要在对话中临时重写排序逻辑。
- 每次执行前先给出基于当前已知运行速度的具体时间、五小时额度和周额度估计，并取得用户的操作授权。

## 开始前收集

必须确认：

1. 源数据文件路径；
2. 需要处理的工作表；存在多个候选工作表时只列出并让用户选择；
3. `城市顺序.xlsx` 路径；
4. 目标年份范围，例如 `2024-2000`；
5. 输出目录，默认源文件所在目录；
6. 如存在，`城市名称映射表.xlsx` 路径，默认与城市顺序文件同目录；
7. 对范围外年份、模糊城市名等未决事项的用户决定。

未达到 90% 理解把握时，一次只问一个关键问题，不执行写入。

## 工作流

1. 完整阅读 [workflow.md](references/workflow.md)、[pause-codes.md](references/pause-codes.md) 与 [output-contract.md](references/output-contract.md)。
2. 只读预检源文件、城市顺序文件和映射表；记录 SHA-256。若目标文件被占用，暂停并让用户关闭。
3. 识别长表或宽表以及城市、年份、指标列。列含义不唯一时暂停，不猜测。
4. 按以下优先级匹配城市：标准名精确匹配；安全空格/全半角归一；唯一行政后缀补全；用户确认过的长期映射。模糊错字只能生成候选并暂停。
5. 用户确认模糊映射后，将主表城市名单元格改成标准名称，并把该记录追加到同目录的 `城市名称映射表.xlsx`。确认同时视为本次追加授权；冲突时仍须暂停。
6. 调用确定性运行器。只在返回 `passed` 且正式文件存在、正向和逐键反向核验均通过后报告完成。
7. 交付独立长表及核验结果；用户自行复制数据列到面板数据。

## 确定性运行器

只把 `.xlsx` 工作簿直接交给运行器。先调用 `codex_app.load_workspace_dependencies`，再把返回的绝对 `node_modules` 路径传给 `node_repl.js_add_node_module_dir`。随后在 `node_repl.js` 中动态导入本 Skill 的绝对 `scripts/run-align.mjs`，调用 `await runAlignment(config)`；不得复制脚本逻辑到临时代码。

配置必须是可序列化对象，至少包含 `inputPath`、`cityOrderPath`、`selectedSheets`、`startYear` 和 `endYear`。`mappingPath` 默认是城市顺序文件同目录的 `城市名称映射表.xlsx`，`outputDir` 默认是源文件目录。所有确认项都要显式写入 `approvedSheetNames`、`approvedExcludedYears` 或 `confirmedMappings`。完整字段、返回值和重跑方式见 [workflow.md](references/workflow.md#运行器配置与调用)。

对结果严格分流：`passed` 才能交付；`paused` 只报告 `code`、最小证据和所需的一个决定；`failed` 报告失败阶段，不得声称已完成。用户确认城市映射后，把原始名、标准名、确认时间、来源文件和来源工作表写入 `confirmedMappings`，从预检开始完整重跑。

## 不可放宽的规则

- 长表移动完整行，保留全部源列；不得删除省份、代码、英文名等辅助列。
- 宽表保留全部非年份列，追加年份和指标值；默认一个工作表对应一个指标。只有指标可靠识别且共享同一城市—年份结构时才允许多指标模式。
- 目标输出为完整城市×年份网格。源数据缺失的键只填写标准城市名与年份，其他字段保持真正空白，并列入核验结果。
- 年份从上到下递减。不得用填充、插值、复制相邻值等方式补数据。
- 重复城市—年份键一律暂停，即使数值相同。
- 范围外年份一律暂停；只有用户明确允许排除后，才保留到 `范围外数据` 工作表并继续。
- 普通公式统一输出当前计算值；公式错误或无可用缓存值一律暂停并让用户处理。
- 尽量保留可安全对应的格式、评论和普通超链接；不美化、不自动列宽、不重排无关内容。合并单元格或复杂结构无法安全展开时暂停。
- 源数据、`城市顺序.xlsx` 和既有映射记录不可修改。正式输出使用新时间戳文件名，绝不覆盖。
- 在 `paused` 或 `failed` 状态下暂停正式输出，仅给出证据和用户需要处理的事项。
- 未单独取得许可，不得启动 Excel、LibreOffice、PDF/PNG 渲染、浏览器预览或外部转换器。

## 暂停与继续

捕获运行器返回的暂停代码，按 [pause-codes.md](references/pause-codes.md) 向用户说明：发生位置、原始值、候选或冲突、需要的唯一决定。不要生成“临时正式结果”规避暂停。得到决定后，以显式配置或确认映射重新运行完整预检与核验。

## 交付说明

最终回复至少列出：输出文件、处理工作表、城市数、目标年份数、输出行数、缺失键数量、公式转值数量、城市匹配明细、范围外数据处理、正向核验、逐键反向核验、源文件哈希未变、是否执行视觉渲染以及 Git/临时文件状态。核验项目必须逐项标记通过、暂停或证据不足。