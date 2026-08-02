# V2 暂停代码

- `INVALID_CONFIG`：路径、模式或年份配置不完整。
- INVALID_WORKBOOK_PACKAGE：OOXML 必需部件、实际引用关系或实际引用工作表缺失；证据包含关系 ID、目标和解析后的部件路径。未引用的孤立关系不属于此错误。
- `MULTIPLE_SHEETS`：源文件或城市顺序文件存在多个候选工作表。
- `INVALID_CITY_ORDER`：固定城市顺序表表头、序号或城市名不合法。
- `AMBIGUOUS_CITY_COLUMN`：存在多个可能的城市列。
- `AMBIGUOUS_YEAR_COLUMN`：年份列或年份值不能唯一可靠识别。
- `INDICATOR_SELECTION_REQUIRED`：精简模式没有指定指标。
- `INDICATOR_NOT_FOUND`：指定指标表头不存在。
- `AMBIGUOUS_INDICATOR`：指定指标对应多个同名列。
- `INDICATOR_SELECTION_UNSUPPORTED`：当前宽表结构不能按列安全提取指标。
- `DUPLICATE_CITY_YEAR`：存在重复城市—年份键。
- `UNSAFE_MERGED_CELLS`：合并单元格与数据区域相交。
- `FORMULA_ERROR`：公式当前值为错误。
- `FORMULA_VALUE_UNAVAILABLE`：公式没有可用当前计算值。
- `OUTPUT_FILE_OCCUPIED`：目标输出已存在或被占用。
- `SOURCE_HASH_CHANGED`：运行期间输入发生变化。
- `REVERSE_VERIFICATION_FAILED`：输出不能逐键追溯到源数据。
- `OUTPUT_WRITE_FAILED`：临时写出、重新导入或正式复制失败。

范围外城市不是暂停条件：非空城市默认按源文件首次出现顺序追加。空城市名仍属于无法安全处理的数据问题。
