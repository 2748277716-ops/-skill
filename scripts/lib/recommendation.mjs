function assertPositiveInteger(value, name, { allowZero = false } = {}) {
  if (!Number.isInteger(value) || value < (allowZero ? 0 : 1)) {
    throw new TypeError(`${name} must be ${allowZero ? "a non-negative" : "a positive"} integer`);
  }
}

function fullRowMinutes(cellCount) {
  if (cellCount <= 250_000) return [4, 8];
  if (cellCount <= 1_000_000) return [5, 10];
  if (cellCount <= 5_000_000) return [8, 18];
  if (cellCount <= 12_000_000) return [15, 35];
  return [25, 50];
}

function selectedMinutes(fullRange, reductionPercent) {
  if (reductionPercent < 25) return fullRange;
  const factor = reductionPercent >= 90 ? 0.42 : reductionPercent >= 70 ? 0.58 : 0.75;
  return [
    Math.max(3, Math.round(fullRange[0] * factor)),
    Math.max(5, Math.round(fullRange[1] * Math.min(0.8, factor + 0.15))),
  ];
}

function rangeText(range, suffix) {
  return `${range[0]}–${range[1]}${suffix}`;
}

function quotaRange(minutes, windowMinutes) {
  return minutes.map((value) => Number((value / windowMinutes * 100).toFixed(1)));
}

function modeSummary(outputCellCount, reductionPercent, minutes) {
  return {
    outputCellCount,
    workloadReductionPercent: reductionPercent,
    timeMinutes: minutes,
    timeEstimate: rangeText(minutes, "分钟"),
    fiveHourQuotaPercent: quotaRange(minutes, 300),
    fiveHourQuotaEstimate: rangeText(quotaRange(minutes, 300), "%"),
    weeklyQuotaPercent: quotaRange(minutes, 7 * 24 * 60),
    weeklyQuotaEstimate: rangeText(quotaRange(minutes, 7 * 24 * 60), "%"),
  };
}

export function recommendProcessingModes({
  rowCount,
  columnCount,
  selectedIndicatorCount = 0,
  keyColumnCount = 2,
}) {
  assertPositiveInteger(rowCount, "rowCount");
  assertPositiveInteger(columnCount, "columnCount");
  assertPositiveInteger(selectedIndicatorCount, "selectedIndicatorCount", { allowZero: true });
  assertPositiveInteger(keyColumnCount, "keyColumnCount");

  const selectedColumnCount = Math.min(
    columnCount,
    keyColumnCount + selectedIndicatorCount,
  );
  const fullCellCount = rowCount * columnCount;
  const selectedCellCount = rowCount * selectedColumnCount;
  const reductionPercent = Number((
    (1 - selectedCellCount / fullCellCount) * 100
  ).toFixed(1));
  const fullMinutes = fullRowMinutes(fullCellCount);
  const slimMinutes = selectedMinutes(fullMinutes, reductionPercent);
  const recommendedMode = selectedIndicatorCount > 0 && reductionPercent >= 25
    ? "selected_indicators"
    : "preserve_rows";

  return {
    basis: {
      rowCount,
      columnCount,
      selectedIndicatorCount,
      keyColumnCount,
      note: "估算包含读取、处理、写出和结构核验；实际时间受公式、样式、工作表数量和文件占用影响。",
    },
    fullRow: modeSummary(fullCellCount, 0, fullMinutes),
    selectedIndicators: modeSummary(selectedCellCount, reductionPercent, slimMinutes),
    recommendedMode,
    reason: recommendedMode === "selected_indicators"
      ? "指定指标模式可显著减少输出单元格，且结果更便于复制和合并。"
      : "精简后的列数优势有限，完整行模式能保留更多源信息。",
  };
}
