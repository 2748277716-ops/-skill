function row(section, item, status, value, details = null) {
  return {
    section,
    item,
    status,
    value,
    details: details === null ? null : JSON.stringify(details),
  };
}

function countBy(items, key) {
  const counts = {};
  for (const item of items) {
    const name = item?.[key] ?? "unknown";
    counts[name] = (counts[name] ?? 0) + 1;
  }
  return counts;
}

export function buildAudit({
  fileMetadata = {},
  config = {},
  sourceModels = [],
  outputModels = [],
  verifications = [],
  mappingEvents = [],
  formulaEvents = [],
  pauseEvents = [],
} = {}) {
  const auditRows = [];
  for (const [item, value] of Object.entries(fileMetadata)) {
    auditRows.push(row("文件", item, value ? "passed" : "failed", value ?? null));
  }

  const minimumYear = Math.min(Number(config.startYear), Number(config.endYear));
  const maximumYear = Math.max(Number(config.startYear), Number(config.endYear));
  auditRows.push(row("范围", "处理工作表", "passed", (config.selectedSheets ?? []).join(", ")));
  auditRows.push(row("范围", "目标年份", "passed", `${maximumYear}-${minimumYear}`));

  sourceModels.forEach((model, index) => {
    auditRows.push(row("结构", `源表${index + 1}`, "passed", model.kind, {
      sheetName: model.sheetName,
      rows: model.rows?.length ?? 0,
      columns: model.headers?.length ?? 0,
      cityColumn: model.cityColumn,
      yearColumn: model.yearColumn ?? null,
      yearColumns: model.yearColumns ?? null,
    }));
  });
  outputModels.forEach((model, index) => {
    auditRows.push(row("结构", `输出${index + 1}`, "passed", model.kind, {
      sheetName: model.sheetName,
      rows: model.rows?.length ?? 0,
      columns: model.headers?.length ?? 0,
    }));
  });

  const allMappingEvents = [
    ...mappingEvents,
    ...outputModels.flatMap((model) => model.mappingEvents ?? []),
  ];
  auditRows.push(row(
    "城市匹配",
    "匹配方法计数",
    "passed",
    allMappingEvents.length,
    countBy(allMappingEvents, "method"),
  ));
  for (const event of allMappingEvents) {
    auditRows.push(row(
      "城市匹配",
      `${event.rawCity} -> ${event.standardName}`,
      "passed",
      event.method,
      { sourceRow: event.sourceRow },
    ));
  }

  const missingKeys = outputModels.flatMap((model) => model.missingKeys ?? []);
  auditRows.push(row("缺失键", "缺失城市＋年份数量", "passed", missingKeys.length));
  for (const missing of missingKeys) {
    auditRows.push(row("缺失键", `${missing.city}|${missing.year}`, "passed", "空白行"));
  }

  auditRows.push(row("公式处理", "公式转当前值数量", "passed", formulaEvents.length));
  for (const event of formulaEvents) {
    auditRows.push(row(
      "公式处理",
      `${event.sheetName ?? event.sourceSheet}!${event.cell ?? `${event.sourceRow},${event.sourceColumn}`}`,
      "passed",
      event.currentValue,
      { formula: event.formula },
    ));
  }

  const outOfRangeCount = outputModels.reduce(
    (total, model) => total + (model.outOfRangeRows?.length ?? 0) + (model.outOfRangeRecords?.length ?? 0),
    0,
  );
  auditRows.push(row("范围", "范围外记录", "passed", outOfRangeCount));
  auditRows.push(row("核验", "暂停事件", pauseEvents.length ? "failed" : "passed", pauseEvents.length));

  let checksPassed = true;
  for (const [verificationIndex, verification] of verifications.entries()) {
    checksPassed &&= verification?.passed === true;
    for (const [checkName, checkValue] of Object.entries(verification?.checks ?? {})) {
      checksPassed &&= checkValue === true;
      auditRows.push(row(
        "核验",
        `${verification.kind ?? verificationIndex + 1}.${checkName}`,
        checkValue === true ? "passed" : "failed",
        checkValue,
      ));
    }
  }
  if (verifications.length === 0) checksPassed = false;

  const passed = checksPassed && pauseEvents.length === 0;
  const verdict = passed ? "通过" : "未通过";
  auditRows.push(row("结论", "最终结论", passed ? "passed" : "failed", verdict));
  return {
    passed,
    verdict,
    checks: {
      verificationCount: verifications.length,
      allVerificationsPassed: checksPassed,
      noPauseEvents: pauseEvents.length === 0,
    },
    auditRows,
  };
}