import { PauseError } from "./pause.mjs";

const ADMINISTRATIVE_SUFFIXES = [
  "特别行政区",
  "自治州",
  "自治县",
  "自治旗",
  "地区",
  "新区",
  "市",
  "盟",
  "县",
  "区",
];

function pause(code, message, evidence) {
  throw new PauseError(code, message, evidence);
}

export function normalizeSafeCityName(value) {
  if (value === null || value === undefined) return "";
  return String(value).normalize("NFKC").replace(/[\s\u00a0\u3000]+/gu, "");
}

function baseCityName(value) {
  const normalized = normalizeSafeCityName(value);
  for (const suffix of ADMINISTRATIVE_SUFFIXES) {
    if (normalized.endsWith(suffix) && normalized.length > suffix.length) {
      return normalized.slice(0, -suffix.length);
    }
  }
  return normalized;
}

function addToIndex(index, key, entry) {
  const existing = index.get(key) ?? [];
  existing.push(entry);
  index.set(key, existing);
}

export function validateCityOrder(rows) {
  if (!Array.isArray(rows) || rows.length === 0) {
    pause("INVALID_CITY_ORDER", "城市顺序表为空", { rowCount: rows?.length ?? 0 });
  }

  const entries = rows.map((row, rowIndex) => {
    if (
      !row ||
      typeof row !== "object" ||
      !Object.hasOwn(row, "序号") ||
      !Object.hasOwn(row, "城市名")
    ) {
      pause("INVALID_CITY_ORDER", "城市顺序表缺少必需表头", {
        row: rowIndex + 2,
        requiredHeaders: ["序号", "城市名"],
      });
    }
    const sequence = Number(row.序号);
    const standardName = String(row.城市名 ?? "");
    if (!Number.isInteger(sequence) || sequence < 1) {
      pause("INVALID_CITY_ORDER", "城市序号必须是正整数", {
        row: rowIndex + 2,
        value: row.序号,
      });
    }
    if (!normalizeSafeCityName(standardName)) {
      pause("INVALID_CITY_ORDER", "标准城市名不得为空", {
        row: rowIndex + 2,
        value: row.城市名,
      });
    }
    return { sequence, standardName, sourceRow: rowIndex + 2 };
  });

  const duplicateSequences = entries
    .filter((entry, index) => entries.findIndex((item) => item.sequence === entry.sequence) !== index)
    .map((entry) => entry.sequence);
  const duplicateNames = entries
    .filter((entry, index) => entries.findIndex((item) => item.standardName === entry.standardName) !== index)
    .map((entry) => entry.standardName);
  if (duplicateSequences.length || duplicateNames.length) {
    pause("INVALID_CITY_ORDER", "城市顺序表存在重复序号或城市名", {
      duplicateSequences: [...new Set(duplicateSequences)],
      duplicateNames: [...new Set(duplicateNames)],
    });
  }

  const sorted = [...entries].sort((left, right) => left.sequence - right.sequence);
  const sequenceErrors = sorted
    .filter((entry, index) => entry.sequence !== index + 1)
    .map((entry, index) => ({ expected: index + 1, actual: entry.sequence }));
  if (sequenceErrors.length) {
    pause("INVALID_CITY_ORDER", "城市序号必须从 1 开始连续排列", {
      sequenceErrors,
    });
  }

  const byExact = new Map();
  const normalizedIndex = new Map();
  const suffixIndex = new Map();
  const rankByName = new Map();
  sorted.forEach((entry, rank) => {
    const ranked = { ...entry, rank };
    byExact.set(entry.standardName, ranked);
    rankByName.set(entry.standardName, rank);
    addToIndex(normalizedIndex, normalizeSafeCityName(entry.standardName), ranked);
    addToIndex(suffixIndex, baseCityName(entry.standardName), ranked);
  });

  return {
    entries: sorted.map((entry, rank) => ({ ...entry, rank })),
    names: sorted.map((entry) => entry.standardName),
    byExact,
    normalizedIndex,
    suffixIndex,
    rankByName,
  };
}

export function validateMappingRows(rows = [], cityOrder) {
  if (rows?.kind === "validated_city_mappings") return rows;
  if (!cityOrder?.byExact) {
    pause("MAPPING_CONFLICT", "缺少已验证的城市顺序上下文", {});
  }
  if (!Array.isArray(rows)) {
    pause("MAPPING_CONFLICT", "城市名称映射记录必须是数组", {
      receivedType: typeof rows,
    });
  }

  const bySourceNormalized = new Map();
  const validatedRows = [];
  for (const [rowIndex, row] of rows.entries()) {
    const sourceName = String(row?.原始城市名 ?? row?.sourceName ?? "");
    const standardName = String(row?.标准城市名 ?? row?.standardName ?? "");
    const sourceKey = normalizeSafeCityName(sourceName);
    if (!sourceKey || !standardName || !cityOrder.byExact.has(standardName)) {
      pause("MAPPING_CONFLICT", "城市名称映射包含空值或未知标准城市", {
        row: rowIndex + 2,
        sourceName,
        standardName,
      });
    }
    const existing = bySourceNormalized.get(sourceKey);
    if (existing && existing.standardName !== standardName) {
      pause("MAPPING_CONFLICT", "同一原始城市名映射到多个标准城市", {
        sourceName,
        standards: [existing.standardName, standardName],
        rows: [existing.sourceRow, rowIndex + 2],
      });
    }
    if (!existing) {
      const mapping = { sourceName, standardName, sourceRow: rowIndex + 2 };
      bySourceNormalized.set(sourceKey, mapping);
      validatedRows.push(mapping);
    }
  }

  return {
    kind: "validated_city_mappings",
    rows: validatedRows,
    bySourceNormalized,
  };
}

function levenshtein(left, right) {
  const a = [...left];
  const b = [...right];
  const previous = Array.from({ length: b.length + 1 }, (_, index) => index);
  for (let row = 1; row <= a.length; row += 1) {
    const current = [row];
    for (let column = 1; column <= b.length; column += 1) {
      current[column] = Math.min(
        current[column - 1] + 1,
        previous[column] + 1,
        previous[column - 1] + (a[row - 1] === b[column - 1] ? 0 : 1),
      );
    }
    previous.splice(0, previous.length, ...current);
  }
  return previous[b.length];
}

function rankedCandidates(rawName, cityOrder) {
  const normalized = normalizeSafeCityName(rawName);
  return cityOrder.entries
    .map((entry) => {
      const candidate = normalizeSafeCityName(entry.standardName);
      const distance = levenshtein(normalized, candidate);
      const length = Math.max([...normalized].length, [...candidate].length, 1);
      return {
        standardName: entry.standardName,
        distance,
        similarity: Number((1 - distance / length).toFixed(4)),
        rank: entry.rank,
      };
    })
    .filter((candidate) => candidate.similarity > 0)
    .sort((left, right) => left.distance - right.distance || left.rank - right.rank)
    .slice(0, 5)
    .map(({ rank, ...candidate }) => candidate);
}

function collisionResult(method, entries) {
  return {
    status: "needs_confirmation",
    reason: `${method}_collision`,
    candidates: entries
      .slice()
      .sort((left, right) => left.rank - right.rank)
      .map((entry) => ({ standardName: entry.standardName })),
  };
}

export function resolveCityName(rawName, cityOrder, mappingRows = []) {
  const raw = String(rawName ?? "");
  const exact = cityOrder.byExact.get(raw);
  if (exact) {
    return { status: "matched", standardName: exact.standardName, method: "exact" };
  }

  const normalized = normalizeSafeCityName(raw);
  const normalizedMatches = cityOrder.normalizedIndex.get(normalized) ?? [];
  if (normalizedMatches.length === 1) {
    return {
      status: "matched",
      standardName: normalizedMatches[0].standardName,
      method: "safe_normalization",
    };
  }
  if (normalizedMatches.length > 1) {
    return collisionResult("safe_normalization", normalizedMatches);
  }

  const suffixMatches = cityOrder.suffixIndex.get(baseCityName(raw)) ?? [];
  if (suffixMatches.length === 1) {
    return {
      status: "matched",
      standardName: suffixMatches[0].standardName,
      method: "unique_suffix",
    };
  }
  if (suffixMatches.length > 1) {
    return collisionResult("unique_suffix", suffixMatches);
  }

  const mappings = validateMappingRows(mappingRows, cityOrder);
  const mapped = mappings.bySourceNormalized.get(normalized);
  if (mapped) {
    return {
      status: "matched",
      standardName: mapped.standardName,
      method: "confirmed_mapping",
    };
  }

  return {
    status: "needs_confirmation",
    reason: normalized ? "fuzzy_candidate_only" : "blank_city_name",
    candidates: normalized ? rankedCandidates(raw, cityOrder) : [],
  };
}