export function parseYear(value) {
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    const year = value.getUTCFullYear();
    return year >= 1000 && year <= 9999 ? year : null;
  }
  if (typeof value === "number" && Number.isInteger(value)) {
    return value >= 1000 && value <= 9999 ? value : null;
  }
  const text = String(value ?? "").trim();
  return /^\d{4}$/u.test(text) ? Number(text) : null;
}
