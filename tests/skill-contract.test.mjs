import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";

const skillUrl = new URL("../SKILL.md", import.meta.url);

test("SKILL.md exposes required metadata and safety text", () => {
  const text = fs.readFileSync(skillUrl, "utf8");
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
  ]) {
    assert.ok(text.includes(phrase), `Missing phrase: ${phrase}`);
  }
});