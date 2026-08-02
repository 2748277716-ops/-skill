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
    "OOXML 定向读取",
    "孤立关系",
    "实际引用部件缺失",
    "INVALID_WORKBOOK_PACKAGE",
    "精简模式只物化",
    "artifact-tool 仅用于结果写出和复核",
    "逐键反向核验",
  ]) {
    assert.ok(text.includes(phrase), `Missing phrase: ${phrase}`);
  }
});
test("runtime helper does not hard-code a Windows user profile", () => {
  const helperUrl = new URL("./helpers/artifact-runtime.mjs", import.meta.url);
  const helperText = fs.readFileSync(helperUrl, "utf8");
  const userProfilePrefix = ["C:", "Users"].join("\\");
  assert.equal(helperText.includes(userProfilePrefix), false);
});
