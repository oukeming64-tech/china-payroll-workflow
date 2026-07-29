#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const ROOT = process.cwd();
const APP_ROOT = path.join(ROOT, "app");
const LIMITS = [
  {
    label: "规则模块",
    directory: path.join(APP_ROOT, "rules"),
    maximum: 400,
  },
  {
    label: "Excel 模块",
    directory: path.join(APP_ROOT, "excel"),
    maximum: 500,
  },
  {
    label: "界面模块",
    directory: path.join(APP_ROOT, "ui"),
    maximum: 500,
  },
];
const FACADES = [
  path.join(APP_ROOT, "xlsx-engine.js"),
  path.join(APP_ROOT, "payroll-engine.js"),
  path.join(APP_ROOT, "app.js"),
];
const EXPECTED_SCRIPT_ORDER = [
  "./vendor/jszip.min.js",
  "./vendor/xlsx.full.min.js",
  "./vendor/officecrypto-tool.min.js",
  "./core/namespace.js",
  "./excel/xml.js",
  "./excel/table.js",
  "./excel/table-regions.js",
  "./excel/mutations.js",
  "./excel/periods.js",
  "./excel/external-links.js",
  "./excel/external-detach.js",
  "./excel/formulas.js",
  "./excel/schema.js",
  "./excel/cumulative.js",
  "./excel/personnel.js",
  "./excel/payroll-sync.js",
  "./excel/workbook.js",
  "./excel/source-workbook.js",
  "./excel/docx-tables.js",
  "./xlsx-engine.js",
  "./rules/common.js",
  "./rules/monthly-routes.js",
  "./rules/monthly-business.js",
  "./rules/january-rollover.js",
  "./rules/natural-language.js",
  "./rules/tabular-changes.js",
  "./rules/external-source.js",
  "./rules/monthly-change-sources.js",
  "./rules/salary-event-proration.js",
  "./rules/employment-events.js",
  "./rules/salary-events.js",
  "./rules/attendance-deductions.js",
  "./rules/business-sources.js",
  "./rules/workbook-evidence-changes.js",
  "./rules/workbook-evidence.js",
  "./rules/attachment-periods.js",
  "./rules/labor-fee-review.js",
  "./rules/attachment-resolution.js",
  "./rules/attachment-batch.js",
  "./rules/workbook-personnel.js",
  "./rules/social-base-july.js",
  "./payroll-engine.js",
  "./ui/state.js",
  "./ui/password-flow.js",
  "./ui/render.js",
  "./ui/requirements-render.js",
  "./ui/workbook-flow.js",
  "./ui/attachment-flow.js",
  "./ui/personnel-sync-flow.js",
  "./ui/source-regions-flow.js",
  "./ui/change-flow.js",
  "./ui/social-base-flow.js",
  "./ui/export-flow.js",
  "./app.js",
];

async function lineCount(filePath) {
  const text = await fs.readFile(filePath, "utf8");
  return text.split(/\r?\n/).length - (text.endsWith("\n") ? 1 : 0);
}

async function javascriptFiles(directory) {
  return (await fs.readdir(directory, { withFileTypes: true }))
    .filter((entry) => entry.isFile() && entry.name.endsWith(".js"))
    .map((entry) => path.join(directory, entry.name))
    .sort();
}

const failures = [];
const checks = [];
for (const group of LIMITS) {
  for (const filePath of await javascriptFiles(group.directory)) {
    const lines = await lineCount(filePath);
    checks.push({
      file: path.relative(ROOT, filePath),
      lines,
      maximum: group.maximum,
    });
    if (lines > group.maximum) {
      failures.push(
        `${group.label} ${path.relative(ROOT, filePath)} 为 ${lines} 行，超过 ${group.maximum} 行`,
      );
    }
  }
}
for (const filePath of FACADES) {
  const lines = await lineCount(filePath);
  checks.push({
    file: path.relative(ROOT, filePath),
    lines,
    maximum: 100,
  });
  if (lines > 100) {
    failures.push(
      `入口文件 ${path.relative(ROOT, filePath)} 为 ${lines} 行，超过 100 行`,
    );
  }
}

const indexHtml = await fs.readFile(path.join(APP_ROOT, "index.html"), "utf8");
const actualScriptOrder = [
  ...indexHtml.matchAll(/<script\s+src="([^"]+)"><\/script>/g),
].map((match) => match[1].replace(/[?#].*$/, ""));
if (JSON.stringify(actualScriptOrder) !== JSON.stringify(EXPECTED_SCRIPT_ORDER)) {
  failures.push("app/index.html 的模块加载顺序与架构契约不一致");
}

if (failures.length) {
  console.error(`架构检查失败：\n- ${failures.join("\n- ")}`);
  process.exitCode = 1;
} else {
  const largest = checks.toSorted((left, right) => right.lines - left.lines)[0];
  console.log(
    `架构检查通过：${checks.length} 个模块；最大文件 ${largest.file} 为 ${largest.lines} 行。`,
  );
}
