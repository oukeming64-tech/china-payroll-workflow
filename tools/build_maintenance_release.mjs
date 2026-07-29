#!/usr/bin/env node

import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

import {
  assertAllowedFiles,
  assertTextIsSafe,
  buildManifest,
  collectTree,
  mappedFile,
  memoryFile,
  writeZipArchive,
} from "./release-utils.mjs";

const toolDirectory = path.dirname(fileURLToPath(import.meta.url));
const projectDirectory = path.dirname(toolDirectory);
const releaseVersion =
  process.env.PAYROLL_RELEASE_VERSION?.trim() || "1.2.0";

if (!/^\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$/.test(releaseVersion)) {
  throw new Error(`发布版本号无效：${releaseVersion}`);
}

const packageName = `工资表工具-维护版-v${releaseVersion}`;
const archiveName = `china-payroll-workflow-maintenance-v${releaseVersion}`;
const archivePath = path.join(
  projectDirectory,
  "output",
  "releases",
  `${archiveName}.zip`,
);

const entries = [
  ...(await collectTree(
    path.join(projectDirectory, "app"),
    "app",
    {
      exclude: (relativePath) =>
        relativePath === "README.md" ||
        path.posix.basename(relativePath) === ".DS_Store",
    },
  )),
  ...(await collectTree(
    path.join(projectDirectory, "tests"),
    "tests",
  )),
  ...(await collectTree(
    path.join(projectDirectory, "release", "universal"),
    "release/universal",
  )),
  mappedFile(
    path.join(toolDirectory, "build_release"),
    "tools/build_release",
  ),
  mappedFile(
    path.join(toolDirectory, "build_release.mjs"),
    "tools/build_release.mjs",
  ),
  mappedFile(
    path.join(toolDirectory, "build_maintenance_release.mjs"),
    "tools/build_maintenance_release.mjs",
  ),
  mappedFile(
    path.join(toolDirectory, "check_app_architecture.mjs"),
    "tools/check_app_architecture.mjs",
  ),
  mappedFile(
    path.join(
      toolDirectory,
      "officecrypto-browser-crypto-shim.cjs",
    ),
    "tools/officecrypto-browser-crypto-shim.cjs",
  ),
  mappedFile(
    path.join(toolDirectory, "release-utils.mjs"),
    "tools/release-utils.mjs",
  ),
  ...[
    ["README.md", "README.md"],
    ["CHANGELOG.md", "CHANGELOG.md"],
    ["CODEX.md", "CODEX.md"],
    ["AGENTS.md", "AGENTS.md"],
    ["package.json", "package.json"],
    ["package-lock.json", "package-lock.json"],
    [".gitignore", ".gitignore"],
    ["LICENSE", "LICENSE"],
    ["NOTICE", "NOTICE"],
    ["THIRD_PARTY_NOTICES.md", "THIRD_PARTY_NOTICES.md"],
    ["CONTRIBUTING.md", "CONTRIBUTING.md"],
    ["给领导的说明.md", "给领导的说明.md"],
    [
      "requirements/PRODUCT_SPEC.md",
      "requirements/PRODUCT_SPEC.md",
    ],
    [
      "requirements/ACCEPTANCE.md",
      "requirements/ACCEPTANCE.md",
    ],
  ].map(([source, destination]) =>
    mappedFile(path.join(projectDirectory, source), destination),
  ),
];

assertAllowedFiles(
  entries,
  new Set([
    ".csv",
    ".doc",
    ".docx",
    ".pdf",
    ".tsv",
    ".xls",
    ".xlsb",
    ".xlsm",
    ".xlsx",
    ".zip",
  ]),
);
await assertTextIsSafe(
  entries,
  [
    ["本机用户路径", /\/Users\/admi\x6e/i],
    ["内部来源目录", /excel_source\x73\//i],
    [
      "内部工作簿名",
      /\u660c\u5e73|\u674e\u603b|\u6d89\u5bc6/,
    ],
    ["真实工作簿指纹", /d006\x6333e/i],
    [
      "身份证号",
      /(?<!\d)[1-9]\d{5}(?:18|19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[0-9Xx](?!\d)/,
    ],
    ["手机号", /(?<!\d)1[3-9]\d{9}(?!\d)/],
    ["银行卡号候选", /(?<!\d)\d{16,19}(?!\d)/],
    [
      "密钥候选",
      /gh[pousr]_[A-Za-z0-9_]{20,}|github_pat_[A-Za-z0-9_]{20,}|sk-[A-Za-z0-9_-]{16,}|AKIA[0-9A-Z]{16}|BEGIN (?:RSA |OPENSSH |EC )?PRIVATE KEY/,
    ],
  ],
  {
    skip: (destination) =>
      destination.includes("/vendor/") &&
      destination.endsWith(".min.js"),
  },
);

const manifest = await buildManifest(entries);
entries.push(
  memoryFile(
    "PACKAGE-MANIFEST.json",
    `${JSON.stringify(manifest, null, 2)}\n`,
  ),
);

const result = await writeZipArchive(
  entries,
  packageName,
  archivePath,
);
console.log(result.archivePath);
console.log(result.checksumPath);
