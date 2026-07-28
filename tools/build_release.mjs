#!/usr/bin/env node

import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

import {
  assertAllowedFiles,
  collectTree,
  mappedFile,
  writeZipArchive,
} from "./release-utils.mjs";

const toolDirectory = path.dirname(fileURLToPath(import.meta.url));
const projectDirectory = path.dirname(toolDirectory);
const releaseVersion =
  process.env.PAYROLL_RELEASE_VERSION?.trim() || "1.1.1";

if (!/^\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$/.test(releaseVersion)) {
  throw new Error(`发布版本号无效：${releaseVersion}`);
}

const packageName = `工资表工具-通用版-v${releaseVersion}`;
const archiveName = `china-payroll-workflow-universal-v${releaseVersion}`;
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
  mappedFile(
    path.join(
      projectDirectory,
      "release",
      "universal",
      "打开工资表工具.html",
    ),
    "打开工资表工具.html",
  ),
  mappedFile(
    path.join(
      projectDirectory,
      "release",
      "universal",
      "使用说明.txt",
    ),
    "使用说明.txt",
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

const result = await writeZipArchive(
  entries,
  packageName,
  archivePath,
);
console.log(result.archivePath);
console.log(result.checksumPath);
