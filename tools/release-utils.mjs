import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "../app/vendor/jszip.min.js";

const ZIP_DATE = new Date("2020-01-01T00:00:00.000Z");

export async function collectTree(
  sourceRoot,
  destinationRoot,
  options = {},
) {
  const entries = [];
  const exclude = options.exclude || (() => false);

  async function visit(currentPath, relativePath = "") {
    const children = await fs.readdir(currentPath, {
      withFileTypes: true,
    });
    children.sort((left, right) =>
      left.name.localeCompare(right.name, "zh-CN"),
    );
    for (const child of children) {
      const childRelative = relativePath
        ? path.posix.join(relativePath, child.name)
        : child.name;
      if (exclude(childRelative, child)) {
        continue;
      }
      const childPath = path.join(currentPath, child.name);
      if (child.isSymbolicLink()) {
        throw new Error(`发布白名单不允许符号链接：${childPath}`);
      }
      if (child.isDirectory()) {
        await visit(childPath, childRelative);
        continue;
      }
      if (!child.isFile()) {
        throw new Error(`发布白名单包含未知文件类型：${childPath}`);
      }
      entries.push({
        sourcePath: childPath,
        destination: path.posix.join(
          destinationRoot,
          childRelative,
        ),
      });
    }
  }

  await visit(sourceRoot);
  return entries;
}

export function mappedFile(sourcePath, destination) {
  return { sourcePath, destination };
}

export function memoryFile(destination, content) {
  return {
    destination,
    content: Buffer.from(content, "utf8"),
  };
}

export async function readEntry(entry) {
  if (entry.content) {
    return entry.content;
  }
  return fs.readFile(entry.sourcePath);
}

export async function buildManifest(entries) {
  const files = [];
  for (const entry of entries) {
    const content = await readEntry(entry);
    files.push({
      path: entry.destination,
      bytes: content.length,
      sha256: crypto.createHash("sha256").update(content).digest("hex"),
    });
  }
  files.sort((left, right) =>
    left.path.localeCompare(right.path, "zh-CN"),
  );
  return {
    format: 1,
    generatedAt: new Date().toISOString(),
    privacy: {
      realPayrollFilesIncluded: false,
      rowLevelPersonalDataIncluded: false,
      gitHistoryIncluded: false,
    },
    files,
  };
}

export function assertUniqueDestinations(entries) {
  const seen = new Set();
  for (const entry of entries) {
    if (seen.has(entry.destination)) {
      throw new Error(`发布路径重复：${entry.destination}`);
    }
    seen.add(entry.destination);
  }
}

export function assertAllowedFiles(entries, forbiddenExtensions) {
  for (const entry of entries) {
    const lower = entry.destination.toLowerCase();
    const extension = path.posix.extname(lower);
    if (
      forbiddenExtensions.has(extension) ||
      path.posix.basename(lower) === ".ds_store"
    ) {
      throw new Error(`发布包包含不允许的数据文件：${entry.destination}`);
    }
  }
}

export async function assertTextIsSafe(
  entries,
  patterns,
  options = {},
) {
  const skip = options.skip || (() => false);
  const textExtensions = new Set([
    ".css",
    ".html",
    ".js",
    ".json",
    ".md",
    ".mjs",
    ".txt",
  ]);
  for (const entry of entries) {
    if (
      skip(entry.destination) ||
      !textExtensions.has(path.posix.extname(entry.destination))
    ) {
      continue;
    }
    const text = (await readEntry(entry)).toString("utf8");
    for (const [label, pattern] of patterns) {
      if (pattern.test(text)) {
        throw new Error(
          `发布隐私检查失败（${label}）：${entry.destination}`,
        );
      }
    }
  }
}

export async function writeZipArchive(
  entries,
  packageName,
  archivePath,
) {
  assertUniqueDestinations(entries);
  const zip = new JSZip();
  const sorted = [...entries].sort((left, right) =>
    left.destination.localeCompare(right.destination, "zh-CN"),
  );
  for (const entry of sorted) {
    zip.file(
      path.posix.join(packageName, entry.destination),
      await readEntry(entry),
      {
        binary: true,
        date: ZIP_DATE,
        createFolders: true,
        unixPermissions: "0644",
      },
    );
  }
  const archive = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
    platform: "UNIX",
  });
  await fs.mkdir(path.dirname(archivePath), { recursive: true });
  await fs.writeFile(archivePath, archive);
  const checksum = crypto
    .createHash("sha256")
    .update(archive)
    .digest("hex");
  const checksumPath = `${archivePath}.sha256`;
  await fs.writeFile(
    checksumPath,
    `${checksum}  ${path.basename(archivePath)}\n`,
    "utf8",
  );
  return { archivePath, checksumPath, checksum };
}
