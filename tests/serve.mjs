import fs from "node:fs/promises";
import http from "node:http";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const testDirectory = path.dirname(fileURLToPath(import.meta.url));
const appDirectory = path.resolve(testDirectory, "../app");
const host = "127.0.0.1";
const port = Number(process.env.PAYROLL_TEST_PORT || 8765);
const contentTypes = new Map([
  [".css", "text/css; charset=utf-8"],
  [".html", "text/html; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".txt", "text/plain; charset=utf-8"],
]);

function safePath(urlValue) {
  const pathname = decodeURIComponent(
    new URL(urlValue, `http://${host}:${port}`).pathname,
  );
  const relative = pathname === "/" ? "index.html" : pathname.slice(1);
  const resolved = path.resolve(appDirectory, relative);
  if (
    resolved !== appDirectory &&
    !resolved.startsWith(`${appDirectory}${path.sep}`)
  ) {
    return null;
  }
  return resolved;
}

const server = http.createServer(async (request, response) => {
  const filePath = safePath(request.url || "/");
  if (!filePath || !["GET", "HEAD"].includes(request.method || "")) {
    response.writeHead(filePath ? 405 : 403);
    response.end();
    return;
  }
  try {
    const content = await fs.readFile(filePath);
    response.writeHead(200, {
      "Content-Type":
        contentTypes.get(path.extname(filePath).toLowerCase()) ||
        "application/octet-stream",
      "Cache-Control": "no-store",
    });
    response.end(request.method === "HEAD" ? undefined : content);
  } catch (error) {
    response.writeHead(error.code === "ENOENT" ? 404 : 500);
    response.end();
  }
});

server.listen(port, host, () => {
  console.log(`payroll test server ready at http://${host}:${port}`);
});

for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => {
    server.close(() => process.exit(0));
  });
}
