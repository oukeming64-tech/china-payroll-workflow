const path = require("node:path");

module.exports = {
  testDir: __dirname,
  testMatch: "**/*.spec.js",
  timeout: 120_000,
  expect: { timeout: 30_000 },
  reporter: "line",
  outputDir: path.resolve(__dirname, "../output/playwright"),
  webServer: {
    command: "node serve.mjs",
    url: "http://127.0.0.1:8765",
    reuseExistingServer: true,
    timeout: 30_000,
  },
  use: {
    baseURL: "http://127.0.0.1:8765",
    channel: process.env.PLAYWRIGHT_CHANNEL || "chrome",
    headless: true,
    locale: "zh-CN",
    timezoneId: "Asia/Shanghai",
    viewport: { width: 1440, height: 900 },
  },
};
