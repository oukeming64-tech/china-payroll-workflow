const fs = require("node:fs");
const path = require("node:path");
const { execFileSync } = require("node:child_process");
const JSZip = require("jszip");
const { test, expect } = require("playwright/test");

const projectRoot = path.resolve(__dirname, "..");
const fixtureRoot = path.join(projectRoot, "output", "test-fixtures");
const regularWorkbook = path.join(
  fixtureRoot,
  "2025.11合成工资表.xlsx",
);
const annualPayrollWorkbooks = Array.from({ length: 12 }, (_, index) =>
  path.join(
    fixtureRoot,
    "2025全年",
    `2025.${index + 1}合成工资表.xlsx`,
  ),
);
const regularAttachments = [
  path.join(
    fixtureRoot,
    "2025.12工资附件",
    "202512_正常工资薪金所得.xls",
  ),
  path.join(
    fixtureRoot,
    "2025.12工资附件",
    "2025.12示例公司社保公积金表.xlsx",
  ),
];
const unmatchedRegularAttachments = [
  path.join(
    fixtureRoot,
    "2025.12工资附件",
    "202512_正常工资薪金所得_含未匹配.xls",
  ),
  regularAttachments[1],
];
const januaryAttachments = [
  path.join(
    fixtureRoot,
    "2026.01工资附件",
    "202601_正常工资薪金所得.xls",
  ),
  path.join(
    fixtureRoot,
    "2026.01工资附件",
    "2026.1示例公司社保公积金表.xlsx",
  ),
  path.join(
    fixtureRoot,
    "2026.01工资附件",
    "2026年劳务费-示例.xlsx",
  ),
];
const regularOutput = path.join(
  projectRoot,
  "output",
  "playwright",
  "2025.12合成工资表_待复核.xlsx",
);
const januaryOutput = path.join(
  projectRoot,
  "output",
  "playwright",
  "2026.01合成工资表_待复核.xlsx",
);

test.beforeAll(() => {
  execFileSync(process.execPath, [
    path.join(__dirname, "create-synthetic-fixtures.mjs"),
  ]);
});

async function selectBaseWorkbook(page, workbook = regularWorkbook) {
  await page.goto("/");
  await page.locator("#baseWorkbookInput").setInputFiles(workbook);
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
}

async function encryptedWorkbookPayload(
  page,
  sourcePath,
  password,
  options = {},
) {
  const encryptedBase64 = await page.evaluate(
    async ({ base64, password: testPassword, type }) => {
      const binary = window.atob(base64);
      const plain = Uint8Array.from(
        binary,
        (character) => character.charCodeAt(0),
      );
      const encryptionOptions = { password: testPassword };
      if (type) {
        encryptionOptions.type = type;
      }
      const encrypted = await Promise.resolve(
        window.OfficeCrypto.encrypt(plain, encryptionOptions),
      );
      let encryptedBinary = "";
      for (let offset = 0; offset < encrypted.length; offset += 0x8000) {
        encryptedBinary += String.fromCharCode(
          ...encrypted.subarray(offset, offset + 0x8000),
        );
      }
      return window.btoa(encryptedBinary);
    },
    {
      base64: fs.readFileSync(sourcePath).toString("base64"),
      password,
      type: options.type || "",
    },
  );
  return {
    name: options.name || path.basename(sourcePath),
    mimeType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    buffer: Buffer.from(encryptedBase64, "base64"),
  };
}

async function openRegularWorkspace(page) {
  await selectBaseWorkbook(page);
  await expect(page.locator("#workspaceView")).toBeVisible();
}

async function openJanuaryWorkspace(page) {
  await selectBaseWorkbook(page, annualPayrollWorkbooks[11]);
  await expect(page.locator("#crossYearSetup")).toBeVisible();
  await expect(page.locator("#workspaceView")).toBeHidden();
  await page
    .locator("#crossYearFilesInput")
    .setInputFiles(annualPayrollWorkbooks);
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await expect(page.locator("#workspaceView")).toBeVisible();
}

async function applyAttachments(page, attachments) {
  await page.locator('[data-tab="attachments"]').click();
  await page.locator("#attachmentFilesInput").setInputFiles(attachments);
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await page.locator("#applyAttachmentsBtn").click();
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await expect(page.locator("#attachmentSummary")).toContainText(
    "无外链",
  );
}

async function saveExport(page, destination) {
  await page.locator('[data-tab="review"]').click();
  await page.locator("#exportAcknowledged").check();
  const downloadPromise = page.waitForEvent("download");
  await page.locator("#exportWorkbookBtn").click();
  const download = await downloadPromise;
  await download.saveAs(destination);
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
}

async function inspectExternalPackage(filePath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(filePath));
  const workbookXml = await zip.file("xl/workbook.xml").async("string");
  const worksheetPaths = Object.keys(zip.files).filter((name) =>
    /^xl\/worksheets\/sheet\d+\.xml$/.test(name),
  );
  const formulas = [];
  let cachedValues = 0;
  for (const worksheetPath of worksheetPaths) {
    const xml = await zip.file(worksheetPath).async("string");
    formulas.push(...[...xml.matchAll(/<f(?:\s[^>]*)?>([\s\S]*?)<\/f>/g)].map(
      (match) => match[1],
    ));
    cachedValues += (xml.match(/<v>/g) || []).length;
  }
  return {
    hasExternalReferences: /<externalReferences\b/.test(workbookXml),
    externalParts: Object.keys(zip.files).filter((name) =>
      name.startsWith("xl/externalLinks/externalLink"),
    ).length,
    externalFormulaCount: formulas.filter((formula) => /\[\d+\]/.test(formula))
      .length,
    cachedValues,
  };
}

test("开始页保持简洁，只接收上月完整工资表", async ({ page }) => {
  const pageErrors = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  await page.goto("/");

  await expect(
    page.getByRole("heading", { name: "生成下月工资表" }),
  ).toBeVisible();
  await expect(page.getByText("选择上月完整工资表")).toBeVisible();
  await expect(page.locator("#targetMonthInput")).toHaveCount(0);
  await expect(page.locator("#targetTemplateInput")).toHaveCount(0);
  await expect(page.locator("#startBtn")).toHaveCount(0);
  await expect(page.getByText(/本地处理|不上传|确认后写入/)).toHaveCount(0);
  expect(pageErrors).toEqual([]);
});

test("加密工作簿弹窗重试并支持不同文件使用不同密码", async ({
  page,
}) => {
  const pageErrors = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  await page.goto("/");

  const encryptedBase = await encryptedWorkbookPayload(
    page,
    regularWorkbook,
    "synthetic-base-password",
    { type: "standard" },
  );
  await page.locator("#baseWorkbookInput").setInputFiles(encryptedBase);
  const dialog = page.locator("#workbookPasswordDialog");
  await expect(dialog).toBeVisible();
  await expect(page.locator("#workbookPasswordFile")).toHaveText(
    encryptedBase.name,
  );
  await page.getByRole("button", { name: "取消" }).click();
  await expect(dialog).toBeHidden();
  await expect(page.locator("#loadingOverlay")).toBeHidden();
  await expect(page.locator("#baseFileLabel")).toHaveText(
    "选择上月完整工资表",
  );

  await page.locator("#baseWorkbookInput").setInputFiles(encryptedBase);
  await expect(dialog).toBeVisible();
  await page.locator("#workbookPasswordInput").fill("wrong-password");
  await page.getByRole("button", { name: "打开工作簿" }).click();
  await expect(page.locator("#workbookPasswordError")).toContainText(
    "密码不正确",
  );
  await expect(page.locator("#workbookPasswordInput")).toHaveValue("");

  await page
    .locator("#workbookPasswordInput")
    .fill("synthetic-base-password");
  await page.getByRole("button", { name: "打开工作簿" }).click();
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await expect(page.locator("#workspaceView")).toBeVisible();
  await expect(page.locator("#targetPeriodText")).toHaveText("2025年12月");

  const encryptedInsurance = await encryptedWorkbookPayload(
    page,
    regularAttachments[1],
    "synthetic-attachment-password",
  );
  const taxPayload = {
    name: path.basename(regularAttachments[0]),
    mimeType: "application/vnd.ms-excel",
    buffer: fs.readFileSync(regularAttachments[0]),
  };
  await page.locator('[data-tab="attachments"]').click();
  await page
    .locator("#attachmentFilesInput")
    .setInputFiles([taxPayload, encryptedInsurance]);
  await expect(dialog).toBeVisible();
  await expect(page.locator("#workbookPasswordFile")).toHaveText(
    encryptedInsurance.name,
  );
  await page
    .locator("#workbookPasswordInput")
    .fill("synthetic-attachment-password");
  await page.getByRole("button", { name: "打开工作簿" }).click();
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });

  const evidence = await page.evaluate(() => {
    const state = window.PayrollLocal.ui.state;
    return {
      attachmentErrors: state.attachments.errors,
      categories: state.attachments.results.map(
        (result) => result.category,
      ),
      passwordState: Object.hasOwn(state, "password"),
      localStorageKeys: Object.keys(window.localStorage),
      sessionStorageKeys: Object.keys(window.sessionStorage),
    };
  });
  expect(evidence).toEqual({
    attachmentErrors: [],
    categories: ["个税工资薪金附件", "社保 / 公积金附件"],
    passwordState: false,
    localStorageKeys: [],
    sessionStorageKeys: [],
  });
  await expect(dialog).toBeHidden();
  await expect(page.locator("#workbookPasswordInput")).toHaveValue("");
  expect(pageErrors).toEqual([]);
});

test("普通月份用目标月附件写入人员字段并生成无外链文件", async ({
  page,
}) => {
  const pageErrors = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  await openRegularWorkspace(page);

  await expect(page.locator("#basePeriodText")).toHaveText("2025年11月");
  await expect(page.locator("#targetPeriodText")).toHaveText("2025年12月");
  await expect(page.locator("#fieldStat")).toHaveText("59");
  await expect(page.locator("#peopleStat")).toHaveText("3");
  await expect(page.locator("#externalStat")).toHaveText("4");
  await expect(page.locator("#exportWorkbookBtn")).toBeDisabled();

  const before = await page.evaluate(() => {
    const state = window.PayrollLocal.ui.state;
    const people = window.PayrollEngine.buildPeople(state.table).people;
    const first = people.find((person) => person.employeeId === "TEST-001");
    return {
      route: state.route.id,
      cumulativeWrites: state.cumulativeResult.written,
      required: state.attachments.required.map((item) => item.category),
      firstOtherPay: first.row.get("其他工资"),
      firstFinalPay: first.row.get("实发合计"),
      unresolvedFormulas: state.formulaDiagnostics.unresolvedFormulaNodes,
    };
  });
  expect(before).toEqual({
    route: "regular-month",
    cumulativeWrites: 14,
    required: [
      "个税工资薪金附件",
      "社保 / 公积金附件",
    ],
    firstOtherPay: 0,
    firstFinalPay: 5600,
    unresolvedFormulas: 0,
  });

  await applyAttachments(page, regularAttachments);
  const after = await page.evaluate(() => {
    const state = window.PayrollLocal.ui.state;
    const first = window.PayrollEngine
      .buildPeople(state.table)
      .people.find((person) => person.employeeId === "TEST-001");
    return {
      errors: state.attachments.errors,
      categories: state.attachments.results.map(
        (result) => result.category,
      ),
      externalLinks: state.externalLinks.length,
      externalFormulas: state.formulaDiagnostics.externalFormulaNodes,
      frozenLegacy: state.attachments.detached.frozenLegacyFormulaCount,
      tax: first.row.get("累计子女教育"),
      pension: first.row.get("代扣养老保险"),
      medical: first.row.get("代扣医疗保险"),
      unemployment: first.row.get("代扣失业保险"),
      housingFund: first.row.get("代扣房积金"),
      personalTotal: first.row.get("个人承担社保部分扣款合计"),
      companyTotal: first.row.get("企业承担社保部分扣款合计"),
    };
  });
  expect(after).toEqual({
    errors: [],
    categories: ["个税工资薪金附件", "社保 / 公积金附件"],
    externalLinks: 0,
    externalFormulas: 0,
    frozenLegacy: 1,
    tax: 120,
    pension: 100,
    medical: 20,
    unemployment: 10,
    housingFund: 200,
    personalTotal: 330,
    companyTotal: 1870,
  });
  await saveExport(page, regularOutput);
  const exported = await inspectExternalPackage(regularOutput);
  expect(exported).toEqual({
    hasExternalReferences: false,
    externalParts: 0,
    externalFormulaCount: 0,
    cachedValues: expect.any(Number),
  });
  expect(exported.cachedValues).toBeGreaterThan(0);
  expect(pageErrors).toEqual([]);
});

test("一月跨年消费全年历史并用三类本地附件生成无外链版本", async ({
  page,
}) => {
  const pageErrors = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  await openJanuaryWorkspace(page);

  const evidence = await page.evaluate(async () => {
    const state = window.PayrollLocal.ui.state;
    const headers = state.table.headers.map((header) => header.name);
    const labor = window.PayrollEngine
      .buildPeople(state.table)
      .people.find((person) => person.name === "测试劳务");
    const current = state.table.headers.find(
      (header) => header.name === "扣减税额",
    );
    const previousFields = state.table.headers.filter((header) =>
      /^之前月份累计/.test(header.name),
    );
    return {
      route: state.route.id,
      annualPeriods: state.annualHistory.periods.length,
      annualErrors: state.annualHistory.errors,
      formulaVersionEvidence:
        state.annualHistory.formulaVersionEvidence,
      fieldCount: headers.length,
      addedFields: state.rolloverPlan.addedFields,
      addedAfter:
        headers.slice(
          headers.indexOf("累计应计工资") + 1,
          headers.indexOf("累计应计工资") + 5,
        ),
      previousFormulaCells: previousFields.reduce(
        (count, header) =>
          count +
          state.table.rows.filter(
            (row) => row.cells.get(header.column)?.hasFormula,
          ).length,
        0,
      ),
      laborCurrent: labor.row.get(current.name),
      required: state.attachments.required.map(
        (item) => item.category,
      ),
      unresolvedFormulas: state.formulaDiagnostics.unresolvedFormulaNodes,
    };
  });
  expect(evidence).toEqual({
    route: "january-rollover",
    annualPeriods: 12,
    annualErrors: [],
    formulaVersionEvidence: 11,
    fieldCount: 63,
    addedFields: [
      "扣减税额",
      "之前月份累计扣减税额",
      "累计扣减税额",
      "累计应纳税所得额",
    ],
    addedAfter: [
      "扣减税额",
      "之前月份累计扣减税额",
      "累计扣减税额",
      "累计应纳税所得额",
    ],
    previousFormulaCells: 0,
    laborCurrent: null,
    required: [
      "个税工资薪金附件",
      "社保 / 公积金附件",
      "劳务费附件",
    ],
    unresolvedFormulas: 0,
  });
  await applyAttachments(page, januaryAttachments);
  const resolved = await page.evaluate(() => {
    const state = window.PayrollLocal.ui.state;
    const people = window.PayrollEngine.buildPeople(state.table).people;
    const employee = people.find(
      (person) => person.employeeId === "TEST-001",
    );
    const labor = people.find((person) => person.name === "测试劳务");
    const accumulated = state.table.headers.find(
      (header) => header.name === "累计扣减税额",
    );
    return {
      errors: state.attachments.errors,
      matched: state.attachments.results.map((result) => [
        result.category,
        result.matchedPeople,
        result.expectedPeople,
      ]),
      externalLinks: state.externalLinks.length,
      externalFormulas: state.formulaDiagnostics.externalFormulaNodes,
      employeeTax: employee.row.get("累计子女教育"),
      laborCurrent: labor.row.get("扣减税额"),
      laborAccumulatedFormula:
        labor.row.cells.get(accumulated.column).formula,
    };
  });
  expect(resolved).toEqual({
    errors: [],
    matched: [
      ["个税工资薪金附件", 2, 2],
      ["社保 / 公积金附件", 2, 2],
      ["劳务费附件", 1, 1],
    ],
    externalLinks: 0,
    externalFormulas: 0,
    employeeTax: 10,
    laborCurrent: 40,
    laborAccumulatedFormula: "T5+U5",
  });
  await saveExport(page, januaryOutput);
  const exported = await inspectExternalPackage(januaryOutput);
  expect(exported.hasExternalReferences).toBe(false);
  expect(exported.externalParts).toBe(0);
  expect(exported.externalFormulaCount).toBe(0);
  expect(exported.cachedValues).toBeGreaterThan(0);
  expect(pageErrors).toEqual([]);

  await selectBaseWorkbook(page, januaryOutput);
  await expect(page.locator("#workspaceView")).toBeVisible();
  const nextMonth = await page.evaluate(() => ({
    route: window.PayrollLocal.ui.state.route.id,
    targetPeriod: window.PayrollLocal.ui.state.targetPeriod,
    links: window.PayrollLocal.ui.state.externalLinks.length,
    cumulativeWrites:
      window.PayrollLocal.ui.state.cumulativeResult.written,
    required: window.PayrollLocal.ui.state.attachments.required.map(
      (item) => item.category,
    ),
  }));
  expect(nextMonth).toEqual({
    route: "regular-month",
    targetPeriod: "2026-02",
    links: 0,
    cumulativeWrites: 16,
    required: [
      "个税工资薪金附件",
      "社保 / 公积金附件",
      "劳务费附件",
    ],
  });
});

test("工资附件缺失或含未匹配人员时停止写入", async ({ page }) => {
  await openRegularWorkspace(page);
  await page.locator('[data-tab="attachments"]').click();
  await page
    .locator("#attachmentFilesInput")
    .setInputFiles([regularAttachments[0]]);
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await expect(page.locator("#attachmentErrors")).toContainText(
    "缺少社保 / 公积金附件",
  );
  await expect(page.locator("#applyAttachmentsBtn")).toBeDisabled();

  await page
    .locator("#attachmentFilesInput")
    .setInputFiles(unmatchedRegularAttachments);
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await expect(page.locator("#attachmentErrors")).toContainText(
    "来源表有 1 人未匹配工资表",
  );
  await expect(page.locator("#applyAttachmentsBtn")).toBeDisabled();
  expect(
    await page.evaluate(
      () => window.PayrollLocal.ui.state.attachments.applied,
    ),
  ).toBe(false);
});

test("跨年历史不完整时停止生成并显示缺月错误", async ({ page }) => {
  await selectBaseWorkbook(page, annualPayrollWorkbooks[11]);
  await page
    .locator("#crossYearFilesInput")
    .setInputFiles(annualPayrollWorkbooks.slice(1));
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await expect(page.locator("#workspaceView")).toBeHidden();
  await expect(page.locator("#crossYearHint")).toContainText("11/12");
  await expect(page.locator("#crossYearHint")).toContainText("1 个问题");
  await expect(page.locator("#toastRegion .toast").last()).toContainText(
    "缺少 2025年1月",
  );
});

test("工资变动同步主表、公式核对表和代发薪，停用同步三张在职表与离职名单", async ({
  page,
}) => {
  await openRegularWorkspace(page);
  await page.locator('[data-tab="changes"]').click();

  const payChanges = [
    "人员编号,姓名,其他工资,实发合计",
    "TEST-001,测试甲,250,5555",
  ].join("\n");
  await page.locator("#sourceFilesInput").setInputFiles({
    name: "合成人员工资变动.csv",
    mimeType: "text/csv",
    buffer: Buffer.from(payChanges),
  });
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await expect(page.locator("#proposalTable tbody tr")).toHaveCount(2);
  await page.locator("#applyProposalsBtn").click();
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await expect(page.locator("#proposalTable tbody tr")).toHaveCount(0);

  const paySync = await page.evaluate(async () => {
    const state = window.PayrollLocal.ui.state;
    const main = await state.workbook.getTable("工资表");
    const disbursement = await state.workbook.getTable("代发薪");
    const person = window.PayrollEngine
      .buildPeople(main)
      .people.find((item) => item.employeeId === "TEST-001");
    const payRow = disbursement.rows.find(
      (row) => row.get("人员编号") === "TEST-001",
    );
    return {
      otherPay: person.row.get("其他工资"),
      finalPay: person.row.get("实发合计"),
      finalPayFormula:
        person.row.cells.get(
          main.headers.find((header) => header.name === "实发合计").column,
        ).formula,
      disbursementAmount: payRow.get("金额(*)"),
      syncEntries: state.workbookSync.length,
    };
  });
  expect(paySync).toEqual({
    otherPay: 250,
    finalPay: 5555,
    finalPayFormula: "AU3-AX3",
    disbursementAmount: 5555,
    syncEntries: 2,
  });

  await page.locator("#changeText").fill("测试乙 停用");
  await page.locator("#parseTextBtn").click();
  await expect(page.locator("#proposalTable tbody tr")).toHaveCount(1);
  await page.locator("#applyProposalsBtn").click();
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });

  const disableSync = await page.evaluate(async () => {
    const state = window.PayrollLocal.ui.state;
    const counts = {};
    for (const sheetName of ["工资表", "工资核对表", "代发薪"]) {
      const table = await state.workbook.getTable(sheetName);
      counts[sheetName] = table.rows.filter((row) =>
        table.headers.some(
          (header) =>
            ["人员编号", "身份证", "姓名", "账户名称(*)"].includes(
              header.name,
            ) &&
            String(row.values.get(header.column) || "").includes("TEST-002"),
        ),
      ).length;
    }
    const archive = await state.workbook.getTable("离职名单");
    const archiveMatch = archive.rows.find(
      (row) => row.get("人员编号") === "TEST-002",
    );
    const hiddenRows = {};
    for (const sheetName of ["工资表", "工资核对表", "代发薪"]) {
      const record = await state.workbook.loadSheetRecord(sheetName);
      hiddenRows[sheetName] = window.PayrollLocal.excel
        .elementsByLocalName(record.document, "row")
        .filter((row) => row.getAttribute("hidden") === "1").length;
    }
    return {
      activeMatches: counts,
      archiveAdded: Boolean(archiveMatch),
      archiveRemark: archiveMatch?.get("备注") || "",
      hiddenRows,
      syncEntries: state.workbookSync.length,
    };
  });
  expect(disableSync.activeMatches).toEqual({
    工资表: 0,
    工资核对表: 0,
    代发薪: 0,
  });
  expect(disableSync.archiveAdded).toBe(true);
  expect(disableSync.archiveRemark).toContain("离职日期待输入");
  expect(disableSync.hiddenRows).toEqual({
    工资表: 1,
    工资核对表: 1,
    代发薪: 1,
  });
  expect(disableSync.syncEntries).toBe(3);
});

test("工资字段变化缺少实发合计时整批停止，不部分写入", async ({
  page,
}) => {
  await openRegularWorkspace(page);
  await page.locator('[data-tab="changes"]').click();
  const incomplete = [
    "人员编号,姓名,其他工资",
    "TEST-001,测试甲,999",
  ].join("\n");
  await page.locator("#sourceFilesInput").setInputFiles({
    name: "缺少实发合计.csv",
    mimeType: "text/csv",
    buffer: Buffer.from(incomplete),
  });
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await page.locator("#applyProposalsBtn").click();
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await expect(page.locator("#toastRegion .toast").last()).toContainText(
    "缺少“实发合计”",
  );
  await expect(page.locator("#proposalTable tbody tr")).toHaveCount(1);
  const currentValue = await page.evaluate(() => {
    const person = window.PayrollEngine
      .buildPeople(window.PayrollLocal.ui.state.table)
      .people.find((item) => item.employeeId === "TEST-001");
    return person.row.get("其他工资");
  });
  expect(currentValue).toBe(0);

  await page.locator("#clearProposalsBtn").click();
  await page
    .locator("#changeText")
    .fill(
      "新增测试丙 人员编号 TEST-003 姓名 测试丙 工资卡号 SYNTHETIC-ACCOUNT-003 实发合计 4000",
    );
  await page.locator("#parseTextBtn").click();
  await expect(page.locator("#proposalTable tbody tr")).toHaveCount(1);
  await page.locator("#applyProposalsBtn").click();
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await expect(page.locator("#toastRegion .toast").last()).toContainText(
    "工资核对表没有覆盖工资表第 6 行",
  );
  const newPersonExists = await page.evaluate(() =>
    window.PayrollEngine
      .buildPeople(window.PayrollLocal.ui.state.table)
      .people.some((person) => person.employeeId === "TEST-003"),
  );
  expect(newPersonExists).toBe(false);
});

test("长表、宽表和月份冲突规则保持独立", async ({ page }) => {
  await openRegularWorkspace(page);
  const result = await page.evaluate(() => {
    const table = window.PayrollLocal.ui.state.table;
    const wide = window.PayrollEngine.proposalsFromChangeTable(
      window.PayrollEngine.tableFromMatrix([
        ["人员编号", "姓名", "其他工资", "实发合计"],
        ["TEST-001", "测试甲", 1, 5500],
      ]),
      table,
      "2025-12",
      "合成宽表",
    );
    const long = window.PayrollEngine.proposalsFromChangeTable(
      window.PayrollEngine.tableFromMatrix([
        ["月份", "人员编号", "姓名", "操作", "字段", "值"],
        ["2025-12", "TEST-001", "测试甲", "设置", "其他工资", 2],
        ["2026-01", "TEST-001", "测试甲", "设置", "其他工资", 3],
      ]),
      table,
      "2025-12",
      "合成长表",
    );
    const coverage = window.PayrollEngine.resolveAttachment(
      "社保 / 公积金附件",
      window.PayrollEngine.tableFromMatrix(
        [
          [
            "姓名",
            "身份证号",
            "养老个人",
            "医疗个人",
            "失业个人",
            "公积金个人",
            "个人合计",
            "公司合计",
          ],
          ["覆盖甲", "SYNTHETIC-ID-A", 1, 1, 1, 1, 4, 8],
          ["新增丙", "SYNTHETIC-ID-C", 1, 1, 1, 1, 4, 8],
        ],
        "2025.12",
      ),
      window.PayrollEngine.tableFromMatrix([
        [
          "人员编号",
          "身份证",
          "姓名",
          "代扣养老保险",
          "代扣医疗保险",
          "代扣失业保险",
          "代扣房积金",
          "个人承担社保部分扣款合计",
          "企业承担社保部分扣款合计",
        ],
        ["SYNTHETIC-A", "SYNTHETIC-ID-A", "覆盖甲", 1, 1, 1, 1, 4, 8],
        ["SYNTHETIC-B", "SYNTHETIC-ID-B", "缺失乙", 1, 1, 1, 1, 4, 8],
        ["SYNTHETIC-C", "SYNTHETIC-ID-C", "新增丙", "", "", "", "", "", ""],
      ]),
      "2025-12",
      "2025.12合成保险表.xlsx",
    );
    let periodConflict = "";
    try {
      window.PayrollLocal.ui.resolveCurrentPeriod("2025.10工资表.xlsx", [
        {
          period: "2025-11",
          titleLike: true,
          cell: "B1",
          text: "2025年11月工资表",
        },
      ]);
    } catch (error) {
      periodConflict = error.message;
    }
    return {
      wideFormat: wide.format,
      wideReady: wide.proposals.every(
        (proposal) => proposal.status === "ready",
      ),
      longFormat: long.format,
      longStatuses: long.proposals.map((proposal) => proposal.status),
      attachmentCoverage: {
        matched: coverage.matchedPeople,
        sourceMatched: coverage.sourceMatchedPeople,
        expected: coverage.expectedPeople,
        errors: coverage.errors,
      },
      periodConflict,
    };
  });
  expect(result).toEqual({
    wideFormat: "wide",
    wideReady: true,
    longFormat: "long",
    longStatuses: ["ready", "error"],
    attachmentCoverage: {
      matched: 1,
      sourceMatched: 2,
      expected: 2,
      errors: ["工资表有 1 人未在该来源表中出现"],
    },
    periodConflict: "文件名月份与工资表标题月份不一致",
  });
});
