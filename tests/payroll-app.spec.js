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
const januaryChangeAttachments = [
  "员工动态表",
  "入职转正薪资",
  "实习生津贴表",
  "考勤表",
].map((label) =>
  path.join(
    fixtureRoot,
    "2026.01工资附件",
    `2026年1月${label}-示例.xlsx`,
  ),
);
const januaryFullAttachments = [
  ...januaryAttachments,
  ...januaryChangeAttachments,
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
  const sourceBuffer = options.buffer || fs.readFileSync(sourcePath);
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
      base64: sourceBuffer.toString("base64"),
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
  const confirmations = page.locator(
    "#monthlyBusinessList [data-monthly-business-id]:not(:disabled)",
  );
  for (let index = 0; index < (await confirmations.count()); index += 1) {
    await confirmations.nth(index).check();
  }
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

function xmlText(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

async function docxTablePayload(name, rows) {
  const zip = new JSZip();
  const tableRows = rows
    .map(
      (row) =>
        `<w:tr>${row
          .map(
            (value) =>
              `<w:tc><w:p><w:r><w:t>${xmlText(value)}</w:t></w:r></w:p></w:tc>`,
          )
          .join("")}</w:tr>`,
    )
    .join("");
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:tbl>${tableRows}</w:tbl></w:body>
      </w:document>`,
  );
  return {
    name,
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: await zip.generateAsync({ type: "nodebuffer" }),
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
  const monthlyRules = await page.evaluate(() =>
    ["2026-01", "2026-02", "2026-03", "2026-04"].map((period) => {
      const plan = window.PayrollEngine.monthlyBusinessPlan(period);
      return {
        period,
        performancePayout: plan.performancePayout,
        confidentialReview: plan.confidentialReview,
        pending: window.PayrollEngine
          .pendingMonthlyBusiness(plan)
          .map((item) => item.id),
      };
    }),
  );
  expect(monthlyRules).toEqual([
    {
      period: "2026-01",
      performancePayout: true,
      confidentialReview: false,
      pending: [
        "hr-changes",
        "attendance",
        "sales-commission",
        "quarterly-performance",
      ],
    },
    {
      period: "2026-02",
      performancePayout: false,
      confidentialReview: false,
      pending: ["hr-changes", "attendance", "sales-commission"],
    },
    {
      period: "2026-03",
      performancePayout: false,
      confidentialReview: true,
      pending: [
        "hr-changes",
        "attendance",
        "sales-commission",
        "confidential-allowance",
      ],
    },
    {
      period: "2026-04",
      performancePayout: true,
      confidentialReview: false,
      pending: [
        "hr-changes",
        "attendance",
        "sales-commission",
        "quarterly-performance",
      ],
    },
  ]);
  expect(pageErrors).toEqual([]);
});

test("考勤需求公式覆盖病假三档、事假、产假和工伤停止项", async ({
  page,
}) => {
  await page.goto("/");
  const audit = await page.evaluate(() => {
    const rules = window.PayrollEngine;
    const target = rules.tableFromMatrix([
      [
        "身份证",
        "姓名",
        "基本工资",
        "岗位工资",
        "绩效工资标准",
        "工资合计",
        "考勤扣款",
      ],
      ["ID-1", "甲", 3000, 5000, 2000, 10000, 0],
      ["ID-2", "乙", 3000, 5000, 2000, 10000, 0],
      ["ID-3", "丙", 3000, 5000, 2000, 10000, 0],
      ["ID-4", "丁", 3000, 5000, 2000, 10000, 0],
      ["ID-5", "戊", 3000, 5000, 2000, 10000, 0],
    ]);
    const source = rules.tableFromMatrix(
      [
        [
          "身份证号",
          "姓名",
          "请假信息",
          "",
          "",
          "",
          "",
        ],
        ["", "", "", "扣款病假", "事假", "产假", "工伤"],
        ["ID-1", "甲", "", 4, 1, 0, 0],
        ["ID-2", "乙", "", 6, 0, 0, 0],
        ["ID-3", "丙", "", 11, 0, 0, 0],
        ["ID-4", "丁", "", 0, 0, 1, 0],
        ["ID-5", "戊", "", 0, 0, 0, 2],
      ],
      "2026.06考勤",
    );
    const profile = rules.matchBusinessSource(
      source,
      "2026年6月考勤表-示例.xlsx",
    );
    const result = rules.proposalsFromBusinessSource(
      source,
      target,
      "2026-06",
      "2026年6月考勤表-示例.xlsx",
      profile,
    );
    return result.proposals.map((proposal) => ({
      status: proposal.status,
      formula: proposal.formula || "",
      errors: proposal.errors,
      warnings: proposal.warnings,
    }));
  });
  expect(audit).toHaveLength(5);
  expect(audit[0].formula).toContain("*1%*4");
  expect(audit[0].formula).toContain("/22*1");
  expect(audit[0].warnings).toEqual(
    expect.arrayContaining([
      expect.stringContaining("不足缴纳社保和公积金"),
    ]),
  );
  expect(audit[1].formula).toContain("*1.5%*6");
  expect(audit[2].formula).toContain("/22*50%*11");
  expect(audit[3].errors).toEqual(
    expect.arrayContaining([expect.stringContaining("产假待遇")]),
  );
  expect(audit[4].errors).toEqual(
    expect.arrayContaining([expect.stringContaining("停工留薪期")]),
  );
});

test("历史业务附件按已证明字段映射，只有名单没有金额时停止", async ({
  page,
}) => {
  const pageErrors = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  await openRegularWorkspace(page);

  const profileEvidence = await page.evaluate(() => {
    const target = window.PayrollLocal.ui.state.table;
    const scenarios = [
      {
        name: "2025.12实习生津贴.xlsx",
        matrix: [
          ["姓名", "身份证号", "津贴总额"],
          ["测试甲", "SYNTHETIC-ID-001", 5100],
        ],
      },
      {
        name: "保密人员统计-25.12.22.xlsx",
        matrix: [
          ["姓名", "津贴"],
          ["测试甲", 300],
        ],
      },
      {
        name: "xgs369人员名单.docx",
        matrix: [
          ["序号", "姓名"],
          [1, "测试甲"],
        ],
      },
    ];
    return scenarios.map((scenario) => {
      const table = window.XlsxEngine.buildTableFromMatrix(
        scenario.matrix,
        "业务表",
      );
      const profile = window.PayrollEngine.matchBusinessSource(
        table,
        scenario.name,
      );
      const result = window.PayrollEngine.proposalsFromBusinessSource(
        table,
        target,
        "2025-12",
        scenario.name,
        profile,
      );
      return {
        profile: profile?.id || "",
        fields: result.proposals.map(
          (proposal) => proposal.field?.name || "",
        ),
        errors: result.errors,
      };
    });
  });
  expect(profileEvidence).toEqual([
    {
      profile: "intern-allowance",
      fields: ["基本工资"],
      errors: [],
    },
    {
      profile: "confidential-allowance-roster",
      fields: ["BM津贴"],
      errors: [],
    },
    {
      profile: "confidential-eligibility",
      fields: [],
      errors: [expect.stringContaining("没有提供可写入工资表的补贴金额")],
    },
  ]);

  const approval = await docxTablePayload(
    "保密补贴发放审批表2025.12.docx",
    [
      ["序号", "姓名", "岗位密级", "应发数额", "实发数额", "备注"],
      ["1", "测试甲", "一般", "300", "300", ""],
    ],
  );
  await page.locator('[data-tab="changes"]').click();
  await page.locator("#sourceFilesInput").setInputFiles(approval);
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await expect(page.locator("#proposalTable tbody tr")).toHaveCount(1);
  const docxEvidence = await page.evaluate(() => {
    const state = window.PayrollLocal.ui.state;
    const proposal = state.proposals.find(
      (item) => item.sourceKind === "confidential-allowance-approval",
    );
    const source = state.sources.at(-1);
    return {
      category: source.category,
      format: source.format,
      mappingCount: source.mappingCount,
      field: proposal?.field?.name || "",
      value: proposal?.inputValue,
      status: proposal?.status || "",
    };
  });
  expect(docxEvidence).toEqual({
    category: "保密补贴审批表",
    format: "docx-business-source",
    mappingCount: 1,
    field: "BM津贴",
    value: "300",
    status: "ready",
  });
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

test("三类附件仅一份加密且工作表名仍为上月时完整读取", async ({
  page,
}) => {
  const pageErrors = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  await openJanuaryWorkspace(page);

  const staleSheetZip = await JSZip.loadAsync(
    fs.readFileSync(januaryAttachments[1]),
  );
  const workbookXmlPath = "xl/workbook.xml";
  const workbookXml = await staleSheetZip
    .file(workbookXmlPath)
    .async("string");
  if (!workbookXml.includes('name="2026.1"')) {
    throw new Error("合成社保附件缺少预期工作表");
  }
  staleSheetZip.file(
    workbookXmlPath,
    workbookXml.replace('name="2026.1"', 'name="2025.12"'),
  );
  const staleSheetBuffer = await staleSheetZip.generateAsync({
    type: "nodebuffer",
  });
  const encryptedInsurance = await encryptedWorkbookPayload(
    page,
    januaryAttachments[1],
    "synthetic-three-file-password",
    { buffer: staleSheetBuffer },
  );
  const filePayload = (sourcePath) => ({
    name: path.basename(sourcePath),
    mimeType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    buffer: fs.readFileSync(sourcePath),
  });

  await page.locator('[data-tab="attachments"]').click();
  await page
    .locator("#attachmentFilesInput")
    .setInputFiles([
      filePayload(januaryAttachments[0]),
      encryptedInsurance,
      filePayload(januaryAttachments[2]),
    ]);
  const dialog = page.locator("#workbookPasswordDialog");
  await expect(dialog).toBeVisible();
  await expect(page.locator("#workbookPasswordFile")).toHaveText(
    encryptedInsurance.name,
  );
  await page
    .locator("#workbookPasswordInput")
    .fill("synthetic-three-file-password");
  await page.getByRole("button", { name: "打开工作簿" }).click();
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });

  const evidence = await page.evaluate(() => {
    const state = window.PayrollLocal.ui.state;
    return {
      errors: state.attachments.errors,
      categories: state.attachments.results.map(
        (result) => result.category,
      ),
      updateCounts: state.attachments.results.map(
        (result) => result.updates.length,
      ),
      insuranceWarnings:
        state.attachments.results.find(
          (result) => result.category === "社保 / 公积金附件",
        )?.warnings || [],
    };
  });
  expect(evidence).toEqual({
    errors: [],
    categories: [
      "个税工资薪金附件",
      "社保 / 公积金附件",
      "劳务费附件",
    ],
    updateCounts: [14, 12, 1],
    insuranceWarnings: [
      expect.stringContaining(
        "工作表名仍为2025年12月，文件名指向2026年1月",
      ),
    ],
  });
  await expect(page.locator("#attachmentSummary")).toHaveText(
    "3/3 类附件已读取，3/3 类通过核对。",
  );
  await expect(page.locator("#applyAttachmentsBtn")).toBeEnabled();
  await expect(page.locator("#attachmentCards")).toContainText(
    "已按唯一字段结构读取",
  );
  expect(pageErrors).toEqual([]);
});

test("七份附件全量读取并以身份证严格匹配人员变动", async ({
  page,
}) => {
  const pageErrors = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  await openJanuaryWorkspace(page);

  await page.locator('[data-tab="diagnostics"]').click();
  await expect(
    page.locator("#requirementCoverageList .diagnostic-card"),
  ).toHaveCount(11);
  await expect(page.locator("#requirementCoverageList")).toContainText(
    "月内转正必须有转正天数",
  );
  await expect(page.locator("#requirementCoverageList")).toContainText(
    "离职前近12个月工资总额月均",
  );

  await page.locator('[data-tab="attachments"]').click();
  await page.locator("#attachmentFilesInput").setInputFiles([
    ...januaryFullAttachments.map((sourcePath) => ({
      name: path.basename(sourcePath),
      mimeType: sourcePath.endsWith(".xls")
        ? "application/vnd.ms-excel"
        : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      buffer: fs.readFileSync(sourcePath),
    })),
    {
      name: "~$2026年1月考勤表-示例.xlsx",
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      buffer: Buffer.from("synthetic temporary lock"),
    },
  ]);
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await expect(page.locator("#attachmentSummary")).toContainText(
    "7/7 个文件已读取",
  );

  const intake = await page.evaluate(() => {
    const state = window.PayrollLocal.ui.state;
    const labor = state.attachments.results.find(
      (result) => result.category === "劳务费附件",
    );
    return {
      files: state.attachments.files.length,
      ignoredFiles: state.attachments.ignoredFiles,
      coreCategories: state.attachments.results.map(
        (result) => result.category,
      ),
      changeProfiles: state.sources
        .filter((source) => source.automaticAttachment)
        .map((source) => source.profileId),
      automaticProposals: state.proposals
        .filter((proposal) => proposal.automaticAttachment)
        .map((proposal) => ({
          field: proposal.field?.name || "",
          matchedBy: proposal.matchedBy || "",
          status: proposal.status,
        })),
      laborWarnings: labor?.warnings || [],
      errors: state.attachments.errors,
    };
  });
  expect(intake.files).toBe(7);
  expect(intake.ignoredFiles).toEqual([
    "~$2026年1月考勤表-示例.xlsx",
  ]);
  expect(intake.coreCategories).toEqual([
    "个税工资薪金附件",
    "社保 / 公积金附件",
    "劳务费附件",
  ]);
  expect(intake.changeProfiles).toEqual([
    "employee-regularization-review,employee-dynamics,employee-salary-adjustment-review",
    "probation-salary,salary-transfer-review,salary-adjustment",
    "intern-allowance",
    "attendance-register",
  ]);
  expect(intake.automaticProposals).toEqual([
    {
      field: "岗位",
      matchedBy: "idCard",
      status: "ready",
    },
  ]);
  expect(intake.laborWarnings).toEqual([
    expect.stringContaining("没有身份证号"),
  ]);
  expect(intake.errors).toEqual([
    expect.stringContaining("1 项人员 / 工资变动待确认"),
  ]);
  await expect(page.locator("#applyAttachmentsBtn")).toBeDisabled();

  const strictIdentity = await page.evaluate(() => {
    const state = window.PayrollLocal.ui.state;
    const table = window.XlsxEngine.buildTableFromMatrix(
      [
        [
          "身份证号",
          "姓名",
          "转正日期",
          "部门",
          "岗位",
          "基本工资",
          "岗位工资",
          "绩效工资",
        ],
        [
          "SYNTHETIC-ID-WRONG",
          "测试甲",
          "2026-01-01",
          "",
          "",
          1800,
          3000,
          1200,
        ],
      ],
      "严格身份证测试",
    );
    const profile = window.PayrollEngine.matchBusinessSource(
      table,
      "2026年1月入职转正薪资-示例.xlsx",
    );
    const result = window.PayrollEngine.proposalsFromBusinessSource(
      table,
      state.table,
      state.targetPeriod,
      "2026年1月入职转正薪资-示例.xlsx",
      profile,
    );
    return result.proposals.map((proposal) => ({
      person: Boolean(proposal.person),
      status: proposal.status,
      errors: proposal.errors,
    }));
  });
  expect(strictIdentity).toEqual([
    {
      person: false,
      status: "error",
      errors: [
        expect.stringContaining(
          "身份证号未匹配到人员，已禁止改用其他字段兜底",
        ),
      ],
    },
  ]);

  const regularizationProration = await page.evaluate(() => {
    const state = window.PayrollLocal.ui.state;
    const build = (regularDays) => {
      const headers = [
        "身份证号",
        "姓名",
        "转正日期",
        "基本工资",
        "岗位工资",
        "绩效工资",
        "合计",
      ];
      if (regularDays !== null) {
        headers.push("转正天数");
      }
      const row = [
        "SYNTHETIC-ID-001",
        "测试甲",
        "2026-01-02",
        2250,
        3750,
        1500,
        7500,
      ];
      if (regularDays !== null) {
        row.push(regularDays);
      }
      const table = window.XlsxEngine.buildTableFromMatrix(
        [headers, row],
        "转正按天测试",
      );
      const profile = window.PayrollEngine.matchBusinessSource(
        table,
        "2026年1月入职转正薪资-示例.xlsx",
      );
      return window.PayrollEngine.proposalsFromBusinessSource(
        table,
        state.table,
        state.targetPeriod,
        "2026年1月入职转正薪资-示例.xlsx",
        profile,
      );
    };
    const missing = build(null);
    const complete = build(21);
    const performance = complete.proposals.find(
      (proposal) => proposal.field?.name === "绩效工资标准",
    );
    return {
      missingError: missing.proposals[0]?.errors[0] || "",
      completeErrors: complete.errors,
      performanceValue: performance?.inputValue,
      performanceFormula: performance?.formula,
    };
  });
  expect(regularizationProration).toEqual({
    missingError: expect.stringContaining("缺少“转正天数”"),
    completeErrors: [],
    performanceValue: 1431.82,
    performanceFormula: "ROUND(1500/22*21,2)",
  });

  const employmentEvents = await page.evaluate(() => {
    const state = window.PayrollLocal.ui.state;
    const source = window.XlsxEngine.buildTableFromMatrix(
      [
        [
          "身份证号",
          "姓名",
          "入职日期",
          "离职日期",
          "试用工资",
          "工作天数",
          "未发放绩效工资",
        ],
        [
          "SYNTHETIC-ID-NEW",
          "新增员工",
          "2026-01-10",
          "",
          4400,
          10,
          "",
        ],
        [
          "SYNTHETIC-ID-001",
          "测试甲",
          "",
          "2026-01-20",
          "",
          15,
          300,
        ],
      ],
      "入离职测试",
    );
    const profile = window.PayrollEngine.matchBusinessSource(
      source,
      "2026年1月员工动态表-示例.xlsx",
    );
    const result = window.PayrollEngine.proposalsFromBusinessSource(
      source,
      state.table,
      state.targetPeriod,
      "2026年1月员工动态表-示例.xlsx",
      profile,
    );
    return result.proposals.map((proposal) => ({
      kind: proposal.kind,
      errors: proposal.errors,
      warnings: proposal.warnings,
    }));
  });
  expect(employmentEvents).toEqual([
    {
      kind: "new-person",
      errors: [
        expect.stringContaining(
          "试用工资、工作天数、完整人员信息、银行账号和实发合计",
        ),
      ],
      warnings: [expect.stringContaining("入职工资预览")],
    },
    {
      kind: "disable-person",
      errors: [
        expect.stringContaining(
          "工作天数、未发绩效、实发合计及适用的近12个月补偿口径",
        ),
      ],
      warnings: [expect.stringContaining("离职工资预览")],
    },
  ]);

  const laborRequirementGuard = await page.evaluate(() => {
    const state = window.PayrollLocal.ui.state;
    const lowTierSource = window.XlsxEngine.buildTableFromMatrix(
      [
        ["姓名", "劳务费", "增值税税额"],
        ["测试劳务", 3000, 0],
      ],
      "劳务费正常变动测试",
    );
    const highTierSource = window.XlsxEngine.buildTableFromMatrix(
      [
        ["姓名", "劳务费", "增值税税额"],
        ["测试劳务", 30000, 0],
      ],
      "劳务费高档测试",
    );
    const resolve = (source) =>
      window.PayrollEngine.resolveAttachment(
        "劳务费附件",
        source,
        state.table,
        state.targetPeriod,
        "2026年劳务费-示例.xlsx",
      );
    const lowTier = resolve(lowTierSource);
    const highTier = resolve(highTierSource);
    return {
      lowTier: {
        errors: lowTier.errors,
        warnings: lowTier.warnings,
      },
      highTier: {
        errors: highTier.errors,
        warnings: highTier.warnings,
      },
    };
  });
  expect(laborRequirementGuard.lowTier.errors).toEqual([]);
  expect(laborRequirementGuard.lowTier.warnings).toEqual(
    expect.arrayContaining([
      expect.stringContaining("劳务费金额复核（不阻断写入）"),
      expect.stringContaining("临时劳务费或本月新增"),
    ]),
  );
  expect(laborRequirementGuard.highTier.errors).toEqual(
    expect.arrayContaining([
      expect.stringContaining("应纳税所得额超过20000元"),
    ]),
  );
  expect(laborRequirementGuard.highTier.warnings).toEqual(
    expect.arrayContaining([
      expect.stringContaining("劳务费金额复核（不阻断写入）"),
    ]),
  );

  await page.locator('[data-tab="changes"]').click();
  await page.locator("#applyProposalsBtn").click();
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await page.locator('[data-tab="attachments"]').click();
  await expect(page.locator("#applyAttachmentsBtn")).toBeEnabled();
  await page.locator("#applyAttachmentsBtn").click();
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  const applied = await page.evaluate(() => {
    const state = window.PayrollLocal.ui.state;
    const person = window.PayrollEngine
      .buildPeople(state.table)
      .people.find((item) => item.idCard === "SYNTHETIC-ID-001");
    return {
      position: person.row.get("岗位"),
      files: state.attachments.files.length,
      externalLinks: state.externalLinks.length,
      externalFormulas:
        state.formulaDiagnostics.externalFormulaNodes,
    };
  });
  expect(applied).toEqual({
    position: "新测试岗位",
    files: 7,
    externalLinks: 0,
    externalFormulas: 0,
  });
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
      firstBmAllowance: first.row.get("BM津贴"),
      resetValues: Object.fromEntries(
        state.monthlyBusiness.resetFields.map((field) => [
          field,
          first.row.get(field),
        ]),
      ),
      pendingBusiness: window.PayrollEngine
        .pendingMonthlyBusiness(state.monthlyBusiness)
        .map((item) => item.id),
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
    firstOtherPay: null,
    firstBmAllowance: 200,
    resetValues: {
      "月度绩效（季度发放）": null,
      销售提成: null,
      其他工资: null,
      考勤扣款: null,
      其他扣款: null,
      年度绩效: null,
      代扣借款: null,
    },
    pendingBusiness: [
      "hr-changes",
      "attendance",
      "sales-commission",
      "confidential-allowance",
    ],
    firstFinalPay: 5600,
    unresolvedFormulas: 0,
  });

  const unchangedAttachmentEvidence = await page.evaluate(() => {
    const state = window.PayrollLocal.ui.state;
    const rules = window.PayrollEngine;
    const fields = [
      ["养老个人", "代扣养老保险"],
      ["医疗个人", "代扣医疗保险"],
      ["失业个人", "代扣失业保险"],
      ["公积金个人", "代扣房积金"],
      ["个人合计", "个人承担社保部分扣款合计"],
      ["公司合计", "企业承担社保部分扣款合计"],
    ];
    const people = rules
      .buildPeople(state.table)
      .people.filter((person) => person.idCard);
    const matrix = [
      ["姓名", "身份证号", ...fields.map(([source]) => source)],
      ...people.map((person) => [
        person.name,
        person.idCard,
        ...fields.map(([, target]) => person.row.get(target)),
      ]),
    ];
    const source = window.XlsxEngine.buildTableFromMatrix(
      matrix,
      "2025.12",
    );
    const result = rules.resolveAttachment(
      "社保 / 公积金附件",
      source,
      state.table,
      state.targetPeriod,
      "2025.12同值保险表.xlsx",
    );
    return {
      errors: result.errors,
      updates: result.updates.length,
      changed: result.fieldSummaries.reduce(
        (total, field) => total + field.changed,
        0,
      ),
    };
  });
  expect(unchangedAttachmentEvidence).toEqual({
    errors: [],
    updates: 12,
    changed: 0,
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
      attachmentWrites: state.attachments.results.map((result) => ({
        category: result.category,
        updates: result.updates.length,
        changed: result.fieldSummaries.reduce(
          (total, field) => total + field.changed,
          0,
        ),
      })),
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
    attachmentWrites: [
      {
        category: "个税工资薪金附件",
        updates: 14,
        changed: expect.any(Number),
      },
      {
        category: "社保 / 公积金附件",
        updates: 12,
        changed: expect.any(Number),
      },
    ],
    tax: 120,
    pension: 100,
    medical: 20,
    unemployment: 10,
    housingFund: 200,
    personalTotal: 330,
    companyTotal: 1870,
  });
  await page.locator('[data-tab="review"]').click();
  await page.locator("#exportAcknowledged").check();
  await expect(page.locator("#exportWorkbookBtn")).toBeDisabled();
  await expect(page.locator("#monthlyBusinessBadge")).toContainText(
    "4 项待",
  );
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
    "人员编号,姓名,其他工资,月度绩效（季度发放）,实发合计",
    "TEST-001,测试甲,250,300,5555",
  ].join("\n");
  await page.locator("#sourceFilesInput").setInputFiles({
    name: "合成人员工资变动.csv",
    mimeType: "text/csv",
    buffer: Buffer.from(payChanges),
  });
  await expect(page.locator("#loadingOverlay")).toBeHidden({
    timeout: 30_000,
  });
  await expect(page.locator("#proposalTable tbody tr")).toHaveCount(3);
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
      settledPerformance: person.row.get("月度绩效（季度发放）"),
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
    settledPerformance: 300,
    finalPay: 5555,
    finalPayFormula: "AU3-AX3",
    disbursementAmount: 5555,
    syncEntries: 3,
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
  expect(disableSync.syncEntries).toBe(4);
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
  expect(currentValue).toBeNull();

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
