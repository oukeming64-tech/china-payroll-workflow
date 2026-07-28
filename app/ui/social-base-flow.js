(() => {
  "use strict";

  const ui = window.PayrollLocal.ui;
  const rules = window.PayrollEngine;

  function periodFromFile(file) {
    return rules.detectPeriod(file.name);
  }

  async function loadAnnualPayrollFiles(event) {
    const files = [...(event.target.files || [])];
    if (!files.length) {
      return;
    }
    const requiredYear = rules.previousCalendarYear(ui.state.targetPeriod);
    ui.setLoading(true, `正在读取 ${files.length} 个月度工资表…`);
    const monthlyTables = [];
    const errors = [];
    try {
      for (const file of files) {
        const period = periodFromFile(file);
        if (!period || Number(period.slice(0, 4)) !== requiredYear) {
          errors.push(`${file.name} 不是 ${requiredYear} 年月度文件`);
          continue;
        }
        const workbook = await ui.loadWorkbookFile(file);
        const sheetName = await ui.chooseMainSheet(workbook);
        monthlyTables.push({
          file,
          period,
          table: await workbook.getTable(sheetName),
        });
      }
      monthlyTables.sort((left, right) =>
        left.period.localeCompare(right.period),
      );
      ui.state.social.monthlyTables = monthlyTables;
      ui.byId("annualFilesHint").textContent =
        `${monthlyTables.length}/12 个有效月份${errors.length ? `，${errors.length} 个未采用` : ""}`;
      populateAnnualWageFields(monthlyTables);
      if (errors.length) {
        ui.toast(errors.join("；"), "error");
      } else {
        ui.toast(`已读取 ${monthlyTables.length} 个上年度工资文件`);
      }
    } catch (error) {
      ui.toast(error.message || "上年度工资文件读取失败", "error");
      console.error(error);
    } finally {
      ui.setLoading(false);
      event.target.value = "";
    }
  }

  function populateAnnualWageFields(monthlyTables) {
    const select = ui.byId("annualWageField");
    select.innerHTML = '<option value="">请选择字段</option>';
    if (!monthlyTables.length) {
      return;
    }
    const firstHeaders = monthlyTables[0].table.headers
      .map((header) => header.name)
      .filter((name) => /工资|薪|收入|应计/.test(name));
    for (const fieldName of firstHeaders) {
      if (
        !monthlyTables.every((item) =>
          rules.targetHeader(item.table, fieldName),
        )
      ) {
        continue;
      }
      const option = document.createElement("option");
      option.value = fieldName;
      option.textContent = fieldName;
      select.appendChild(option);
    }
  }

  async function loadSocialTarget(event) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    ui.setLoading(true, "正在识别社保模板…");
    try {
      const workbook = await ui.loadWorkbookFile(file);
      const periodVariants = [
        ui.state.targetPeriod,
        ui.state.targetPeriod.replace("-", "."),
        `${Number(ui.state.targetPeriod.slice(5))}月`,
      ];
      const exact = workbook.sheets.find((sheet) =>
        periodVariants.some((variant) => sheet.name.includes(variant)),
      );
      const sheetName = exact
        ? exact.name
        : (await workbook.scoreSheetsAgainst([
            "姓名",
            "身份证",
            "人员编号",
            "缴费工资",
            "社保基数",
          ]))[0]?.name;
      if (!sheetName) {
        throw new Error("社保模板中没有可识别的人员工作表");
      }
      const table = await workbook.getTable(sheetName);
      ui.state.social.targetFile = file;
      ui.state.social.targetWorkbook = workbook;
      ui.state.social.targetTable = table;
      ui.state.social.targetSheetName = sheetName;
      ui.byId("socialTargetLabel").textContent = file.name;
      populateSocialTargetFields(table);
      ui.toast(`已识别社保模板工作表：${sheetName}`);
    } catch (error) {
      ui.toast(error.message || "社保模板读取失败", "error");
      console.error(error);
    } finally {
      ui.setLoading(false);
      event.target.value = "";
    }
  }

  function populateSocialTargetFields(table) {
    const select = ui.byId("socialTargetField");
    select.innerHTML = '<option value="">请选择目标字段</option>';
    const likely = table.headers.filter((header) =>
      /基数|缴费工资|申报工资/.test(header.name),
    );
    const candidates = likely.length ? likely : table.headers;
    for (const header of candidates) {
      const option = document.createElement("option");
      option.value = header.name;
      option.textContent = header.displayName;
      select.appendChild(option);
    }
  }

  function adjustedValue(candidate) {
    let value = candidate.candidate;
    const lower = rules.asNumber(ui.byId("socialLowerLimit").value);
    const upper = rules.asNumber(ui.byId("socialUpperLimit").value);
    if (lower !== null) {
      value = Math.max(value, lower);
    }
    if (upper !== null) {
      value = Math.min(value, upper);
    }
    const rounding = ui.byId("socialRounding").value;
    if (rounding === "round") {
      value = Math.round(value);
    } else if (rounding === "floor") {
      value = Math.floor(value);
    } else if (rounding === "ceil") {
      value = Math.ceil(value);
    }
    return value;
  }

  function calculateSocialCandidates() {
    const fieldName = ui.byId("annualWageField").value;
    if (!fieldName) {
      ui.toast("请先人工确认“工资总额”的来源字段", "error");
      return;
    }
    if (!ui.state.social.monthlyTables.length) {
      ui.toast("请先选择上一年度月度工资文件", "error");
      return;
    }
    const divisorMode = ui.byId("socialDivisorMode").value;
    const candidates = rules.socialBaseCandidates(
      ui.state.social.monthlyTables,
      fieldName,
      divisorMode,
    );
    ui.state.social.candidates = candidates;
    ui.state.social.adjustedCandidates = candidates.map((candidate) => ({
      ...candidate,
      adjusted: adjustedValue(candidate),
    }));
    ui.state.social.updatesApplied = false;
    ui.renderSocialCandidates();
    updateSocialExportState();
    ui.toast(`已形成 ${candidates.length} 人的社保基数候选`);
  }

  function socialHasUnconfirmedExceptions() {
    const incomplete = ui.state.social.adjustedCandidates.some(
      (candidate) => candidate.coverage < 12,
    );
    const configured =
      ui.byId("socialLowerLimit").value ||
      ui.byId("socialUpperLimit").value ||
      ui.byId("socialRounding").value !== "none";
    return (
      (incomplete || configured) &&
      !ui.byId("socialExceptionConfirmed").checked
    );
  }

  function updateSocialExportState() {
    const ready = ui.state.social.adjustedCandidates.length > 0;
    ui.byId("exportSocialWorkbookBtn").disabled =
      !ready ||
      !ui.state.social.targetWorkbook ||
      !ui.byId("socialTargetField").value ||
      socialHasUnconfirmedExceptions();
  }

  function downloadSocialCsv() {
    const rows = [
      [
        "人员编号",
        "身份证",
        "姓名",
        "覆盖月份",
        "上年度工资总额",
        "分母",
        "候选社保基数",
        "调整后社保基数",
        "状态",
      ],
      ...ui.state.social.adjustedCandidates.map((item) => [
        item.employeeId,
        item.idCard,
        item.name,
        item.coverage,
        item.annualTotal,
        item.divisor,
        item.candidate,
        item.adjusted,
        item.note,
      ]),
    ];
    ui.downloadCsv(
      rows,
      `${ui.state.targetPeriod.replace("-", "")}_社保基数候选_待复核.csv`,
    );
  }

  async function exportSocialWorkbook() {
    if (socialHasUnconfirmedExceptions()) {
      ui.toast("请先确认不足 12 个月、上下限和取整方式", "error");
      return;
    }
    const social = ui.state.social;
    const targetField = rules.targetHeader(
      social.targetTable,
      ui.byId("socialTargetField").value,
    );
    if (!targetField) {
      ui.toast("请确认社保模板的目标基数字段", "error");
      return;
    }
    ui.setLoading(true, "正在写入社保基数工作副本…");
    try {
      const targetIndex = rules.indexPeople(social.targetTable);
      let updated = 0;
      const unmatched = [];
      for (const candidate of social.adjustedCandidates) {
        const match = rules.matchPerson(targetIndex, candidate);
        if (match.status !== "matched") {
          unmatched.push(candidate.maskedName);
          continue;
        }
        await social.targetWorkbook.updateCell(
          social.targetSheetName,
          match.person.rowNumber,
          targetField.column,
          candidate.adjusted,
          { preserveFormula: true },
        );
        updated += 1;
      }
      if (unmatched.length) {
        throw new Error(
          `${unmatched.length} 人未匹配到社保模板，已停止导出，请先核对人员`,
        );
      }
      const result = await social.targetWorkbook.export();
      ui.downloadBlob(
        result.blob,
        `${ui.state.targetPeriod.replace("-", ".")}社保基数_待复核.xlsx`,
      );
      social.updatesApplied = true;
      ui.toast(`已生成社保工作副本，共更新 ${updated} 人`);
    } catch (error) {
      ui.toast(error.message || "社保表导出失败", "error");
      console.error(error);
    } finally {
      ui.setLoading(false);
    }
  }

  function bindSocialBaseFlow() {
    ui.byId("annualPayrollFilesInput").addEventListener(
      "change",
      loadAnnualPayrollFiles,
    );
    ui.byId("socialTargetInput").addEventListener(
      "change",
      loadSocialTarget,
    );
    ui.byId("calculateSocialBtn").addEventListener(
      "click",
      calculateSocialCandidates,
    );
    ui.byId("downloadSocialCsvBtn").addEventListener(
      "click",
      downloadSocialCsv,
    );
    ui.byId("exportSocialWorkbookBtn").addEventListener(
      "click",
      exportSocialWorkbook,
    );
    for (const id of [
      "socialTargetField",
      "socialExceptionConfirmed",
      "socialDivisorMode",
      "socialLowerLimit",
      "socialUpperLimit",
      "socialRounding",
    ]) {
      ui.byId(id).addEventListener("change", updateSocialExportState);
    }
  }

  Object.assign(ui, {
    bindSocialBaseFlow,
    calculateSocialCandidates,
  });
})();
