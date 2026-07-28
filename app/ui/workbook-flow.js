(() => {
  "use strict";

  const ui = window.PayrollLocal.ui;
  const rules = window.PayrollEngine;

  async function chooseMainSheet(workbook) {
    if (workbook.sheetByName.has("工资表")) {
      return "工资表";
    }
    const scores = await workbook.scoreSheetsAgainst([
      "人员编号",
      "身份证",
      "姓名",
      "部门",
      "基本工资",
      "实发合计",
    ]);
    if (!scores.length || scores[0].score < 20) {
      throw new Error("没有找到可安全识别的主工资表工作表");
    }
    return scores[0].name;
  }

  function uniquePeriods(candidates) {
    return [...new Set(candidates.map((candidate) => candidate.period))];
  }

  function resolveCurrentPeriod(fileName, candidates) {
    const filenamePeriod = rules.detectPeriod(fileName);
    const titleCandidates = candidates.filter(
      (candidate) => candidate.titleLike,
    );
    const titlePeriods = uniquePeriods(titleCandidates);
    const allPeriods = uniquePeriods(candidates);
    let contentPeriod = "";
    if (titlePeriods.length === 1) {
      [contentPeriod] = titlePeriods;
    } else if (titlePeriods.length > 1) {
      contentPeriod = titlePeriods.includes(filenamePeriod)
        ? filenamePeriod
        : "";
    } else if (allPeriods.length === 1) {
      [contentPeriod] = allPeriods;
    } else if (allPeriods.length > 1) {
      contentPeriod = allPeriods.includes(filenamePeriod)
        ? filenamePeriod
        : "";
    }
    if (
      (titlePeriods.length > 1 ||
        (!titlePeriods.length && allPeriods.length > 1)) &&
      !contentPeriod
    ) {
      throw new Error("工资表内识别到多个月份，无法确定上月");
    }
    if (filenamePeriod && contentPeriod && filenamePeriod !== contentPeriod) {
      throw new Error("文件名月份与工资表标题月份不一致");
    }
    const period = contentPeriod || filenamePeriod;
    if (!period) {
      throw new Error("无法从工资表标题或文件名识别月份");
    }
    return {
      period,
      evidence: candidates.filter(
        (candidate) => candidate.period === period,
      ),
      filenamePeriod,
      contentPeriod,
    };
  }

  function setSelectedFile(input) {
    const file = input.files?.[0] || null;
    ui.state.baseFile = file;
    ui.byId("baseFileLabel").textContent =
      file?.name || "选择上月完整工资表";
    return file;
  }

  async function readBaseWorkbook(file) {
    ui.state.workingSourceFile = file;
    ui.state.workbook = await ui.loadWorkbookFile(file);
    ui.state.mainSheetName = await chooseMainSheet(ui.state.workbook);
    const candidates = await ui.state.workbook.findPeriodCandidates(
      ui.state.mainSheetName,
    );
    const resolved = resolveCurrentPeriod(file.name, candidates);
    ui.state.basePeriod = resolved.period;
    ui.state.workingPeriod = resolved.period;
    ui.state.targetPeriod = rules.nextPeriod(resolved.period);
    ui.state.periodEvidence = resolved.evidence;
    ui.state.table = await ui.state.workbook.getTable(
      ui.state.mainSheetName,
    );
    ui.state.externalLinks =
      await ui.state.workbook.analyzeExternalLinks();
    ui.state.formulaDiagnostics =
      await ui.state.workbook.analyzeFormulas(
        ui.state.mainSheetName,
      );
    ui.state.baseHash = await ui.sha256File(file);
    ui.state.route = rules.generationRoute(
      ui.state.basePeriod,
      ui.state.targetPeriod,
    );
  }

  async function readAnnualHistory(files) {
    const items = [];
    const errors = [];
    for (const file of files) {
      try {
        const workbook = await ui.loadWorkbookFile(file);
        const sheetName = await chooseMainSheet(workbook);
        const candidates = await workbook.findPeriodCandidates(sheetName);
        const resolved = resolveCurrentPeriod(file.name, candidates);
        items.push({
          file,
          workbook,
          sheetName,
          period: resolved.period,
          table: await workbook.getTable(sheetName),
          links: await workbook.analyzeExternalLinks(),
          formulas: await workbook.analyzeFormulas(sheetName),
          sha256: await ui.sha256File(file),
        });
      } catch (error) {
        errors.push(`${file.name}：${error.message}`);
      }
    }
    const audit = rules.validateAnnualHistory(
      items,
      ui.state.route.historyYear,
    );
    audit.errors.unshift(...errors);
    return audit;
  }

  async function applyRegularRoute() {
    ui.state.cumulativeResult =
      await ui.state.workbook.materializePreviousCumulative(
        ui.state.mainSheetName,
        ui.state.table,
      );
  }

  async function applyJanuaryRoute() {
    const rollover = rules.buildJanuaryRolloverPlan(
      ui.state.table,
      ui.state.annualHistory,
      ui.state.targetPeriod,
    );
    ui.state.rolloverPlan = rollover;
    if (rollover.errors.length) {
      throw new Error(rollover.errors.join("；"));
    }
    ui.state.table = await ui.state.workbook.addSchemaFields(
      ui.state.mainSheetName,
      rollover.insertColumn,
      rollover.addedFields,
    );
    await ui.state.workbook.clearFieldsForPeople(
      ui.state.mainSheetName,
      ui.state.table,
      rollover.resetFields,
    );
    await ui.state.workbook.applyJanuaryInternalFormulas(
      ui.state.mainSheetName,
      ui.state.table,
      rollover.nameOnlyRows,
    );
  }

  async function applyMonthlyBusinessRules() {
    const plan = rules.monthlyBusinessPlan(
      ui.state.targetPeriod,
      ui.state.table,
    );
    if (plan.errors.length) {
      throw new Error(plan.errors.join("；"));
    }
    if (plan.resetFields.length) {
      await ui.state.workbook.clearFieldsForPeople(
        ui.state.mainSheetName,
        ui.state.table,
        plan.resetFields,
      );
    }
    ui.state.monthlyBusiness = plan;
  }

  async function finalizeWorkspace() {
    if (ui.state.route.id === "january-rollover") {
      await applyJanuaryRoute();
    } else {
      await applyRegularRoute();
    }
    await applyMonthlyBusinessRules();
    const markers = await ui.state.workbook.findMonthMarkers(
      ui.state.basePeriod,
    );
    ui.state.monthMarkers = await ui.state.workbook.updateMonthMarkers(
      markers,
      ui.state.basePeriod,
      ui.state.targetPeriod,
    );
    if (!ui.state.monthMarkers.length) {
      throw new Error("没有找到可更新的月份标题");
    }
    ui.state.table = await ui.state.workbook.getTable(
      ui.state.mainSheetName,
    );
    ui.state.externalLinks =
      await ui.state.workbook.analyzeExternalLinks();
    ui.state.formulaDiagnostics =
      await ui.state.workbook.analyzeFormulas(
        ui.state.mainSheetName,
      );
    ui.state.personnelSheets =
      await ui.state.workbook.analyzePersonnelSheets();
    ui.resetAttachmentState();
    ui.state.history = [
      {
        time: new Date().toLocaleTimeString("zh-CN"),
        label: `${rules.formatPeriod(ui.state.basePeriod)} → ${rules.formatPeriod(ui.state.targetPeriod)}`,
        detail: `${ui.state.route.label}；更新 ${ui.state.monthMarkers.length} 个月份标题，重置 ${ui.state.monthlyBusiness.resetFields.length} 个目标月份专属字段，等待本月资料与工资附件核对。`,
        kind: "month-route",
      },
    ];
    showWorkspace();
  }

  function showHistoryRequest() {
    ui.byId("crossYearSetup").hidden = false;
    ui.byId("crossYearTitle").textContent =
      `请选择 ${ui.state.route.historyYear} 年 1—12 月工资表`;
    ui.byId("crossYearHint").textContent =
      "需要完整 12 个月，系统将核对一月重置、字段版本和工资附件来源。";
    ui.byId("crossYearFilesLabel").textContent = "选择全年工资表";
  }

  function showWorkspace() {
    ui.byId("workingFileName").textContent =
      ui.state.workingSourceFile.name;
    ui.byId("workingFileMeta").textContent =
      `${ui.formatBytes(ui.state.workingSourceFile.size)} · ${ui.state.mainSheetName}`;
    ui.byId("basePeriodText").textContent =
      rules.formatPeriod(ui.state.basePeriod);
    ui.byId("targetPeriodText").textContent =
      rules.formatPeriod(ui.state.targetPeriod);
    ui.byId("outputFileName").value =
      `${ui.state.targetPeriod.replace("-", ".")}工资表_待复核.xlsx`;
    ui.byId("exportTitle").textContent =
      `生成 ${rules.formatPeriod(ui.state.targetPeriod)}工资表`;
    const july = rules.isJuly(ui.state.targetPeriod);
    ui.byId("julyBanner").hidden = !july;
    ui.byId("socialTab").hidden = !july;
    ui.byId("socialInactive").hidden = july;
    ui.byId("socialWorkflow").hidden = !july;
    if (july) {
      ui.byId("socialYearHint").textContent =
        `本次使用 ${rules.previousCalendarYear(ui.state.targetPeriod)} 年 1—12 月工资数据，逐人计算月均。`;
    }
    ui.byId("setupView").hidden = true;
    ui.byId("workspaceView").hidden = false;
    ui.renderAll();
    ui.openTab("overview");
    ui.toast(`已建立 ${rules.formatPeriod(ui.state.targetPeriod)}草案`);
  }

  async function onBaseSelected(event) {
    const file = setSelectedFile(event.target);
    if (!file) {
      return;
    }
    ui.setLoading(true, "正在读取工资表…");
    try {
      await readBaseWorkbook(file);
      ui.byId("crossYearSetup").hidden = true;
      if (ui.state.route.requiresFullYearHistory) {
        showHistoryRequest();
        ui.toast("跨年生成需要上一年度 12 个月工资表");
      } else {
        await finalizeWorkspace();
      }
    } catch (error) {
      ui.toast(error.message || "工资表读取失败", "error");
      console.error(error);
      ui.state.baseFile = null;
      ui.state.workingSourceFile = null;
      ui.byId("baseFileLabel").textContent = "选择上月完整工资表";
      event.target.value = "";
    } finally {
      ui.setLoading(false);
    }
  }

  async function onCrossYearFilesSelected(event) {
    const files = [...(event.target.files || [])];
    if (!files.length) {
      return;
    }
    ui.setLoading(true, `正在核对 ${files.length} 个月度工资表…`);
    try {
      const audit = await readAnnualHistory(files);
      const baseItem = audit.items.find(
        (item) => item.period === ui.state.basePeriod,
      );
      if (!baseItem || baseItem.sha256 !== ui.state.baseHash) {
        audit.errors.push("全年文件中的十二月工资表与已选基线不是同一文件");
      }
      ui.state.annualHistory = audit;
      ui.byId("crossYearHint").textContent =
        `${audit.periods.length}/12 个有效月份${audit.errors.length ? `，${audit.errors.length} 个问题` : ""}`;
      if (audit.errors.length) {
        throw new Error(audit.errors.join("；"));
      }
      await finalizeWorkspace();
    } catch (error) {
      ui.toast(error.message || "全年工资表核对失败", "error");
      console.error(error);
    } finally {
      ui.setLoading(false);
      event.target.value = "";
    }
  }

  function restart() {
    window.location.reload();
  }

  function togglePrivacy() {
    ui.state.showPrivateData = !ui.state.showPrivateData;
    ui.byId("privacyToggle").textContent = ui.state.showPrivateData
      ? "恢复脱敏显示"
      : "显示完整数据";
    ui.renderAll();
  }

  function bindWorkbookFlow() {
    ui.byId("baseWorkbookInput").addEventListener("change", onBaseSelected);
    ui.byId("crossYearFilesInput").addEventListener(
      "change",
      onCrossYearFilesSelected,
    );
    ui.byId("restartBtn").addEventListener("click", restart);
    ui.byId("privacyToggle").addEventListener("click", togglePrivacy);
    ui.byId("peopleSearch").addEventListener("input", ui.renderPeopleTable);
  }

  Object.assign(ui, {
    chooseMainSheet,
    resolveCurrentPeriod,
    readAnnualHistory,
    finalizeWorkspace,
    bindWorkbookFlow,
  });
})();
