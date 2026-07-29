(() => {
  "use strict";

  const ui = window.PayrollLocal.ui;
  const rules = window.PayrollEngine;

  function appendProposals(result) {
    ui.state.proposals.push(...(result.proposals || []));
    for (const error of result.errors || []) {
      ui.state.proposals.push({
        ...rules.proposalBase("来源诊断", "—", ""),
        status: "error",
        selected: false,
        kind: "source-error",
        errors: [error],
      });
    }
    rules.markConflicts(ui.state.proposals);
    ui.renderProposals();
  }

  function parseTextChanges() {
    if (!ui.state.table) {
      return;
    }
    const result = rules.parseNaturalLanguage(
      ui.byId("changeText").value,
      ui.state.table,
      ui.state.targetPeriod,
    );
    appendProposals(result);
    ui.openTab("changes");
    if (result.proposals.length) {
      ui.toast(`已形成 ${result.proposals.length} 项文字变动预览`);
    } else {
      ui.toast(result.errors[0] || "没有识别到变动", "error");
    }
  }

  function matchingDiagnostic(fileName, category) {
    const exact = ui.state.externalLinks.find(
      (link) =>
        rules.normalizeText(link.filename) === rules.normalizeText(fileName),
    );
    if (exact) {
      return exact;
    }
    const categoryKey = rules.normalizeText(category).replace("工资模板", "");
    return [...ui.state.externalLinks]
      .filter((link) => {
        const linkKey = rules.normalizeText(link.category);
        return linkKey.includes(categoryKey) || categoryKey.includes(linkKey);
      })
      .sort(
        (left, right) =>
          right.formulaReferenceCount - left.formulaReferenceCount,
      )[0];
  }

  function looksLikeChangeTable(table) {
    const headerKeys = new Set(
      table.headers.map((header) => rules.normalizeText(header.name)),
    );
    const operation = ["操作", "变动类型", "类型"].some((name) =>
      headerKeys.has(rules.normalizeText(name)),
    );
    const field = ["字段", "目标字段"].some((name) =>
      headerKeys.has(rules.normalizeText(name)),
    );
    const value = ["值", "新值", "数值"].some((name) =>
      headerKeys.has(rules.normalizeText(name)),
    );
    return operation && field && value;
  }

  async function importCsvSource(file, role) {
    const matrix = rules.parseCsv(await file.text());
    const table = rules.tableFromMatrix(matrix, file.name);
    const profile = looksLikeChangeTable(table)
      ? null
      : rules.matchBusinessSource(table, file.name);
    const result = profile
      ? rules.proposalsFromBusinessSource(
          table,
          ui.state.table,
          ui.state.targetPeriod,
          role || file.name,
          profile,
        )
      : rules.proposalsFromChangeTable(
          table,
          ui.state.table,
          ui.state.targetPeriod,
          role || file.name,
        );
    appendProposals(result);
    ui.state.sources.push({
      name: file.name,
      category: profile
        ? profile.label
        : result.format === "wide"
          ? "人员 / 工资变动表（宽表）"
          : "人员 / 工资变动表（长表）",
      profileId: profile?.id || "",
      sheetName: table.sheetName,
      format: result.format,
      mappingCount: result.mappings?.length || 0,
      proposalCount: result.proposals.length,
      errors: result.errors,
      warnings: result.warnings || [],
    });
  }

  async function importWorkbookSource(file, options = {}) {
    if (!ui.state.table) {
      return;
    }
    const role = options.role || file.name;
    const category = rules.classifySource(file.name);
    const workbook = await ui.loadWorkbookFile(file);
    const diagnostic = matchingDiagnostic(file.name, category);
    const inspection = await ui.inspectWorkbookBusinessRegions(
      workbook,
      diagnostic,
      ui.state.table,
      ui.state.targetPeriod,
      file.name,
      role,
    );
    const table = inspection.table;
    if (!table) {
      throw new Error(`${file.name} 没有可识别的工作表`);
    }
    let result;
    let resolvedCategory = category;
    if (inspection.result) {
      result = inspection.result;
      resolvedCategory = result.labels.join("、");
    } else if (looksLikeChangeTable(table)) {
      result = rules.proposalsFromChangeTable(
        table,
        ui.state.table,
        ui.state.targetPeriod,
        role,
      );
      resolvedCategory = "人员 / 工资变动表（长表）";
    } else if (diagnostic) {
      result = rules.proposalsFromExternalSource(
        table,
        ui.state.table,
        diagnostic,
        ui.state.targetPeriod,
        role,
      );
    } else {
      result = rules.proposalsFromChangeTable(
        table,
        ui.state.table,
        ui.state.targetPeriod,
        role,
      );
      resolvedCategory =
        result.format === "wide"
          ? "人员 / 工资变动表（宽表）"
          : "人员 / 工资变动表（长表）";
    }
    if (options.automaticAttachment) {
      for (const proposal of result.proposals || []) {
        proposal.automaticAttachment = true;
      }
    }
    appendProposals(result);
    ui.state.sources.push({
      name: file.name,
      category: resolvedCategory,
      profileId:
        result.profileIds?.join(",") || result.profile?.id || "",
      sheetName: result.sheetName || table.sheetName,
      mappingCount: result.mappings?.length || 0,
      format: result.format || "external",
      proposalCount: result.proposals.length,
      errors: result.errors,
      warnings: result.warnings || [],
      automatic: options.automatic,
      automaticAttachment: Boolean(options.automaticAttachment),
    });
    ui.renderSources();
    if (!options.automatic) {
      ui.toast(`${file.name} 已完成本地匹配预览`);
    }
  }

  async function importDocxSource(file, options = {}) {
    if (!ui.state.table) {
      return;
    }
    const role = options.role || file.name;
    const document = await window.XlsxEngine.parseDocxTables(file);
    const candidates = document.tables
      .map((item) => {
        const table = rules.tableFromMatrix(item.matrix, item.name);
        return {
          table,
          profile: rules.matchBusinessSource(table, file.name),
        };
      })
      .filter((item) => item.profile);
    if (!candidates.length) {
      throw new Error(
        `${file.name} 没有找到经历史证明的人员或工资金额表格`,
      );
    }
    if (candidates.length > 1) {
      throw new Error(`${file.name} 有多个可识别业务表格，请拆分后导入`);
    }
    const { table, profile } = candidates[0];
    const result = rules.proposalsFromBusinessSource(
      table,
      ui.state.table,
      ui.state.targetPeriod,
      role,
      profile,
    );
    appendProposals(result);
    ui.state.sources.push({
      name: file.name,
      category: profile.label,
      sheetName: table.sheetName,
      mappingCount: result.mappings.length,
      format: "docx-business-source",
      proposalCount: result.proposals.length,
      errors: result.errors,
      automatic: options.automatic,
    });
    ui.renderSources();
    if (!options.automatic) {
      ui.toast(`${file.name} 已完成本地匹配预览`);
    }
  }

  async function importSelectedSources(event) {
    const files = [...(event.target.files || [])];
    if (!files.length) {
      return;
    }
    ui.setLoading(true, `正在分析 ${files.length} 个来源文件…`);
    try {
      for (const file of files) {
        if (/\.csv$/i.test(file.name)) {
          await importCsvSource(file);
        } else if (/\.docx$/i.test(file.name)) {
          await importDocxSource(file);
        } else {
          await importWorkbookSource(file);
        }
      }
      ui.renderAll();
    } catch (error) {
      ui.toast(error.message || "来源文件识别失败", "error");
      console.error(error);
    } finally {
      ui.setLoading(false);
      event.target.value = "";
    }
  }

  async function restoreApplicationSnapshot(snapshot) {
    ui.state.workbook = await window.XlsxEngine.XlsxWorkbook.load(
      snapshot.bytes,
      ui.state.workingSourceFile?.name || "当前草案.xlsx",
    );
    ui.state.table = await ui.state.workbook.getTable(
      ui.state.mainSheetName,
    );
    ui.state.externalLinks =
      await ui.state.workbook.analyzeExternalLinks();
    ui.state.formulaDiagnostics =
      await ui.state.workbook.analyzeFormulas(
        ui.state.mainSheetName,
      );
    ui.state.history = snapshot.history;
    ui.state.disabledRows = snapshot.disabledRows;
    ui.state.workbookSync = snapshot.workbookSync;
    ui.state.personnelSheets =
      await ui.state.workbook.analyzePersonnelSheets();
    ui.renderAll();
  }

  async function applySelectedProposals() {
    const selected = ui.state.proposals.filter(
      (proposal) => proposal.selected && proposal.status !== "error",
    );
    if (!selected.length) {
      return;
    }
    ui.setLoading(true, `正在应用 ${selected.length} 项已确认变动…`);
    const appliedIds = new Set();
    let snapshot = null;
    try {
      await ui.preflightWorkbookProposals(selected);
      const exported = await ui.state.workbook.export();
      snapshot = {
        bytes: new Uint8Array(await exported.blob.arrayBuffer()),
        history: [...ui.state.history],
        disabledRows: new Set(ui.state.disabledRows),
        workbookSync: [...ui.state.workbookSync],
      };
      const reuseProviderIds = new Set(
        selected.flatMap((proposal) =>
          Object.values(proposal.auxiliaryReuse || {})
            .map((reuse) => reuse.providerId)
            .filter(Boolean),
        ),
      );
      const ordered = selected
        .map((proposal, index) => ({ proposal, index }))
        .sort((left, right) => {
          const leftProvider = reuseProviderIds.has(left.proposal.id)
            ? 0
            : 1;
          const rightProvider = reuseProviderIds.has(right.proposal.id)
            ? 0
            : 1;
          return leftProvider - rightProvider || left.index - right.index;
        })
        .map((item) => item.proposal);
      for (const proposal of ordered) {
        if (proposal.kind === "cell-change") {
          await ui.applyWorkbookCellProposal(proposal);
        } else if (proposal.kind === "new-person") {
          await ui.applyWorkbookNewPerson(proposal);
        } else if (proposal.kind === "disable-person") {
          await ui.applyWorkbookDisableProposal(proposal);
        }
        appliedIds.add(proposal.id);
      }
      ui.state.proposals = ui.state.proposals.filter(
        (proposal) => !appliedIds.has(proposal.id),
      );
      ui.state.table = await ui.state.workbook.getTable(
        ui.state.mainSheetName,
      );
      ui.state.formulaDiagnostics =
        await ui.state.workbook.analyzeFormulas(
          ui.state.mainSheetName,
        );
      ui.state.personnelSheets =
        await ui.state.workbook.analyzePersonnelSheets();
      const attachmentWasChecked =
        ui.state.attachments.applied ||
        ui.state.attachments.results.length > 0;
      if (
        attachmentWasChecked &&
        ui.state.attachments.inputs?.length
      ) {
        await ui.refreshAttachmentResolution(
          "草案发生变动，已用本批附件重新核对",
        );
      } else {
        ui.invalidateAttachmentResolution(
          attachmentWasChecked
            ? "草案发生变动，请重新核对目标月份工资附件"
            : "",
        );
      }
      ui.renderAll();
      ui.toast(`${appliedIds.size} 项变动已写入当前草案`);
    } catch (error) {
      if (snapshot) {
        await restoreApplicationSnapshot(snapshot);
      }
      ui.toast(
        `${error.message || "应用变动失败"}${snapshot ? "；本批变动已全部回滚" : ""}`,
        "error",
      );
      console.error(error);
    } finally {
      ui.setLoading(false);
    }
  }

  function downloadChangeTemplate() {
    ui.downloadCsv(
      [
        ["月份", "人员编号", "身份证", "姓名", "操作", "字段", "值", "备注"],
        [ui.state.targetPeriod || "2026-08", "", "", "", "设置", "其他工资", "", ""],
      ],
      "工资变动模板.csv",
    );
  }

  function clearProposals() {
    ui.state.proposals = ui.state.proposals.filter(
      (proposal) => proposal.workbookEvidence,
    );
    ui.renderProposals();
  }

  function bindChangeFlow() {
    ui.byId("parseTextBtn").addEventListener("click", parseTextChanges);
    ui.byId("sourceFilesInput").addEventListener(
      "change",
      importSelectedSources,
    );
    ui.byId("applyProposalsBtn").addEventListener(
      "click",
      applySelectedProposals,
    );
    ui.byId("clearProposalsBtn").addEventListener("click", clearProposals);
    ui.byId("downloadChangeTemplateBtn").addEventListener(
      "click",
      downloadChangeTemplate,
    );
  }

  Object.assign(ui, {
    importWorkbookSource,
    importDocxSource,
    applySelectedProposals,
    bindChangeFlow,
  });
})();
