(() => {
  "use strict";

  const ui = window.PayrollLocal.ui;
  const rules = window.PayrollEngine;

  function attachmentPlan(required, results = []) {
    const byCategory = new Map(
      results.map((result) => [result.category, result]),
    );
    return {
      route: ui.state.route?.id || "",
      basePeriod: ui.state.basePeriod,
      targetPeriod: ui.state.targetPeriod,
      errors: results.flatMap((result) => result.errors || []),
      plans: required.map((item) => {
        const result = byCategory.get(item.category);
        return {
          category: item.category,
          action: "localize",
          status: result
            ? result.errors.length
              ? "error"
              : ui.state.attachments.applied
                ? "applied"
                : "ready"
            : "missing",
          beforeFilename:
            ui.state.externalLinks.find(
              (link) => link.category === item.category,
            )?.filename || "工资表内现有值",
          afterFilename: result?.sourceName || "待选择",
          basis:
            result?.basis ||
            item.profile?.basis ||
            "等待本地附件核对",
        };
      }),
    };
  }

  function resetAttachmentState() {
    const required = rules.requiredAttachments(
      ui.state.table,
    );
    ui.state.attachments = {
      required,
      inputs: [],
      results: [],
      files: [],
      ignoredFiles: [],
      inputErrors: [],
      updates: [],
      errors: [],
      applied: false,
      detached: null,
    };
    ui.state.sourceRouting = attachmentPlan(required);
  }

  async function chooseAttachmentTable(
    workbook,
    category,
    fileName,
  ) {
    const candidates = [];
    for (const sheet of workbook.sheets) {
      const table = await workbook.getTable(sheet.name);
      const periodEvidence = rules.attachmentPeriodEvidence(
        table,
        ui.state.targetPeriod,
        fileName,
      );
      candidates.push({
        table,
        periodEvidence,
        score: rules.attachmentTableScore(
          category,
          table,
          ui.state.targetPeriod,
          fileName,
        ),
      });
    }
    candidates.sort((left, right) => right.score - left.score);
    if (!candidates.length || candidates[0].score < 100) {
      throw new Error(`${fileName} 没有找到目标月份的可识别工作表`);
    }
    if (
      candidates[1] &&
      candidates[0].score === candidates[1].score
    ) {
      throw new Error(`${fileName} 有多个同等匹配的工作表`);
    }
    return candidates[0];
  }

  async function inspectAttachment(file) {
    const category = rules.classifySource(file.name);
    if (!rules.profileForAttachment(category)) {
      throw new Error(`${file.name} 不是本次需要的工资附件`);
    }
    const workbook = await ui.loadWorkbookFile(file);
    const selected = await chooseAttachmentTable(
      workbook,
      category,
      file.name,
    );
    const warning = rules.attachmentPeriodWarning(
      selected.periodEvidence,
      ui.state.targetPeriod,
    );
    return {
      category,
      sourceName: file.name,
      sourceTable: selected.table,
      periodEvidence: selected.periodEvidence,
      warnings: warning ? [warning] : [],
    };
  }

  function clearAutomaticAttachmentSources() {
    ui.state.proposals = ui.state.proposals.filter(
      (proposal) => !proposal.automaticAttachment,
    );
    ui.state.sources = ui.state.sources.filter(
      (source) => !source.automaticAttachment,
    );
  }

  function automaticAttachmentProblems() {
    rules.markConflicts(ui.state.proposals);
    const sources = ui.state.sources.filter(
      (source) => source.automaticAttachment,
    );
    const proposals = ui.state.proposals.filter(
      (proposal) => proposal.automaticAttachment,
    );
    const errors = [
      ...sources.flatMap((source) =>
        (source.errors || []).map(
          (error) => `${source.name}：${error}`,
        ),
      ),
      ...proposals.flatMap((proposal) =>
        (proposal.errors || []).map(
          (error) => `${proposal.source}：${error}`,
        ),
      ),
    ];
    const pending = proposals.filter(
      (proposal) =>
        proposal.status !== "error" &&
        proposal.selected &&
        !proposal.redundant,
    );
    if (pending.length) {
      errors.push(
        `有 ${pending.length} 项人员 / 工资变动待确认，请先处理后再写入工资附件`,
      );
    }
    return { errors, pending };
  }

  async function refreshAttachmentResolution(message = "") {
    const audit = ui.state.attachments;
    const batch = rules.resolveAttachmentBatch(
      audit.inputs || [],
      ui.state.table,
      ui.state.targetPeriod,
    );
    const categoryOrder = new Map(
      audit.required.map((item, index) => [item.category, index]),
    );
    batch.results.sort(
      (left, right) =>
        (categoryOrder.get(left.category) ?? Number.MAX_SAFE_INTEGER) -
        (categoryOrder.get(right.category) ?? Number.MAX_SAFE_INTEGER),
    );
    const foundCategories = new Set(
      batch.results.map((result) => result.category),
    );
    const missing = audit.required
      .filter((item) => !foundCategories.has(item.category))
      .map((item) => `缺少${item.category}`);
    const automatic = automaticAttachmentProblems();
    audit.results = batch.results;
    audit.updates = batch.updates;
    audit.errors = [
      ...(audit.inputErrors || []),
      ...batch.errors,
      ...missing,
      ...automatic.errors,
    ].filter((error, index, items) => items.indexOf(error) === index);
    audit.applied = false;
    audit.detached = null;
    ui.state.sourceRouting = attachmentPlan(
      audit.required,
      audit.results,
    );
    ui.renderAll();
    ui.updateExportState();
    if (message) {
      ui.toast(message);
    }
    return audit;
  }

  async function inspectSelectedAttachments(event) {
    const selectedFiles = [...(event.target.files || [])];
    if (!selectedFiles.length) {
      return;
    }
    const ignoredFiles = selectedFiles
      .filter((file) => /^~\$/i.test(file.name))
      .map((file) => file.name);
    const files = selectedFiles.filter(
      (file) => !/^~\$/i.test(file.name),
    );
    if (!files.length) {
      ui.toast("所选文件均为 Excel 临时文件，已忽略", "error");
      event.target.value = "";
      return;
    }
    clearAutomaticAttachmentSources();
    resetAttachmentState();
    ui.state.attachments.ignoredFiles = ignoredFiles;
    ui.setLoading(true, `正在核对 ${files.length} 个工资附件…`);
    try {
      for (const file of files) {
        const category = rules.classifySource(file.name);
        const coreProfile = rules.profileForAttachment(category);
        try {
          if (coreProfile) {
            const input = await inspectAttachment(file);
            ui.state.attachments.inputs.push(input);
            ui.state.attachments.files.push({
              name: file.name,
              role: "core",
              category,
              mappingCount: coreProfile.mappings.length,
              proposalCount: 0,
              read: true,
            });
          } else {
            await ui.importWorkbookSource(file, {
              automatic: true,
              automaticAttachment: true,
              role: file.name,
            });
            const source = ui.state.sources.at(-1);
            ui.state.attachments.files.push({
              name: file.name,
              role: "change",
              category: source?.category || category,
              profileId: source?.profileId || "",
              mappingCount: source?.mappingCount || 0,
              proposalCount: source?.proposalCount || 0,
              warnings: source?.warnings || [],
              errors: source?.errors || [],
              read: true,
            });
          }
        } catch (error) {
          const problem = error.message || `${file.name} 无法读取`;
          ui.state.attachments.inputErrors.push(
            `${file.name}：${problem}`,
          );
          ui.state.attachments.files.push({
            name: file.name,
            role: coreProfile ? "core" : "change",
            category,
            mappingCount: 0,
            proposalCount: 0,
            errors: [problem],
            read: false,
          });
        }
      }
      await refreshAttachmentResolution();
      ui.openTab("attachments");
      const readCount = ui.state.attachments.files.filter(
        (file) => file.read,
      ).length;
      if (ui.state.attachments.errors.length) {
        ui.toast(
          `已读取 ${readCount}/${files.length} 个附件，核对发现 ${ui.state.attachments.errors.length} 个待处理项`,
          "error",
        );
      } else {
        ui.toast(`已读取 ${readCount} 个附件，均通过核对`);
      }
    } finally {
      ui.setLoading(false);
      event.target.value = "";
    }
  }

  async function restoreAttachmentSnapshot(snapshot) {
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
    ui.state.attachments.applied = false;
    ui.state.attachments.detached = null;
    ui.renderAll();
  }

  async function applyAttachments() {
    const audit = ui.state.attachments;
    if (audit.errors.length || !audit.updates.length) {
      ui.toast("工资附件仍有问题，不能写入", "error");
      return;
    }
    ui.setLoading(true, "正在写入本地附件并移除外链…");
    let snapshot = null;
    try {
      const exported = await ui.state.workbook.export();
      snapshot = {
        bytes: new Uint8Array(await exported.blob.arrayBuffer()),
        history: [...ui.state.history],
      };
      for (const update of audit.updates) {
        await ui.state.workbook.updateCell(
          update.targetSheet,
          update.targetRow,
          update.targetColumn,
          update.value,
          { preserveFormula: false },
        );
      }
      ui.state.table = await ui.state.workbook.getTable(
        ui.state.mainSheetName,
      );
      const detached = await ui.state.workbook.detachExternalLinks(
        ui.state.mainSheetName,
        ui.state.disabledRows,
      );
      ui.state.externalLinks =
        await ui.state.workbook.analyzeExternalLinks();
      ui.state.formulaDiagnostics =
        await ui.state.workbook.analyzeFormulas(
          ui.state.mainSheetName,
        );
      if (
        ui.state.externalLinks.length ||
        ui.state.formulaDiagnostics.externalFormulaNodes
      ) {
        throw new Error("外链移除后复检失败");
      }
      ui.state.attachments.applied = true;
      ui.state.attachments.detached = detached;
      ui.state.attachments.errors = [];
      ui.state.sourceRouting = attachmentPlan(
        audit.required,
        audit.results,
      );
      ui.state.history.push({
        time: new Date().toLocaleTimeString("zh-CN"),
        label: "写入工资附件",
        detail: `写入 ${audit.updates.length} 个人员字段，移除 ${detached.removedLinks} 条外链；${detached.frozenLegacyFormulaCount} 个停用行或旧辅助公式固定为完整缓存值。`,
        kind: "attachments",
      });
      ui.renderAll();
      ui.updateExportState();
      ui.toast("工资附件已写入，导出文件不再包含外链");
    } catch (error) {
      if (snapshot) {
        await restoreAttachmentSnapshot(snapshot);
      }
      ui.toast(
        `${error.message || "工资附件写入失败"}${snapshot ? "；本次写入已回滚" : ""}`,
        "error",
      );
      console.error(error);
    } finally {
      ui.setLoading(false);
    }
  }

  function invalidateAttachmentResolution(message = "") {
    if (!ui.state.attachments.required.length) {
      return;
    }
    resetAttachmentState();
    if (message) {
      ui.toast(message, "error");
    }
    ui.renderAll();
    ui.updateExportState();
  }

  function bindAttachmentFlow() {
    ui.byId("attachmentFilesInput").addEventListener(
      "change",
      inspectSelectedAttachments,
    );
    ui.byId("applyAttachmentsBtn").addEventListener(
      "click",
      applyAttachments,
    );
  }

  Object.assign(ui, {
    attachmentPlan,
    resetAttachmentState,
    inspectAttachment,
    refreshAttachmentResolution,
    applyAttachments,
    invalidateAttachmentResolution,
    bindAttachmentFlow,
  });
})();
