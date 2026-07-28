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
      results: [],
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

  async function inspectAttachment(file, priorResults = []) {
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
    const result = rules.resolveAttachment(
      category,
      selected.table,
      ui.state.table,
      ui.state.targetPeriod,
      file.name,
      {
        excludedTargetRows: rules.attachmentExclusionRows(
          priorResults,
          category,
        ),
      },
    );
    const warning = rules.attachmentPeriodWarning(
      selected.periodEvidence,
      ui.state.targetPeriod,
    );
    result.warnings = warning ? [warning] : [];
    return result;
  }

  async function inspectSelectedAttachments(event) {
    const files = [...(event.target.files || [])];
    if (!files.length) {
      return;
    }
    ui.setLoading(true, `正在核对 ${files.length} 个工资附件…`);
    try {
      const results = [];
      const errors = [];
      const seen = new Set();
      const orderedFiles = [...files].sort((left, right) => {
        const leftProfile = rules.profileForAttachment(
          rules.classifySource(left.name),
        );
        const rightProfile = rules.profileForAttachment(
          rules.classifySource(right.name),
        );
        return Number(Boolean(rightProfile?.excludesFrom?.length)) -
          Number(Boolean(leftProfile?.excludesFrom?.length));
      });
      for (const file of orderedFiles) {
        try {
          const result = await inspectAttachment(file, results);
          if (seen.has(result.category)) {
            errors.push(`${result.category} 选择了多个文件`);
          } else {
            seen.add(result.category);
            results.push(result);
          }
        } catch (error) {
          errors.push(error.message || `${file.name} 无法读取`);
        }
      }
      const categoryOrder = new Map(
        ui.state.attachments.required.map((item, index) => [
          item.category,
          index,
        ]),
      );
      results.sort(
        (left, right) =>
          (categoryOrder.get(left.category) ?? Number.MAX_SAFE_INTEGER) -
          (categoryOrder.get(right.category) ?? Number.MAX_SAFE_INTEGER),
      );
      const requiredCategories = new Set(
        ui.state.attachments.required.map((item) => item.category),
      );
      for (const result of results) {
        if (!requiredCategories.has(result.category)) {
          errors.push(`${result.category} 不是本次生成的必要附件`);
        }
      }
      for (const category of requiredCategories) {
        if (!seen.has(category)) {
          errors.push(`缺少${category}`);
        }
      }
      errors.push(
        ...results.flatMap((result) =>
          (result.errors || []).map(
            (error) => `${result.category}：${error}`,
          ),
        ),
      );
      if (
        results.some((result) =>
          (result.errors || []).some((error) =>
            /工资表有 \d+ 人未在该来源表中出现/.test(error),
          ),
        )
      ) {
        errors.push(
          "请先到“人员 / 工资变动”录入停用、新增或身份信息更正，再重新选择附件。",
        );
      }
      ui.state.attachments.results = results;
      ui.state.attachments.updates = results.flatMap(
        (result) => result.updates,
      );
      ui.state.attachments.errors = [...new Set(errors)];
      ui.state.attachments.applied = false;
      ui.state.attachments.detached = null;
      ui.state.sourceRouting = attachmentPlan(
        ui.state.attachments.required,
        results,
      );
      ui.renderAll();
      ui.openTab("attachments");
      if (ui.state.attachments.errors.length) {
        ui.toast(
          `已读取 ${results.length}/${files.length} 个附件，核对发现 ${ui.state.attachments.errors.length} 个待处理项`,
          "error",
        );
      } else {
        ui.toast(
          `已读取 ${results.length} 个附件，均通过核对`,
        );
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
    applyAttachments,
    invalidateAttachmentResolution,
    bindAttachmentFlow,
  });
})();
