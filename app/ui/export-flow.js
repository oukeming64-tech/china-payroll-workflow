(() => {
  "use strict";

  const ui = window.PayrollLocal.ui;
  const rules = window.PayrollEngine;

  function updateExportState() {
    ui.byId("exportWorkbookBtn").disabled =
      !ui.state.workbook ||
      !ui.state.attachments.applied ||
      !ui.byId("exportAcknowledged").checked;
  }

  async function exportWorkbook() {
    const filename = ui.byId("outputFileName").value.trim();
    if (!filename) {
      ui.toast("请填写新文件名", "error");
      return;
    }
    if (!/\.xlsx$/i.test(filename)) {
      ui.toast("新文件名必须以 .xlsx 结尾", "error");
      return;
    }
    const unresolved = ui.state.proposals.filter(
      (proposal) => proposal.status === "error",
    ).length;
    if (unresolved) {
      ui.toast(
        `仍有 ${unresolved} 项匹配错误，已停止生成`,
        "error",
      );
      return;
    }
    if (!ui.state.attachments.applied) {
      ui.toast("请先完成目标月份工资附件核对", "error");
      return;
    }
    if (
      ui.state.externalLinks.length ||
      ui.state.formulaDiagnostics?.externalFormulaNodes
    ) {
      ui.toast("当前草案仍含外链，已停止生成", "error");
      return;
    }
    ui.setLoading(true, "正在生成新的工资表…");
    try {
      const currentHash = await ui.sha256File(ui.state.baseFile);
      if (currentHash !== ui.state.baseHash) {
        throw new Error("基线原文件哈希发生变化，已停止导出");
      }
      const result = await ui.state.workbook.export();
      ui.downloadBlob(result.blob, filename);
      ui.state.history.push({
        time: new Date().toLocaleTimeString("zh-CN"),
        label: "生成新工资表",
        detail: "本地附件已写入，导出文件无外链；原文件未改动。",
        kind: "export",
      });
      ui.renderHistory();
      ui.toast("新工资表已生成，请用 Excel 完成最终复核");
    } catch (error) {
      ui.toast(error.message || "工资表导出失败", "error");
      console.error(error);
    } finally {
      ui.setLoading(false);
    }
  }

  async function resetWorkingChanges() {
    if (!ui.state.workingSourceFile) {
      return;
    }
    ui.setLoading(true, "正在恢复本次草案…");
    try {
      ui.state.workbook = await window.XlsxEngine.XlsxWorkbook.load(
        ui.state.workingSourceFile,
      );
      ui.state.mainSheetName = await ui.chooseMainSheet(ui.state.workbook);
      ui.state.table = await ui.state.workbook.getTable(
        ui.state.mainSheetName,
      );
      ui.state.externalLinks =
        await ui.state.workbook.analyzeExternalLinks();
      ui.state.formulaDiagnostics =
        await ui.state.workbook.analyzeFormulas(
          ui.state.mainSheetName,
        );
      ui.state.proposals = [];
      ui.state.sources = [];
      ui.state.sourceRouting = null;
      ui.state.disabledRows = new Set();
      ui.state.workbookSync = [];
      await ui.finalizeWorkspace();
      ui.toast("本次修改已重置");
    } catch (error) {
      ui.toast(error.message || "重置失败", "error");
      console.error(error);
    } finally {
      ui.setLoading(false);
    }
  }

  function bindExportFlow() {
    ui.byId("exportAcknowledged").addEventListener(
      "change",
      updateExportState,
    );
    ui.byId("exportWorkbookBtn").addEventListener("click", exportWorkbook);
    ui.byId("resetChangesBtn").addEventListener(
      "click",
      resetWorkingChanges,
    );
  }

  Object.assign(ui, {
    updateExportState,
    exportWorkbook,
    resetWorkingChanges,
    bindExportFlow,
  });
})();
