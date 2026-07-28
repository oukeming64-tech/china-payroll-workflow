(() => {
  "use strict";

  const api = window.PayrollLocal.rules;

  function attachmentPeriodEvidence(table, targetPeriod, fileName = "") {
    const sheetPeriod = api.detectPeriod(table.sheetName);
    const filePeriod = api.detectPeriod(fileName);
    if (sheetPeriod === targetPeriod) {
      return { status: "sheet-exact", sheetPeriod, filePeriod, score: 300 };
    }
    if (
      filePeriod === targetPeriod &&
      sheetPeriod &&
      api.nextPeriod(sheetPeriod) === targetPeriod
    ) {
      return {
        status: "previous-sheet-name",
        sheetPeriod,
        filePeriod,
        score: 240,
      };
    }
    if (sheetPeriod) {
      return { status: "sheet-conflict", sheetPeriod, filePeriod, score: -300 };
    }
    if (filePeriod === targetPeriod) {
      return { status: "filename-exact", sheetPeriod, filePeriod, score: 30 };
    }
    return {
      status: filePeriod ? "filename-conflict" : "unknown",
      sheetPeriod,
      filePeriod,
      score: filePeriod ? -30 : 0,
    };
  }

  function attachmentPeriodWarning(evidence, targetPeriod) {
    if (evidence?.status !== "previous-sheet-name") {
      return "";
    }
    return `工作表名仍为${api.formatPeriod(evidence.sheetPeriod)}，文件名指向${api.formatPeriod(targetPeriod)}；已按唯一字段结构读取，请确认该附件确为目标月份数据。`;
  }

  Object.assign(api, {
    attachmentPeriodEvidence,
    attachmentPeriodWarning,
  });
})();
