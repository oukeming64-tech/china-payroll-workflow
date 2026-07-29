(() => {
  "use strict";

  const rules = window.PayrollLocal?.rules;
  if (
    !rules?.parseNaturalLanguage ||
    !rules?.socialBaseCandidates ||
    !rules?.resolveAttachment ||
    !rules?.resolveAttachmentBatch ||
    !rules?.monthlyBusinessPlan ||
    !rules?.MONTHLY_CHANGE_SOURCE_PROFILES ||
    !rules?.salaryEventTiming ||
    !rules?.proposalsFromEmploymentEventSource ||
    !rules?.proposalsFromSalaryEventSource ||
    !rules?.proposalsFromAttendanceSource ||
    !rules?.proposalsFromBusinessSource ||
    !rules?.structuredWorkbookChangeResults ||
    !rules?.collectWorkbookBusinessEvidence ||
    !rules?.auditLaborFeeAmounts
  ) {
    throw new Error("工资规则模块加载顺序不完整");
  }
  window.PayrollEngine = rules;
})();
