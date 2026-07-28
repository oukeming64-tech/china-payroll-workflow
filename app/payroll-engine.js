(() => {
  "use strict";

  const rules = window.PayrollLocal?.rules;
  if (
    !rules?.parseNaturalLanguage ||
    !rules?.socialBaseCandidates ||
    !rules?.resolveAttachment ||
    !rules?.monthlyBusinessPlan ||
    !rules?.proposalsFromBusinessSource
  ) {
    throw new Error("工资规则模块加载顺序不完整");
  }
  window.PayrollEngine = rules;
})();
