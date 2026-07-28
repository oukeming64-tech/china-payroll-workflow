(() => {
  "use strict";

  const api = window.PayrollLocal.rules;
  const JULY_SOCIAL_BASE_RULE = Object.freeze({
    id: "social-base-annual-july",
    triggerMonth: 7,
    sourcePeriod: "previous-calendar-year",
    formula: "annual-wage-total / confirmed-divisor",
    defaultDivisorForFullYear: 12,
    requiresUserConfirmation: [
      "工资总额来源字段",
      "不足十二个月的分母",
      "上下限",
      "取整方式",
      "目标社保基数字段",
    ],
  });

  function isJuly(period) {
    return Number(String(period || "").slice(5, 7)) === 7;
  }

  function previousCalendarYear(period) {
    const year = Number(String(period || "").slice(0, 4));
    return Number.isFinite(year) ? year - 1 : null;
  }

  function socialBaseCandidates(
    monthlyTables,
    wageFieldName,
    divisorMode = "12",
  ) {
    const periods = new Set();
    const records = new Map();
    for (const item of monthlyTables) {
      if (periods.has(item.period)) {
        continue;
      }
      periods.add(item.period);
      const field = api.targetHeader(item.table, wageFieldName);
      if (!field) {
        continue;
      }
      for (const person of api.buildPeople(item.table).people) {
        const stableKey =
          api.normalizeText(person.employeeId) ||
          api.normalizeText(person.idCard) ||
          api.normalizeText(person.name);
        if (!stableKey) {
          continue;
        }
        if (!records.has(stableKey)) {
          records.set(stableKey, {
            key: stableKey,
            employeeId: person.employeeId,
            idCard: person.idCard,
            name: person.name,
            maskedName: person.maskedName,
            values: [],
            periods: [],
          });
        }
        const numeric = api.asNumber(person.row.get(field.name));
        if (numeric !== null) {
          const record = records.get(stableKey);
          record.values.push(numeric);
          record.periods.push(item.period);
        }
      }
    }

    return [...records.values()].map((record) => {
      const annualTotal = record.values.reduce((sum, value) => sum + value, 0);
      const coverage = new Set(record.periods).size;
      const divisor = divisorMode === "actual" ? coverage : 12;
      return {
        ...record,
        annualTotal,
        coverage,
        divisor,
        candidate: divisor > 0 ? annualTotal / divisor : null,
        status: coverage === 12 ? "ready" : "needs-confirmation",
        note:
          coverage === 12
            ? "上一自然年度 12 个月数据齐全"
            : `仅识别到 ${coverage} 个月；分母和新入职/中途入职处理需确认`,
      };
    });
  }

  Object.assign(api, {
    JULY_SOCIAL_BASE_RULE,
    isJuly,
    previousCalendarYear,
    socialBaseCandidates,
  });
})();
