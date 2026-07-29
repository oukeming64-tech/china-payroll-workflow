(() => {
  "use strict";

  const api = window.PayrollLocal.rules;
  const REGULAR_DAYS = Object.freeze([
    "转正天数",
    "转正后工作天数",
    "转正工作天数",
  ]);
  const TRIAL_DAYS = Object.freeze([
    "试用天数",
    "转正前工作天数",
    "试用工作天数",
  ]);
  const SALARY_EVENT_PRORATION_META = Object.freeze({
    id: "salary-event-proration",
    divisor: 22,
    policy:
      "当月转正工资按试用工资÷22×试用天数＋转正工资×80%÷22×转正天数计算；转正当月绩效工资标准按转正工资的20%÷22×转正天数计算。",
  });

  function sourceHeader(table, aliases) {
    return api.headerForAliases(table, aliases);
  }

  function excelDate(value) {
    if (value instanceof Date && !Number.isNaN(value.valueOf())) {
      return value;
    }
    const numeric = api.asNumber(value);
    if (
      numeric !== null &&
      numeric >= 1 &&
      numeric <= 2958465
    ) {
      return new Date(Date.UTC(1899, 11, 30) + numeric * 86400000);
    }
    const match = api.asText(value).match(
      /^(20\d{2})[年./-]\s*(\d{1,2})[月./-]\s*(\d{1,2})日?$/,
    );
    if (!match) {
      return null;
    }
    const date = new Date(
      Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])),
    );
    return Number.isNaN(date.valueOf()) ? null : date;
  }

  function datePeriod(date) {
    return date
      ? `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`
      : "";
  }

  function eventDate(sourceRow, sourceTable, kind) {
    const aliases =
      kind === "regularization"
        ? ["转正日期"]
        : ["调薪日期", "生效日期"];
    const header = sourceHeader(sourceTable, aliases);
    return header ? excelDate(sourceRow.get(header.name)) : null;
  }

  function workingDays(sourceRow, sourceTable, aliases) {
    const header = sourceHeader(sourceTable, aliases);
    const value = header
      ? api.asNumber(sourceRow.get(header.name))
      : null;
    return value !== null && value >= 0 && value <= 22
      ? value
      : null;
  }

  function currentWageTotal(person) {
    const components = [
      "基本工资",
      "岗位工资",
      "绩效工资标准",
    ].map((field) => api.asNumber(person.row.get(field)));
    return components.every((value) => value !== null)
      ? components.reduce((sum, value) => sum + value, 0)
      : api.asNumber(person.row.get("工资合计"));
  }

  function formulaNumber(value) {
    return Number(value.toFixed(8)).toString();
  }

  function salaryEventTiming(
    sourceRow,
    sourceTable,
    profile,
    targetPeriod,
    match,
    total,
  ) {
    const date = eventDate(
      sourceRow,
      sourceTable,
      profile.eventKind,
    );
    if (!date) {
      return {
        date,
        period: "",
        error: `第 ${sourceRow.rowNumber} 行缺少有效的${profile.eventKind === "regularization" ? "转正" : "调薪"}日期`,
      };
    }
    const period = datePeriod(date);
    if (period > targetPeriod) {
      return { date, period, future: true, error: "" };
    }
    if (profile.eventKind === "salary-adjustment") {
      return {
        date,
        period,
        error:
          period === targetPeriod && date.getUTCDate() !== 1
            ? `第 ${sourceRow.rowNumber} 行为月中调薪，需求没有提供月中调薪折算公式，已停止`
            : "",
      };
    }
    if (period !== targetPeriod || date.getUTCDate() === 1) {
      return { date, period, error: "" };
    }
    const regularDays = workingDays(
      sourceRow,
      sourceTable,
      REGULAR_DAYS,
    );
    if (regularDays === null) {
      return {
        date,
        period,
        error: `第 ${sourceRow.rowNumber} 行为月内转正，但附件缺少“转正天数”；无法按天计算转正绩效，已停止`,
      };
    }
    const trialWage = currentWageTotal(match.person);
    const regularCash = total * 0.8;
    if (
      trialWage !== null &&
      Math.abs(trialWage - regularCash) <= 0.05
    ) {
      return {
        date,
        period,
        regularDays,
        cashAdjustment: 0,
        error: "",
      };
    }
    const trialDays = workingDays(
      sourceRow,
      sourceTable,
      TRIAL_DAYS,
    );
    if (
      trialWage === null ||
      trialDays === null ||
      Math.abs(trialDays + regularDays - 22) > 0.001
    ) {
      return {
        date,
        period,
        error: `第 ${sourceRow.rowNumber} 行需同时提供合计为22天的“试用天数”和“转正天数”，才能折算当月转正工资`,
      };
    }
    const cash =
      (trialWage / 22) * trialDays +
      (regularCash / 22) * regularDays;
    return {
      date,
      period,
      regularDays,
      trialDays,
      cashAdjustment: Math.round((cash - regularCash) * 100) / 100,
      cashFormula:
        `ROUND(${formulaNumber(trialWage)}/22*${formulaNumber(trialDays)}` +
        `+${formulaNumber(total)}*80%/22*${formulaNumber(regularDays)}` +
        `-${formulaNumber(total)}*80%,2)`,
      error: "",
    };
  }

  Object.assign(api, {
    SALARY_EVENT_PRORATION_META,
    parseSalaryEventDate: excelDate,
    salaryEventDatePeriod: datePeriod,
    salaryEventTiming,
  });
})();
