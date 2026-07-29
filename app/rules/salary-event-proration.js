(() => {
  "use strict";

  const api = window.PayrollLocal.rules;
  const SALARY_EVENT_PRORATION_META = Object.freeze({
    id: "salary-event-proration",
    divisor: 22,
    policy:
      "转正统一按生效月份1日处理，直接使用整月转正工资构成，不要求转正天数或试用天数，也不做月中折算；月中调薪仍因缺少折算口径而停止。",
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

  function firstDay(period) {
    const match = String(period || "").match(
      /^(\d{4})-(0[1-9]|1[0-2])$/,
    );
    return match
      ? new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, 1))
      : null;
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
    if (!date && profile.eventKind !== "regularization") {
      return {
        date,
        period: "",
        error: `第 ${sourceRow.rowNumber} 行缺少有效的调薪日期`,
      };
    }
    const period = datePeriod(date) || targetPeriod;
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
    return {
      date: firstDay(period),
      sourceDate: date,
      period,
      defaultedRegularizationDay:
        !date || date.getUTCDate() !== 1,
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
