(() => {
  "use strict";

  const api = window.PayrollLocal.rules;

  function periodParts(period) {
    const match = String(period || "").match(/^(\d{4})-(0[1-9]|1[0-2])$/);
    return match
      ? { year: Number(match[1]), month: Number(match[2]) }
      : null;
  }

  function isJanuary(period) {
    return periodParts(period)?.month === 1;
  }

  function generationRoute(basePeriod, targetPeriod = api.nextPeriod(basePeriod)) {
    if (!basePeriod || !targetPeriod) {
      return {
        id: "invalid",
        errors: ["无法确定月度生成路线"],
      };
    }
    if (isJanuary(targetPeriod)) {
      return {
        id: "january-rollover",
        label: "一月跨年生成",
        basePeriod,
        targetPeriod,
        historyYear: Number(targetPeriod.slice(0, 4)) - 1,
        requiresFullYearHistory: true,
        errors: [],
      };
    }
    return {
      id: "regular-month",
      label: "普通月份递推",
      basePeriod,
      targetPeriod,
      requiresFullYearHistory: false,
      errors: [],
    };
  }

  function internalFormulaFieldSignature(table) {
    return table.headers
      .filter(
        (header) =>
          !/^之前月份累计/.test(header.name) &&
          table.rows.some((row) => {
            const formula = row.cells.get(header.column)?.formula || "";
            return formula && !/\[[^\]]+\]/.test(formula);
          }),
      )
      .map((header) => header.name)
      .join("\u001f");
  }

  function validateAnnualHistory(items, requiredYear) {
    const errors = [];
    const byPeriod = new Map();
    for (const item of items) {
      if (!item.period || Number(item.period.slice(0, 4)) !== requiredYear) {
        errors.push(`${item.file?.name || "未命名文件"} 不属于 ${requiredYear} 年`);
        continue;
      }
      if (byPeriod.has(item.period)) {
        errors.push(`${api.formatPeriod(item.period)} 出现重复文件`);
        continue;
      }
      byPeriod.set(item.period, item);
    }
    const expected = Array.from({ length: 12 }, (_, index) =>
      `${requiredYear}-${String(index + 1).padStart(2, "0")}`,
    );
    const missing = expected.filter((period) => !byPeriod.has(period));
    if (missing.length) {
      errors.push(`缺少 ${missing.map(api.formatPeriod).join("、")}`);
    }
    const layouts = new Set(
      [...byPeriod.values()].map((item) =>
        item.table.headers.map((header) => header.name).join("\u001f"),
      ),
    );
    if (layouts.size > 1) {
      errors.push("上一年度 12 个月主工资表字段布局不一致");
    }
    const january = byPeriod.get(`${requiredYear}-01`);
    const resetFields = [];
    if (january) {
      for (const header of january.table.headers) {
        if (!/^之前月份累计/.test(header.name)) {
          continue;
        }
        const formulaCells = january.table.rows.filter(
          (row) => row.cells.get(header.column)?.hasFormula,
        ).length;
        if (formulaCells === 0) {
          resetFields.push(header.name);
        }
      }
    }
    const regularItems = expected
      .slice(1)
      .map((period) => byPeriod.get(period))
      .filter(Boolean);
    const formulaSignatures = new Map();
    for (const item of regularItems) {
      const signature = internalFormulaFieldSignature(item.table);
      formulaSignatures.set(
        signature,
        (formulaSignatures.get(signature) || 0) + 1,
      );
    }
    const formulaVersionEvidence = Math.max(
      0,
      ...formulaSignatures.values(),
    );
    return {
      requiredYear,
      items: expected.map((period) => byPeriod.get(period)).filter(Boolean),
      periods: expected.filter((period) => byPeriod.has(period)),
      january,
      resetFields,
      formulaVersionEvidence,
      errors,
    };
  }

  Object.assign(api, {
    periodParts,
    isJanuary,
    generationRoute,
    internalFormulaFieldSignature,
    validateAnnualHistory,
  });
})();
