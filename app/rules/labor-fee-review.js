(() => {
  "use strict";

  const api = window.PayrollLocal.rules;
  const LABOR_FEE_REVIEW_META = Object.freeze({
    id: "labor-fee-amount-review",
    policy:
      "劳务费具体金额由人力提供；附件金额与当前草案不同时作为正常月度变动提醒，不阻断写入。劳务发票增值税先从应发中扣减，再按劳务三级预扣率计算；应纳税所得额超过20000元时不能继续使用只证明过20%档的公式。",
  });

  function laborTaxableIncome(amount, vat) {
    const net = amount - vat;
    if (net <= 0) {
      return 0;
    }
    return net <= 4000 ? Math.max(net - 800, 0) : net * 0.8;
  }

  function auditLaborFeeAmounts(
    category,
    sourceTable,
    targetTable,
    profile,
  ) {
    if (category !== "劳务费附件") {
      return { errors: [], warnings: [] };
    }
    const amountHeader = api.headerForAliases(sourceTable, ["劳务费"]);
    const vatHeader = api.headerForAliases(sourceTable, ["增值税税额"]);
    const targetHeader = api.targetHeader(targetTable, "基本工资");
    const identities = api.identityHeaders(sourceTable);
    const errors = [];
    const warnings = [];
    if (!amountHeader || !targetHeader) {
      errors.push("劳务费金额无法与工资表“基本工资”核对");
      return { errors, warnings };
    }
    const targetIndex = api.indexPeople(targetTable);
    const amountDifferences = [];
    let highTier = 0;
    for (const sourceRow of sourceTable.rows) {
      const amount = api.asNumber(sourceRow.get(amountHeader.name));
      if (amount === null) {
        continue;
      }
      const identity = api.identityFromRow(sourceRow, identities);
      const match = api.matchPerson(targetIndex, identity, {
        matchBy: profile.matchBy,
        stopAfterPresent: true,
      });
      if (match.status !== "matched") {
        continue;
      }
      const currentAmount = api.asNumber(
        match.person.row.get(targetHeader.name),
      );
      if (!api.sameValue(currentAmount, amount)) {
        amountDifferences.push({
          name: api.asText(identity.name) ||
            api.asText(match.person.name) ||
            `第 ${sourceRow.rowNumber} 行`,
          amount,
          currentAmount,
        });
      }
      const vat = vatHeader
        ? api.asNumber(sourceRow.get(vatHeader.name)) || 0
        : 0;
      if (laborTaxableIncome(amount, vat) > 20000) {
        highTier += 1;
      }
    }
    if (amountDifferences.length) {
      const details = amountDifferences
        .slice(0, 8)
        .map((item) =>
          `${item.name}：附件 ${item.amount}，当前草案 ${
            item.currentAmount === null ? "空白" : item.currentAmount
          }`,
        )
        .join("；");
      const remaining =
        amountDifferences.length > 8
          ? `；另有 ${amountDifferences.length - 8} 人`
          : "";
      warnings.push(
        `劳务费金额复核（不阻断写入）：${details}${remaining}。临时劳务费或本月新增可能形成正常差异，请财务复核`,
      );
    }
    if (highTier) {
      errors.push(
        `劳务费附件有 ${highTier} 人的应纳税所得额超过20000元，必须按30%/40%档复核税额`,
      );
    }
    return { errors, warnings };
  }

  Object.assign(api, {
    LABOR_FEE_REVIEW_META,
    auditLaborFeeAmounts,
  });
})();
