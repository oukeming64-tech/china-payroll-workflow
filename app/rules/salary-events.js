(() => {
  "use strict";
  const api = window.PayrollLocal.rules;
  const COMPONENTS = Object.freeze([
    Object.freeze({
      source: Object.freeze(["基本工资"]),
      target: "基本工资",
      ratio: 0.3,
    }),
    Object.freeze({
      source: Object.freeze(["岗位工资"]),
      target: "岗位工资",
      ratio: 0.5,
    }),
    Object.freeze({
      source: Object.freeze(["绩效工资", "绩效工资标准"]),
      target: "绩效工资标准",
      ratio: 0.2,
    }),
  ]);
  const SALARY_EVENT_RULE_META = Object.freeze({
    id: "salary-events",
    trigger: "目标月份入职转正薪资表包含转正或调薪分区",
    policy:
      "按身份证匹配；工资构成必须符合基本30%、岗位50%、绩效20%。月内转正必须提供转正天数以按天计算绩效；试用工资与转正工资80%不同时还要提供试用天数。月中调薪没有确认折算规则时停止。",
  });
  function sourceHeader(table, aliases) {
    return api.headerForAliases(table, aliases);
  }
  function resolvedMappings(profile, sourceTable, targetTable) {
    return profile.mappings.map((mapping) => ({
      sourceHeader: sourceHeader(sourceTable, mapping.source),
      targetHeader: api.targetHeader(targetTable, mapping.target),
      sourceField: mapping.source[0],
      targetField: mapping.target,
      basis: mapping.basis,
    }));
  }
  function errorProposal(
    sourceName,
    sourceRow,
    targetPeriod,
    mapping,
    message,
    match = null,
  ) {
    const proposal = api.proposalBase(
      sourceName,
      sourceRow.rowNumber,
      "",
    );
    Object.assign(proposal, {
      kind: "cell-change",
      operation: "设置",
      period: targetPeriod,
      field: mapping?.targetHeader || null,
      mapping,
      status: "error",
      selected: false,
    });
    if (match?.status === "matched") {
      proposal.person = match.person;
      proposal.matchedBy = match.matchedBy;
      proposal.currentValue = mapping?.targetHeader
        ? match.person.row.get(mapping.targetHeader.name)
        : null;
    }
    proposal.errors.push(message);
    return proposal;
  }
  function salaryComponentValues(
    sourceRow,
    sourceTable,
    profile,
  ) {
    if (profile.eventKind === "salary-adjustment") {
      const totalHeader = sourceHeader(sourceTable, ["调整后薪酬"]);
      const total = totalHeader
        ? api.asNumber(sourceRow.get(totalHeader.name))
        : null;
      return {
        total,
        values:
          total === null
            ? []
            : COMPONENTS.map((component) => ({
                ...component,
                value: Math.round(total * component.ratio * 100) / 100,
                sourceHeader: totalHeader,
              })),
      };
    }
    const values = COMPONENTS.map((component) => {
      const header = sourceHeader(sourceTable, component.source);
      return {
        ...component,
        sourceHeader: header,
        value: header ? api.asNumber(sourceRow.get(header.name)) : null,
      };
    });
    return {
      total: values.every((item) => item.value !== null)
        ? values.reduce((sum, item) => sum + item.value, 0)
        : null,
      values,
    };
  }
  function componentError(
    sourceRow,
    sourceTable,
    profile,
    components,
  ) {
    if (
      components.total === null ||
      components.values.some((item) => item.value === null)
    ) {
      return `第 ${sourceRow.rowNumber} 行缺少可读取的工资金额`;
    }
    const totalHeader =
      profile.eventKind === "regularization"
        ? sourceHeader(sourceTable, ["合计", "工资合计"])
        : null;
    const statedTotal = totalHeader
      ? api.asNumber(sourceRow.get(totalHeader.name))
      : null;
    if (
      statedTotal !== null &&
      Math.abs(statedTotal - components.total) > 0.05
    ) {
      return `第 ${sourceRow.rowNumber} 行工资构成之和与合计不一致`;
    }
    if (
      components.values.some(
        (item) =>
          Math.abs(item.value - components.total * item.ratio) > 0.05,
      )
    ) {
      return `第 ${sourceRow.rowNumber} 行工资构成不符合基本30%、岗位50%、绩效20%`;
    }
    return "";
  }
  function applyPerformanceProration(components, timing) {
    if (timing.regularDays === undefined) {
      return;
    }
    const performance = components.values.find(
      (component) => component.target === "绩效工资标准",
    );
    if (!performance) {
      return;
    }
    performance.fullValue = performance.value;
    performance.value =
      Math.round(
        ((performance.fullValue / 22) * timing.regularDays +
          Number.EPSILON) *
          100,
      ) / 100;
    performance.formula =
      `ROUND(${performance.fullValue}/22*${timing.regularDays},2)`;
  }
  function cashAdjustmentProposal(
    sourceName,
    sourceRow,
    targetPeriod,
    targetTable,
    match,
    timing,
  ) {
    if (!timing.cashAdjustment) {
      return null;
    }
    const field = api.targetHeader(targetTable, "其他工资");
    const mapping = {
      sourceField: "试用天数、转正天数",
      targetField: "其他工资",
      targetHeader: field,
      basis: api.SALARY_EVENT_PRORATION_META.policy,
    };
    if (!field) {
      return errorProposal(
        sourceName,
        sourceRow,
        targetPeriod,
        mapping,
        "工资表缺少“其他工资”，不能承接转正当月折算差额",
        match,
      );
    }
    const proposal = api.proposalBase(
      sourceName,
      sourceRow.rowNumber,
      "",
    );
    Object.assign(proposal, {
      kind: "cell-change",
      operation: "设置",
      period: targetPeriod,
      field,
      inputValue: timing.cashAdjustment,
      formula: timing.cashFormula,
      currentValue: match.person.row.get(field.name),
      person: match.person,
      matchedBy: match.matchedBy,
      mapping,
      sourceKind: "regularization-proration",
    });
    return proposal;
  }
  function proposalsFromSalaryEventSource(
    sourceTable,
    targetTable,
    targetPeriod,
    sourceName,
    profile,
    sourcePeriod,
  ) {
    const mappings = resolvedMappings(
      profile,
      sourceTable,
      targetTable,
    );
    const errors = [];
    const warnings = [];
    const proposals = [];
    const identities = api.identityHeaders(sourceTable);
    if (!identities.idCard) {
      errors.push("入职转正薪资表缺少身份证号，不能按身份证匹配");
    }
    for (const mapping of mappings) {
      if (!mapping.sourceHeader) {
        errors.push(`业务附件缺少“${mapping.sourceField}”`);
      }
      if (!mapping.targetHeader) {
        errors.push(`工资表缺少“${mapping.targetField}”`);
      }
    }
    if (errors.length) {
      return {
        profile,
        sourcePeriod,
        format: "salary-events",
        mappings,
        proposals,
        warnings,
        errors,
      };
    }
    const targetIndex = api.indexPeople(targetTable);
    let matchedPeople = 0;
    let unmatchedPeople = 0;
    for (const sourceRow of sourceTable.rows) {
      const identity = api.identityFromRow(sourceRow, identities);
      if (!api.asText(identity.idCard)) {
        continue;
      }
      const match = api.matchPerson(targetIndex, identity, {
        matchBy: ["idCard"],
        stopAfterPresent: true,
      });
      if (match.status !== "matched") {
        unmatchedPeople += 1;
        proposals.push(
          errorProposal(
            sourceName,
            sourceRow,
            targetPeriod,
            mappings[0],
            `${match.message}；如为当月新增人员，请先用完整人员变动资料新增`,
          ),
        );
        continue;
      }
      matchedPeople += 1;
      const components = salaryComponentValues(
        sourceRow,
        sourceTable,
        profile,
      );
      const timing = api.salaryEventTiming(
        sourceRow,
        sourceTable,
        profile,
        targetPeriod,
        match,
        components.total,
      );
      const rowError =
        componentError(sourceRow, sourceTable, profile, components) ||
        timing.error;
      if (timing.future) {
        warnings.push(
          `第 ${sourceRow.rowNumber} 行尚未到生效月份，未写入本月工资表`,
        );
        continue;
      }
      if (rowError) {
        proposals.push(
          errorProposal(
            sourceName,
            sourceRow,
            targetPeriod,
            mappings[0],
            rowError,
            match,
          ),
        );
        continue;
      }
      applyPerformanceProration(components, timing);
      if (timing.regularDays !== undefined) {
        warnings.push(
          `第 ${sourceRow.rowNumber} 行转正绩效按“20%绩效标准÷22×${timing.regularDays}天”计算`,
        );
      }
      for (const [index, component] of components.values.entries()) {
        const mapping = mappings[index];
        const currentValue = match.person.row.get(
          mapping.targetHeader.name,
        );
        if (
          profile.onlyChanged &&
          api.sameValue(currentValue, component.value)
        ) {
          continue;
        }
        const proposal = api.proposalBase(
          sourceName,
          sourceRow.rowNumber,
          "",
        );
        Object.assign(proposal, {
          kind: "cell-change",
          operation: "设置",
          period: targetPeriod,
          field: mapping.targetHeader,
          inputValue: component.value,
          currentValue,
          person: match.person,
          matchedBy: match.matchedBy,
          mapping,
          sourceKind: profile.id,
          formula: component.formula || "",
        });
        if (component.sourceHeader) {
          const sourceCell = sourceRow.cells.get(
            component.sourceHeader.column,
          );
          if (sourceCell?.hasFormula) {
            proposal.warnings.push(
              "来源金额含公式；只采用文件保存的结果值，不复制外链公式",
            );
          }
        }
        proposals.push(proposal);
      }
      const adjustment = cashAdjustmentProposal(
        sourceName,
        sourceRow,
        targetPeriod,
        targetTable,
        match,
        timing,
      );
      if (adjustment) {
        proposals.push(adjustment);
      }
    }
    if (unmatchedPeople) {
      errors.push(`来源表有 ${unmatchedPeople} 人未按身份证匹配工资表`);
    }
    if (!proposals.length && !warnings.length && !errors.length) {
      warnings.push("工资标准与当前草案一致，没有形成重复写入项");
    }
    return {
      profile,
      sourcePeriod,
      format: "salary-events",
      mappings,
      proposals,
      matchedPeople,
      unmatchedPeople,
      warnings,
      errors,
    };
  }
  Object.assign(api, {
    SALARY_EVENT_RULE_META,
    proposalsFromSalaryEventSource,
  });
})();
