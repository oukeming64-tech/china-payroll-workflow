(() => {
  "use strict";

  const api = window.PayrollLocal.rules;
  const ENTRY_DATE_ALIASES = Object.freeze([
    "入职日期",
    "入司日期",
    "入职时间",
  ]);
  const EXIT_DATE_ALIASES = Object.freeze([
    "离职日期",
    "离司日期",
    "离职时间",
  ]);
  const WORKDAY_ALIASES = Object.freeze([
    "工作天数",
    "出勤天数",
    "计薪天数",
  ]);
  const UNPAID_PERFORMANCE_ALIASES = Object.freeze([
    "未发放绩效工资",
    "未发绩效",
    "未结绩效",
  ]);
  const TRIAL_WAGE_ALIASES = Object.freeze([
    "试用工资",
    "试用期工资",
  ]);
  const EMPLOYMENT_EVENT_RULE_META = Object.freeze({
    id: "employment-entry-exit",
    policy:
      "入职工资=试用工资÷22×工作天数；离职工资=试用或转正工资÷22×工作天数＋未发绩效。入离职还需完整人员、银行、实发和结算资料，不能仅凭缺人自动新增或停用。",
  });

  function header(table, aliases) {
    return api.headerForAliases(table, aliases);
  }

  function wageTotal(person) {
    const components = [
      "基本工资",
      "岗位工资",
      "绩效工资标准",
    ].map((field) => api.asNumber(person?.row.get(field)));
    return components.every((value) => value !== null)
      ? components.reduce((sum, value) => sum + value, 0)
      : api.asNumber(person?.row.get("工资合计"));
  }

  function eventProposal(
    sourceName,
    sourceRow,
    targetPeriod,
    kind,
    match,
    message,
    warning = "",
  ) {
    const proposal = api.proposalBase(
      sourceName,
      sourceRow.rowNumber,
      "",
    );
    Object.assign(proposal, {
      kind,
      operation: kind === "new-person" ? "新增员工" : "停用",
      period: targetPeriod,
      status: "error",
      selected: false,
    });
    if (match?.status === "matched") {
      proposal.person = match.person;
      proposal.matchedBy = match.matchedBy;
    }
    if (kind === "new-person") {
      proposal.newPersonValues = {};
    }
    proposal.errors.push(message);
    if (warning) {
      proposal.warnings.push(warning);
    }
    return proposal;
  }

  function dateValue(sourceRow, sourceTable, aliases) {
    const dateHeader = header(sourceTable, aliases);
    return dateHeader
      ? api.parseSalaryEventDate(sourceRow.get(dateHeader.name))
      : null;
  }

  function numberValue(sourceRow, sourceTable, aliases) {
    const valueHeader = header(sourceTable, aliases);
    return valueHeader
      ? api.asNumber(sourceRow.get(valueHeader.name))
      : null;
  }

  function proposalsFromEmploymentEventSource(
    sourceTable,
    targetTable,
    targetPeriod,
    sourceName,
    profile,
    sourcePeriod,
  ) {
    const identities = api.identityHeaders(sourceTable);
    const errors = [];
    const warnings = [];
    const proposals = [];
    if (!identities.idCard) {
      errors.push("员工动态表缺少身份证号，不能处理入离职");
      return {
        profile,
        sourcePeriod,
        format: "employment-events",
        mappings: [],
        proposals,
        warnings,
        errors,
      };
    }
    const targetIndex = api.indexPeople(targetTable);
    for (const sourceRow of sourceTable.rows) {
      const identity = api.identityFromRow(sourceRow, identities);
      if (!api.asText(identity.idCard)) {
        continue;
      }
      const entryDate = dateValue(
        sourceRow,
        sourceTable,
        ENTRY_DATE_ALIASES,
      );
      const exitDate = dateValue(
        sourceRow,
        sourceTable,
        EXIT_DATE_ALIASES,
      );
      const eventDate = entryDate || exitDate;
      if (
        !eventDate ||
        api.salaryEventDatePeriod(eventDate) !== targetPeriod
      ) {
        continue;
      }
      const match = api.matchPerson(targetIndex, identity, {
        matchBy: ["idCard"],
        stopAfterPresent: true,
      });
      const workdays = numberValue(
        sourceRow,
        sourceTable,
        WORKDAY_ALIASES,
      );
      if (entryDate) {
        const trialWage = numberValue(
          sourceRow,
          sourceTable,
          TRIAL_WAGE_ALIASES,
        );
        const preview =
          trialWage !== null && workdays !== null
            ? `入职工资预览：${trialWage}÷22×${workdays}=${Math.round((trialWage / 22) * workdays * 100) / 100}`
            : "";
        proposals.push(
          eventProposal(
            sourceName,
            sourceRow,
            targetPeriod,
            "new-person",
            match,
            `第 ${sourceRow.rowNumber} 行为当月入职；必须补齐试用工资、工作天数、完整人员信息、银行账号和实发合计后再新增`,
            preview,
          ),
        );
        continue;
      }
      const unpaidPerformance = numberValue(
        sourceRow,
        sourceTable,
        UNPAID_PERFORMANCE_ALIASES,
      );
      const wage = match.status === "matched"
        ? wageTotal(match.person)
        : null;
      const preview =
        wage !== null &&
        workdays !== null &&
        unpaidPerformance !== null
          ? `离职工资预览：${wage}÷22×${workdays}+${unpaidPerformance}=${Math.round(((wage / 22) * workdays + unpaidPerformance) * 100) / 100}`
          : "";
      proposals.push(
        eventProposal(
          sourceName,
          sourceRow,
          targetPeriod,
          "disable-person",
          match,
          `第 ${sourceRow.rowNumber} 行为当月离职；必须补齐工作天数、未发绩效、实发合计及适用的近12个月补偿口径，先完成结算再显式停用`,
          preview,
        ),
      );
    }
    if (!proposals.length) {
      warnings.push("员工动态表没有目标月份的入职或离职记录");
    }
    return {
      profile,
      sourcePeriod,
      format: "employment-events",
      mappings: [],
      proposals,
      warnings,
      errors,
    };
  }

  Object.assign(api, {
    EMPLOYMENT_EVENT_RULE_META,
    proposalsFromEmploymentEventSource,
  });
})();
