(() => {
  "use strict";
  const api = window.PayrollLocal.rules;
  const HISTORICAL_BUSINESS_SOURCE_PROFILES = Object.freeze([
    Object.freeze({
      id: "intern-allowance",
      label: "实习生津贴表",
      filenameAny: Object.freeze(["实习生津贴"]),
      requiredHeaders: Object.freeze([
        Object.freeze(["姓名"]),
        Object.freeze(["身份证号", "身份证"]),
        Object.freeze(["津贴总额"]),
      ]),
      mappings: Object.freeze([
        Object.freeze({
          source: Object.freeze(["津贴总额"]),
          target: "基本工资",
          basis:
            "2025年2月历史附件与正式工资表逐人核对，2/2 个可比人员的津贴总额对应基本工资",
        }),
      ]),
      matchBy: Object.freeze(["idCard"]),
      onlyChanged: true,
      requirePeriod: true,
    }),
    Object.freeze({
      id: "confidential-allowance-roster",
      label: "保密补贴统计表",
      filenameAny: Object.freeze([
        "\u6d89\u5bc6人员统计",
        "保密人员统计",
      ]),
      requiredHeaders: Object.freeze([
        Object.freeze(["姓名"]),
        Object.freeze(["津贴"]),
      ]),
      mappings: Object.freeze([
        Object.freeze({
          source: Object.freeze(["津贴"]),
          target: "BM津贴",
          basis:
            "2025年4月历史保密津贴附件与正式工资表逐人核对，17/17 人的津贴对应BM津贴",
        }),
      ]),
      requirePeriod: true,
    }),
    Object.freeze({
      id: "confidential-allowance-approval",
      label: "保密补贴审批表",
      filenameAny: Object.freeze(["保密补贴发放审批", "保密补贴"]),
      requiredHeaders: Object.freeze([
        Object.freeze(["姓名"]),
        Object.freeze(["实发数额"]),
      ]),
      mappings: Object.freeze([
        Object.freeze({
          source: Object.freeze(["实发数额"]),
          target: "BM津贴",
          basis:
            "2025年9月历史审批表与正式工资表逐人核对，16/16 人的实发数额对应BM津贴",
        }),
      ]),
      requirePeriod: true,
    }),
    Object.freeze({
      id: "confidential-eligibility",
      label: "保密人员范围资料",
      filenameAny: Object.freeze([
        "新增\u6d89\u5bc6人员信息",
        "xgs369人员名单",
        "xgs369",
      ]),
      requiredHeaders: Object.freeze([Object.freeze(["姓名"])]),
      mappings: Object.freeze([]),
      nonActionable:
        "该资料只证明保密人员范围或生效时间，没有提供可写入工资表的补贴金额；请同时提供保密补贴金额表或在文字变动中明确金额。",
      requirePeriod: false,
    }),
  ]);
  const BUSINESS_SOURCE_PROFILES = Object.freeze([
    ...(api.MONTHLY_CHANGE_SOURCE_PROFILES || []),
    ...HISTORICAL_BUSINESS_SOURCE_PROFILES,
  ]);
  const BUSINESS_SOURCE_RULE_META = Object.freeze({
    id: "business-source-adapters",
    trigger: "选择目标月份全量附件或在人员 / 工资变动入口选择业务附件",
    policy:
      "身份证存在时只按身份证匹配；只转换附件明确给出的目标字段，缺金额资料只提示，不推算金额",
    profiles: Object.freeze(
      BUSINESS_SOURCE_PROFILES.map((profile) => profile.id),
    ),
  });
  function headerForAny(table, aliases) {
    return api.headerForAliases(table, aliases);
  }
  function normalizedFilenameMatches(profile, fileName) {
    const key = api.normalizeText(fileName);
    return profile.filenameAny.some((term) =>
      key.includes(api.normalizeText(term)),
    );
  }
  function profileHeadersMatch(profile, table) {
    return profile.requiredHeaders.every((aliases) =>
      headerForAny(table, aliases),
    );
  }
  function matchBusinessSource(table, fileName) {
    return (
      BUSINESS_SOURCE_PROFILES.find(
        (profile) =>
          (profile.structureMatch === "attendance" &&
            api.isAttendanceSourceTable(table)) ||
          (normalizedFilenameMatches(profile, fileName) &&
            profileHeadersMatch(profile, table)),
      ) || null
    );
  }
  function businessSourcePeriod(fileName, sheetName = "") {
    const regular = api.detectPeriod(fileName, sheetName);
    if (regular) {
      return regular;
    }
    for (const value of [fileName, sheetName]) {
      const match = String(value || "").match(
        /(?:^|[^\d])(2\d)\s*[./-]\s*(0?[1-9]|1[0-2])\s*[./-]\s*(?:0?[1-9]|[12]\d|3[01])(?:[^\d]|$)/,
      );
      if (match) {
        return `20${match[1]}-${String(Number(match[2])).padStart(2, "0")}`;
      }
    }
    return "";
  }
  function resolvedMappings(profile, sourceTable, targetTable) {
    return profile.mappings.map((mapping) => ({
      sourceHeader: headerForAny(sourceTable, mapping.source),
      targetHeader: api.targetHeader(targetTable, mapping.target),
      sourceField: mapping.source[0],
      targetField: mapping.target,
      basis: mapping.basis,
    }));
  }
  function matchSourcePerson(targetIndex, identity, profile) {
    return api.matchPerson(targetIndex, identity, {
      matchBy: profile.matchBy || ["employeeId", "idCard", "name"],
      stopAfterPresent: Boolean(profile.matchBy),
    });
  }

  function proposalsFromBusinessSource(
    sourceTable,
    targetTable,
    targetPeriod,
    sourceName,
    profile = matchBusinessSource(sourceTable, sourceName),
  ) {
    if (!profile) {
      return {
        profile: null,
        format: "business-source",
        mappings: [],
        proposals: [],
        errors: ["没有找到经历史证明的业务附件规则"],
      };
    }
    const sourcePeriod = businessSourcePeriod(
      sourceName,
      sourceTable.sheetName,
    );
    if (profile.requirePeriod && !sourcePeriod) {
      return {
        profile,
        format: "business-source",
        mappings: [],
        proposals: [],
        errors: ["无法从业务附件确认生效月份"],
      };
    }
    if (
      sourcePeriod &&
      targetPeriod &&
      sourcePeriod !== targetPeriod
    ) {
      return {
        profile,
        format: "business-source",
        mappings: [],
        proposals: [],
        errors: [
          `业务附件月份 ${api.formatPeriod(sourcePeriod)} 与目标月份 ${api.formatPeriod(targetPeriod)} 不一致`,
        ],
      };
    }
    if (profile.adapter === "attendance-deductions") {
      return api.proposalsFromAttendanceSource(
        sourceTable,
        targetTable,
        targetPeriod,
        sourceName,
        profile,
        sourcePeriod,
      );
    }
    if (profile.adapter === "salary-events") {
      return api.proposalsFromSalaryEventSource(
        sourceTable,
        targetTable,
        targetPeriod,
        sourceName,
        profile,
        sourcePeriod,
      );
    }
    if (profile.adapter === "employment-events") {
      return api.proposalsFromEmploymentEventSource(
        sourceTable,
        targetTable,
        targetPeriod,
        sourceName,
        profile,
        sourcePeriod,
      );
    }
    const identities = api.identityHeaders(sourceTable);
    const errors = [];
    const warnings = [];
    const requiredIdentity = profile.matchBy?.[0];
    if (requiredIdentity && !identities[requiredIdentity]) {
      errors.push(
        requiredIdentity === "idCard"
          ? "业务附件缺少身份证号，不能按身份证匹配"
          : "业务附件缺少必要的人员匹配字段",
      );
    } else if (
      !identities.employeeId &&
      !identities.idCard &&
      !identities.name
    ) {
      errors.push("业务附件缺少人员编号、身份证或姓名");
    }
    if (profile.nonActionable) {
      errors.push(profile.nonActionable);
    }
    if (errors.length) {
      return {
        profile,
        sourcePeriod,
        format: "business-source",
        mappings: [],
        proposals: [],
        warnings,
        errors,
      };
    }
    if (profile.reviewOnly) {
      const targetIndex = api.indexPeople(targetTable);
      let matchedPeople = 0;
      let unmatchedPeople = 0;
      for (const sourceRow of sourceTable.rows) {
        const identity = api.identityFromRow(sourceRow, identities);
        if (!profile.matchBy.some((kind) => api.asText(identity[kind]))) {
          continue;
        }
        const match = matchSourcePerson(targetIndex, identity, profile);
        if (match.status === "matched") {
          matchedPeople += 1;
        } else {
          unmatchedPeople += 1;
        }
      }
      if (unmatchedPeople) {
        errors.push(`来源表有 ${unmatchedPeople} 人未按身份证匹配工资表`);
      }
      if (!matchedPeople && !unmatchedPeople) {
        errors.push("业务附件没有可读取的身份证人员行");
      }
      warnings.push(profile.reviewOnly);
      return {
        profile,
        sourcePeriod,
        format: "business-source-review",
        mappings: [],
        proposals: [],
        matchedPeople,
        unmatchedPeople,
        warnings,
        errors,
      };
    }
    const mappings = resolvedMappings(
      profile,
      sourceTable,
      targetTable,
    );
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
        format: "business-source",
        mappings,
        proposals: [],
        warnings,
        errors,
      };
    }

    const targetIndex = api.indexPeople(targetTable);
    const proposals = [];
    let mappedValues = 0;
    for (const sourceRow of sourceTable.rows) {
      const populatedMappings = mappings.filter((mapping) =>
        api.asText(sourceRow.get(mapping.sourceHeader.name)),
      );
      if (!populatedMappings.length) {
        continue;
      }
      mappedValues += populatedMappings.length;
      const identity = api.identityFromRow(sourceRow, identities);
      const match = matchSourcePerson(targetIndex, identity, profile);
      for (const mapping of populatedMappings) {
        const inputValue = sourceRow.get(mapping.sourceHeader.name);
        const currentValue =
          match.status === "matched"
            ? match.person.row.get(mapping.targetHeader.name)
            : null;
        if (
          profile.onlyChanged &&
          match.status === "matched" &&
          api.sameValue(currentValue, inputValue)
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
          inputValue,
          mapping,
          sourceKind: profile.id,
          sourceValues: Object.fromEntries(sourceTable.headers.map((header) => [header.name, sourceRow.values.get(header.column)])),
        });
        if (match.status !== "matched") {
          proposal.status = "error";
          proposal.selected = false;
          proposal.errors.push(
            `${match.message}；如为当月新增人员，请先用人员变动表完成新增，再重新导入本资料`,
          );
        } else {
          proposal.person = match.person;
          proposal.matchedBy = match.matchedBy;
          proposal.currentValue = currentValue;
        }
        const sourceCell = sourceRow.cells.get(
          mapping.sourceHeader.column,
        );
        if (sourceCell?.hasFormula) {
          proposal.warnings.push(
            "来源金额含公式；仅使用工作簿保存的缓存值，不复制来源公式",
          );
        }
        proposals.push(proposal);
      }
    }
    if (!proposals.length && !mappedValues) {
      errors.push("业务附件没有形成可核对的人员金额");
    } else if (!proposals.length) {
      warnings.push("附件字段与当前草案一致，没有形成重复写入项");
    }
    return {
      profile,
      sourcePeriod,
      format: "business-source",
      mappings,
      proposals,
      warnings,
      errors,
    };
  }

  Object.assign(api, {
    BUSINESS_SOURCE_PROFILES,
    BUSINESS_SOURCE_RULE_META,
    matchBusinessSource,
    businessSourcePeriod,
    proposalsFromBusinessSource,
  });
})();
