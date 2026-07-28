(() => {
  "use strict";

  const api = window.PayrollLocal.rules;

  const ATTACHMENT_PROFILES = Object.freeze({
    "个税工资薪金附件": Object.freeze({
      id: "tax-salary",
      matchBy: Object.freeze(["idCard", "employeeId", "name"]),
      sourceIdentity: Object.freeze(["证件号码"]),
      mappings: Object.freeze([
        ["累计子女教育", "累计子女教育"],
        ["累计继续教育", "累计继续教育"],
        ["累计住房贷款利息", "累计住房贷款利息"],
        ["累计住房租金", "累计住房租金"],
        ["累计赡养老人", "累计赡养老人"],
        ["累计3岁以下婴幼儿照护", "累计婴幼儿照护费用"],
        ["累计个人养老金", "累计个人养老金扣除"],
      ]),
      basis: "工资表精确 VLOOKUP 已证明证件号码与 7 项累计专项扣除列",
    }),
    "社保 / 公积金附件": Object.freeze({
      id: "social-insurance",
      matchBy: Object.freeze(["idCard", "name"]),
      sourceIdentity: Object.freeze(["身份证号"]),
      mappings: Object.freeze([
        ["养老个人", "代扣养老保险"],
        ["医疗个人", "代扣医疗保险"],
        ["失业个人", "代扣失业保险"],
        ["公积金个人", "代扣房积金"],
        ["个人合计", "个人承担社保部分扣款合计"],
        ["公司合计", "企业承担社保部分扣款合计"],
      ]),
      basis: "逐月保险表字段与工资表实际值、公司合计外链公式共同证明",
    }),
    "劳务费附件": Object.freeze({
      id: "labor-fee",
      matchBy: Object.freeze(["name"]),
      sourceIdentity: Object.freeze(["姓名"]),
      mappings: Object.freeze([["增值税税额", "扣减税额"]]),
      excludesFrom: Object.freeze([
        "个税工资薪金附件",
        "社保 / 公积金附件",
      ]),
      basis: "已验证的一月结构证据证明姓名与扣减税额列的精确查找关系",
    }),
  });

  function profileForAttachment(category) {
    return ATTACHMENT_PROFILES[category] || null;
  }

  function requiredAttachments(targetTable) {
    const required = [
      "个税工资薪金附件",
      "社保 / 公积金附件",
    ];
    const hasLaborField = Boolean(
      api.targetHeader(targetTable, "扣减税额"),
    );
    if (hasLaborField) {
      required.push("劳务费附件");
    }
    return required.map((category) => ({
      category,
      profile: profileForAttachment(category),
      status: "missing",
    }));
  }

  function headerForSource(table, name) {
    const key = api.normalizeText(name);
    return (
      table.headers.find(
        (header) => api.normalizeText(header.name) === key,
      ) || null
    );
  }

  function attachmentTableScore(
    category,
    table,
    targetPeriod,
    fileName = "",
  ) {
    const profile = profileForAttachment(category);
    if (!profile) {
      return -Infinity;
    }
    const mappedHeaders = profile.mappings.filter(([sourceField]) =>
      headerForSource(table, sourceField),
    ).length;
    const identityHeaders = profile.sourceIdentity.filter((name) =>
      headerForSource(table, name),
    ).length;
    const sheetPeriod = api.detectPeriod(table.sheetName);
    const filePeriod = api.detectPeriod(fileName);
    const periodScore = sheetPeriod
      ? sheetPeriod === targetPeriod
        ? 300
        : -300
      : filePeriod === targetPeriod
        ? 30
        : filePeriod
          ? -30
          : 0;
    return (
      periodScore +
      mappedHeaders * 20 +
      identityHeaders * 100 +
      Math.min(table.rows.length, 20)
    );
  }

  function sourceIdentityKey(identity, matchBy) {
    for (const kind of matchBy) {
      const key = api.normalizeText(identity[kind]);
      if (key) {
        return `${kind}:${key}`;
      }
    }
    return "";
  }

  function sameValue(left, right) {
    const leftNumber = api.asNumber(left);
    const rightNumber = api.asNumber(right);
    if (leftNumber !== null && rightNumber !== null) {
      return Math.abs(leftNumber - rightNumber) < 0.001;
    }
    return api.normalizeText(left) === api.normalizeText(right);
  }

  function expectedPeople(
    targetTable,
    category,
    mappings,
    excludedRows,
  ) {
    return api.buildPeople(targetTable).people.filter((person) => {
      if (excludedRows.has(person.rowNumber)) {
        return false;
      }
      if (category !== "劳务费附件" && !api.asText(person.idCard)) {
        return false;
      }
      return mappings.some((mapping) => {
        const cell = person.row.cells.get(mapping.targetHeader.column);
        const value = person.row.values.get(mapping.targetHeader.column);
        return (
          cell?.hasFormula ||
          (value !== null &&
            value !== undefined &&
            String(value).trim() !== "")
        );
      });
    });
  }

  function sourceValueError(cell, value) {
    if (cell?.hasFormula && !cell.hasCachedValue) {
      return "来源公式缺少缓存值";
    }
    if (
      cell?.type === "e" ||
      /^#(?:REF|VALUE|N\/A|NAME|DIV\/0|NUM|NULL)!?$/i.test(
        api.asText(value),
      )
    ) {
      return "来源单元格为 Excel 错误值";
    }
    return "";
  }

  function resolveAttachment(
    category,
    sourceTable,
    targetTable,
    targetPeriod,
    sourceName,
    options = {},
  ) {
    const profile = profileForAttachment(category);
    const errors = [];
    if (!profile) {
      return {
        category,
        sourceName,
        mappings: [],
        updates: [],
        fieldSummaries: [],
        matchedPeople: 0,
        errors: [`${category} 没有经验证的字段规则`],
      };
    }
    const sourcePeriod = api.detectPeriod(
      sourceName,
      sourceTable.sheetName,
    );
    if (sourcePeriod && sourcePeriod !== targetPeriod) {
      errors.push(
        `来源月份 ${api.formatPeriod(sourcePeriod)} 与目标月份 ${api.formatPeriod(targetPeriod)} 不一致`,
      );
    }
    if (!sourcePeriod && category !== "劳务费附件") {
      errors.push("无法从文件名或工作表确认来源月份");
    }
    const mappings = profile.mappings
      .map(([sourceNameValue, targetName]) => ({
        sourceHeader: headerForSource(sourceTable, sourceNameValue),
        targetHeader: api.targetHeader(targetTable, targetName),
        sourceField: sourceNameValue,
        targetField: targetName,
        basis: profile.basis,
      }))
      .filter((mapping) => {
        if (!mapping.sourceHeader) {
          errors.push(`来源表缺少“${mapping.sourceField}”`);
        }
        if (!mapping.targetHeader) {
          errors.push(`工资表缺少“${mapping.targetField}”`);
        }
        return mapping.sourceHeader && mapping.targetHeader;
      });
    const sourceIdentities = api.identityHeaders(sourceTable);
    if (
      !profile.matchBy.some((kind) => sourceIdentities[kind])
    ) {
      errors.push("来源表缺少可验证的人员匹配字段");
    }

    const targetIndex = api.indexPeople(targetTable);
    const expected = expectedPeople(
      targetTable,
      category,
      mappings,
      new Set(options.excludedTargetRows || []),
    );
    const expectedRows = new Set(
      expected.map((person) => person.rowNumber),
    );
    const matchedRows = new Set();
    const duplicateKeys = new Set();
    const sourceKeys = new Set();
    const updates = [];
    const fieldSummaryMap = new Map(
      mappings.map((mapping) => [
        mapping.targetField,
        {
          sourceField: mapping.sourceField,
          targetField: mapping.targetField,
          matched: 0,
          changed: 0,
          formulaValues: 0,
        },
      ]),
    );
    let unmatchedSource = 0;
    let invalidValues = 0;

    for (const sourceRow of sourceTable.rows) {
      const identity = api.identityFromRow(
        sourceRow,
        sourceIdentities,
      );
      const key = sourceIdentityKey(identity, profile.matchBy);
      if (!key) {
        continue;
      }
      if (sourceKeys.has(key)) {
        duplicateKeys.add(key);
        continue;
      }
      sourceKeys.add(key);
      const match = api.matchPerson(targetIndex, identity);
      if (match.status !== "matched") {
        unmatchedSource += 1;
        continue;
      }
      matchedRows.add(match.person.rowNumber);
      for (const mapping of mappings) {
        const value = sourceRow.values.get(mapping.sourceHeader.column);
        const sourceCell = sourceRow.cells.get(
          mapping.sourceHeader.column,
        );
        const valueError = sourceValueError(sourceCell, value);
        if (valueError) {
          invalidValues += 1;
          continue;
        }
        const currentValue = match.person.row.values.get(
          mapping.targetHeader.column,
        );
        const summary = fieldSummaryMap.get(mapping.targetField);
        summary.matched += 1;
        summary.changed += sameValue(currentValue, value) ? 0 : 1;
        summary.formulaValues += sourceCell?.hasFormula ? 1 : 0;
        updates.push({
          category,
          sourceName,
          sourceSheet: sourceTable.sheetName,
          sourceRow: sourceRow.rowNumber,
          sourceField: mapping.sourceField,
          targetSheet: targetTable.sheetName,
          targetRow: match.person.rowNumber,
          targetColumn: mapping.targetHeader.column,
          targetField: mapping.targetField,
          value,
          matchedBy: match.matchedBy,
          basis: mapping.basis,
        });
      }
    }

    const missingTarget = [...expectedRows].filter(
      (rowNumber) => !matchedRows.has(rowNumber),
    ).length;
    const matchedExpectedPeople = [...expectedRows].filter(
      (rowNumber) => matchedRows.has(rowNumber),
    ).length;
    if (duplicateKeys.size) {
      errors.push(`来源表有 ${duplicateKeys.size} 个重复人员标识`);
    }
    if (unmatchedSource) {
      errors.push(`来源表有 ${unmatchedSource} 人未匹配工资表`);
    }
    if (missingTarget) {
      errors.push(`工资表有 ${missingTarget} 人未在该来源表中出现`);
    }
    if (invalidValues) {
      errors.push(`来源表有 ${invalidValues} 个公式缓存或错误值不可用`);
    }
    if (!updates.length) {
      errors.push("来源表没有形成可写入的人员字段");
    }
    return {
      category,
      sourceName,
      sourceSheet: sourceTable.sheetName,
      sourcePeriod,
      mappings,
      updates,
      fieldSummaries: [...fieldSummaryMap.values()],
      matchedPeople:
        category === "劳务费附件" && !expectedRows.size
          ? matchedRows.size
          : matchedExpectedPeople,
      sourceMatchedPeople: matchedRows.size,
      matchedTargetRows: [...matchedRows],
      expectedPeople:
        category === "劳务费附件" && !expectedRows.size
          ? matchedRows.size
          : expectedRows.size,
      errors: [...new Set(errors)],
      basis: profile.basis,
      excludesFrom: [...(profile.excludesFrom || [])],
    };
  }

  function attachmentExclusionRows(results, category) {
    const rows = new Set();
    for (const result of results) {
      if (!result.excludesFrom?.includes(category)) {
        continue;
      }
      for (const rowNumber of result.matchedTargetRows || []) {
        rows.add(rowNumber);
      }
    }
    return [...rows];
  }

  Object.assign(api, {
    ATTACHMENT_PROFILES,
    profileForAttachment,
    requiredAttachments,
    attachmentTableScore,
    resolveAttachment,
    attachmentExclusionRows,
  });
})();
