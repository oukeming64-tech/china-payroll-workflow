(() => {
  "use strict";

  const api = window.PayrollLocal.rules;

  function externalMappings(sourceTable, targetTable, diagnostic) {
    const mappings = [];
    const seen = new Set();
    for (const reference of diagnostic?.references || []) {
      if (
        reference.formulaKind !== "exact-vlookup" ||
        reference.transformed
      ) {
        continue;
      }
      if (
        reference.sourceSheet &&
        api.normalizeText(reference.sourceSheet) !==
          api.normalizeText(sourceTable.sheetName)
      ) {
        continue;
      }
      const sourceHeader = sourceTable.headerByColumn.get(
        reference.sourceValueColumn,
      );
      const target = api.targetHeader(targetTable, reference.targetField);
      if (!sourceHeader || !target) {
        continue;
      }
      const key = `${sourceHeader.column}:${target.column}`;
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      mappings.push({
        sourceHeader,
        targetHeader: target,
        confidence: "高",
        basis: "原工资表外链公式的来源列与目标列",
      });
    }
    for (const sourceHeader of sourceTable.headers) {
      const target = api.targetHeader(targetTable, sourceHeader.name);
      if (!target) {
        continue;
      }
      const key = `${sourceHeader.column}:${target.column}`;
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      mappings.push({
        sourceHeader,
        targetHeader: target,
        confidence: "中",
        basis: "字段名完全一致",
      });
    }
    return mappings;
  }

  function proposalsFromExternalSource(
    sourceTable,
    targetTable,
    diagnostic,
    targetPeriod,
    sourceName,
  ) {
    const mappings = externalMappings(sourceTable, targetTable, diagnostic);
    const targetIndex = api.indexPeople(targetTable);
    const sourceIdentities = api.identityHeaders(sourceTable);
    const proposals = [];
    const errors = [];
    if (!mappings.length) {
      return {
        mappings,
        proposals,
        errors: ["没有找到可证明的来源字段与目标字段对应关系"],
      };
    }
    if (
      !sourceIdentities.name &&
      !sourceIdentities.idCard &&
      !sourceIdentities.employeeId
    ) {
      return {
        mappings,
        proposals,
        errors: ["来源表缺少可识别的人员编号、身份证或姓名列"],
      };
    }

    for (const sourceRow of sourceTable.rows) {
      const identity = api.identityFromRow(sourceRow, sourceIdentities);
      const match = api.matchPerson(targetIndex, identity);
      for (const mapping of mappings) {
        const value = sourceRow.get(mapping.sourceHeader.name);
        if (
          value === null ||
          value === undefined ||
          api.asText(value) === ""
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
          inputValue: value,
          mapping,
        });
        if (match.status !== "matched") {
          proposal.status = "error";
          proposal.errors.push(match.message);
        } else {
          proposal.person = match.person;
          proposal.matchedBy = match.matchedBy;
          proposal.currentValue = match.person.row.get(
            mapping.targetHeader.name,
          );
        }
        const sourceCell = sourceRow.cells?.get(mapping.sourceHeader.column);
        if (sourceCell?.hasFormula) {
          proposal.warnings.push(
            "来源单元格含公式；默认只使用其缓存值，不复制来源公式",
          );
        }
        proposals.push(proposal);
      }
    }
    return { mappings, proposals, errors };
  }

  Object.assign(api, {
    externalMappings,
    proposalsFromExternalSource,
  });
})();
