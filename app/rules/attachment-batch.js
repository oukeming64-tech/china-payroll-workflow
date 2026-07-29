(() => {
  "use strict";

  const api = window.PayrollLocal.rules;

  function uniqueMappings(parts) {
    const mappings = [];
    const seen = new Set();
    for (const mapping of parts.flatMap((part) => part.mappings || [])) {
      const key =
        `${mapping.sourceHeader?.column || 0}:` +
        `${mapping.targetHeader?.column || 0}:` +
        `${mapping.sourceField}:${mapping.targetField}`;
      if (!seen.has(key)) {
        seen.add(key);
        mappings.push(mapping);
      }
    }
    return mappings;
  }

  function mergeUpdates(category, parts, errors) {
    const updates = [];
    const byTarget = new Map();
    for (const update of parts.flatMap((part) => part.updates || [])) {
      const key =
        `${update.targetSheet}:${update.targetRow}:` +
        `${update.targetColumn}`;
      const prior = byTarget.get(key);
      if (!prior) {
        byTarget.set(key, update);
        updates.push(update);
        continue;
      }
      if (api.sameValue(prior.value, update.value)) {
        continue;
      }
      errors.push(
        `${category}的多个文件对同一人员同一字段给出了不同值`,
      );
    }
    return updates;
  }

  function fieldSummaries(updates) {
    const summaries = new Map();
    for (const update of updates) {
      if (!summaries.has(update.targetField)) {
        summaries.set(update.targetField, {
          sourceFields: new Set(),
          targetField: update.targetField,
          rows: new Set(),
          changed: 0,
          formulaValues: 0,
        });
      }
      const summary = summaries.get(update.targetField);
      summary.sourceFields.add(update.sourceField);
      summary.rows.add(update.targetRow);
      summary.changed += update.changed ? 1 : 0;
      summary.formulaValues += update.formulaValue ? 1 : 0;
    }
    return [...summaries.values()].map((summary) => ({
      sourceField: [...summary.sourceFields].join("、"),
      targetField: summary.targetField,
      matched: summary.rows.size,
      changed: summary.changed,
      formulaValues: summary.formulaValues,
    }));
  }

  function mergeCategoryResults(category, parts) {
    const errors = parts.flatMap((part) =>
      (part.errors || []).map((error) => `${part.sourceName}：${error}`),
    );
    const warnings = parts.flatMap((part) =>
      (part.warnings || []).map(
        (warning) => `${part.sourceName}：${warning}`,
      ),
    );
    const expectedRows = new Set(
      parts.flatMap((part) => part.expectedTargetRows || []),
    );
    const matchedRows = new Set(
      parts.flatMap((part) => part.matchedTargetRows || []),
    );
    if (category === "劳务费附件" && !expectedRows.size) {
      for (const rowNumber of matchedRows) {
        expectedRows.add(rowNumber);
      }
    }
    const missingTarget = [...expectedRows].filter(
      (rowNumber) => !matchedRows.has(rowNumber),
    ).length;
    if (missingTarget) {
      errors.push(`工资表有 ${missingTarget} 人未在该类附件中出现`);
    }
    const updates = mergeUpdates(category, parts, errors);
    if (!updates.length && !errors.length) {
      errors.push("该类附件没有形成可写入的人员字段");
    }
    const sourceNames = parts.map((part) => part.sourceName);
    const sourceSheets = [
      ...new Set(parts.map((part) => part.sourceSheet).filter(Boolean)),
    ];
    return {
      category,
      sourceName:
        sourceNames.length === 1
          ? sourceNames[0]
          : `${sourceNames.length} 个文件`,
      sourceNames,
      sourceSheet: sourceSheets.join("、"),
      sourcePeriod: parts[0]?.sourcePeriod || "",
      mappings: uniqueMappings(parts),
      updates,
      fieldSummaries: fieldSummaries(updates),
      matchedPeople: [...expectedRows].filter((rowNumber) =>
        matchedRows.has(rowNumber),
      ).length,
      sourceMatchedPeople: matchedRows.size,
      matchedTargetRows: [...matchedRows],
      expectedTargetRows: [...expectedRows],
      expectedPeople: expectedRows.size,
      errors: [...new Set(errors)],
      warnings: [...new Set(warnings)],
      basis: parts[0]?.basis || "",
      excludesFrom: [
        ...new Set(parts.flatMap((part) => part.excludesFrom || [])),
      ],
    };
  }

  function resolveAttachmentBatch(inputs, targetTable, targetPeriod) {
    const orderedInputs = [...inputs].sort((left, right) => {
      const leftProfile = api.profileForAttachment(left.category);
      const rightProfile = api.profileForAttachment(right.category);
      return (
        Number(Boolean(rightProfile?.excludesFrom?.length)) -
        Number(Boolean(leftProfile?.excludesFrom?.length))
      );
    });
    const parts = [];
    for (const input of orderedInputs) {
      const result = api.resolveAttachment(
        input.category,
        input.sourceTable,
        targetTable,
        targetPeriod,
        input.sourceName,
        {
          deferMissingTarget: true,
          excludedTargetRows: api.attachmentExclusionRows(
            parts,
            input.category,
          ),
        },
      );
      result.warnings = [
        ...(result.warnings || []),
        ...(input.warnings || []),
      ];
      parts.push(result);
    }
    const grouped = new Map();
    for (const part of parts) {
      if (!grouped.has(part.category)) {
        grouped.set(part.category, []);
      }
      grouped.get(part.category).push(part);
    }
    const results = [...grouped.entries()].map(([category, items]) =>
      mergeCategoryResults(category, items),
    );
    return {
      parts,
      results,
      updates: results.flatMap((result) => result.updates || []),
      errors: results.flatMap((result) =>
        (result.errors || []).map(
          (error) => `${result.category}：${error}`,
        ),
      ),
    };
  }

  Object.assign(api, {
    resolveAttachmentBatch,
  });
})();
