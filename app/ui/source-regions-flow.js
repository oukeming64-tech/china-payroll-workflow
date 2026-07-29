(() => {
  "use strict";

  const ui = window.PayrollLocal.ui;
  const rules = window.PayrollEngine;

  async function sourceTables(workbook, diagnostic, targetTable) {
    const referencedSheetNames = [
      ...new Set(
        (diagnostic?.references || [])
          .map((reference) => reference.sourceSheet)
          .filter(Boolean),
      ),
    ];
    let sheetName = referencedSheetNames.find((name) =>
      workbook.sheetByName.has(name),
    );
    if (!sheetName) {
      const scores = await workbook.scoreSheetsAgainst(
        targetTable.headers.map((header) => header.name),
      );
      sheetName = scores[0]?.name || workbook.sheets[0]?.name;
    }
    if (!sheetName) {
      return [];
    }
    return workbook.getTables
      ? workbook.getTables(sheetName)
      : [await workbook.getTable(sheetName)];
  }

  function combineBusinessRegions(parts, sheetName) {
    const mappings = parts.flatMap((part) => part.mappings || []);
    const proposals = parts.flatMap((part) => part.proposals || []);
    const errors = parts.flatMap((part) => part.errors || []);
    const warnings = parts.flatMap((part) => part.warnings || []);
    const profiles = parts.map((part) => part.profile).filter(Boolean);
    return {
      format: "business-regions",
      sheetName,
      mappings,
      proposals,
      errors: [...new Set(errors)],
      warnings: [...new Set(warnings)],
      profileIds: profiles.map((profile) => profile.id),
      labels: [...new Set(profiles.map((profile) => profile.label))],
      regions: parts.map((part) => ({
        headerRow: part.sourceHeaderRow,
        profileId: part.profile?.id || "",
        proposals: part.proposals?.length || 0,
      })),
    };
  }

  async function inspectWorkbookBusinessRegions(
    workbook,
    diagnostic,
    targetTable,
    targetPeriod,
    fileName,
    role,
  ) {
    const tables = await sourceTables(
      workbook,
      diagnostic,
      targetTable,
    );
    const matched = tables
      .map((table) => ({
        table,
        profile: rules.matchBusinessSource(table, fileName),
      }))
      .filter((item) => item.profile);
    if (!matched.length) {
      return { table: tables[0] || null, result: null };
    }
    const parts = matched.map(({ table, profile }) => {
      const result = rules.proposalsFromBusinessSource(
        table,
        targetTable,
        targetPeriod,
        role,
        profile,
      );
      result.sourceHeaderRow = table.headerRow;
      return result;
    });
    return {
      table: matched[0].table,
      result: combineBusinessRegions(
        parts,
        matched[0].table.sheetName,
      ),
    };
  }

  Object.assign(ui, {
    inspectWorkbookBusinessRegions,
  });
})();
