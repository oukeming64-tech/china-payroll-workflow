(() => {
  "use strict";

  const api = window.PayrollLocal.excel;

  function externalFormulaRecords(sheetRecord, sharedStrings) {
    const formulaNodes = api.elementsByLocalName(
      sheetRecord.document,
      "f",
    );
    const shared = new Map();
    for (const formulaNode of formulaNodes) {
      const sharedId = formulaNode.getAttribute("si");
      const formula = formulaNode.textContent || "";
      if (sharedId && formula) {
        shared.set(sharedId, formula);
      }
    }
    return formulaNodes
      .map((formulaNode) => {
        const cell = formulaNode.parentElement;
        const sharedId = formulaNode.getAttribute("si");
        const formula =
          formulaNode.textContent || shared.get(sharedId) || "";
        return {
          formulaNode,
          formula,
          descriptor: api.cellDescriptor(cell, sharedStrings),
          address: cell?.getAttribute("r") || "",
        };
      })
      .filter((record) => api.isExternalFormula(record.formula));
  }

  async function materializeLegacyExternalFormulas(
    workbook,
    mainSheetName,
    allowedMainRows = [],
  ) {
    const allowedRows = new Set(
      [...allowedMainRows].map(Number).filter(Number.isFinite),
    );
    const bySheet = [];
    const failures = [];
    const recordsBySheet = [];
    for (const sheet of workbook.sheets) {
      const record = await workbook.loadSheetRecord(sheet.name);
      const external = externalFormulaRecords(
        record,
        workbook.sharedStrings,
      );
      if (!external.length) {
        continue;
      }
      const permitted =
        sheet.name === mainSheetName
          ? external.filter((item) => {
              const row = api.parseCellReference(item.address)?.row;
              return allowedRows.has(row);
            })
          : external;
      const blocked =
        sheet.name === mainSheetName
          ? external.length - permitted.length
          : 0;
      if (blocked) {
        failures.push(
          `主工资表仍有 ${blocked} 个外部公式未由本地附件解析`,
        );
      }
      const unusable = permitted.filter(
        (item) =>
          !item.descriptor.hasCachedValue ||
          item.descriptor.type === "e",
      );
      if (unusable.length) {
        failures.push(
          `${sheet.name} 有 ${unusable.length} 个旧外部公式缺少可用缓存值`,
        );
        continue;
      }
      if (permitted.length) {
        recordsBySheet.push({
          sheet,
          record,
          external: permitted,
          basis:
            sheet.name === mainSheetName
              ? "已确认停用人员的隐藏行保留完整缓存结果"
              : "旧辅助公式保留完整缓存结果并移除不可访问的来源关系",
        });
      }
    }
    if (failures.length) {
      throw new Error(failures.join("；"));
    }
    let frozen = 0;
    for (const { sheet, record, external, basis } of recordsBySheet) {
      for (const item of external) {
        item.formulaNode.remove();
        frozen += 1;
      }
      workbook.dirtySheetPaths.add(record.path);
      bySheet.push({
        sheet: sheet.name,
        formulaCells: external.length,
        basis,
      });
    }
    return { frozen, bySheet };
  }

  async function remainingExternalFormulaCount(workbook) {
    let count = 0;
    for (const sheet of workbook.sheets) {
      const record = await workbook.loadSheetRecord(sheet.name);
      count += externalFormulaRecords(
        record,
        workbook.sharedStrings,
      ).length;
    }
    return count;
  }

  async function removeExternalPackage(workbook) {
    for (const node of api.elementsByLocalName(
      workbook.workbookDocument,
      "externalReferences",
    )) {
      node.remove();
    }
    api.setCalcMode(workbook.workbookDocument);
    api.removeRelationshipsByType(
      workbook.workbookRelationshipsDocument,
      ["/externalLink", "/calcChain"],
    );
    const contentTypesXml = await workbook.zip
      .file("[Content_Types].xml")
      ?.async("string");
    if (!contentTypesXml) {
      throw new Error("工作簿缺少 Content Types，无法移除外部来源");
    }
    const contentTypes = api.parseXml(
      contentTypesXml,
      "[Content_Types].xml",
    );
    api.removeContentTypeOverrides(contentTypes, [
      (partName) => partName.startsWith("/xl/externalLinks/"),
      (partName) => partName === "/xl/calcChain.xml",
    ]);
    for (const filePath of Object.keys(workbook.zip.files)) {
      if (
        filePath.startsWith("xl/externalLinks/") ||
        filePath === "xl/calcChain.xml"
      ) {
        workbook.zip.remove(filePath);
      }
    }
    workbook.zip.file(
      "xl/_rels/workbook.xml.rels",
      api.serializeXml(workbook.workbookRelationshipsDocument),
    );
    workbook.zip.file(
      "[Content_Types].xml",
      api.serializeXml(contentTypes),
    );
    workbook.workbookRelationships = api.relationshipMap(
      workbook.workbookRelationshipsDocument,
      "xl/workbook.xml",
    );
    workbook.externalDiagnostics = null;
  }

  async function detachExternalLinks(
    workbook,
    mainSheetName,
    allowedMainRows = [],
  ) {
    const linkCount = (await api.analyzeExternalLinks(workbook)).length;
    const legacy = await materializeLegacyExternalFormulas(
      workbook,
      mainSheetName,
      allowedMainRows,
    );
    const remaining = await remainingExternalFormulaCount(workbook);
    if (remaining) {
      throw new Error(`仍有 ${remaining} 个外部公式，已停止移除外链`);
    }
    await removeExternalPackage(workbook);
    return {
      removedLinks: linkCount,
      frozenLegacyFormulaCount: legacy.frozen,
      frozenSheets: legacy.bySheet,
    };
  }

  Object.assign(api, {
    externalFormulaRecords,
    materializeLegacyExternalFormulas,
    remainingExternalFormulaCount,
    removeExternalPackage,
    detachExternalLinks,
  });
})();
