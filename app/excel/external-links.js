(() => {
  "use strict";

  const api = window.PayrollLocal.excel;

  function safeDecodeUri(value) {
    try {
      return decodeURIComponent(String(value || ""));
    } catch {
      return String(value || "");
    }
  }

  function workbookFileNameFromUri(uri) {
    const decoded = safeDecodeUri(uri);
    return decoded.split(/[\\/]/).at(-1) || decoded;
  }

  function classifyExternalSource(uri) {
    const filename = workbookFileNameFromUri(uri);
    const normalized = api.normalizeText(filename);
    if (normalized.includes("正常工资薪金所得")) {
      return {
        category: "个税工资薪金附件",
        confidence: "高",
        basis: "文件名含“正常工资薪金所得”",
      };
    }
    if (
      normalized.includes("保险表") ||
      normalized.includes("社保") ||
      normalized.includes("公积金")
    ) {
      return {
        category: "社保 / 公积金附件",
        confidence: "高",
        basis: "文件名含保险、社保或公积金标识",
      };
    }
    if (normalized.includes("劳务费")) {
      return {
        category: "劳务费附件",
        confidence: "高",
        basis: "文件名含“劳务费”",
      };
    }
    if (normalized.includes("工资")) {
      return {
        category: "历史工资表",
        confidence: "中",
        basis: "文件名含“工资”，且不是已识别附件类型",
      };
    }
    return {
      category: "未分类外部工作簿",
      confidence: "低",
      basis: "仅能确认它是外部工作簿",
    };
  }

  function externalLinkRelationshipsPath(linkPath) {
    const segments = linkPath.split("/");
    const filename = segments.pop();
    return [...segments, "_rels", `${filename}.rels`].join("/");
  }

  function cachedExternalStructure(documentNode) {
    if (!documentNode) {
      return {
        sheetNames: [],
        cachedCells: 0,
        cachedRows: 0,
        headerFields: [],
      };
    }
    const sheetNames = api
      .elementsByLocalName(documentNode, "sheetName")
      .map((node) => node.getAttribute("val") || "")
      .filter(Boolean);
    const rows = api.elementsByLocalName(documentNode, "row");
    const headerFields = [];
    const firstRow = rows.find((row) => row.getAttribute("r") === "1") || rows[0];
    if (firstRow) {
      for (const cell of api.directChildrenByLocalName(firstRow, "cell")) {
        const value = api.childByLocalName(cell, "v")?.textContent || "";
        if (value) {
          headerFields.push(value);
        }
      }
    }
    return {
      sheetNames,
      cachedCells: api.elementsByLocalName(documentNode, "cell").length,
      cachedRows: rows.length,
      headerFields: [...new Set(headerFields)].slice(0, 24),
    };
  }

  function isExternalFormula(formula) {
    return /\[[^\]]+\]/.test(String(formula || ""));
  }

  function parseExternalVlookup(formula) {
    const text = String(formula || "").trim().replace(/^=/, "");
    const match = text.match(
      /VLOOKUP\(\s*(\$?[A-Z]{1,3}\$?\d+)\s*,\s*'?(\[(\d+)\])?([^'!,]+)'?!\s*\$?([A-Z]{1,3})(?:\$?\d+)?\s*:\s*\$?([A-Z]{1,3})(?:\$?\d+)?\s*,\s*(\d+)\s*,\s*(?:0|FALSE)\s*\)/i,
    );
    if (!match) {
      return null;
    }
    const startColumn = api.columnLettersToNumber(match[5]);
    const endColumn = api.columnLettersToNumber(match[6]);
    const returnIndex = Number(match[7]);
    const valueColumn = startColumn + returnIndex - 1;
    const lookup = api.parseCellReference(match[1].replaceAll("$", ""));
    if (
      !lookup ||
      !Number.isFinite(startColumn) ||
      !Number.isFinite(endColumn) ||
      !Number.isInteger(returnIndex) ||
      returnIndex < 1 ||
      valueColumn > endColumn
    ) {
      return null;
    }
    return {
      kind: "exact-vlookup",
      externalIndex: Number(match[3]),
      sourceSheet: String(match[4] || "").trim(),
      lookupCell: match[1].replaceAll("$", ""),
      lookupColumn: lookup.column,
      sourceStartColumn: startColumn,
      sourceEndColumn: endColumn,
      sourceValueColumn: valueColumn,
      returnIndex,
      transformed: text.replace(/\s+/g, "") !== match[0].replace(/\s+/g, ""),
    };
  }

  async function analyzeExternalLinks(workbook) {
    if (workbook.externalDiagnostics) {
      return workbook.externalDiagnostics;
    }
    const externalReferences = api.elementsByLocalName(
      workbook.workbookDocument,
      "externalReference",
    );
    const diagnostics = [];
    for (let index = 0; index < externalReferences.length; index += 1) {
      const reference = externalReferences[index];
      const relationshipId =
        reference.getAttributeNS(api.REL_NS, "id") ||
        reference.getAttribute("r:id") ||
        "";
      const workbookRelationship =
        workbook.workbookRelationships.get(relationshipId);
      const linkPath = workbookRelationship?.resolvedTarget || "";
      const linkXml = linkPath
        ? await workbook.zip.file(linkPath)?.async("string")
        : "";
      const linkDocument = linkXml ? api.parseXml(linkXml, linkPath) : null;
      const relsPath = externalLinkRelationshipsPath(linkPath);
      const relsXml = relsPath
        ? await workbook.zip.file(relsPath)?.async("string")
        : "";
      const linkRelationships = relsXml
        ? api.relationshipMap(api.parseXml(relsXml, relsPath), linkPath)
        : new Map();
      const externalPathRelationship = [...linkRelationships.values()].find(
        (relationship) => relationship.type.endsWith("/externalLinkPath"),
      );
      const uri = externalPathRelationship?.target || "";
      const classification = classifyExternalSource(uri);
      diagnostics.push({
        index: index + 1,
        relationshipId,
        linkPath,
        uri: safeDecodeUri(uri),
        rawUri: uri,
        relationshipsPath: relsPath,
        filename: workbookFileNameFromUri(uri),
        ...classification,
        macStatus: /^file:\/\/\/?[A-Za-z]:/i.test(uri)
          ? "Windows 盘符路径，当前 Mac 不可直接访问"
          : "需由用户确认本机是否可访问",
        references: [],
        targetFields: [],
        targetSheets: [],
        ...cachedExternalStructure(linkDocument),
      });
    }

    for (const sheet of workbook.sheets) {
      const sheetRecord = await workbook.loadSheetRecord(sheet.name);
      const table = api.buildTableFromSheet(
        sheetRecord,
        workbook.sharedStrings,
      );
      for (const formulaNode of api.elementsByLocalName(
        sheetRecord.document,
        "f",
      )) {
        const formula = formulaNode.textContent || "";
        const indices = [...formula.matchAll(/\[(\d+)\]/g)].map((match) =>
          Number(match[1]),
        );
        if (!indices.length) {
          continue;
        }
        const cell = formulaNode.parentElement;
        const address = cell?.getAttribute("r") || "";
        const parsed = api.parseCellReference(address);
        const targetField =
          parsed && table.headerByColumn.get(parsed.column)
            ? table.headerByColumn.get(parsed.column).displayName
            : parsed
              ? `${api.columnNumberToLetters(parsed.column)}列`
              : "未知列";
        const lookup = parseExternalVlookup(formula);
        for (const linkIndex of new Set(indices)) {
          const diagnostic = diagnostics[linkIndex - 1];
          if (!diagnostic) {
            continue;
          }
          diagnostic.references.push({
            sheet: sheet.name,
            cell: address,
            targetField,
            sourceSheet: lookup?.sourceSheet || "",
            sourceCell: lookup
              ? `${api.columnNumberToLetters(lookup.sourceValueColumn)}1`
              : "",
            lookupCell: lookup?.lookupCell || "",
            lookupColumn: lookup?.lookupColumn || null,
            sourceStartColumn: lookup?.sourceStartColumn || null,
            sourceEndColumn: lookup?.sourceEndColumn || null,
            sourceValueColumn: lookup?.sourceValueColumn || null,
            returnIndex: lookup?.returnIndex || null,
            formulaKind: lookup?.kind || "unsupported",
            transformed: lookup?.transformed ?? true,
          });
        }
      }
    }
    for (const diagnostic of diagnostics) {
      diagnostic.targetFields = [
        ...new Set(diagnostic.references.map((reference) => reference.targetField)),
      ];
      diagnostic.targetSheets = [
        ...new Set(diagnostic.references.map((reference) => reference.sheet)),
      ];
      diagnostic.formulaReferenceCount = diagnostic.references.length;
    }
    workbook.externalDiagnostics = diagnostics;
    return diagnostics;
  }

  function removeRelationshipsByType(documentNode, suffixes) {
    for (const relationship of api.elementsByLocalName(
      documentNode,
      "Relationship",
    )) {
      const type = relationship.getAttribute("Type") || "";
      if (suffixes.some((suffix) => type.endsWith(suffix))) {
        relationship.remove();
      }
    }
  }

  function removeContentTypeOverrides(documentNode, predicates) {
    for (const override of api.elementsByLocalName(documentNode, "Override")) {
      const partName = override.getAttribute("PartName") || "";
      if (predicates.some((predicate) => predicate(partName))) {
        override.remove();
      }
    }
  }

  Object.assign(api, {
    safeDecodeUri,
    workbookFileNameFromUri,
    classifyExternalSource,
    isExternalFormula,
    parseExternalVlookup,
    analyzeExternalLinks,
    removeRelationshipsByType,
    removeContentTypeOverrides,
  });
})();
