(() => {
  "use strict";

  const api = window.PayrollLocal.excel;

  function formulaCell(formulaNode) {
    let current = formulaNode?.parentElement || null;
    while (current && current.localName !== "c") {
      current = current.parentElement;
    }
    return current;
  }

  function formulaSources(formula) {
    const text = String(formula || "");
    const sources = new Set();
    for (const match of text.matchAll(
      /(?:'([^']+)'|([^\s'!()+\-*/^,]+))!\$?[A-Z]{1,3}\$?\d+/gi,
    )) {
      const value = (match[1] || match[2] || "").replaceAll("''", "'");
      if (value) {
        sources.add(value);
      }
    }
    for (const match of text.matchAll(/\[(\d+)\]/g)) {
      sources.add(`外部工作簿 ${match[1]}`);
    }
    return [...sources];
  }

  async function analyzeFormulas(workbook, sheetName) {
    const sheetRecord = await workbook.loadSheetRecord(sheetName);
    const table = api.buildTableFromSheet(
      sheetRecord,
      workbook.sharedStrings,
    );
    const formulaNodes = api.elementsByLocalName(
      sheetRecord.document,
      "f",
    );
    const sharedFormulas = new Map();
    for (const formulaNode of formulaNodes) {
      const sharedId = formulaNode.getAttribute("si");
      const text = formulaNode.textContent || "";
      if (sharedId && text) {
        sharedFormulas.set(sharedId, text);
      }
    }

    const groups = new Map();
    let externalFormulaNodes = 0;
    let unresolvedFormulaNodes = 0;
    for (const formulaNode of formulaNodes) {
      const cell = formulaCell(formulaNode);
      const address = cell?.getAttribute("r") || "";
      const parsed = api.parseCellReference(address);
      const header = parsed
        ? table.headerByColumn.get(parsed.column)
        : null;
      const field =
        header?.displayName ||
        (parsed ? `${api.columnNumberToLetters(parsed.column)}列` : "未知列");
      const sharedId = formulaNode.getAttribute("si");
      const formula =
        formulaNode.textContent || sharedFormulas.get(sharedId) || "";
      const external = api.isExternalFormula(formula);
      const unresolved = !formula && Boolean(sharedId);
      externalFormulaNodes += external ? 1 : 0;
      unresolvedFormulaNodes += unresolved ? 1 : 0;

      if (!groups.has(field)) {
        groups.set(field, {
          field,
          formulaNodes: 0,
          internalFormulaNodes: 0,
          externalFormulaNodes: 0,
          unresolvedFormulaNodes: 0,
          sourceSheets: new Set(),
          sampleCells: [],
        });
      }
      const group = groups.get(field);
      group.formulaNodes += 1;
      group.externalFormulaNodes += external ? 1 : 0;
      group.internalFormulaNodes += external ? 0 : 1;
      group.unresolvedFormulaNodes += unresolved ? 1 : 0;
      for (const source of formulaSources(formula)) {
        group.sourceSheets.add(source);
      }
      if (address && group.sampleCells.length < 6) {
        group.sampleCells.push(address);
      }
    }

    return {
      sheetName,
      totalFormulaNodes: formulaNodes.length,
      internalFormulaNodes: formulaNodes.length - externalFormulaNodes,
      externalFormulaNodes,
      unresolvedFormulaNodes,
      fields: [...groups.values()]
        .map((group) => ({
          ...group,
          sourceSheets: [...group.sourceSheets].slice(0, 8),
        }))
        .sort(
          (left, right) =>
            right.formulaNodes - left.formulaNodes ||
            left.field.localeCompare(right.field, "zh-CN"),
        ),
    };
  }

  Object.assign(api, {
    formulaSources,
    analyzeFormulas,
  });
})();
