(() => {
  "use strict";

  const api = window.PayrollLocal.excel;

  function setFormulaCache(cell, value) {
    for (const inline of api.directChildrenByLocalName(cell, "is")) {
      inline.remove();
    }
    let valueNode = api.childByLocalName(cell, "v");
    if (value === null || value === undefined || value === "") {
      valueNode?.remove();
      if (cell.getAttribute("t") === "str") {
        cell.removeAttribute("t");
      }
      return;
    }
    if (!valueNode) {
      valueNode = cell.ownerDocument.createElementNS(api.MAIN_NS, "v");
      cell.appendChild(valueNode);
    }
    if (typeof value === "number" && Number.isFinite(value)) {
      cell.removeAttribute("t");
      valueNode.textContent = String(value);
    } else if (typeof value === "boolean") {
      cell.setAttribute("t", "b");
      valueNode.textContent = value ? "1" : "0";
    } else {
      cell.setAttribute("t", "str");
      valueNode.textContent = String(value);
    }
  }

  function setDirectCellValue(cell, value) {
    for (const child of [...cell.childNodes]) {
      if (
        child.nodeType === Node.ELEMENT_NODE &&
        ["f", "v", "is"].includes(child.localName)
      ) {
        child.remove();
      }
    }
    if (value === null || value === undefined || value === "") {
      cell.removeAttribute("t");
      return;
    }
    if (typeof value === "number" && Number.isFinite(value)) {
      cell.removeAttribute("t");
      const valueNode = cell.ownerDocument.createElementNS(api.MAIN_NS, "v");
      valueNode.textContent = String(value);
      cell.appendChild(valueNode);
      return;
    }
    if (typeof value === "boolean") {
      cell.setAttribute("t", "b");
      const valueNode = cell.ownerDocument.createElementNS(api.MAIN_NS, "v");
      valueNode.textContent = value ? "1" : "0";
      cell.appendChild(valueNode);
      return;
    }
    cell.setAttribute("t", "inlineStr");
    const inline = cell.ownerDocument.createElementNS(api.MAIN_NS, "is");
    const text = cell.ownerDocument.createElementNS(api.MAIN_NS, "t");
    const stringValue = String(value);
    if (/^\s|\s$/.test(stringValue)) {
      text.setAttribute("xml:space", "preserve");
    }
    text.textContent = stringValue;
    inline.appendChild(text);
    cell.appendChild(inline);
  }

  function setCellFormula(cell, formula, cachedValue = null) {
    for (const child of [...cell.childNodes]) {
      if (
        child.nodeType === Node.ELEMENT_NODE &&
        ["f", "v", "is"].includes(child.localName)
      ) {
        child.remove();
      }
    }
    cell.removeAttribute("t");
    const formulaNode = cell.ownerDocument.createElementNS(
      api.MAIN_NS,
      "f",
    );
    formulaNode.textContent = String(formula || "").replace(/^=/, "");
    cell.appendChild(formulaNode);
    setFormulaCache(cell, cachedValue);
  }

  function getOrCreateCell(documentNode, rowNumber, columnNumber) {
    const reference = `${api.columnNumberToLetters(columnNumber)}${rowNumber}`;
    const sheetData = api.firstByLocalName(documentNode, "sheetData");
    if (!sheetData) {
      throw new Error("工作表缺少 sheetData，无法安全写入");
    }
    let row = api
      .directChildrenByLocalName(sheetData, "row")
      .find(
        (candidate) => Number(candidate.getAttribute("r")) === Number(rowNumber),
      );
    if (!row) {
      row = documentNode.createElementNS(api.MAIN_NS, "row");
      row.setAttribute("r", String(rowNumber));
      const following = api
        .directChildrenByLocalName(sheetData, "row")
        .find(
          (candidate) =>
            Number(candidate.getAttribute("r")) > Number(rowNumber),
        );
      sheetData.insertBefore(row, following || null);
    }
    let cell = api
      .directChildrenByLocalName(row, "c")
      .find((candidate) => candidate.getAttribute("r") === reference);
    if (!cell) {
      cell = documentNode.createElementNS(api.MAIN_NS, "c");
      cell.setAttribute("r", reference);
      const following = api
        .directChildrenByLocalName(row, "c")
        .find((candidate) => {
          const parsed = api.parseCellReference(candidate.getAttribute("r"));
          return parsed && parsed.column > columnNumber;
        });
      row.insertBefore(cell, following || null);
    }
    return cell;
  }

  function setCalcMode(documentNode) {
    const workbook = documentNode.documentElement;
    let calcPr = api.firstByLocalName(documentNode, "calcPr");
    if (!calcPr) {
      calcPr = documentNode.createElementNS(api.MAIN_NS, "calcPr");
      workbook.appendChild(calcPr);
    }
    calcPr.setAttribute("calcMode", "auto");
    calcPr.setAttribute("fullCalcOnLoad", "1");
    calcPr.setAttribute("forceFullCalc", "1");
  }

  async function getCell(workbook, sheetName, rowNumber, columnNumber) {
    const sheetRecord = await workbook.loadSheetRecord(sheetName);
    const cellsByRow = api.rowCellMap(sheetRecord.document);
    return api.cellDescriptor(
      cellsByRow.get(Number(rowNumber))?.get(Number(columnNumber)),
      workbook.sharedStrings,
    );
  }

  async function updateCell(
    workbook,
    sheetName,
    rowNumber,
    columnNumber,
    value,
    options = {},
  ) {
    const sheetRecord = await workbook.loadSheetRecord(sheetName);
    const cell = getOrCreateCell(
      sheetRecord.document,
      Number(rowNumber),
      Number(columnNumber),
    );
    const before = api.cellDescriptor(cell, workbook.sharedStrings);
    const preserveFormula =
      options.preserveFormula !== false && before.hasFormula;
    if (preserveFormula) {
      setFormulaCache(cell, value);
    } else {
      setDirectCellValue(cell, value);
    }
    workbook.dirtySheetPaths.add(sheetRecord.path);
    return {
      before,
      after: api.cellDescriptor(cell, workbook.sharedStrings),
      preservedFormula: preserveFormula,
    };
  }

  async function updateFormulaCell(
    workbook,
    sheetName,
    rowNumber,
    columnNumber,
    formula,
    cachedValue,
  ) {
    const sheetRecord = await workbook.loadSheetRecord(sheetName);
    const cell = getOrCreateCell(
      sheetRecord.document,
      Number(rowNumber),
      Number(columnNumber),
    );
    const before = api.cellDescriptor(cell, workbook.sharedStrings);
    setCellFormula(cell, formula, cachedValue);
    setCalcMode(workbook.workbookDocument);
    workbook.dirtySheetPaths.add(sheetRecord.path);
    return {
      before,
      after: api.cellDescriptor(cell, workbook.sharedStrings),
      wroteFormula: true,
    };
  }

  async function findReservedBlankRows(workbook, sheetName, table = null) {
    const resolvedTable = table || (await workbook.getTable(sheetName));
    const identityColumns = resolvedTable.headers
      .filter((header) =>
        api.IDENTITY_HEADER_KEYS.has(api.normalizeText(header.name)),
      )
      .map((header) => header.column);
    const identityRows = resolvedTable.rows
      .filter((row) =>
        identityColumns.some((column) => {
          const value = row.values.get(column);
          return value !== null && value !== undefined && String(value).trim();
        }),
      )
      .map((row) => row.rowNumber);
    if (!identityRows.length) {
      return [];
    }
    const lastIdentityRow = Math.max(...identityRows);
    const sheetRecord = await workbook.loadSheetRecord(sheetName);
    const cellsByRow = api.rowCellMap(sheetRecord.document);
    const reserved = [];
    for (
      let rowNumber = lastIdentityRow + 1;
      rowNumber <= lastIdentityRow + 12;
      rowNumber += 1
    ) {
      const cells = cellsByRow.get(rowNumber) || new Map();
      const hasValueOrFormula = [...cells.values()]
        .map((cell) => api.cellDescriptor(cell, workbook.sharedStrings))
        .some(
          (descriptor) =>
            descriptor.hasFormula ||
            (descriptor.value !== null &&
              descriptor.value !== undefined &&
              String(descriptor.value).trim() !== ""),
        );
      if (hasValueOrFormula) {
        break;
      }
      reserved.push(rowNumber);
    }
    return reserved;
  }

  async function cloneEmployeeRow(
    workbook,
    sheetName,
    sourceRowNumber,
    targetRowNumber,
    options = {},
  ) {
    const sheetRecord = await workbook.loadSheetRecord(sheetName);
    const sheetData = api.firstByLocalName(sheetRecord.document, "sheetData");
    if (!sheetData) {
      throw new Error("工作表缺少 sheetData，无法复制人员模板行");
    }
    const rows = api.directChildrenByLocalName(sheetData, "row");
    const sourceRow = rows.find(
      (row) => Number(row.getAttribute("r")) === Number(sourceRowNumber),
    );
    if (!sourceRow) {
      throw new Error("找不到可复制的人员模板行");
    }
    const existingTarget = rows.find(
      (row) => Number(row.getAttribute("r")) === Number(targetRowNumber),
    );
    if (existingTarget) {
      const replaceHiddenTarget =
        options.replaceHiddenTarget &&
        existingTarget.getAttribute("hidden") === "1";
      const hasContent = api
        .directChildrenByLocalName(existingTarget, "c")
        .some((cell) => {
          const descriptor = api.cellDescriptor(
            cell,
            workbook.sharedStrings,
          );
          return (
            descriptor.hasFormula ||
            (descriptor.value !== null &&
              descriptor.value !== undefined &&
              String(descriptor.value).trim() !== "")
          );
        });
      if (hasContent && !replaceHiddenTarget) {
        throw new Error("目标行已有内容，不能作为新增人员模板行");
      }
    }

    const rowDelta = Number(targetRowNumber) - Number(sourceRowNumber);
    const clonedRow = sourceRow.cloneNode(true);
    clonedRow.setAttribute("r", String(targetRowNumber));
    clonedRow.removeAttribute("hidden");
    for (const cell of api.directChildrenByLocalName(clonedRow, "c")) {
      const parsed = api.parseCellReference(cell.getAttribute("r"));
      if (!parsed) {
        continue;
      }
      cell.setAttribute(
        "r",
        `${api.columnNumberToLetters(parsed.column)}${targetRowNumber}`,
      );
      const formulaNode = api.childByLocalName(cell, "f");
      if (formulaNode) {
        formulaNode.textContent = api.translateFormulaA1(
          formulaNode.textContent || "",
          rowDelta,
          0,
        );
        formulaNode.removeAttribute("t");
        formulaNode.removeAttribute("si");
        formulaNode.removeAttribute("ref");
        api.childByLocalName(cell, "v")?.remove();
        continue;
      }
      api.childByLocalName(cell, "v")?.remove();
      api.childByLocalName(cell, "is")?.remove();
      cell.removeAttribute("t");
    }
    if (existingTarget) {
      existingTarget.parentNode.replaceChild(clonedRow, existingTarget);
    } else {
      const following = rows.find(
        (row) => Number(row.getAttribute("r")) > Number(targetRowNumber),
      );
      sheetData.insertBefore(clonedRow, following || null);
    }
    workbook.dirtySheetPaths.add(sheetRecord.path);
    return {
      sourceRow: Number(sourceRowNumber),
      targetRow: Number(targetRowNumber),
    };
  }

  async function setRowHidden(workbook, sheetName, rowNumber, hidden = true) {
    const sheetRecord = await workbook.loadSheetRecord(sheetName);
    const sheetData = api.firstByLocalName(sheetRecord.document, "sheetData");
    const row = api
      .directChildrenByLocalName(sheetData, "row")
      .find(
        (candidate) => Number(candidate.getAttribute("r")) === Number(rowNumber),
      );
    if (!row) {
      throw new Error("找不到需要停用的人员行");
    }
    if (hidden) {
      row.setAttribute("hidden", "1");
    } else {
      row.removeAttribute("hidden");
    }
    workbook.dirtySheetPaths.add(sheetRecord.path);
  }

  Object.assign(api, {
    setFormulaCache,
    setDirectCellValue,
    setCellFormula,
    getOrCreateCell,
    setCalcMode,
    getCell,
    updateCell,
    updateFormulaCell,
    findReservedBlankRows,
    cloneEmployeeRow,
    setRowHidden,
  });
})();
