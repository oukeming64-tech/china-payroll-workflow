(() => {
  "use strict";
  const api = window.PayrollLocal.excel;

  function shiftedColumnLetters(letters, insertColumn, amount) {
    const column = api.columnLettersToNumber(letters);
    return api.columnNumberToLetters(
      column >= insertColumn ? column + amount : column,
    );
  }

  function shiftRangeReference(reference, insertColumn, amount) {
    return String(reference || "")
      .split(/\s+/)
      .map((part) =>
        part.replace(
          /(\$?)([A-Z]{1,3})(\$?\d+)?(?::(\$?)([A-Z]{1,3})(\$?\d+)?)?/gi,
          (
            match,
            firstDollar,
            firstLetters,
            firstRow,
            secondDollar,
            secondLetters,
            secondRow,
          ) => {
            if (!firstRow && !secondLetters) {
              return match;
            }
            const first = `${firstDollar}${shiftedColumnLetters(
              firstLetters,
              insertColumn,
              amount,
            )}${firstRow || ""}`;
            if (!secondLetters) {
              return first;
            }
            return `${first}:${secondDollar}${shiftedColumnLetters(
              secondLetters,
              insertColumn,
              amount,
            )}${secondRow || ""}`;
          },
        ),
      )
      .join(" ");
  }

  function formulaPrefixTargetsSheet(prefix, targetSheetName) {
    if (!prefix) {
      return false;
    }
    const value = String(prefix)
      .slice(0, -1)
      .replace(/^'/, "")
      .replace(/'$/, "")
      .replaceAll("''", "'");
    return !value.includes("[") && value === targetSheetName;
  }

  function shiftFormulaColumns(
    formula,
    insertColumn,
    amount,
    contextSheetName,
    targetSheetName,
  ) {
    const allowUnqualified = contextSheetName === targetSheetName;
    const cellPattern =
      /((?:'[^']+'|(?:\[\d+\])?[\w.\u3400-\u9fff]+)!)?(\$?)([A-Z]{1,3})(\$?\d+)(?::(\$?)([A-Z]{1,3})(\$?\d+))?/gi;
    const shiftedCells = String(formula || "").replace(
      cellPattern,
      (
        match,
        prefix,
        firstDollar,
        firstLetters,
        firstRow,
        secondDollar,
        secondLetters,
        secondRow,
      ) => {
        const shouldShift =
          (!prefix && allowUnqualified) ||
          formulaPrefixTargetsSheet(prefix, targetSheetName);
        if (!shouldShift) {
          return match;
        }
        const first = `${firstDollar}${shiftedColumnLetters(
          firstLetters,
          insertColumn,
          amount,
        )}${firstRow}`;
        if (!secondLetters) {
          return `${prefix || ""}${first}`;
        }
        return `${prefix || ""}${first}:${secondDollar}${shiftedColumnLetters(
          secondLetters,
          insertColumn,
          amount,
        )}${secondRow}`;
      },
    );
    const columnPattern =
      /((?:'[^']+'|(?:\[\d+\])?[\w.\u3400-\u9fff]+)!)?(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})/gi;
    return shiftedCells.replace(
      columnPattern,
      (
        match,
        prefix,
        firstDollar,
        firstLetters,
        secondDollar,
        secondLetters,
      ) => {
        const shouldShift =
          (!prefix && allowUnqualified) ||
          formulaPrefixTargetsSheet(prefix, targetSheetName);
        if (!shouldShift) {
          return match;
        }
        return `${prefix || ""}${firstDollar}${shiftedColumnLetters(
          firstLetters,
          insertColumn,
          amount,
        )}:${secondDollar}${shiftedColumnLetters(
          secondLetters,
          insertColumn,
          amount,
        )}`;
      },
    );
  }

  function shiftWorksheetAttributes(documentNode, insertColumn, amount) {
    const referenceElements = new Set([
      "dimension",
      "mergeCell",
      "autoFilter",
      "hyperlink",
      "formula1",
      "formula2",
    ]);
    for (const element of api.elementsByLocalName(documentNode, "*")) {
      if (referenceElements.has(element.localName) && element.hasAttribute("ref")) {
        element.setAttribute(
          "ref",
          shiftRangeReference(
            element.getAttribute("ref"),
            insertColumn,
            amount,
          ),
        );
      }
      if (element.hasAttribute("sqref")) {
        element.setAttribute(
          "sqref",
          shiftRangeReference(
            element.getAttribute("sqref"),
            insertColumn,
            amount,
          ),
        );
      }
      if (element.localName === "selection" && element.hasAttribute("activeCell")) {
        element.setAttribute(
          "activeCell",
          shiftRangeReference(
            element.getAttribute("activeCell"),
            insertColumn,
            amount,
          ),
        );
      }
      if (element.localName === "f" && element.hasAttribute("ref")) {
        element.setAttribute(
          "ref",
          shiftRangeReference(
            element.getAttribute("ref"),
            insertColumn,
            amount,
          ),
        );
      }
    }
  }

  function shiftColumnDefinitions(documentNode, insertColumn, amount) {
    for (const column of api.elementsByLocalName(documentNode, "col")) {
      const minimum = Number(column.getAttribute("min"));
      const maximum = Number(column.getAttribute("max"));
      if (!Number.isFinite(minimum) || !Number.isFinite(maximum)) {
        continue;
      }
      if (minimum >= insertColumn) {
        column.setAttribute("min", String(minimum + amount));
        column.setAttribute("max", String(maximum + amount));
      } else if (maximum >= insertColumn) {
        column.setAttribute("max", String(maximum + amount));
      }
    }
  }

  function shiftRowSpans(documentNode, insertColumn, amount) {
    for (const row of api.elementsByLocalName(documentNode, "row")) {
      const spans = row.getAttribute("spans");
      const match = spans?.match(/^(\d+):(\d+)$/);
      if (!match) {
        continue;
      }
      const first = Number(match[1]);
      const last = Number(match[2]);
      row.setAttribute(
        "spans",
        `${first >= insertColumn ? first + amount : first}:${
          last >= insertColumn ? last + amount : last
        }`,
      );
    }
  }

  function copyCellStyle(sourceCell, targetCell) {
    const style = sourceCell?.getAttribute("s");
    if (style) {
      targetCell.setAttribute("s", style);
    } else {
      targetCell.removeAttribute("s");
    }
  }

  function setFormula(cell, formula) {
    for (const child of [...cell.childNodes]) {
      if (
        child.nodeType === Node.ELEMENT_NODE &&
        ["f", "v", "is"].includes(child.localName)
      ) {
        child.remove();
      }
    }
    cell.removeAttribute("t");
    const formulaNode = cell.ownerDocument.createElementNS(api.MAIN_NS, "f");
    formulaNode.textContent = String(formula || "");
    cell.appendChild(formulaNode);
  }

  async function insertColumns(
    workbook,
    sheetName,
    insertColumn,
    amount,
  ) {
    const targetRecord = await workbook.loadSheetRecord(sheetName);
    for (const sheet of workbook.sheets) {
      const record = await workbook.loadSheetRecord(sheet.name);
      for (const formulaNode of api.elementsByLocalName(record.document, "f")) {
        const before = formulaNode.textContent || "";
        const after = shiftFormulaColumns(
          before,
          insertColumn,
          amount,
          sheet.name,
          sheetName,
        );
        if (before !== after) {
          formulaNode.textContent = after;
          workbook.dirtySheetPaths.add(record.path);
        }
      }
    }

    const cells = api
      .elementsByLocalName(targetRecord.document, "c")
      .map((cell) => ({
        cell,
        parsed: api.parseCellReference(cell.getAttribute("r")),
      }))
      .filter((item) => item.parsed && item.parsed.column >= insertColumn)
      .sort((left, right) => right.parsed.column - left.parsed.column);
    for (const item of cells) {
      item.cell.setAttribute(
        "r",
        `${api.columnNumberToLetters(item.parsed.column + amount)}${item.parsed.row}`,
      );
    }
    shiftWorksheetAttributes(
      targetRecord.document,
      insertColumn,
      amount,
    );
    shiftColumnDefinitions(targetRecord.document, insertColumn, amount);
    shiftRowSpans(targetRecord.document, insertColumn, amount);

    for (const definedName of api.elementsByLocalName(
      workbook.workbookDocument,
      "definedName",
    )) {
      const before = definedName.textContent || "";
      const after = shiftFormulaColumns(
        before,
        insertColumn,
        amount,
        "",
        sheetName,
      );
      if (before !== after) {
        definedName.textContent = after;
      }
    }
    workbook.dirtySheetPaths.add(targetRecord.path);
  }

  async function addSchemaFields(
    workbook,
    sheetName,
    insertColumn,
    fieldNames,
  ) {
    await insertColumns(
      workbook,
      sheetName,
      insertColumn,
      fieldNames.length,
    );
    const record = await workbook.loadSheetRecord(sheetName);
    const table = api.buildTableFromSheet(record, workbook.sharedStrings);
    const cellsByRow = api.rowCellMap(record.document);
    const rowNumbers = [...cellsByRow.keys()];
    for (const rowNumber of rowNumbers) {
      const styleSource =
        cellsByRow.get(rowNumber)?.get(insertColumn - 1) ||
        cellsByRow.get(rowNumber)?.get(insertColumn + fieldNames.length);
      for (let offset = 0; offset < fieldNames.length; offset += 1) {
        const cell = api.getOrCreateCell(
          record.document,
          rowNumber,
          insertColumn + offset,
        );
        copyCellStyle(styleSource, cell);
      }
    }
    for (let offset = 0; offset < fieldNames.length; offset += 1) {
      const cell = api.getOrCreateCell(
        record.document,
        table.headerRow,
        insertColumn + offset,
      );
      api.setDirectCellValue(cell, fieldNames[offset]);
    }
    workbook.dirtySheetPaths.add(record.path);
    return api.buildTableFromSheet(record, workbook.sharedStrings);
  }

  async function clearFieldsForPeople(workbook, sheetName, table, fieldNames) {
    const record = await workbook.loadSheetRecord(sheetName);
    const headers = fieldNames
      .map((name) =>
        table.headers.find((header) => header.name === name),
      )
      .filter(Boolean);
    const rows = table.rows;
    for (const row of rows) {
      for (const header of headers) {
        const cell = api.getOrCreateCell(
          record.document,
          row.rowNumber,
          header.column,
        );
        api.setDirectCellValue(cell, null);
      }
    }
    workbook.dirtySheetPaths.add(record.path);
  }

  async function applyJanuaryInternalFormulas(
    workbook,
    sheetName,
    table,
    nameOnlyRows,
  ) {
    const record = await workbook.loadSheetRecord(sheetName);
    const byName = new Map(
      table.headers.map((header) => [header.name, header]),
    );
    const current = byName.get("扣减税额");
    const previous = byName.get("之前月份累计扣减税额");
    const accumulated = byName.get("累计扣减税额");
    const taxable = byName.get("累计应纳税所得额");
    const income = byName.get("累计应计工资");
    if (![current, previous, accumulated, taxable, income].every(Boolean)) {
      throw new Error("跨年新增字段公式定位不完整");
    }
    for (const rowNumber of nameOnlyRows) {
      const accumulatedCell = api.getOrCreateCell(
        record.document,
        rowNumber,
        accumulated.column,
      );
      setFormula(
        accumulatedCell,
        `${api.columnNumberToLetters(current.column)}${rowNumber}+${api.columnNumberToLetters(previous.column)}${rowNumber}`,
      );
      const taxableCell = api.getOrCreateCell(
        record.document,
        rowNumber,
        taxable.column,
      );
      setFormula(
        taxableCell,
        `${api.columnNumberToLetters(income.column)}${rowNumber}-${api.columnNumberToLetters(accumulated.column)}${rowNumber}`,
      );
    }
    workbook.dirtySheetPaths.add(record.path);
  }

  Object.assign(api, {
    shiftedColumnLetters,
    shiftRangeReference,
    shiftFormulaColumns,
    copyCellStyle,
    setFormula,
    insertColumns,
    addSchemaFields,
    clearFieldsForPeople,
    applyJanuaryInternalFormulas,
  });
})();
