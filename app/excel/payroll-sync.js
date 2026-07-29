(() => {
  "use strict";

  const api = window.PayrollLocal.excel;
  const ACCOUNT_ALIASES = [
    "账号(*)",
    "账号",
    "工资卡号",
    "银行卡号",
    "新工资卡号",
  ];
  const AMOUNT_ALIASES = ["金额(*)", "金额", "实发合计"];

  function headerForAliases(table, aliases) {
    const keys = new Set(aliases.map(api.normalizeText));
    return (
      table.headers.find((header) =>
        keys.has(api.normalizeText(header.name)),
      ) || null
    );
  }

  function writeCellByAliases(
    record,
    table,
    rowNumber,
    aliases,
    value,
  ) {
    const header = headerForAliases(table, aliases);
    if (
      !header ||
      value === null ||
      value === undefined ||
      String(value).trim() === ""
    ) {
      return false;
    }
    const cell = api.getOrCreateCell(
      record.document,
      rowNumber,
      header.column,
    );
    api.setDirectCellValue(cell, value);
    return true;
  }

  function quotedSheetReference(sheetName, columnNumber, rowNumber) {
    const escaped = String(sheetName || "").replaceAll("'", "''");
    return `'${escaped}'!$${api.columnNumberToLetters(columnNumber)}$${rowNumber}`;
  }

  function exactInternalVlookups(formula, mainSheetName) {
    const matches = [];
    const matcher =
      /VLOOKUP\(\s*[^,]+,\s*(?:'((?:[^']|'')+)'|([^!,]+))!\s*\$?([A-Z]{1,3})(?:\$?\d+)?\s*:\s*\$?([A-Z]{1,3})(?:\$?\d+)?\s*,\s*(\d+)\s*,\s*(?:0|FALSE)\s*\)/gi;
    for (const match of String(formula || "").matchAll(matcher)) {
      const sheetName = String(match[1] || match[2] || "")
        .replaceAll("''", "'")
        .trim();
      if (api.normalizeText(sheetName) !== api.normalizeText(mainSheetName)) {
        continue;
      }
      const startColumn = api.columnLettersToNumber(match[3]);
      const endColumn = api.columnLettersToNumber(match[4]);
      const returnIndex = Number(match[5]);
      const valueColumn = startColumn + returnIndex - 1;
      if (
        !Number.isInteger(returnIndex) ||
        returnIndex < 1 ||
        valueColumn > endColumn
      ) {
        continue;
      }
      matches.push({
        raw: match[0],
        index: match.index,
        end: match.index + match[0].length,
        valueColumn,
      });
    }
    return matches;
  }

  function additiveFormulaTerms(formula, mainSheetName, rowNumber) {
    const lookups = exactInternalVlookups(formula, mainSheetName);
    let cursor = 0;
    let expression = "";
    for (let index = 0; index < lookups.length; index += 1) {
      const lookup = lookups[index];
      expression += String(formula).slice(cursor, lookup.index);
      expression += `#${index}#`;
      cursor = lookup.end;
    }
    expression += String(formula).slice(cursor);
    expression = expression.replace(/\s+/g, "").replace(/^=/, "");
    const tokenMatcher =
      /([+-]?)(#[0-9]+#|\$?[A-Z]{1,3}\$?\d+)/g;
    const tokens = [...expression.matchAll(tokenMatcher)].map((match) => ({
      sign: match[1] === "-" ? -1 : 1,
      token: match[2],
      raw: match[0],
    }));
    if (!tokens.length || tokens.map((item) => item.raw).join("") !== expression) {
      return null;
    }
    const terms = [];
    for (const token of tokens) {
      if (/^#[0-9]+#$/.test(token.token)) {
        const lookup = lookups[Number(token.token.slice(1, -1))];
        if (!lookup) {
          return null;
        }
        terms.push({
          kind: "main",
          coefficient: token.sign,
          column: lookup.valueColumn,
        });
        continue;
      }
      const reference = api.parseCellReference(
        token.token.replaceAll("$", ""),
      );
      if (!reference || reference.row !== Number(rowNumber)) {
        return null;
      }
      terms.push({
        kind: "local",
        coefficient: token.sign,
        column: reference.column,
      });
    }
    return terms;
  }

  function numericCellValue(record, sharedStrings, rowNumber, columnNumber) {
    const cell = api
      .rowCellMap(record.document)
      .get(Number(rowNumber))
      ?.get(Number(columnNumber));
    const value = api.cellDescriptor(cell, sharedStrings).value;
    if (value === null || value === undefined || value === "") {
      return 0;
    }
    const numeric = Number(value);
    return Number.isFinite(numeric) ? numeric : null;
  }

  function populateCheckFormulaInputs(
    record,
    sharedStrings,
    checkRowNumber,
    mainSheetName,
    mainRow,
  ) {
    const row = api
      .elementsByLocalName(record.document, "row")
      .find(
        (candidate) =>
          Number(candidate.getAttribute("r")) === Number(checkRowNumber),
      );
    const formulas = row
      ? api
          .directChildrenByLocalName(row, "c")
          .map((cell) => api.cellDescriptor(cell, sharedStrings))
          .filter(
            (descriptor) =>
              descriptor.hasFormula && descriptor.formula,
          )
      : [];
    const planned = formulas
      .map((descriptor) => ({
        descriptor,
        terms: additiveFormulaTerms(
          descriptor.formula,
          mainSheetName,
          checkRowNumber,
        ),
      }))
      .sort((left, right) => {
        const leftCount =
          left.terms?.filter((term) => term.kind === "local").length ??
          Number.MAX_SAFE_INTEGER;
        const rightCount =
          right.terms?.filter((term) => term.kind === "local").length ??
          Number.MAX_SAFE_INTEGER;
        return leftCount - rightCount;
      });
    let mapped = 0;
    const unsupported = [];
    for (const item of planned) {
      const localTerms =
        item.terms?.filter((term) => term.kind === "local") || [];
      if (!localTerms.length) {
        if (!item.terms) {
          unsupported.push(item.descriptor.address);
        }
        continue;
      }
      const target = localTerms.find(
        (term) => term.coefficient === -1,
      );
      if (!target) {
        unsupported.push(item.descriptor.address);
        continue;
      }
      let known = 0;
      let valid = true;
      for (const term of item.terms) {
        if (term === target) {
          continue;
        }
        const value =
          term.kind === "main"
            ? mainRow.values.get(term.column)
            : numericCellValue(
                record,
                sharedStrings,
                checkRowNumber,
                term.column,
              );
        const numeric =
          value === null || value === undefined || value === ""
            ? 0
            : Number(value);
        if (!Number.isFinite(numeric)) {
          valid = false;
          break;
        }
        known += term.coefficient * numeric;
      }
      if (!valid) {
        unsupported.push(item.descriptor.address);
        continue;
      }
      const cell = api.getOrCreateCell(
        record.document,
        checkRowNumber,
        target.column,
      );
      api.setDirectCellValue(cell, -known / target.coefficient);
      mapped += 1;
    }
    return {
      formulaCount: formulas.length,
      mappedInputCount: mapped,
      unsupported,
    };
  }

  async function matchingPersonnelRows(workbook, sheetName, identity) {
    const table = await workbook.getTable(sheetName);
    const match = api.rowsMatchingIdentity(table, identity);
    return {
      sheetName,
      matchedBy: match.matchedBy,
      rows: match.rows.map((row) => row.rowNumber),
    };
  }

  async function findPersonnelTemplateRow(
    workbook,
    sheetName,
    options = {},
  ) {
    const table = await workbook.getTable(sheetName);
    const identities = api.personnelIdentityHeaders(table);
    const candidates = table.rows.filter((row) => {
      const identityCount = Object.values(identities)
        .filter(Boolean)
        .filter((header) =>
          String(row.values.get(header.column) ?? "").trim(),
        ).length;
      const formulaCount = [...row.cells.values()].filter(
        (cell) => cell.hasFormula,
      ).length;
      return (
        !row.hidden &&
        identityCount >= 2 &&
        (!options.requireFormula || formulaCount > 0)
      );
    });
    return candidates.at(-1)?.rowNumber || null;
  }

  async function recyclePayrollCheckRow(workbook, options) {
    const mainTable = await workbook.getTable(options.mainSheetName);
    const checkTable = await workbook.getTable(options.sheetName);
    const mainRow = mainTable.rows.find(
      (row) => row.rowNumber === Number(options.mainRowNumber),
    );
    if (!mainRow) {
      throw new Error("工资表找不到需要补入核对表的人员行");
    }
    await api.cloneEmployeeRow(
      workbook,
      options.sheetName,
      options.sourceRow,
      options.targetRow,
      { replaceHiddenTarget: true },
    );
    const record = await workbook.loadSheetRecord(options.sheetName);
    writeCellByAliases(
      record,
      checkTable,
      options.targetRow,
      api.EXCEL_IDENTITY_ALIASES.employeeId,
      options.identity.employeeId,
    );
    writeCellByAliases(
      record,
      checkTable,
      options.targetRow,
      api.EXCEL_IDENTITY_ALIASES.idCard,
      options.identity.idCard,
    );
    writeCellByAliases(
      record,
      checkTable,
      options.targetRow,
      api.EXCEL_IDENTITY_ALIASES.name,
      options.identity.name,
    );
    writeCellByAliases(
      record,
      checkTable,
      options.targetRow,
      ["部门", "所属部门"],
      options.department,
    );
    const formulaInputs = populateCheckFormulaInputs(
      record,
      workbook.sharedStrings,
      options.targetRow,
      options.mainSheetName,
      mainRow,
    );
    if (formulaInputs.unsupported.length) {
      throw new Error(
        `工资核对表复用行有 ${formulaInputs.unsupported.length} 个公式无法安全建立校验输入`,
      );
    }
    api.setCalcMode(workbook.workbookDocument);
    workbook.dirtySheetPaths.add(record.path);
    return {
      sheetName: options.sheetName,
      rowNumber: options.targetRow,
      ...formulaInputs,
    };
  }

  async function linkDisbursementAmount(workbook, options) {
    const table = await workbook.getTable(options.sheetName);
    const match = api.rowsMatchingIdentity(table, options.identity);
    const amountHeader = headerForAliases(table, AMOUNT_ALIASES);
    if (!match.rows.length || !amountHeader) {
      return {
        sheetName: options.sheetName,
        rows: match.rows.map((row) => row.rowNumber),
        updated: 0,
      };
    }
    const formula = quotedSheetReference(
      options.mainSheetName,
      options.finalPayColumn,
      options.mainRowNumber,
    );
    for (const row of match.rows) {
      await api.updateFormulaCell(
        workbook,
        options.sheetName,
        row.rowNumber,
        amountHeader.column,
        formula,
        options.cachedValue,
      );
    }
    return {
      sheetName: options.sheetName,
      rows: match.rows.map((row) => row.rowNumber),
      updated: match.rows.length,
      formula,
    };
  }

  async function recycleDisbursementRow(workbook, options) {
    const table = await workbook.getTable(options.sheetName);
    await api.cloneEmployeeRow(
      workbook,
      options.sheetName,
      options.sourceRow,
      options.targetRow,
      { replaceHiddenTarget: true },
    );
    const record = await workbook.loadSheetRecord(options.sheetName);
    writeCellByAliases(
      record,
      table,
      options.targetRow,
      api.EXCEL_IDENTITY_ALIASES.employeeId,
      options.identity.employeeId,
    );
    writeCellByAliases(
      record,
      table,
      options.targetRow,
      api.EXCEL_IDENTITY_ALIASES.idCard,
      options.identity.idCard,
    );
    writeCellByAliases(
      record,
      table,
      options.targetRow,
      ["账户名称(*)", ...api.EXCEL_IDENTITY_ALIASES.name],
      options.identity.name,
    );
    if (
      !writeCellByAliases(
        record,
        table,
        options.targetRow,
        ACCOUNT_ALIASES,
        options.account,
      )
    ) {
      throw new Error("代发薪缺少可验证的银行账号");
    }
    workbook.dirtySheetPaths.add(record.path);
    const linked = await linkDisbursementAmount(workbook, options);
    if (!linked.updated) {
      throw new Error("代发薪复用行未能建立实发公式引用");
    }
    return {
      ...linked,
      rowNumber: options.targetRow,
      reused: true,
    };
  }

  Object.assign(api, {
    matchingPersonnelRows,
    findPersonnelTemplateRow,
    recyclePayrollCheckRow,
    linkDisbursementAmount,
    recycleDisbursementRow,
  });
})();
