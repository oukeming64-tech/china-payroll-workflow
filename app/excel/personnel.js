(() => {
  "use strict";

  const api = window.PayrollLocal.excel;
  const IDENTITY_ALIASES = {
    employeeId: ["人员编号", "员工编号", "工号", "编号"],
    idCard: ["身份证", "身份证号", "身份证号码", "证件号", "证件号码"],
    name: ["姓名", "员工姓名", "账户名称(*)"],
  };
  const FIELD_ALIASES = {
    人员编号: IDENTITY_ALIASES.employeeId,
    身份证: IDENTITY_ALIASES.idCard,
    姓名: IDENTITY_ALIASES.name,
    部门: ["部门", "所属部门"],
    岗位: ["岗位", "职务"],
    入司时间: ["入司时间", "入职日期"],
    实发合计: ["实发合计", "金额(*)", "金额"],
    银行账号: ["账号(*)", "工资卡号", "银行卡号", "新工资卡号"],
  };
  const PERSONNEL_SHEET_ROLES = Object.freeze({
    离职名单: "archive",
    工资表: "active-primary",
    工资核对表: "active-derived",
    代发薪: "active-static",
    备忘: "optional-notes",
    BM津贴: "unstructured-source",
  });

  function normalized(value) {
    return api.normalizeText(value);
  }

  function headerForAliases(table, aliases) {
    const keys = new Set(aliases.map(normalized));
    return (
      table.headers.find((header) => keys.has(normalized(header.name))) || null
    );
  }

  function identityHeaders(table) {
    return {
      employeeId: headerForAliases(table, IDENTITY_ALIASES.employeeId),
      idCard: headerForAliases(table, IDENTITY_ALIASES.idCard),
      name: headerForAliases(table, IDENTITY_ALIASES.name),
    };
  }

  function identityFromRow(table, row) {
    const headers = identityHeaders(table);
    return {
      employeeId: headers.employeeId
        ? row.values.get(headers.employeeId.column)
        : "",
      idCard: headers.idCard ? row.values.get(headers.idCard.column) : "",
      name: headers.name ? row.values.get(headers.name.column) : "",
    };
  }

  function rowsMatchingIdentity(table, identity) {
    const headers = identityHeaders(table);
    for (const key of ["employeeId", "idCard", "name"]) {
      const value = normalized(identity?.[key]);
      const header = headers[key];
      if (!value || !header) {
        continue;
      }
      const matches = table.rows.filter(
        (row) =>
          !row.hidden &&
          normalized(row.values.get(header.column)) === value,
      );
      if (matches.length) {
        return { matchedBy: key, rows: matches };
      }
    }
    return { matchedBy: "", rows: [] };
  }

  async function analyzePersonnelSheets(workbook) {
    const sheets = [];
    for (const sheet of workbook.sheets) {
      const tables = workbook.getTables
        ? await workbook.getTables(sheet.name)
        : [await workbook.getTable(sheet.name)];
      const table = tables[0];
      const identityHeaderCount = Math.max(
        0,
        ...tables.map((item) =>
          Object.values(identityHeaders(item)).filter(Boolean).length
        ),
      );
      sheets.push({
        name: sheet.name,
        role: PERSONNEL_SHEET_ROLES[sheet.name] || "unknown",
        table,
        tables,
        regionCount: tables.length,
        identityHeaderCount,
        formulaCount: Math.max(
          0,
          ...tables.map((item) => item.formulaCount),
        ),
      });
    }
    return sheets;
  }

  function lastSheetRow(record) {
    const rows = api.elementsByLocalName(record.document, "row");
    return Math.max(
      0,
      ...rows.map((row) => Number(row.getAttribute("r")) || 0),
    );
  }

  function cloneRowStyleOnly(record, sourceRowNumber, targetRowNumber) {
    const sheetData = api.firstByLocalName(record.document, "sheetData");
    const rows = api.directChildrenByLocalName(sheetData, "row");
    const source = rows.find(
      (row) => Number(row.getAttribute("r")) === Number(sourceRowNumber),
    );
    if (!source) {
      throw new Error("找不到可复制的辅助表模板行");
    }
    const clone = source.cloneNode(true);
    clone.setAttribute("r", String(targetRowNumber));
    clone.removeAttribute("hidden");
    for (const cell of api.directChildrenByLocalName(clone, "c")) {
      const parsed = api.parseCellReference(cell.getAttribute("r"));
      if (!parsed) {
        continue;
      }
      cell.setAttribute(
        "r",
        `${api.columnNumberToLetters(parsed.column)}${targetRowNumber}`,
      );
      api.setDirectCellValue(cell, null);
    }
    const following = rows.find(
      (row) => Number(row.getAttribute("r")) > Number(targetRowNumber),
    );
    sheetData.insertBefore(clone, following || null);
    return clone;
  }

  function updateDimensionForCell(documentNode, rowNumber, columnNumber) {
    const dimension = api.firstByLocalName(documentNode, "dimension");
    if (!dimension) {
      return;
    }
    const reference = dimension.getAttribute("ref") || "A1";
    const parts = reference.split(":");
    const start = api.parseCellReference(parts[0]) || { row: 1, column: 1 };
    const end =
      api.parseCellReference(parts.at(-1)) || { row: 1, column: 1 };
    dimension.setAttribute(
      "ref",
      `${api.columnNumberToLetters(Math.min(start.column, columnNumber))}${Math.min(
        start.row,
        rowNumber,
      )}:${api.columnNumberToLetters(Math.max(end.column, columnNumber))}${Math.max(
        end.row,
        rowNumber,
      )}`,
    );
  }

  function valueForTargetHeader(targetHeader, values) {
    const exact = Object.entries(values || {}).find(
      ([name]) => normalized(name) === normalized(targetHeader.name),
    );
    if (exact) {
      return exact[1];
    }
    for (const [canonical, aliases] of Object.entries(FIELD_ALIASES)) {
      if (!aliases.some((alias) => normalized(alias) === normalized(targetHeader.name))) {
        continue;
      }
      const candidate = Object.entries(values || {}).find(
        ([name]) =>
          normalized(name) === normalized(canonical) ||
          aliases.some((alias) => normalized(alias) === normalized(name)),
      );
      if (candidate) {
        return candidate[1];
      }
    }
    return undefined;
  }

  async function appendMappedPersonRow(
    workbook,
    sheetName,
    values,
    options = {},
  ) {
    const table = await workbook.getTable(sheetName);
    if (!table.headers.length) {
      throw new Error(`${sheetName} 没有可验证的字段头`);
    }
    const record = await workbook.loadSheetRecord(sheetName);
    const sourceRow =
      options.sourceRow ||
      table.rows.at(-1)?.rowNumber ||
      Math.max(table.headerRow + 1, lastSheetRow(record));
    const targetRow = Math.max(
      lastSheetRow(record) + 1,
      table.headerRow + 1,
    );
    cloneRowStyleOnly(record, sourceRow, targetRow);
    const missing = [];
    for (const header of table.headers) {
      const value = valueForTargetHeader(header, values);
      if (
        options.requiredFields?.some(
          (name) => normalized(name) === normalized(header.name),
        ) &&
        (value === undefined || value === null || String(value).trim() === "")
      ) {
        missing.push(header.name);
      }
      if (value === undefined || value === null || value === "") {
        continue;
      }
      const cell = api.getOrCreateCell(
        record.document,
        targetRow,
        header.column,
      );
      api.setDirectCellValue(cell, value);
      updateDimensionForCell(record.document, targetRow, header.column);
    }
    if (missing.length) {
      throw new Error(`${sheetName} 缺少必填字段：${missing.join("、")}`);
    }
    workbook.dirtySheetPaths.add(record.path);
    return { sheetName, rowNumber: targetRow };
  }

  async function clearIdentityAndHideMatches(
    workbook,
    sheetName,
    identity,
  ) {
    const table = await workbook.getTable(sheetName);
    const match = rowsMatchingIdentity(table, identity);
    if (!match.rows.length) {
      return { sheetName, matchedBy: "", rows: [] };
    }
    const record = await workbook.loadSheetRecord(sheetName);
    const headers = identityHeaders(table);
    for (const row of match.rows) {
      for (const header of Object.values(headers).filter(Boolean)) {
        const cell = api.getOrCreateCell(
          record.document,
          row.rowNumber,
          header.column,
        );
        api.setDirectCellValue(cell, null);
      }
      await api.setRowHidden(workbook, sheetName, row.rowNumber, true);
    }
    workbook.dirtySheetPaths.add(record.path);
    return {
      sheetName,
      matchedBy: match.matchedBy,
      rows: match.rows.map((row) => row.rowNumber),
    };
  }

  async function updateMatchingRows(
    workbook,
    sheetName,
    identity,
    fieldName,
    value,
  ) {
    const table = await workbook.getTable(sheetName);
    const match = rowsMatchingIdentity(table, identity);
    if (!match.rows.length) {
      return { sheetName, rows: [], updated: 0 };
    }
    const aliases = FIELD_ALIASES[fieldName] || [fieldName];
    const header = headerForAliases(table, aliases);
    if (!header) {
      return {
        sheetName,
        rows: match.rows.map((row) => row.rowNumber),
        updated: 0,
      };
    }
    for (const row of match.rows) {
      await api.updateCell(
        workbook,
        sheetName,
        row.rowNumber,
        header.column,
        value,
        { preserveFormula: true },
      );
    }
    return {
      sheetName,
      rows: match.rows.map((row) => row.rowNumber),
      updated: match.rows.length,
    };
  }

  async function formulaReferencesMainRow(
    workbook,
    sheetName,
    mainSheetName,
    rowNumber,
  ) {
    const record = await workbook.loadSheetRecord(sheetName);
    const escaped = mainSheetName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const matcher = new RegExp(
      `(?:'${escaped}'|${escaped})!\\$?[A-Z]{1,3}\\$?${rowNumber}(?!\\d)`,
      "i",
    );
    return api
      .elementsByLocalName(record.document, "f")
      .filter((formula) => matcher.test(formula.textContent || ""))
      .map((formula) => formula.parentElement?.getAttribute("r") || "")
      .filter(Boolean);
  }

  async function formulaReferencesIdentity(
    workbook,
    sheetName,
    mainSheetName,
    identity,
  ) {
    const table = await workbook.getTable(sheetName);
    const match = rowsMatchingIdentity(table, identity);
    if (!match.rows.length) {
      return { matchedBy: "", rows: [], cells: [] };
    }
    const escaped = mainSheetName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const matcher = new RegExp(`(?:'${escaped}'|${escaped})!`, "i");
    const cells = [];
    for (const row of match.rows) {
      for (const descriptor of row.cells.values()) {
        if (matcher.test(descriptor.formula || "")) {
          cells.push(descriptor.address);
        }
      }
    }
    return {
      matchedBy: match.matchedBy,
      rows: match.rows.map((row) => row.rowNumber),
      cells,
    };
  }

  Object.assign(api, {
    EXCEL_IDENTITY_ALIASES: IDENTITY_ALIASES,
    EXCEL_FIELD_ALIASES: FIELD_ALIASES,
    PERSONNEL_SHEET_ROLES,
    personnelIdentityHeaders: identityHeaders,
    identityFromTableRow: identityFromRow,
    rowsMatchingIdentity,
    analyzePersonnelSheets,
    valueForTargetHeader,
    appendMappedPersonRow,
    clearIdentityAndHideMatches,
    updateMatchingRows,
    formulaReferencesMainRow,
    formulaReferencesIdentity,
  });
})();
