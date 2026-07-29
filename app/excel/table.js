(() => {
  "use strict";

  const api = window.PayrollLocal.excel;
  const IDENTITY_HEADER_KEYS = new Set(
    [
      "人员编号",
      "员工编号",
      "工号",
      "编号",
      "身份证",
      "身份证号",
      "证件号码",
      "证件号",
      "姓名",
      "员工姓名",
      "name",
      "employeeid",
      "staffid",
      "idcard",
    ].map(api.normalizeText),
  );

  function detectHeaderRow(cellsByRow, sharedStrings, preferredRow = null) {
    if (preferredRow && cellsByRow.has(Number(preferredRow))) {
      return Number(preferredRow);
    }
    let best = { row: 1, score: -Infinity };
    const candidateRows = [...cellsByRow.keys()]
      .filter((row) => row <= 15)
      .sort((left, right) => left - right);
    for (const row of candidateRows) {
      const descriptors = [...cellsByRow.get(row).values()]
        .map((cell) => api.cellDescriptor(cell, sharedStrings))
        .filter(
          (descriptor) =>
            descriptor.value !== null && descriptor.value !== "",
        );
      const identityHits = descriptors.filter((descriptor) =>
        IDENTITY_HEADER_KEYS.has(api.normalizeText(descriptor.value)),
      ).length;
      const textual = descriptors.filter(
        (descriptor) => typeof descriptor.value === "string",
      ).length;
      const score =
        identityHits * 100 +
        Math.min(descriptors.length, 70) +
        textual * 0.2 -
        (descriptors.length <= 1 ? 20 : 0);
      if (score > best.score) {
        best = { row, score };
      }
    }
    return best.row;
  }

  function makeUniqueHeaderLabels(headers) {
    const counts = new Map();
    return headers.map((header) => {
      const base = String(header.name || `列 ${header.columnLabel}`).trim();
      const count = (counts.get(base) || 0) + 1;
      counts.set(base, count);
      return {
        ...header,
        displayName: count === 1 ? base : `${base}（${header.columnLabel}列）`,
      };
    });
  }

  function hasIdentityValue(valuesByColumn, identityColumns) {
    return identityColumns.some((column) => {
      const value = valuesByColumn.get(column);
      return value !== null &&
        value !== undefined &&
        String(value).trim() !== "";
    });
  }

  function sheetSubheaders(
    cellsByRow,
    headerRow,
    headerColumns,
    identityColumns,
    sharedStrings,
  ) {
    const cells = cellsByRow.get(headerRow + 1);
    if (!cells) {
      return [];
    }
    const values = new Map(
      [...cells.entries()].map(([column, cell]) => [
        column,
        api.cellDescriptor(cell, sharedStrings).value,
      ]),
    );
    const candidates = [...values.entries()]
      .filter(
        ([, value]) =>
          typeof value === "string" && String(value).trim(),
      )
      .map(([column, value]) => ({
        column,
        columnLabel: api.columnNumberToLetters(column),
        name: String(value).trim(),
      }));
    if (
      candidates.length < 2 ||
      !candidates.some((header) => !headerColumns.has(header.column)) ||
      hasIdentityValue(values, identityColumns)
    ) {
      return [];
    }
    return makeUniqueHeaderLabels(candidates);
  }

  function readableColumns(headers, subheaders) {
    return [
      ...new Set(
        [...headers, ...subheaders].map((header) => header.column),
      ),
    ];
  }

  function columnForHeader(headers, subheaders, headerOrColumn) {
    if (typeof headerOrColumn === "number") {
      return headerOrColumn;
    }
    return [...headers, ...subheaders].find(
      (header) =>
        header.name === headerOrColumn ||
        header.displayName === headerOrColumn,
    )?.column;
  }

  function buildTableFromSheet(sheetRecord, sharedStrings, options = {}) {
    const cellsByRow = api.rowCellMap(sheetRecord.document);
    const headerRow = detectHeaderRow(
      cellsByRow,
      sharedStrings,
      options.headerRow,
    );
    const headerCells = cellsByRow.get(headerRow) || new Map();
    const headers = makeUniqueHeaderLabels(
      [...headerCells.entries()]
        .map(([column, cell]) => ({
          column,
          columnLabel: api.columnNumberToLetters(column),
          name: String(api.cellDescriptor(cell, sharedStrings).value ?? "").trim(),
        }))
        .filter((header) => header.name),
    );
    const headerByColumn = new Map(
      headers.map((header) => [header.column, header]),
    );
    const identityColumns = headers
      .filter((header) =>
        IDENTITY_HEADER_KEYS.has(api.normalizeText(header.name)),
      )
      .map((header) => header.column);
    const subheaders = sheetSubheaders(
      cellsByRow,
      headerRow,
      new Set(headers.map((header) => header.column)),
      identityColumns,
      sharedStrings,
    );
    const subheaderByColumn = new Map(
      subheaders.map((header) => [header.column, header]),
    );
    const columns = readableColumns(headers, subheaders);
    const dataStartRow = headerRow + (subheaders.length ? 1 : 0);

    const rows = [];
    const hiddenRows = new Set(
      api
        .elementsByLocalName(sheetRecord.document, "row")
        .filter((row) => row.getAttribute("hidden") === "1")
        .map((row) => Number(row.getAttribute("r")))
        .filter(Number.isFinite),
    );
    const sortedRows = [...cellsByRow.keys()]
      .filter(
        (row) =>
          row > dataStartRow &&
          (!options.endRow || row <= Number(options.endRow)),
      )
      .sort((left, right) => left - right);
    for (const rowNumber of sortedRows) {
      const cellMap = cellsByRow.get(rowNumber);
      const values = new Map();
      const cells = new Map();
      let anyValue = false;
      for (const column of columns) {
        const descriptor = api.cellDescriptor(
          cellMap.get(column),
          sharedStrings,
        );
        values.set(column, descriptor.value);
        cells.set(column, descriptor);
        if (descriptor.value !== null && descriptor.value !== "") {
          anyValue = true;
        }
      }
      const hasIdentity = identityColumns.length
        ? hasIdentityValue(values, identityColumns)
        : anyValue;
      if (!hasIdentity) {
        continue;
      }
      rows.push({
        rowNumber,
        hidden: hiddenRows.has(rowNumber),
        values,
        cells,
        get(headerOrColumn) {
          const column = columnForHeader(
            headers,
            subheaders,
            headerOrColumn,
          );
          return column ? values.get(column) : null;
        },
      });
    }

    return {
      sheetName: sheetRecord.name,
      sheetPath: sheetRecord.path,
      headerRow,
      headers,
      headerByColumn,
      subheaders,
      subheaderByColumn,
      dataStartRow,
      rows,
      formulaCount: api.elementsByLocalName(
        sheetRecord.document,
        "f",
      ).length,
      rowCount: rows.length,
    };
  }

  function matrixHeaderRow(matrix, preferredRow = null) {
    if (
      preferredRow &&
      Array.isArray(matrix[Number(preferredRow) - 1])
    ) {
      return Number(preferredRow);
    }
    let best = { row: 1, score: -Infinity };
    for (let index = 0; index < Math.min(matrix.length, 15); index += 1) {
      const values = Array.isArray(matrix[index]) ? matrix[index] : [];
      const populated = values.filter(
        (value) =>
          value !== null &&
          value !== undefined &&
          String(value).trim() !== "",
      );
      const identityHits = populated.filter((value) =>
        IDENTITY_HEADER_KEYS.has(api.normalizeText(value)),
      ).length;
      const textual = populated.filter(
        (value) => typeof value === "string",
      ).length;
      const score =
        identityHits * 100 +
        Math.min(populated.length, 70) +
        textual * 0.2 -
        (populated.length <= 1 ? 20 : 0);
      if (score > best.score) {
        best = { row: index + 1, score };
      }
    }
    return best.row;
  }

  function matrixCellDescriptor(value, metadata = null) {
    return {
      address: metadata?.address || "",
      value,
      formula: metadata?.formula || "",
      hasFormula: Boolean(metadata?.formula),
      hasCachedValue:
        value !== null &&
        value !== undefined &&
        String(value).trim() !== "",
      type: metadata?.type || "",
      styleId: "",
    };
  }

  function matrixSubheaders(
    rows,
    headerRow,
    headerColumns,
    identityColumns,
  ) {
    const values = Array.isArray(rows[headerRow])
      ? rows[headerRow]
      : [];
    const byColumn = new Map(
      values.map((value, index) => [index + 1, value]),
    );
    const candidates = values
      .map((value, index) => ({
        column: index + 1,
        columnLabel: api.columnNumberToLetters(index + 1),
        name: typeof value === "string" ? value.trim() : "",
      }))
      .filter((header) => header.name);
    if (
      candidates.length < 2 ||
      !candidates.some((header) => !headerColumns.has(header.column)) ||
      hasIdentityValue(byColumn, identityColumns)
    ) {
      return [];
    }
    return makeUniqueHeaderLabels(candidates);
  }

  function buildTableFromMatrix(
    matrix,
    sheetName = "Sheet1",
    options = {},
  ) {
    const rows = Array.isArray(matrix) ? matrix : [];
    const headerRow = matrixHeaderRow(rows, options.headerRow);
    const headerValues = Array.isArray(rows[headerRow - 1])
      ? rows[headerRow - 1]
      : [];
    const headers = makeUniqueHeaderLabels(
      headerValues
        .map((value, index) => ({
          column: index + 1,
          columnLabel: api.columnNumberToLetters(index + 1),
          name: String(value ?? "").trim(),
        }))
        .filter((header) => header.name),
    );
    const headerByColumn = new Map(
      headers.map((header) => [header.column, header]),
    );
    const identityColumns = headers
      .filter((header) =>
        IDENTITY_HEADER_KEYS.has(api.normalizeText(header.name)),
      )
      .map((header) => header.column);
    const subheaders = matrixSubheaders(
      rows,
      headerRow,
      new Set(headers.map((header) => header.column)),
      identityColumns,
    );
    const subheaderByColumn = new Map(
      subheaders.map((header) => [header.column, header]),
    );
    const columns = readableColumns(headers, subheaders);
    const dataStartRow = headerRow + (subheaders.length ? 1 : 0);
    const dataRows = [];
    const endIndex = options.endRow
      ? Math.min(rows.length, Number(options.endRow))
      : rows.length;
    for (let index = dataStartRow; index < endIndex; index += 1) {
      const sourceRow = Array.isArray(rows[index]) ? rows[index] : [];
      const metadataRow = options.metadata?.[index] || [];
      const values = new Map();
      const cells = new Map();
      let anyValue = false;
      for (const column of columns) {
        const value = sourceRow[column - 1] ?? null;
        values.set(column, value);
        cells.set(
          column,
          matrixCellDescriptor(
            value,
            metadataRow[column - 1],
          ),
        );
        if (
          value !== null &&
          value !== undefined &&
          String(value).trim() !== ""
        ) {
          anyValue = true;
        }
      }
      const hasIdentity = identityColumns.length
        ? hasIdentityValue(values, identityColumns)
        : anyValue;
      if (!hasIdentity) {
        continue;
      }
      dataRows.push({
        rowNumber: index + 1,
        hidden: Boolean(metadataRow.hidden),
        values,
        cells,
        get(headerOrColumn) {
          const column = columnForHeader(
            headers,
            subheaders,
            headerOrColumn,
          );
          return column ? values.get(column) : null;
        },
      });
    }
    return {
      sheetName,
      sheetPath: "",
      headerRow,
      headers,
      headerByColumn,
      subheaders,
      subheaderByColumn,
      dataStartRow,
      rows: dataRows,
      formulaCount: dataRows.reduce(
        (count, row) =>
          count +
          [...row.cells.values()].filter((cell) => cell.hasFormula)
            .length,
        0,
      ),
      rowCount: dataRows.length,
    };
  }

  Object.assign(api, {
    IDENTITY_HEADER_KEYS,
    detectHeaderRow,
    buildTableFromSheet,
    buildTableFromMatrix,
  });
})();
