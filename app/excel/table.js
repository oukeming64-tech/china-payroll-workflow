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

    const rows = [];
    const sortedRows = [...cellsByRow.keys()]
      .filter((row) => row > headerRow)
      .sort((left, right) => left - right);
    for (const rowNumber of sortedRows) {
      const cellMap = cellsByRow.get(rowNumber);
      const values = new Map();
      const cells = new Map();
      let anyValue = false;
      for (const header of headers) {
        const descriptor = api.cellDescriptor(
          cellMap.get(header.column),
          sharedStrings,
        );
        values.set(header.column, descriptor.value);
        cells.set(header.column, descriptor);
        if (descriptor.value !== null && descriptor.value !== "") {
          anyValue = true;
        }
      }
      const hasIdentity = identityColumns.length
        ? identityColumns.some((column) => {
            const value = values.get(column);
            return value !== null && String(value).trim() !== "";
          })
        : anyValue;
      if (!hasIdentity) {
        continue;
      }
      rows.push({
        rowNumber,
        values,
        cells,
        get(headerOrColumn) {
          const column =
            typeof headerOrColumn === "number"
              ? headerOrColumn
              : headers.find(
                  (header) =>
                    header.name === headerOrColumn ||
                    header.displayName === headerOrColumn,
                )?.column;
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
    const dataRows = [];
    for (let index = headerRow; index < rows.length; index += 1) {
      const sourceRow = Array.isArray(rows[index]) ? rows[index] : [];
      const metadataRow = options.metadata?.[index] || [];
      const values = new Map();
      const cells = new Map();
      let anyValue = false;
      for (const header of headers) {
        const value = sourceRow[header.column - 1] ?? null;
        values.set(header.column, value);
        cells.set(
          header.column,
          matrixCellDescriptor(
            value,
            metadataRow[header.column - 1],
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
        ? identityColumns.some((column) => {
            const value = values.get(column);
            return (
              value !== null &&
              value !== undefined &&
              String(value).trim() !== ""
            );
          })
        : anyValue;
      if (!hasIdentity) {
        continue;
      }
      dataRows.push({
        rowNumber: index + 1,
        values,
        cells,
        get(headerOrColumn) {
          const column =
            typeof headerOrColumn === "number"
              ? headerOrColumn
              : headers.find(
                  (header) =>
                    header.name === headerOrColumn ||
                    header.displayName === headerOrColumn,
                )?.column;
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
