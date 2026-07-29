(() => {
  "use strict";

  const api = window.PayrollLocal.excel;

  function isRegionHeader(values) {
    const populated = values.filter(
      (value) =>
        value !== null &&
        value !== undefined &&
        String(value).trim(),
    );
    const identityHits = populated.filter((value) =>
      api.IDENTITY_HEADER_KEYS.has(api.normalizeText(value)),
    ).length;
    const textual = populated.filter(
      (value) => typeof value === "string",
    ).length;
    return identityHits > 0 && textual >= 3;
  }

  function sheetRegionHeaderRows(sheetRecord, sharedStrings) {
    const cellsByRow = api.rowCellMap(sheetRecord.document);
    return [...cellsByRow.entries()]
      .filter(([row]) => row <= 100)
      .filter(([, cells]) =>
        isRegionHeader(
          [...cells.values()].map(
            (cell) => api.cellDescriptor(cell, sharedStrings).value,
          ),
        ),
      )
      .map(([row]) => row)
      .sort((left, right) => left - right);
  }

  function matrixRegionHeaderRows(matrix) {
    return (Array.isArray(matrix) ? matrix : [])
      .slice(0, 100)
      .map((values, index) => ({
        row: index + 1,
        values: Array.isArray(values) ? values : [],
      }))
      .filter((item) => isRegionHeader(item.values))
      .map((item) => item.row);
  }

  function regionBounds(headerRows, index) {
    return {
      headerRow: headerRows[index],
      endRow: headerRows[index + 1]
        ? headerRows[index + 1] - 1
        : undefined,
    };
  }

  function buildTablesFromSheet(sheetRecord, sharedStrings) {
    const headerRows = sheetRegionHeaderRows(
      sheetRecord,
      sharedStrings,
    );
    if (!headerRows.length) {
      return [api.buildTableFromSheet(sheetRecord, sharedStrings)];
    }
    return headerRows.map((headerRow, index) =>
      api.buildTableFromSheet(
        sheetRecord,
        sharedStrings,
        regionBounds(headerRows, index),
      ),
    );
  }

  function buildTablesFromMatrix(
    matrix,
    sheetName = "Sheet1",
    options = {},
  ) {
    const headerRows = matrixRegionHeaderRows(matrix);
    if (!headerRows.length) {
      return [api.buildTableFromMatrix(matrix, sheetName, options)];
    }
    return headerRows.map((headerRow, index) =>
      api.buildTableFromMatrix(matrix, sheetName, {
        ...options,
        ...regionBounds(headerRows, index),
      }),
    );
  }

  Object.assign(api, {
    sheetRegionHeaderRows,
    matrixRegionHeaderRows,
    buildTablesFromSheet,
    buildTablesFromMatrix,
  });
})();
