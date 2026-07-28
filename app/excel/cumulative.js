(() => {
  "use strict";

  const api = window.PayrollLocal.excel;
  const CUMULATIVE_MAPPINGS = Object.freeze([
    ["之前月份累计应计工资", "累计应计工资"],
    ["之前月份累计扣减税额", "累计扣减税额", "累计应计工资"],
    ["之前月份累计代扣养老保险", "累计代扣养老保险"],
    ["之前月份累计代扣医疗保险", "累计代扣医疗保险"],
    ["之前月份累计代扣失业保险", "累计代扣失业保险"],
    ["之前月份累计代扣房积金", "累计代扣房积金"],
    ["之前月份累计基本扣除费用", "累计基本扣除费用"],
    ["之前月份累计代扣税", "累计实缴代扣税"],
  ]);

  function populated(row, header) {
    const value = header ? row.values.get(header.column) : null;
    return Boolean(
      header &&
        (row.cells.get(header.column)?.hasFormula ||
          (value !== null &&
            value !== undefined &&
            String(value).trim() !== "")),
    );
  }

  async function materializePreviousCumulative(
    workbook,
    sheetName,
    table,
  ) {
    const record = await workbook.loadSheetRecord(sheetName);
    const identity = table.headers.find((header) =>
      ["身份证", "身份证号"].includes(header.name),
    );
    if (!identity) {
      throw new Error("工资表缺少身份证字段，无法接续上月累计值");
    }
    const errors = [];
    const mappings = [];
    let written = 0;
    for (const [previousName, accumulatedName, applicabilityName] of
      CUMULATIVE_MAPPINGS) {
      const previous = table.headers.find(
        (header) => header.name === previousName,
      );
      const accumulated = table.headers.find(
        (header) => header.name === accumulatedName,
      );
      const applicability = applicabilityName
        ? table.headers.find((header) => header.name === applicabilityName)
        : null;
      if (!previous || !accumulated) {
        continue;
      }
      let fieldWrites = 0;
      for (const row of table.rows) {
        if (
          !populated(row, identity) ||
          (applicabilityName && !populated(row, applicability))
        ) {
          continue;
        }
        const source = row.cells.get(accumulated.column);
        if (
          source?.hasFormula &&
          !source.hasCachedValue
        ) {
          errors.push(
            `${accumulatedName} 第 ${row.rowNumber} 行缺少可用缓存值`,
          );
          continue;
        }
        const cell = api.getOrCreateCell(
          record.document,
          row.rowNumber,
          previous.column,
        );
        api.setDirectCellValue(
          cell,
          row.values.get(accumulated.column),
        );
        written += 1;
        fieldWrites += 1;
      }
      mappings.push({
        targetField: previousName,
        sourceField: accumulatedName,
        written: fieldWrites,
      });
    }
    if (errors.length) {
      throw new Error(errors.slice(0, 8).join("；"));
    }
    workbook.dirtySheetPaths.add(record.path);
    return { written, mappings };
  }

  Object.assign(api, {
    CUMULATIVE_MAPPINGS,
    materializePreviousCumulative,
  });
})();
