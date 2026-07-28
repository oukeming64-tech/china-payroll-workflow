(() => {
  "use strict";

  const excel = window.PayrollLocal?.excel;
  if (!excel?.XlsxWorkbook) {
    throw new Error("Excel 模块加载顺序不完整");
  }

  window.XlsxEngine = {
    XlsxWorkbook: excel.XlsxWorkbook,
    SourceWorkbook: excel.SourceWorkbook,
    buildTableFromMatrix: excel.buildTableFromMatrix,
    normalizeText: excel.normalizeText,
    parseCellReference: excel.parseCellReference,
    columnLettersToNumber: excel.columnLettersToNumber,
    columnNumberToLetters: excel.columnNumberToLetters,
    translateFormulaA1: excel.translateFormulaA1,
    constants: {
      MAIN_NS: excel.MAIN_NS,
      REL_NS: excel.REL_NS,
      PACKAGE_REL_NS: excel.PACKAGE_REL_NS,
      CONTENT_TYPES_NS: excel.CONTENT_TYPES_NS,
    },
  };
})();
