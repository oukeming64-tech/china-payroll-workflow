(() => {
  "use strict";

  const api = window.PayrollLocal.excel;

  function legacyMatrix(worksheet) {
    return window.XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: true,
      defval: null,
      blankrows: true,
    });
  }

  function legacyMetadata(worksheet, matrix) {
    const metadata = matrix.map((row) =>
      Array.from({ length: row.length }, () => null),
    );
    const range = window.XLSX.utils.decode_range(
      worksheet["!ref"] || "A1:A1",
    );
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      if (!metadata[row]) {
        metadata[row] = [];
      }
      for (let column = range.s.c; column <= range.e.c; column += 1) {
        const address = window.XLSX.utils.encode_cell({ r: row, c: column });
        const cell = worksheet[address];
        if (!cell) {
          continue;
        }
        metadata[row][column] = {
          address,
          formula: cell.f || "",
          type: cell.t || "",
        };
      }
    }
    return metadata;
  }

  class LegacySourceWorkbook {
    constructor(fileName, workbook) {
      this.fileName = fileName;
      this.legacy = true;
      this.workbook = workbook;
      this.sheets = workbook.SheetNames.map((name) => ({ name }));
      this.sheetByName = new Map(
        this.sheets.map((sheet) => [sheet.name, sheet]),
      );
      this.tableCache = new Map();
    }

    async getTable(sheetName) {
      if (!this.sheetByName.has(sheetName)) {
        throw new Error(`找不到工作表：${sheetName}`);
      }
      if (this.tableCache.has(sheetName)) {
        return this.tableCache.get(sheetName);
      }
      const worksheet = this.workbook.Sheets[sheetName];
      const matrix = legacyMatrix(worksheet);
      const table = api.buildTableFromMatrix(matrix, sheetName, {
        metadata: legacyMetadata(worksheet, matrix),
      });
      this.tableCache.set(sheetName, table);
      return table;
    }

    async scoreSheetsAgainst(targetHeaders = []) {
      const targetKeys = new Set(
        targetHeaders.map((header) =>
          api.normalizeText(header.name || header),
        ),
      );
      const output = [];
      for (const sheet of this.sheets) {
        const table = await this.getTable(sheet.name);
        const headerKeys = table.headers.map((header) =>
          api.normalizeText(header.name),
        );
        const exactMatches = headerKeys.filter((key) =>
          targetKeys.has(key),
        ).length;
        const identityMatches = headerKeys.filter((key) =>
          api.IDENTITY_HEADER_KEYS.has(key),
        ).length;
        output.push({
          name: sheet.name,
          score:
            exactMatches * 4 +
            identityMatches * 20 +
            Math.min(table.rows.length, 5),
          headerRow: table.headerRow,
          rows: table.rows.length,
          columns: table.headers.length,
          formulas: table.formulaCount,
        });
      }
      return output.sort((left, right) => right.score - left.score);
    }
  }

  class SourceWorkbook {
    static async load(file) {
      if (!file) {
        throw new Error("没有选择来源文件");
      }
      const prefix = new Uint8Array(
        await file.slice(0, 8).arrayBuffer(),
      );
      const isCompoundFile =
        prefix.length === 8 &&
        [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every(
          (value, index) => prefix[index] === value,
        );
      const legacyExtension = /\.xls$/i.test(file.name || "");
      if (!legacyExtension && !isCompoundFile) {
        return api.XlsxWorkbook.load(file);
      }
      if (!window.XLSX?.read) {
        throw new Error("旧版 Excel 读取组件未加载");
      }
      const bytes = await file.arrayBuffer();
      if (bytes.byteLength > api.MAX_WORKBOOK_BYTES) {
        throw new Error("来源工作簿超过 100 MB，已停止读取");
      }
      let workbook;
      try {
        workbook = window.XLSX.read(bytes, {
          type: "array",
          cellFormula: true,
          cellDates: false,
          cellNF: true,
          dense: false,
        });
      } catch (error) {
        if (isCompoundFile && !legacyExtension) {
          throw new Error(
            `${file.name} 已加密；请先在 Excel 输入密码并另存为无密码 .xlsx 后再选择`,
          );
        }
        throw new Error(`旧版 Excel 无法读取：${error.message}`);
      }
      if (!workbook.SheetNames?.length) {
        throw new Error("旧版 Excel 没有可读取的工作表");
      }
      return new LegacySourceWorkbook(file.name, workbook);
    }
  }

  Object.assign(api, {
    LegacySourceWorkbook,
    SourceWorkbook,
  });
})();
