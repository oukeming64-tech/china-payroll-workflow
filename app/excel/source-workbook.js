(() => {
  "use strict";

  const api = window.PayrollLocal.excel;
  const PASSWORD_REQUIRED = "WORKBOOK_PASSWORD_REQUIRED";
  const PASSWORD_INCORRECT = "WORKBOOK_PASSWORD_INCORRECT";

  function workbookPasswordError(code, fileName) {
    const error = new Error(
      code === PASSWORD_INCORRECT
        ? `${fileName} 密码不正确`
        : `${fileName} 已加密，需要输入打开密码`,
    );
    error.code = code;
    return error;
  }

  function hasZipSignature(bytes) {
    return bytes.length >= 4 &&
      bytes[0] === 0x50 &&
      bytes[1] === 0x4b &&
      [0x03, 0x05, 0x07].includes(bytes[2]) &&
      [0x04, 0x06, 0x08].includes(bytes[3]);
  }

  async function decryptWorkbook(bytes, fileName, password) {
    if (
      !window.OfficeCrypto?.isEncrypted ||
      !window.OfficeCrypto?.decrypt
    ) {
      throw new Error("本地工作簿解密组件未加载，请重新打开页面");
    }
    let encrypted;
    try {
      encrypted = window.OfficeCrypto.isEncrypted(bytes);
    } catch (error) {
      throw new Error(`${fileName} 的加密结构无法识别：${error.message}`);
    }
    if (!encrypted) {
      return null;
    }
    if (!password) {
      throw workbookPasswordError(PASSWORD_REQUIRED, fileName);
    }
    try {
      const decrypted = await window.OfficeCrypto.decrypt(bytes, {
        password,
      });
      const view = decrypted instanceof Uint8Array
        ? decrypted
        : new Uint8Array(decrypted);
      return view.slice();
    } catch (error) {
      if (/password is incorrect/i.test(error.message || "")) {
        throw workbookPasswordError(PASSWORD_INCORRECT, fileName);
      }
      throw new Error(`${fileName} 解密失败：${error.message}`);
    }
  }

  function readLegacyWorkbook(bytes, fileName) {
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
      throw new Error(`旧版 Excel 无法读取：${error.message}`);
    }
    if (!workbook.SheetNames?.length) {
      throw new Error("旧版 Excel 没有可读取的工作表");
    }
    return new LegacySourceWorkbook(fileName, workbook);
  }

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
    static async load(file, options = {}) {
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
      const decrypted = await decryptWorkbook(
        new Uint8Array(bytes),
        file.name,
        options.password,
      );
      if (decrypted && hasZipSignature(decrypted)) {
        return api.XlsxWorkbook.load(decrypted, file.name);
      }
      return readLegacyWorkbook(decrypted || bytes, file.name);
    }
  }

  Object.assign(api, {
    LegacySourceWorkbook,
    PASSWORD_INCORRECT,
    PASSWORD_REQUIRED,
    SourceWorkbook,
  });
})();
