(() => {
  "use strict";

  const api = window.PayrollLocal.excel;
  const MAX_WORKBOOK_BYTES = 100 * 1024 * 1024;
  const MAX_ZIP_ENTRIES = 5000;

  class XlsxWorkbook {
    constructor(options) {
      Object.assign(this, options);
      this.sheetByName = new Map(
        this.sheets.map((sheet) => [sheet.name, sheet]),
      );
      this.sheetCache = new Map();
      this.dirtySheetPaths = new Set();
      this.externalDiagnostics = null;
    }

    static async load(fileOrBuffer, fileName = "workbook.xlsx") {
      if (!window.JSZip) {
        throw new Error("本地 Excel 引擎未加载，请重新打开页面");
      }
      const bytes =
        fileOrBuffer instanceof ArrayBuffer
          ? new Uint8Array(fileOrBuffer)
          : fileOrBuffer instanceof Uint8Array
            ? fileOrBuffer
            : new Uint8Array(await fileOrBuffer.arrayBuffer());
      const resolvedName = fileOrBuffer?.name || fileName;
      if (bytes.byteLength > MAX_WORKBOOK_BYTES) {
        throw new Error("工作簿超过 100 MB，本地浏览器为安全起见未打开");
      }
      const zip = await window.JSZip.loadAsync(bytes);
      if (Object.keys(zip.files).length > MAX_ZIP_ENTRIES) {
        throw new Error("工作簿内部文件数量异常，本地浏览器为安全起见未打开");
      }
      const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
      const relationshipsXml = await zip
        .file("xl/_rels/workbook.xml.rels")
        ?.async("string");
      if (!workbookXml || !relationshipsXml) {
        throw new Error("不是可识别的 Excel OOXML 工作簿");
      }
      const workbookDocument = api.parseXml(workbookXml, "workbook.xml");
      const workbookRelationshipsDocument = api.parseXml(
        relationshipsXml,
        "workbook.xml.rels",
      );
      const workbookRelationships = api.relationshipMap(
        workbookRelationshipsDocument,
        "xl/workbook.xml",
      );
      const sharedStringsXml = await zip
        .file("xl/sharedStrings.xml")
        ?.async("string");
      const sharedStrings = api.readSharedStrings(
        sharedStringsXml
          ? api.parseXml(sharedStringsXml, "sharedStrings.xml")
          : null,
      );
      const sheets = api
        .elementsByLocalName(workbookDocument, "sheet")
        .map((sheetNode) => {
          const relationshipId =
            sheetNode.getAttributeNS(api.REL_NS, "id") ||
            sheetNode.getAttribute("r:id") ||
            "";
          const relationship = workbookRelationships.get(relationshipId);
          return {
            name: sheetNode.getAttribute("name") || "",
            sheetId: sheetNode.getAttribute("sheetId") || "",
            relationshipId,
            path: relationship?.resolvedTarget || "",
          };
        });
      if (!sheets.length || sheets.some((sheet) => !sheet.path)) {
        throw new Error("工作簿的工作表关系不完整");
      }
      return new XlsxWorkbook({
        fileName: resolvedName,
        fileSize: bytes.byteLength,
        originalBytes: bytes.slice(),
        zip,
        workbookDocument,
        workbookRelationshipsDocument,
        workbookRelationships,
        sharedStrings,
        sheets,
      });
    }

    async loadSheetRecord(sheetName) {
      const sheet = this.sheetByName.get(sheetName);
      if (!sheet) {
        throw new Error(`找不到工作表：${sheetName}`);
      }
      if (this.sheetCache.has(sheet.path)) {
        return this.sheetCache.get(sheet.path);
      }
      const xml = await this.zip.file(sheet.path)?.async("string");
      if (!xml) {
        throw new Error(`工作表文件缺失：${sheetName}`);
      }
      const record = {
        ...sheet,
        document: api.parseXml(xml, sheet.path),
      };
      this.sheetCache.set(sheet.path, record);
      return record;
    }

    async getTable(sheetName, options = {}) {
      const sheetRecord = await this.loadSheetRecord(sheetName);
      return api.buildTableFromSheet(
        sheetRecord,
        this.sharedStrings,
        options,
      );
    }

    async getCell(sheetName, rowNumber, columnNumber) {
      return api.getCell(this, sheetName, rowNumber, columnNumber);
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

    async analyzeExternalLinks() {
      return api.analyzeExternalLinks(this);
    }

    async detachExternalLinks(mainSheetName, allowedMainRows = []) {
      return api.detachExternalLinks(
        this,
        mainSheetName,
        allowedMainRows,
      );
    }

    async analyzeFormulas(sheetName) {
      return api.analyzeFormulas(this, sheetName);
    }

    async updateCell(
      sheetName,
      rowNumber,
      columnNumber,
      value,
      options = {},
    ) {
      return api.updateCell(
        this,
        sheetName,
        rowNumber,
        columnNumber,
        value,
        options,
      );
    }

    async findMonthMarkers(period) {
      return api.findMonthMarkers(this, period);
    }

    async findPeriodCandidates(sheetNames = []) {
      return api.findPeriodCandidates(this, sheetNames);
    }

    async updateMonthMarkers(markers, fromPeriod, toPeriod) {
      return api.updateMonthMarkers(
        this,
        markers,
        fromPeriod,
        toPeriod,
      );
    }

    async addSchemaFields(sheetName, insertColumn, fieldNames) {
      return api.addSchemaFields(
        this,
        sheetName,
        insertColumn,
        fieldNames,
      );
    }

    async clearFieldsForPeople(sheetName, table, fieldNames) {
      return api.clearFieldsForPeople(
        this,
        sheetName,
        table,
        fieldNames,
      );
    }

    async applyJanuaryInternalFormulas(
      sheetName,
      table,
      nameOnlyRows,
    ) {
      return api.applyJanuaryInternalFormulas(
        this,
        sheetName,
        table,
        nameOnlyRows,
      );
    }

    async materializePreviousCumulative(sheetName, table) {
      return api.materializePreviousCumulative(
        this,
        sheetName,
        table,
      );
    }

    async analyzePersonnelSheets() {
      return api.analyzePersonnelSheets(this);
    }

    async appendMappedPersonRow(sheetName, values, options = {}) {
      return api.appendMappedPersonRow(
        this,
        sheetName,
        values,
        options,
      );
    }

    async clearIdentityAndHideMatches(sheetName, identity) {
      return api.clearIdentityAndHideMatches(
        this,
        sheetName,
        identity,
      );
    }

    async updateMatchingRows(
      sheetName,
      identity,
      fieldName,
      value,
    ) {
      return api.updateMatchingRows(
        this,
        sheetName,
        identity,
        fieldName,
        value,
      );
    }

    async formulaReferencesMainRow(sheetName, mainSheetName, rowNumber) {
      return api.formulaReferencesMainRow(
        this,
        sheetName,
        mainSheetName,
        rowNumber,
      );
    }

    async formulaReferencesIdentity(
      sheetName,
      mainSheetName,
      identity,
    ) {
      return api.formulaReferencesIdentity(
        this,
        sheetName,
        mainSheetName,
        identity,
      );
    }

    async findReservedBlankRows(sheetName, table = null) {
      return api.findReservedBlankRows(this, sheetName, table);
    }

    async cloneEmployeeRow(sheetName, sourceRowNumber, targetRowNumber) {
      return api.cloneEmployeeRow(
        this,
        sheetName,
        sourceRowNumber,
        targetRowNumber,
      );
    }

    async setRowHidden(sheetName, rowNumber, hidden = true) {
      return api.setRowHidden(this, sheetName, rowNumber, hidden);
    }

    async flushToZip() {
      for (const sheetPath of this.dirtySheetPaths) {
        const sheetRecord = this.sheetCache.get(sheetPath);
        if (sheetRecord) {
          this.zip.file(
            sheetPath,
            api.serializeXml(sheetRecord.document),
          );
        }
      }
      api.setCalcMode(this.workbookDocument);
      this.zip.file(
        "xl/workbook.xml",
        api.serializeXml(this.workbookDocument),
      );
    }

    async export() {
      await this.flushToZip();
      const bytes = await this.zip.generateAsync({
        type: "uint8array",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
      });
      return {
        blob: new Blob([bytes], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
      };
    }
  }

  Object.assign(api, {
    MAX_WORKBOOK_BYTES,
    MAX_ZIP_ENTRIES,
    XlsxWorkbook,
  });
})();
