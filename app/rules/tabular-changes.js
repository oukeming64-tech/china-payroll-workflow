(() => {
  "use strict";

  const api = window.PayrollLocal.rules;

  function parseCsv(text) {
    const rows = [];
    let row = [];
    let value = "";
    let quoted = false;
    const input = String(text || "").replace(/^\uFEFF/, "");
    for (let index = 0; index < input.length; index += 1) {
      const character = input[index];
      if (quoted) {
        if (character === '"' && input[index + 1] === '"') {
          value += '"';
          index += 1;
        } else if (character === '"') {
          quoted = false;
        } else {
          value += character;
        }
      } else if (character === '"') {
        quoted = true;
      } else if (character === ",") {
        row.push(value);
        value = "";
      } else if (character === "\n") {
        row.push(value.replace(/\r$/, ""));
        rows.push(row);
        row = [];
        value = "";
      } else {
        value += character;
      }
    }
    row.push(value.replace(/\r$/, ""));
    if (row.some((cell) => cell !== "") || rows.length === 0) {
      rows.push(row);
    }
    return rows;
  }

  function tableFromMatrix(matrix, sheetName = "CSV") {
    return window.XlsxEngine.buildTableFromMatrix(matrix, sheetName);
  }

  function operationHeader(table) {
    return api.headerForAliases(table, ["操作", "变动类型", "类型"]);
  }

  function fieldHeader(table) {
    return api.headerForAliases(table, ["字段", "目标字段"]);
  }

  function valueHeader(table) {
    return api.headerForAliases(table, ["值", "数值", "新值"]);
  }

  function monthHeader(table) {
    return api.headerForAliases(table, ["月份", "目标月份", "生效月份"]);
  }

  function newPersonProposal(base, sourceTable, sourceRow, period) {
    Object.assign(base, {
      kind: "new-person",
      operation: "新增员工",
      period,
      newPersonValues: Object.fromEntries(
        sourceTable.headers
          .filter((header) => !api.META_HEADERS.has(api.normalizeText(header.name)))
          .map((header) => [header.name, sourceRow.get(header.name)])
          .filter(([, value]) => api.asText(value)),
      ),
    });
    return base;
  }

  function proposalsFromChangeTable(
    sourceTable,
    targetTable,
    targetPeriod,
    sourceName,
  ) {
    const targetIndex = api.indexPeople(targetTable);
    const sourceIdentities = api.identityHeaders(sourceTable);
    const operationColumn = operationHeader(sourceTable);
    const fieldColumn = fieldHeader(sourceTable);
    const valueColumn = valueHeader(sourceTable);
    const monthColumn = monthHeader(sourceTable);
    const isLong = Boolean(operationColumn && fieldColumn && valueColumn);
    const proposals = [];

    for (const sourceRow of sourceTable.rows) {
      const base = api.proposalBase(sourceName, sourceRow.rowNumber, "");
      const period = monthColumn
        ? api.detectPeriod(sourceRow.get(monthColumn.name)) ||
          api.periodInText(sourceRow.get(monthColumn.name), targetPeriod)
        : targetPeriod;
      if (targetPeriod && period && period !== targetPeriod) {
        base.status = "error";
        base.errors.push(
          `生效月份 ${api.formatPeriod(period)} 与目标月份 ${api.formatPeriod(targetPeriod)} 不一致`,
        );
      }
      const identity = api.identityFromRow(sourceRow, sourceIdentities);
      const personMatch = api.matchPerson(targetIndex, identity);
      const operation = api.normalizeOperation(
        operationColumn ? sourceRow.get(operationColumn.name) : "设置",
      ) || "设置";
      if (operation === "新增员工") {
        proposals.push(
          newPersonProposal(base, sourceTable, sourceRow, period),
        );
        continue;
      }
      if (personMatch.status !== "matched") {
        base.status = "error";
        base.errors.push(personMatch.message);
      } else {
        base.person = personMatch.person;
        base.matchedBy = personMatch.matchedBy;
      }
      if (operation === "停用") {
        Object.assign(base, {
          kind: "disable-person",
          operation,
          period,
        });
        base.warnings.push(
          "停用会同步离职名单、工资表、工资核对表和代发薪；不自动推断离职结算金额，如有结算请在变动表另列字段",
        );
        proposals.push(base);
        continue;
      }
      if (isLong) {
        const requestedField = sourceRow.get(fieldColumn.name);
        const field = api.targetHeader(targetTable, requestedField);
        if (!field) {
          base.status = "error";
          base.errors.push(`目标工资表不存在字段“${api.asText(requestedField)}”`);
        }
        Object.assign(base, {
          kind: "cell-change",
          operation,
          period,
          field,
          inputValue: sourceRow.get(valueColumn.name),
          currentValue:
            base.person && field ? base.person.row.get(field.name) : null,
        });
        proposals.push(base);
        continue;
      }
      for (const sourceHeader of sourceTable.headers) {
        if (
          api.META_HEADERS.has(api.normalizeText(sourceHeader.name)) ||
          Object.values(sourceIdentities).includes(sourceHeader)
        ) {
          continue;
        }
        const inputValue = sourceRow.get(sourceHeader.name);
        if (
          inputValue === null ||
          inputValue === undefined ||
          api.asText(inputValue) === ""
        ) {
          continue;
        }
        const item = {
          ...base,
          id: crypto.randomUUID(),
          errors: [...base.errors],
          warnings: [],
        };
        const field = api.targetHeader(targetTable, sourceHeader.name);
        if (!field) {
          item.status = "error";
          item.errors.push(`目标工资表不存在字段“${sourceHeader.name}”`);
        }
        Object.assign(item, {
          kind: "cell-change",
          operation: "设置",
          period,
          field,
          inputValue,
          currentValue:
            item.person && field ? item.person.row.get(field.name) : null,
        });
        proposals.push(item);
      }
    }
    return {
      format: isLong ? "long" : "wide",
      proposals,
      errors: sourceTable.rows.length ? [] : ["变动表没有可读取的数据行"],
    };
  }

  Object.assign(api, {
    parseCsv,
    tableFromMatrix,
    proposalsFromChangeTable,
  });
})();
