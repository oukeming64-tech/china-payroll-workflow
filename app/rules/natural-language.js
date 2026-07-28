(() => {
  "use strict";

  const api = window.PayrollLocal.rules;

  function findPersonMention(line, peopleIndex) {
    const matches = [];
    for (const person of peopleIndex.people) {
      for (const kind of ["employeeId", "idCard", "name"]) {
        const value = api.asText(person[kind]);
        if (value && line.includes(value)) {
          matches.push({ person, kind, length: value.length });
        }
      }
    }
    matches.sort((left, right) => right.length - left.length);
    if (!matches.length) {
      return { status: "unmatched" };
    }
    const longest = matches[0].length;
    const people = [
      ...new Map(
        matches
          .filter((match) => match.length === longest)
          .map((match) => [match.person.rowNumber, match]),
      ).values(),
    ];
    return people.length === 1
      ? { status: "matched", ...people[0] }
      : { status: "ambiguous" };
  }

  function findFieldMentions(line, table) {
    const identityKeys = new Set(
      Object.values(api.IDENTITY_ALIASES).flat().map(api.normalizeText),
    );
    return table.headers
      .filter(
        (header) =>
          header.name &&
          !identityKeys.has(api.normalizeText(header.name)) &&
          line.includes(header.name),
      )
      .sort((left, right) => right.name.length - left.name.length);
  }

  function lastNumericToken(line, excludedValues = []) {
    const excluded = new Set(excludedValues.map(api.asText));
    const matches = [
      ...String(line).matchAll(/[+-]?\d[\d,]*(?:\.\d+)?/g),
    ].map((match) => match[0]);
    for (let index = matches.length - 1; index >= 0; index -= 1) {
      if (excluded.has(matches[index])) {
        continue;
      }
      const numeric = api.asNumber(matches[index]);
      if (numeric !== null) {
        return numeric;
      }
    }
    return null;
  }

  function operationInLine(line) {
    if (/新增|入职/.test(line)) {
      return "新增员工";
    }
    if (/离职|停用|移除|删除/.test(line)) {
      return "停用";
    }
    if (/增加|加发|上调|调增/.test(line)) {
      return "增加";
    }
    if (/减少|扣减|下调|调减/.test(line)) {
      return "减少";
    }
    if (/设置为|调整为|改为|变更为|调到/.test(line)) {
      return "设置";
    }
    return "";
  }

  function parseNewPerson(line, table, base, period, peopleIndex) {
    const data = {};
    const fields = [
      ...table.headers.map((header) => header.name),
      ...(api.ACCOUNT_ALIASES || []),
      ...(api.AMOUNT_ALIASES || []),
    ];
    for (const fieldName of [...new Set(fields.filter(Boolean))]) {
      const escaped = fieldName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const match = line.match(
        new RegExp(`${escaped}\\s*[：:=]?\\s*([^，,\\s]+)`),
      );
      if (match) {
        data[fieldName] = api.asNumber(match[1]) ?? match[1];
      }
    }
    const nameMatch =
      line.match(/姓名\s*[：:=]?\s*([^，,\s]+)/) ||
      line.match(/新增(?:员工)?\s*([^，,\s]+)/);
    const nameHeader = peopleIndex.identities.name;
    if (nameHeader && nameMatch && !data[nameHeader.name]) {
      data[nameHeader.name] = nameMatch[1];
    }
    if (!nameHeader || !api.asText(data[nameHeader.name])) {
      base.status = "error";
      base.errors.push("新增员工缺少姓名");
    }
    Object.assign(base, {
      kind: "new-person",
      operation: "新增员工",
      period,
      newPersonValues: data,
    });
    return base;
  }

  function parseNaturalLanguage(text, table, targetPeriod) {
    const lines = String(text || "")
      .split(/\r?\n|；|;/)
      .map((line) => line.trim())
      .filter(Boolean);
    const peopleIndex = api.indexPeople(table);
    const proposals = [];
    const errors = [];
    for (let lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
      const line = lines[lineIndex];
      const base = api.proposalBase("文字输入", lineIndex + 1, line);
      const period = api.periodInText(line, targetPeriod);
      if (targetPeriod && period && period !== targetPeriod) {
        base.status = "error";
        base.errors.push(
          `生效月份 ${api.formatPeriod(period)} 与目标月份 ${api.formatPeriod(targetPeriod)} 不一致`,
        );
      }
      const operation = operationInLine(line);
      if (operation === "新增员工") {
        proposals.push(
          parseNewPerson(line, table, base, period, peopleIndex),
        );
        continue;
      }
      const personMatch = findPersonMention(line, peopleIndex);
      if (personMatch.status !== "matched") {
        base.status = "error";
        base.errors.push(
          personMatch.status === "ambiguous"
            ? "人员匹配不唯一"
            : "未识别到现有人员",
        );
      } else {
        base.person = personMatch.person;
        base.matchedBy = personMatch.kind;
      }
      if (operation === "停用") {
        Object.assign(base, {
          kind: "disable-person",
          operation: "停用",
          period,
        });
        base.warnings.push(
          "停用会同步离职名单、工资表、工资核对表和代发薪；不自动推断离职结算金额，如有当月结算请另列明字段变动",
        );
        proposals.push(base);
        continue;
      }
      if (!operation) {
        base.status = "error";
        base.errors.push(
          "未识别操作，请使用“设置为 / 增加 / 减少 / 新增 / 停用”",
        );
      }
      const fields = findFieldMentions(line, table);
      if (!fields.length) {
        base.status = "error";
        base.errors.push("未识别目标字段，请使用工资表中的完整字段名");
      } else if (fields.length > 1) {
        base.status = "error";
        base.errors.push("一行识别到多个字段，请拆成一项变动一行");
      }
      const field = fields[0] || null;
      let value = null;
      if (field) {
        const currentValue = base.person?.row.get(field.name);
        if (
          typeof currentValue === "number" ||
          ["增加", "减少"].includes(operation)
        ) {
          value = lastNumericToken(line, [
            base.person?.employeeId,
            base.person?.idCard,
            targetPeriod?.slice(0, 4),
            String(Number(targetPeriod?.slice(5) || 0)),
          ]);
          if (value === null) {
            base.status = "error";
            base.errors.push("未识别到数值");
          }
        } else {
          const escaped = field.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
          const match = line.match(
            new RegExp(
              `${escaped}.*?(?:设置为|调整为|改为|变更为|调到)\\s*([^，,]+)`,
            ),
          );
          value = match?.[1]?.trim() || "";
          if (!value) {
            base.status = "error";
            base.errors.push("未识别到新内容");
          }
        }
      }
      Object.assign(base, {
        kind: "cell-change",
        operation,
        period,
        field,
        inputValue: value,
        currentValue:
          base.person && field ? base.person.row.get(field.name) : null,
      });
      proposals.push(base);
    }
    if (!lines.length) {
      errors.push("请先输入至少一行变动");
    }
    return { proposals, errors };
  }

  Object.assign(api, {
    findPersonMention,
    findFieldMentions,
    parseNaturalLanguage,
  });
})();
