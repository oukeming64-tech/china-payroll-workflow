(() => {
  "use strict";

  const api = window.PayrollLocal.rules;
  const IDENTITY_ALIASES = {
    employeeId: [
      "人员编号",
      "员工编号",
      "工号",
      "编号",
      "职工编号",
      "雇员编号",
    ],
    idCard: [
      "身份证",
      "身份证号",
      "身份证号码",
      "证件号",
      "证件号码",
      "身份证件号码",
      "证照号码",
      "纳税人识别号",
    ],
    name: ["姓名", "员工姓名", "职工姓名", "纳税人姓名", "人员姓名"],
    department: ["部门", "所属部门", "单位", "组织", "组织名称"],
  };
  const META_HEADERS = new Set(
    [
      "月份",
      "目标月份",
      "生效月份",
      "操作",
      "变动类型",
      "类型",
      "字段",
      "目标字段",
      "值",
      "数值",
      "新值",
      "备注",
      "说明",
      "来源",
    ].map(normalizeText),
  );

  function normalizeText(value) {
    return String(value ?? "")
      .normalize("NFKC")
      .trim()
      .toLowerCase()
      .replace(/[\s_\-—–·:：()（）[\]【】/\\]+/g, "");
  }

  function asText(value) {
    return value === null || value === undefined ? "" : String(value).trim();
  }

  function asNumber(value) {
    if (typeof value === "number" && Number.isFinite(value)) {
      return value;
    }
    const normalized = asText(value).replaceAll(",", "").replace(/[元￥¥]/g, "");
    if (!normalized || !/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(normalized)) {
      return null;
    }
    const result = Number(normalized);
    return Number.isFinite(result) ? result : null;
  }

  function nextPeriod(period) {
    const match = String(period || "").match(/^(\d{4})-(0?[1-9]|1[0-2])$/);
    if (!match) {
      return "";
    }
    const date = new Date(Number(match[1]), Number(match[2]) - 1, 1);
    date.setMonth(date.getMonth() + 1);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  }

  function detectPeriod(...values) {
    for (const value of values) {
      const text = asText(value);
      const match =
        text.match(/(20\d{2})\s*年\s*(0?[1-9]|1[0-2])\s*月/) ||
        text.match(/(20\d{2})\s*[./-]\s*(0?[1-9]|1[0-2])(?!\d)/) ||
        text.match(/(20\d{2})(0[1-9]|1[0-2])(?=[^\d]|$)/);
      if (match) {
        return `${match[1]}-${String(Number(match[2])).padStart(2, "0")}`;
      }
    }
    return "";
  }

  function formatPeriod(period) {
    const match = String(period || "").match(/^(\d{4})-(0?[1-9]|1[0-2])$/);
    return match ? `${match[1]}年${Number(match[2])}月` : "未指定";
  }

  function headerForAliases(table, aliases) {
    const keys = new Set(aliases.map(normalizeText));
    return (
      table.headers.find((header) => keys.has(normalizeText(header.name))) || null
    );
  }

  function identityHeaders(table) {
    return {
      employeeId: headerForAliases(table, IDENTITY_ALIASES.employeeId),
      idCard: headerForAliases(table, IDENTITY_ALIASES.idCard),
      name: headerForAliases(table, IDENTITY_ALIASES.name),
      department: headerForAliases(table, IDENTITY_ALIASES.department),
    };
  }

  function maskName(value) {
    const text = asText(value);
    if (!text) {
      return "未命名";
    }
    return text.length === 1
      ? `${text}*`
      : `${text[0]}${"*".repeat(Math.min(3, text.length - 1))}`;
  }

  function maskIdentifier(value) {
    const text = asText(value);
    if (text.length <= 4) {
      return text ? `${text[0]}***` : "—";
    }
    return `${text.slice(0, 2)}${"*".repeat(Math.min(10, text.length - 4))}${text.slice(-2)}`;
  }

  function buildPeople(table) {
    const identities = identityHeaders(table);
    const people = [];
    for (const row of table.rows) {
      if (row.hidden) {
        continue;
      }
      const employeeId = identities.employeeId
        ? row.get(identities.employeeId.name)
        : "";
      const idCard = identities.idCard ? row.get(identities.idCard.name) : "";
      const name = identities.name ? row.get(identities.name.name) : "";
      const department = identities.department
        ? row.get(identities.department.name)
        : "";
      if (![employeeId, idCard, name].some((value) => asText(value))) {
        continue;
      }
      people.push({
        row,
        rowNumber: row.rowNumber,
        employeeId,
        idCard,
        name,
        department,
        maskedName: maskName(name),
        maskedEmployeeId: maskIdentifier(employeeId),
        maskedIdCard: maskIdentifier(idCard),
      });
    }
    return { people, identities };
  }

  function indexPeople(table) {
    const { people, identities } = buildPeople(table);
    const maps = {
      employeeId: new Map(),
      idCard: new Map(),
      name: new Map(),
    };
    for (const person of people) {
      for (const kind of ["employeeId", "idCard", "name"]) {
        const key = normalizeText(person[kind]);
        if (!key) {
          continue;
        }
        if (!maps[kind].has(key)) {
          maps[kind].set(key, []);
        }
        maps[kind].get(key).push(person);
      }
    }
    return { people, identities, maps };
  }

  function matchPerson(index, identity = {}, options = {}) {
    const matchBy = options.matchBy || ["employeeId", "idCard", "name"];
    for (const kind of matchBy) {
      const key = normalizeText(identity[kind]);
      if (!key) {
        continue;
      }
      const candidates = index.maps[kind].get(key) || [];
      if (candidates.length === 1) {
        return { status: "matched", person: candidates[0], matchedBy: kind };
      }
      if (candidates.length > 1) {
        return {
          status: "ambiguous",
          matchedBy: kind,
          candidates,
          message: `${kind === "name" ? "姓名" : "人员标识"}存在重名或重复`,
        };
      }
      if (options.stopAfterPresent) {
        const label = { idCard: "身份证号", employeeId: "人员编号", name: "姓名" }[kind];
        return { status: "unmatched", matchedBy: kind, message:
          `${label}未匹配到人员，已禁止改用其他字段兜底` };
      }
    }
    return { status: "unmatched", message: "未匹配到人员" };
  }

  function identityFromRow(row, identities) {
    return {
      employeeId: identities.employeeId
        ? row.get(identities.employeeId.name)
        : "",
      idCard: identities.idCard ? row.get(identities.idCard.name) : "",
      name: identities.name ? row.get(identities.name.name) : "",
    };
  }

  function targetHeader(table, name) {
    const key = normalizeText(name);
    return key
      ? table.headers.find((header) => normalizeText(header.name) === key) || null
      : null;
  }

  function normalizeOperation(value) {
    const key = normalizeText(value);
    if (["设置", "设置为", "改为", "调整为", "变更为"].includes(key)) {
      return "设置";
    }
    if (["增加", "加", "加发", "上调", "调增"].includes(key)) {
      return "增加";
    }
    if (["减少", "减", "扣减", "下调", "调减"].includes(key)) {
      return "减少";
    }
    if (["新增", "新增员工", "入职", "新入职"].includes(key)) {
      return "新增员工";
    }
    if (["离职", "停用", "删除", "移除"].includes(key)) {
      return "停用";
    }
    return "";
  }

  function periodInText(text, targetPeriod) {
    const explicit = detectPeriod(text);
    if (explicit) {
      return explicit;
    }
    const monthOnly = String(text).match(/(?:^|[^\d])([1-9]|1[0-2])\s*月/);
    return monthOnly && targetPeriod
      ? `${targetPeriod.slice(0, 4)}-${String(Number(monthOnly[1])).padStart(2, "0")}`
      : targetPeriod;
  }

  function proposalBase(source, sourceRow, raw) {
    return {
      id: crypto.randomUUID(),
      source,
      sourceRow,
      raw,
      status: "ready",
      selected: true,
      warnings: [],
      errors: [],
    };
  }

  function computeOperation(currentValue, operation, inputValue) {
    if (operation === "设置") {
      const numeric = asNumber(inputValue);
      return typeof currentValue === "number" && numeric !== null
        ? numeric
        : inputValue;
    }
    const current = asNumber(currentValue);
    const delta = asNumber(inputValue);
    if (current === null || delta === null) {
      throw new Error("增加或减少操作要求当前值与变动值均为数字");
    }
    return operation === "增加" ? current + delta : current - delta;
  }

  function sameValue(left, right) {
    const leftNumber = asNumber(left);
    const rightNumber = asNumber(right);
    if (leftNumber !== null && rightNumber !== null)
      return Math.abs(leftNumber - rightNumber) < 0.001;
    return normalizeText(left) === normalizeText(right);
  }

  function markConflicts(proposals) {
    const conflictMessage = "同一人员同一字段出现多项不同变动，请删除冲突后再确认";
    const duplicateMessage = "与同批其他附件的值相同，已合并为一项";
    for (const proposal of proposals) {
      proposal.errors = (proposal.errors || []).filter((error) =>
        error !== conflictMessage);
      proposal.warnings = (proposal.warnings || []).filter((warning) =>
        warning !== duplicateMessage);
      proposal.redundant = false;
      if (proposal.status === "error" && !proposal.errors.length) {
        proposal.status = "ready";
      }
    }
    const groups = new Map();
    for (const proposal of proposals) {
      if (
        proposal.status === "error" ||
        proposal.kind !== "cell-change" ||
        !proposal.person ||
        !proposal.field
      ) {
        continue;
      }
      const key = `${proposal.person.rowNumber}:${proposal.field.column}`;
      if (!groups.has(key)) {
        groups.set(key, []);
      }
      groups.get(key).push(proposal);
    }
    for (const items of groups.values()) {
      if (items.length <= 1) {
        continue;
      }
      const identical =
        items.every((item) => item.operation === "设置") &&
        items.every((item) => sameValue(item.inputValue, items[0].inputValue));
      if (identical) {
        items.slice(1).forEach((item) => {
          item.selected = false;
          item.redundant = true;
          item.warnings.push(duplicateMessage);
        });
        continue;
      }
      for (const item of items) {
        item.status = "error";
        item.selected = false;
        item.errors.push(conflictMessage);
      }
    }
    return proposals;
  }

  function classifySource(filename) {
    const normalized = normalizeText(filename);
    if (normalized.includes("正常工资薪金所得")) {
      return "个税工资薪金附件";
    }
    if (
      normalized.includes("保险表") ||
      normalized.includes("社保") ||
      normalized.includes("公积金")
    ) {
      return "社保 / 公积金附件";
    }
    if (normalized.includes("劳务费")) {
      return "劳务费附件";
    }
    if (normalized.includes("工资")) {
      return "历史工资表 / 工资模板";
    }
    return "未分类表格";
  }

  Object.assign(api, {
    IDENTITY_ALIASES,
    META_HEADERS,
    normalizeText,
    asText,
    asNumber,
    nextPeriod,
    detectPeriod,
    formatPeriod,
    headerForAliases,
    identityHeaders,
    buildPeople,
    indexPeople,
    matchPerson,
    identityFromRow,
    targetHeader,
    normalizeOperation,
    periodInText,
    proposalBase,
    computeOperation,
    sameValue,
    markConflicts,
    classifySource,
    maskName,
    maskIdentifier,
  });
})();
