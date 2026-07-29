(() => {
  "use strict";

  const api = window.PayrollLocal.rules;
  const NON_PAY_FIELDS = new Set(
    [
      "序号",
      "身份证",
      "岗位",
      "人员编号",
      "姓名",
      "部门",
      "毕业院校",
      "学历",
      "入司时间",
      "TRS司龄",
      "首次参加工作时间",
      "工作年限",
    ].map(api.normalizeText),
  );
  const ACCOUNT_ALIASES = [
    "账号(*)",
    "账号",
    "工资卡号",
    "银行卡号",
    "新工资卡号",
  ];
  const AMOUNT_ALIASES = ["金额(*)", "金额", "实发合计"];

  function valueByAliases(values, aliases) {
    const keys = new Set(aliases.map(api.normalizeText));
    const match = Object.entries(values || {}).find(([name]) =>
      keys.has(api.normalizeText(name)),
    );
    return match?.[1];
  }

  function proposalPersonKey(proposal) {
    if (proposal.person?.rowNumber) {
      return `row:${proposal.person.rowNumber}`;
    }
    const values = proposal.newPersonValues || {};
    for (const aliases of Object.values(api.IDENTITY_ALIASES)) {
      const value = valueByAliases(values, aliases);
      if (api.asText(value)) {
        return `new:${api.normalizeText(value)}`;
      }
    }
    return `proposal:${proposal.id}`;
  }

  function isPayImpactingProposal(proposal) {
    return (
      proposal.kind === "cell-change" &&
      proposal.field &&
      !NON_PAY_FIELDS.has(api.normalizeText(proposal.field.name)) &&
      api.normalizeText(proposal.field.name) !== api.normalizeText("实发合计")
    );
  }

  function isFinalPayProposal(proposal) {
    return (
      proposal.kind === "cell-change" &&
      api.normalizeText(proposal.field?.name) ===
        api.normalizeText("实发合计")
    );
  }

  function proposalInputValue(proposal) {
    if (proposal.operation !== "设置") {
      return proposal.inputValue;
    }
    const numeric = api.asNumber(proposal.inputValue);
    if (numeric === null) {
      return proposal.inputValue;
    }
    const isFinalAmount =
      api.normalizeText(proposal.field?.name) ===
      api.normalizeText("实发合计");
    return (
      typeof proposal.currentValue === "number" ||
      isFinalAmount ||
      isPayImpactingProposal(proposal)
    )
      ? numeric
      : proposal.inputValue;
  }

  function proposalSourceAccount(proposals) {
    for (const proposal of proposals) {
      const account = valueByAliases(
        proposal.sourceValues,
        ACCOUNT_ALIASES,
      );
      if (api.asText(account)) {
        return account;
      }
    }
    return "";
  }

  function validateWorkbookProposalBatch(proposals) {
    const errors = [];
    const selected = proposals.filter(
      (proposal) => proposal.selected && proposal.status !== "error",
    );
    const byPerson = new Map();
    for (const proposal of selected) {
      const key = proposalPersonKey(proposal);
      if (!byPerson.has(key)) {
        byPerson.set(key, []);
      }
      byPerson.get(key).push(proposal);
    }
    for (const [key, items] of byPerson) {
      for (const item of items.filter(
        (proposal) => proposal.kind === "new-person",
      )) {
        const account = valueByAliases(
          item.newPersonValues,
          ACCOUNT_ALIASES,
        );
        const amount = valueByAliases(
          item.newPersonValues,
          AMOUNT_ALIASES,
        );
        if (!api.asText(account)) {
          errors.push("新增员工缺少银行账号，不能同步代发薪");
        }
        if (api.asNumber(amount) === null) {
          errors.push("新增员工缺少明确的实发金额，不能同步代发薪");
        }
        const bm = valueByAliases(item.newPersonValues, ["BM津贴"]);
        if (api.asNumber(bm) > 0) {
          errors.push(
            "新增员工包含 BM 津贴，但 BM津贴 工作表没有可验证的人员键，已停止自动处理",
          );
        }
      }
    }
    return { selected, errors };
  }

  function personIdentity(person) {
    return {
      employeeId: person?.employeeId || "",
      idCard: person?.idCard || "",
      name: person?.name || "",
    };
  }

  function rowValuesByHeader(table, row) {
    return Object.fromEntries(
      table.headers.map((header) => [
        header.name,
        row.values.get(header.column),
      ]),
    );
  }

  function disableArchiveValues(proposal, targetPeriod, table) {
    const values = rowValuesByHeader(table, proposal.person.row);
    values.备注 = `月度停用：${api.formatPeriod(targetPeriod)}；离职日期待输入`;
    return values;
  }

  function auxiliaryNewPersonValues(proposal) {
    const values = { ...(proposal.newPersonValues || {}) };
    const name = valueByAliases(values, api.IDENTITY_ALIASES.name);
    if (api.asText(name) && !api.asText(valueByAliases(values, ["账户名称(*)"]))) {
      values["账户名称(*)"] = name;
    }
    const account = valueByAliases(values, ACCOUNT_ALIASES);
    if (api.asText(account)) {
      values["账号(*)"] = account;
      values.新工资卡号 = account;
    }
    const amount = valueByAliases(values, AMOUNT_ALIASES);
    if (api.asNumber(amount) !== null) {
      values["金额(*)"] = api.asNumber(amount);
    }
    return values;
  }

  function workbookSyncSurfaces(proposal) {
    if (proposal.kind === "disable-person") {
      return [
        "工资表移出在职名单",
        "工资核对表移出匹配行",
        "代发薪移出匹配行",
        proposal.archiveExisting
          ? "离职名单沿用已有记录"
          : "离职名单新增待补日期记录",
      ];
    }
    if (proposal.kind === "new-person") {
      return [
        "工资表新增模板行",
        "工资核对表验证公式覆盖",
        "代发薪新增明确账户与金额",
        "备忘新增人员记录",
      ];
    }
    if (proposal.kind === "cell-change") {
      return isFinalPayProposal(proposal)
        ? ["工资表更新", "工资核对表公式联动", "代发薪金额同步"]
        : isPayImpactingProposal(proposal)
          ? [
              "工资表更新",
              "工资核对表公式联动",
              "代发薪引用实发公式",
            ]
          : ["工资表更新", "工资核对表公式联动"];
    }
    return [];
  }

  Object.assign(api, {
    ACCOUNT_ALIASES,
    AMOUNT_ALIASES,
    valueByAliases,
    proposalPersonKey,
    isPayImpactingProposal,
    isFinalPayProposal,
    proposalInputValue,
    proposalSourceAccount,
    validateWorkbookProposalBatch,
    personIdentity,
    rowValuesByHeader,
    disableArchiveValues,
    auxiliaryNewPersonValues,
    workbookSyncSurfaces,
  });
})();
