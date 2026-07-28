(() => {
  "use strict";

  const ui = window.PayrollLocal.ui;
  const rules = window.PayrollEngine;

  async function applyWorkbookCellProposal(proposal) {
    const currentTable = await ui.state.workbook.getTable(
      ui.state.mainSheetName,
    );
    const currentPerson = rules
      .buildPeople(currentTable)
      .people.find(
        (person) => person.rowNumber === proposal.person.rowNumber,
      );
    const identity = rules.personIdentity(
      currentPerson || proposal.person,
    );
    const after = rules.computeOperation(
      proposal.currentValue,
      proposal.operation,
      rules.proposalInputValue(proposal),
    );
    const result = await ui.state.workbook.updateCell(
      ui.state.mainSheetName,
      proposal.person.rowNumber,
      proposal.field.column,
      after,
      { preserveFormula: true },
    );
    const fieldName = proposal.field.name;
    const directSyncFields = new Set([
      "身份证",
      "岗位",
      "人员编号",
      "姓名",
      "部门",
      "入司时间",
    ]);
    const syncResults = [];
    if (directSyncFields.has(fieldName)) {
      for (const sheetName of ["工资核对表", "代发薪", "备忘"]) {
        const sync = await ui.state.workbook.updateMatchingRows(
          sheetName,
          identity,
          fieldName,
          after,
        );
        if (sheetName !== "备忘" && !sync.rows.length) {
          throw new Error(
            `${sheetName} 未找到该人员，不能同步 ${fieldName}`,
          );
        }
        syncResults.push(sync);
      }
    } else if (
      rules.normalizeText(fieldName) === rules.normalizeText("实发合计")
    ) {
      syncResults.push(
        await ui.state.workbook.updateMatchingRows(
          "代发薪",
          identity,
          "实发合计",
          after,
        ),
      );
      if (!syncResults[0].updated) {
        throw new Error("代发薪未找到该人员，不能同步实发金额");
      }
    }
    ui.state.workbookSync.push({
      proposalId: proposal.id,
      kind: proposal.kind,
      surfaces: rules.workbookSyncSurfaces(proposal),
      details: syncResults,
    });
    ui.state.history.push({
      time: new Date().toLocaleTimeString("zh-CN"),
      label: `${proposal.person.maskedName} · ${proposal.field.displayName}`,
      detail: `${proposal.operation}：${ui.formatValue(proposal.currentValue)} → ${ui.formatValue(after)}；来源 ${proposal.source}${result.preservedFormula ? "；目标公式已保留，仅更新缓存值" : ""}${syncResults.length ? `；同步 ${syncResults.filter((item) => item.updated).length} 张辅助表` : ""}`,
      kind: "cell-change",
    });
  }

  async function applyWorkbookNewPerson(proposal) {
    if (!proposal.targetRow) {
      throw new Error("模板没有可证明的空白人员行，不能安全新增员工");
    }
    const sourceRow = proposal.sourceTemplateRow;
    const targetRow = proposal.targetRow;
    await ui.state.workbook.cloneEmployeeRow(
      ui.state.mainSheetName,
      sourceRow,
      targetRow,
    );
    for (const [fieldName, value] of Object.entries(
      proposal.newPersonValues || {},
    )) {
      const header = rules.targetHeader(ui.state.table, fieldName);
      if (!header || value === null || value === undefined || value === "") {
        continue;
      }
      await ui.state.workbook.updateCell(
        ui.state.mainSheetName,
        targetRow,
        header.column,
        value,
        { preserveFormula: true },
      );
    }
    const auxiliaryValues = rules.auxiliaryNewPersonValues(proposal);
    const references = await ui.state.workbook.formulaReferencesMainRow(
      "工资核对表",
      ui.state.mainSheetName,
      targetRow,
    );
    if (!references.length) {
      throw new Error(
        `工资核对表没有覆盖工资表第 ${targetRow} 行，不能安全新增员工`,
      );
    }
    const disbursement = await ui.state.workbook.appendMappedPersonRow(
      "代发薪",
      auxiliaryValues,
      { requiredFields: ["账户名称(*)", "账号(*)", "金额(*)"] },
    );
    const memo = await ui.state.workbook.appendMappedPersonRow(
      "备忘",
      auxiliaryValues,
    );
    ui.state.workbookSync.push({
      proposalId: proposal.id,
      kind: proposal.kind,
      surfaces: rules.workbookSyncSurfaces(proposal),
      details: [
        { sheetName: "工资表", rowNumber: targetRow },
        { sheetName: "工资核对表", references: references.length },
        disbursement,
        memo,
      ],
    });
    ui.state.history.push({
      time: new Date().toLocaleTimeString("zh-CN"),
      label: "新增员工",
      detail: `使用工资表保留行 ${targetRow}，并同步工资核对表、代发薪和备忘；未提供字段保持空白。`,
      kind: "new-person",
    });
  }

  async function applyWorkbookDisableProposal(proposal) {
    const identity = rules.personIdentity(proposal.person);
    const archive = await ui.state.workbook.appendMappedPersonRow(
      "离职名单",
      rules.disableArchiveValues(
        proposal,
        ui.state.targetPeriod,
        ui.state.table,
      ),
    );
    const cleared = [];
    for (const sheetName of [
      ui.state.mainSheetName,
      "工资核对表",
      "代发薪",
    ]) {
      const result = await ui.state.workbook.clearIdentityAndHideMatches(
        sheetName,
        identity,
      );
      if (!result.rows.length) {
        throw new Error(`${sheetName} 未找到该人员，不能完成整本同步停用`);
      }
      cleared.push(result);
    }
    ui.state.disabledRows.add(proposal.person.rowNumber);
    ui.state.workbookSync.push({
      proposalId: proposal.id,
      kind: proposal.kind,
      surfaces: rules.workbookSyncSurfaces(proposal),
      details: [archive, ...cleared],
    });
    ui.state.history.push({
      time: new Date().toLocaleTimeString("zh-CN"),
      label: `${proposal.person.maskedName} · 停用`,
      detail:
        "已写入离职名单待补日期记录，并从工资表、工资核对表和代发薪的在职名单移出；未推断离职结算金额。",
      kind: "disable-person",
    });
  }

  async function preflightWorkbookProposals(selected) {
    const validation = rules.validateWorkbookProposalBatch(selected);
    if (validation.errors.length) {
      throw new Error(validation.errors.join("；"));
    }
    const newPeople = selected.filter(
      (proposal) => proposal.kind === "new-person",
    );
    if (newPeople.length) {
      const reserved = await ui.state.workbook.findReservedBlankRows(
        ui.state.mainSheetName,
        ui.state.table,
      );
      if (newPeople.length > reserved.length) {
        throw new Error(
          `本次选择了 ${newPeople.length} 名新增员工，但模板仅有 ${reserved.length} 条可证明的空白人员行`,
        );
      }
      const people = rules.buildPeople(ui.state.table).people;
      const sourceRow = Math.max(
        ...people.map((person) => person.rowNumber),
      );
      for (let index = 0; index < newPeople.length; index += 1) {
        newPeople[index].sourceTemplateRow = sourceRow;
        newPeople[index].targetRow = reserved[index];
        const references =
          await ui.state.workbook.formulaReferencesMainRow(
            "工资核对表",
            ui.state.mainSheetName,
            reserved[index],
          );
        if (!references.length) {
          throw new Error(
            `工资核对表没有覆盖工资表第 ${reserved[index]} 行，不能安全新增员工`,
          );
        }
      }
    }
    for (const proposal of selected) {
      if (
        proposal.kind === "cell-change" &&
        (!proposal.person || !proposal.field)
      ) {
        throw new Error("存在缺少人员或目标字段的变动，已停止整批应用");
      }
      if (proposal.kind === "disable-person" && !proposal.person) {
        throw new Error("存在未匹配人员的停用变动，已停止整批应用");
      }
      if (
        proposal.kind === "cell-change" &&
        rules.isPayImpactingProposal(proposal)
      ) {
        const references =
          await ui.state.workbook.formulaReferencesIdentity(
            "工资核对表",
            ui.state.mainSheetName,
            rules.personIdentity(proposal.person),
          );
        if (!references.rows.length || !references.cells.length) {
          throw new Error(
            "工资核对表未证明该人员与主工资表联动，已停止工资变动",
          );
        }
      }
    }
  }

  Object.assign(ui, {
    applyWorkbookCellProposal,
    applyWorkbookNewPerson,
    applyWorkbookDisableProposal,
    preflightWorkbookProposals,
  });
})();
