(() => {
  "use strict";

  const api = window.PayrollLocal.rules;
  const EXIT_DATE_ALIASES = Object.freeze([
    "离职日期",
    "离司日期",
    "离职时间",
  ]);
  const ENTRY_DATE_ALIASES = Object.freeze([
    "入职日期",
    "入司日期",
    "入职时间",
  ]);
  const REGULARIZATION_DATE_ALIASES = Object.freeze([
    "转正日期",
  ]);
  const SALARY_DATE_ALIASES = Object.freeze([
    "调薪日期",
    "生效日期",
  ]);
  const NOTE_ALIASES = Object.freeze(["备注", "说明"]);
  const WORKBOOK_EVIDENCE_RULE_META = Object.freeze({
    id: "base-workbook-business-evidence",
    trigger: "读取上月完整工资表后",
    policy:
      "先遍历整本工作簿及多表区，吸收离职名单、备忘等分表中的明确生效记录，再核对目标月附件；已在目标月前生效的离职不得继续报为附件缺人。",
  });

  function header(table, aliases) {
    return api.headerForAliases(table, aliases);
  }

  function dateText(date) {
    return date
      ? [
          date.getUTCFullYear(),
          String(date.getUTCMonth() + 1).padStart(2, "0"),
          String(date.getUTCDate()).padStart(2, "0"),
        ].join("-")
      : "";
  }

  function targetStart(period) {
    const match = String(period || "").match(
      /^(\d{4})-(0[1-9]|1[0-2])$/,
    );
    return match
      ? new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, 1))
      : null;
  }

  function identityMatch(targetIndex, table, row) {
    const identities = api.identityHeaders(table);
    const identity = api.identityFromRow(row, identities);
    const kind = identities.idCard
      ? "idCard"
      : identities.employeeId
        ? "employeeId"
        : identities.name
          ? "name"
          : "";
    if (!kind || !api.asText(identity[kind])) {
      return { identity, match: { status: "unmatched" }, matchedBy: "" };
    }
    const match = api.matchPerson(targetIndex, identity, {
      matchBy: [kind],
      stopAfterPresent: true,
    });
    return { identity, match, matchedBy: match.matchedBy || kind };
  }

  function regionTopics(table) {
    const topics = [];
    if (header(table, EXIT_DATE_ALIASES)) {
      topics.push("离职");
    }
    if (header(table, ENTRY_DATE_ALIASES)) {
      topics.push("入职");
    }
    if (header(table, REGULARIZATION_DATE_ALIASES)) {
      topics.push("转正");
    }
    if (header(table, SALARY_DATE_ALIASES)) {
      topics.push("调薪");
    }
    if (header(table, ["金额(*)", "实发合计"])) {
      topics.push("代发金额");
    }
    if (header(table, ["BM津贴", "津贴", "实发数额"])) {
      topics.push("津贴");
    }
    return topics;
  }

  function collectWorkbookBusinessEvidence(
    personnelSheets,
    targetTable,
    basePeriod,
    targetPeriod,
  ) {
    const errors = [];
    const warnings = [];
    const recognizedRegions = [];
    const departures = new Map();
    const targetDepartures = new Map();
    const changeResults = [];
    let historicalDepartureRows = 0;
    const start = targetStart(targetPeriod);
    if (!start) {
      return {
        sheetCount: personnelSheets?.length || 0,
        recognizedRegions,
        departures: [],
        targetDepartures: [],
        historicalDepartureRows,
        warnings,
        errors: ["目标月份格式无效，不能核对上月工资表分表"],
      };
    }
    const targetIndex = api.indexPeople(targetTable);
    for (const surface of personnelSheets || []) {
      for (const table of surface.tables || [surface.table].filter(Boolean)) {
        const topics = regionTopics(table);
        if (topics.length) {
          recognizedRegions.push({
            sheetName: surface.name,
            headerRow: table.headerRow,
            rows: table.rows.length,
            topics,
          });
        }
        changeResults.push(
          ...api.structuredWorkbookChangeResults(
            surface,
            table,
            targetTable,
            targetPeriod,
          ),
        );
        const exitHeader = header(table, EXIT_DATE_ALIASES);
        if (!exitHeader) {
          continue;
        }
        const noteHeader = header(table, NOTE_ALIASES);
        for (const row of table.rows) {
          const exitDate = api.parseSalaryEventDate(
            row.get(exitHeader.name),
          );
          if (!exitDate) {
            continue;
          }
          const matched = identityMatch(targetIndex, table, row);
          if (matched.match.status !== "matched") {
            continue;
          }
          const evidence = {
            sourceSheet: surface.name,
            sourceRow: row.rowNumber,
            exitDate,
            exitDateText: dateText(exitDate),
            note: noteHeader ? api.asText(row.get(noteHeader.name)) : "",
            identity: matched.identity,
            person: matched.match.person,
            matchedBy: matched.matchedBy,
          };
          const key = matched.match.person.rowNumber;
          const exitPeriod = api.salaryEventDatePeriod(exitDate);
          if (
            exitDate.valueOf() < start.valueOf() &&
            exitPeriod === basePeriod
          ) {
            const existing = departures.get(key);
            if (
              !existing ||
              surface.name === "离职名单" ||
              existing.sourceSheet !== "离职名单"
            ) {
              departures.set(key, evidence);
            }
          } else if (exitPeriod === targetPeriod) {
            const existing = targetDepartures.get(key);
            if (
              !existing ||
              surface.name === "离职名单" ||
              existing.sourceSheet !== "离职名单"
            ) {
              targetDepartures.set(key, evidence);
            }
          } else if (exitDate.valueOf() < start.valueOf()) {
            historicalDepartureRows += 1;
          }
        }
      }
    }
    if (departures.size) {
      warnings.push(
        `上月工资表分表有 ${departures.size} 项已在目标月前生效的离职或转出记录`,
      );
    }
    if (targetDepartures.size) {
      warnings.push(
        `上月工资表分表有 ${targetDepartures.size} 项目标月离职记录，仍需补齐当月结算资料`,
      );
    }
    if (historicalDepartureRows) {
      warnings.push(
        `分表另有 ${historicalDepartureRows} 条更早月份离职记录，仅作为历史留存，不重复形成停用项`,
      );
    }
    const changeProposals = changeResults.flatMap(
      (result) => result.proposals || [],
    );
    errors.push(
      ...changeResults.flatMap((result) => result.errors || []),
    );
    warnings.push(
      ...changeResults.flatMap((result) => result.warnings || []),
    );
    return {
      sheetCount: personnelSheets?.length || 0,
      recognizedRegions,
      departures: [...departures.values()],
      targetDepartures: [...targetDepartures.values()],
      historicalDepartureRows,
      changeProposals,
      warnings: [...new Set(warnings)],
      errors: [...new Set(errors)],
    };
  }

  function workbookEvidenceProposals(evidence, targetPeriod) {
    const proposals = [];
    for (const departure of evidence?.departures || []) {
      const proposal = api.proposalBase(
        `上月工资表·${departure.sourceSheet}`,
        departure.sourceRow,
        "",
      );
      Object.assign(proposal, {
        kind: "disable-person",
        operation: "停用",
        period: targetPeriod,
        person: departure.person,
        matchedBy: departure.matchedBy,
        workbookEvidence: true,
        archiveExisting: true,
        archiveSheet: departure.sourceSheet,
        archiveRow: departure.sourceRow,
        effectiveDate: departure.exitDateText,
      });
      proposal.warnings.push(
        `分表已明确 ${departure.exitDateText} 离职或转出；确认后从目标月份在职表移出，并沿用现有离职记录`,
      );
      proposals.push(proposal);
    }
    for (const departure of evidence?.targetDepartures || []) {
      const proposal = api.proposalBase(
        `上月工资表·${departure.sourceSheet}`,
        departure.sourceRow,
        "",
      );
      Object.assign(proposal, {
        kind: "disable-person",
        operation: "停用",
        period: targetPeriod,
        person: departure.person,
        matchedBy: departure.matchedBy,
        workbookEvidence: true,
        archiveExisting: true,
        archiveSheet: departure.sourceSheet,
        archiveRow: departure.sourceRow,
        effectiveDate: departure.exitDateText,
        status: "error",
        selected: false,
      });
      proposal.errors.push(
        `分表已明确 ${departure.exitDateText} 为目标月离职，但缺少工作天数、未发绩效和实发合计，需完成当月结算后再停用`,
      );
      proposals.push(proposal);
    }
    for (const sourceProposal of evidence?.changeProposals || []) {
      sourceProposal.workbookEvidence = true;
      sourceProposal.warnings = [
        ...(sourceProposal.warnings || []),
        "来源为上月完整工资表中的结构化业务分表",
      ];
      proposals.push(sourceProposal);
    }
    return proposals;
  }

  function unresolvedWorkbookDepartures(evidence, targetTable) {
    const targetIndex = api.indexPeople(targetTable);
    return (evidence?.departures || []).filter((departure) => {
      const kind = api.asText(departure.identity?.idCard)
        ? "idCard"
        : api.asText(departure.identity?.employeeId)
          ? "employeeId"
          : "name";
      return (
        api.matchPerson(targetIndex, departure.identity, {
          matchBy: [kind],
          stopAfterPresent: true,
        }).status === "matched"
      );
    });
  }

  function workbookEvidenceExcludedRows(evidence, targetTable) {
    return unresolvedWorkbookDepartures(evidence, targetTable).map(
      (departure) => departure.person.rowNumber,
    );
  }

  Object.assign(api, {
    WORKBOOK_EVIDENCE_RULE_META,
    collectWorkbookBusinessEvidence,
    workbookEvidenceProposals,
    unresolvedWorkbookDepartures,
    workbookEvidenceExcludedRows,
  });
})();
