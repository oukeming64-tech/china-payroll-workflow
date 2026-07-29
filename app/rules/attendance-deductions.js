(() => {
  "use strict";

  const api = window.PayrollLocal.rules;
  const LEAVE_FIELDS = Object.freeze([
    "年假",
    "婚假",
    "丧假",
    "带薪病假",
    "产检",
    "陪产假",
    "扣款病假",
    "产假",
    "事假",
    "无薪休息",
    "工伤",
  ]);
  const MANUAL_REVIEW_MESSAGES = Object.freeze({
    产检: "需求未明确产检扣款规则，请人工确认",
    产假: "产假待遇取决于当地生育津贴发放方式，请人工确认",
    无薪休息: "需求未提供无薪休息的计算公式，请人工确认",
    工伤: "停工留薪期及工资需按国家和当地认定结果处理，请人工确认",
  });
  const ATTENDANCE_DEDUCTION_RULE_META = Object.freeze({
    id: "attendance-deductions",
    trigger: "目标月份考勤表包含按假别列示的请假天数",
    targetField: "考勤扣款",
    wageBaseField: "工资合计",
    policy:
      "年假、婚假、丧假、带薪病假和陪产假不扣款；扣款病假按天数分档，事假按工资合计除以22乘天数；产检、产假、无薪休息缺少完整规则时停止",
  });

  function sourceHeader(table, name) {
    const key = api.normalizeText(name);
    return [...(table.subheaders || []), ...(table.headers || [])].find(
      (header) => api.normalizeText(header.name) === key,
    );
  }

  function roundMoney(value) {
    return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
  }

  function dayText(value) {
    return Number(value).toString();
  }

  function leaveDays(sourceRow, header, label, errors) {
    if (!header) {
      return 0;
    }
    const raw = sourceRow.get(header.column);
    if (raw === null || raw === undefined || api.asText(raw) === "") {
      return 0;
    }
    const days = api.asNumber(raw);
    if (days === null || days < 0) {
      errors.push(`第 ${sourceRow.rowNumber} 行“${label}”天数无效`);
      return 0;
    }
    return days;
  }

  function wageBase(person, targetTable) {
    const totalHeader = api.targetHeader(targetTable, "工资合计");
    const components = [
      "基本工资",
      "岗位工资",
      "绩效工资标准",
    ].map((name) => api.targetHeader(targetTable, name));
    const total = totalHeader
      ? api.asNumber(person.row.get(totalHeader.name))
      : null;
    const componentValues = components.map((header) =>
      header ? api.asNumber(person.row.get(header.name)) : null,
    );
    const componentTotal = componentValues.every(
      (value) => value !== null,
    )
      ? componentValues.reduce((sum, value) => sum + value, 0)
      : null;
    return {
      header: totalHeader,
      value:
        componentTotal !== null &&
        (total === null || !api.sameValue(componentTotal, total))
          ? componentTotal
          : total,
      usedComponentTotal:
        componentTotal !== null &&
        (total === null || !api.sameValue(componentTotal, total)),
    };
  }

  function sickExpression(reference, days) {
    const text = dayText(days);
    if (days <= 5) {
      return {
        expression: `${reference}*1%*${text}`,
        value: (wage) => wage * 0.01 * days,
        basis: `扣款病假 ${text} 天，按工资合计×1%×天数`,
      };
    }
    if (days <= 10) {
      return {
        expression: `${reference}*1.5%*${text}`,
        value: (wage) => wage * 0.015 * days,
        basis: `扣款病假 ${text} 天，按工资合计×1.5%×天数`,
      };
    }
    return {
      expression: `${reference}/22*50%*${text}`,
      value: (wage) => (wage / 22) * 0.5 * days,
      basis: `扣款病假 ${text} 天，按工资合计÷22×50%×天数`,
    };
  }

  function deductionCalculation(
    wage,
    wageReference,
    sickDays,
    personalDays,
  ) {
    const parts = [];
    const basis = [];
    let value = 0;
    if (sickDays > 0) {
      const sick = sickExpression(wageReference, sickDays);
      parts.push(sick.expression);
      basis.push(sick.basis);
      value += sick.value(wage);
    }
    if (personalDays > 0) {
      const days = dayText(personalDays);
      parts.push(`${wageReference}/22*${days}`);
      basis.push(`事假 ${days} 天，按工资合计÷22×天数`);
      value += (wage / 22) * personalDays;
    }
    return {
      value: roundMoney(value),
      formula: `ROUND(${parts.join("+")},2)`,
      basis: basis.join("；"),
    };
  }

  function attendanceMappings(sourceTable, targetTable) {
    const targetHeader = api.targetHeader(targetTable, "考勤扣款");
    return ["扣款病假", "事假"]
      .map((name) => ({
        sourceHeader: sourceHeader(sourceTable, name),
        targetHeader,
        sourceField: name,
        targetField: "考勤扣款",
        basis: ATTENDANCE_DEDUCTION_RULE_META.policy,
      }))
      .filter((mapping) => mapping.sourceHeader && mapping.targetHeader);
  }

  function isEmptyLeaveSummary(value) {
    const text = api.normalizeText(value);
    return !text || ["无", "无请假", "无请假信息", "0"].includes(text);
  }

  function proposalsFromAttendanceSource(
    sourceTable,
    targetTable,
    targetPeriod,
    sourceName,
    profile,
    sourcePeriod,
  ) {
    const mappings = attendanceMappings(sourceTable, targetTable);
    const errors = [];
    const warnings = [];
    const proposals = [];
    const identities = api.identityHeaders(sourceTable);
    const targetHeader = api.targetHeader(targetTable, "考勤扣款");
    const wageHeader = api.targetHeader(targetTable, "工资合计");
    const headers = Object.fromEntries(
      LEAVE_FIELDS.map((name) => [name, sourceHeader(sourceTable, name)]),
    );
    const leaveSummaryHeader = sourceHeader(sourceTable, "请假信息");
    const hasDetailedLeaveHeaders = Object.values(headers).some(Boolean);
    if (!identities.idCard) {
      errors.push("考勤表缺少身份证号，不能按身份证匹配");
    }
    if (!targetHeader) {
      errors.push("目标工资表缺少“考勤扣款”字段");
    }
    if (!wageHeader) {
      errors.push("目标工资表缺少“工资合计”字段");
    }
    if (!hasDetailedLeaveHeaders && !leaveSummaryHeader) {
      errors.push("考勤表没有找到按假别列示的请假天数");
    }
    if (!hasDetailedLeaveHeaders && leaveSummaryHeader) {
      const unresolvedSummaries = sourceTable.rows.filter(
        (row) => !isEmptyLeaveSummary(row.get(leaveSummaryHeader.column)),
      );
      if (unresolvedSummaries.length) {
        errors.push(
          "考勤表的“请假信息”没有按假别和天数拆分，不能套用扣款公式",
        );
      } else {
        warnings.push(
          "考勤表只有“请假信息”总栏且均为无请假，不形成考勤扣款",
        );
      }
    }
    if (errors.length) {
      return {
        profile,
        sourcePeriod,
        format: "attendance-deductions",
        mappings,
        proposals,
        warnings,
        errors,
      };
    }
    if (!hasDetailedLeaveHeaders) {
      return {
        profile,
        sourcePeriod,
        format: "attendance-deductions",
        mappings,
        proposals,
        matchedPeople: 0,
        unmatchedPeople: 0,
        warnings,
        errors,
      };
    }

    const targetIndex = api.indexPeople(targetTable);
    let matchedPeople = 0;
    let unmatchedPeople = 0;
    let noDeductionPeople = 0;
    for (const sourceRow of sourceTable.rows) {
      const rowErrors = [];
      const days = Object.fromEntries(
        LEAVE_FIELDS.map((name) => [
          name,
          leaveDays(sourceRow, headers[name], name, rowErrors),
        ]),
      );
      const totalLeave = Object.values(days).reduce(
        (sum, value) => sum + value,
        0,
      );
      if (!totalLeave && !rowErrors.length) {
        continue;
      }
      const identity = api.identityFromRow(sourceRow, identities);
      const match = api.matchPerson(targetIndex, identity, {
        matchBy: ["idCard"],
        stopAfterPresent: true,
      });
      if (match.status !== "matched") {
        unmatchedPeople += 1;
        continue;
      }
      matchedPeople += 1;
      for (const [name, message] of Object.entries(
        MANUAL_REVIEW_MESSAGES,
      )) {
        if (days[name] > 0) {
          rowErrors.push(`第 ${sourceRow.rowNumber} 行“${name}”：${message}`);
        }
      }
      if (days.扣款病假 >= 22) {
        rowErrors.push(
          `第 ${sourceRow.rowNumber} 行扣款病假达到全月，需提供当地最低工资标准后按80%计发`,
        );
      }
      const wage = wageBase(match.person, targetTable);
      if (wage.value === null || wage.value < 0) {
        rowErrors.push(
          `第 ${sourceRow.rowNumber} 行无法取得有效的“工资合计”`,
        );
      }
      const deductible =
        days.扣款病假 > 0 || days.事假 > 0;
      if (!deductible && !rowErrors.length) {
        noDeductionPeople += 1;
        continue;
      }
      const proposal = api.proposalBase(
        sourceName,
        sourceRow.rowNumber,
        "",
      );
      Object.assign(proposal, {
        kind: "cell-change",
        operation: "设置",
        period: targetPeriod,
        field: targetHeader,
        person: match.person,
        matchedBy: match.matchedBy,
        currentValue: match.person.row.get(targetHeader.name),
        mapping: {
          sourceField: "扣款病假、事假",
          targetField: "考勤扣款",
          basis: ATTENDANCE_DEDUCTION_RULE_META.policy,
        },
      });
      if (rowErrors.length) {
        proposal.status = "error";
        proposal.selected = false;
        proposal.errors.push(...rowErrors);
      } else {
        const reference =
          `${window.XlsxEngine.columnNumberToLetters(wageHeader.column)}` +
          `${match.person.rowNumber}`;
        const calculation = deductionCalculation(
          wage.value,
          reference,
          days.扣款病假,
          days.事假,
        );
        proposal.inputValue = calculation.value;
        proposal.formula = calculation.formula;
        proposal.calculation = calculation;
        proposal.warnings.push(
          `${calculation.basis}；写入公式 =${calculation.formula}`,
        );
        if (wage.usedComponentTotal) {
          proposal.warnings.push(
            "工资合计缓存尚未重算，预览金额按基本工资、岗位工资和绩效工资标准之和计算",
          );
        }
        if (days.事假 > 0) {
          proposal.warnings.push(
            "事假扣款后如当月工资不足缴纳社保和公积金个人部分，需人工确认补交或下月扣回",
          );
        }
      }
      proposals.push(proposal);
    }
    if (unmatchedPeople) {
      errors.push(`来源表有 ${unmatchedPeople} 人未按身份证匹配工资表`);
    }
    if (noDeductionPeople) {
      warnings.push(
        `${noDeductionPeople} 人只有年假等正常出勤假期，不形成考勤扣款`,
      );
    }
    if (!proposals.length && !errors.length) {
      warnings.push("考勤表没有需要扣款的病假或事假");
    }
    return {
      profile,
      sourcePeriod,
      format: "attendance-deductions",
      mappings,
      proposals,
      matchedPeople,
      unmatchedPeople,
      warnings,
      errors,
    };
  }

  Object.assign(api, {
    ATTENDANCE_DEDUCTION_RULE_META,
    proposalsFromAttendanceSource,
  });
})();
