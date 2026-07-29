(() => {
  "use strict";

  const api = window.PayrollLocal.rules;

  const PERFORMANCE_PAYOUT_MONTHS = Object.freeze([1, 4, 7, 10]);
  const CONFIDENTIAL_REVIEW_MONTHS = Object.freeze([3, 6, 9, 12]);
  const MONTHLY_RESET_FIELDS = Object.freeze([
    "月度绩效（季度发放）",
    "销售提成",
    "其他工资",
    "考勤扣款",
    "其他扣款",
    "年度绩效",
    "代扣借款",
  ]);
  const MONTHLY_BUSINESS_RULE_META = Object.freeze({
    id: "monthly-business-inputs",
    trigger: "建立目标月份草案后",
    automaticAction: "清空仅属于上月的一次性工资和扣款字段",
    requiredInputs: Object.freeze([
      "人力员工动态及当月金额",
      "行政考勤",
      "销售提成邮件",
      "季度绩效汇总（1、4、7、10 月）",
      "保密补贴邮件（3、6、9、12 月）",
    ]),
    pendingPolicy: "所需资料未逐项确认时停止生成",
  });
  const PAYROLL_REQUIREMENT_COVERAGE = Object.freeze([
    Object.freeze({
      id: "staff-events",
      label: "1. 入职、转正、离职、转岗和调薪",
      mode: "source",
      detail:
        "员工动态和入职转正薪资表按身份证、日期和分区读取。转正统一按生效月份1日处理，直接使用整月转正工资，不再要求转正或试用天数。离职工资与未发绩效必须有当月明确金额。",
    }),
    Object.freeze({
      id: "salary-composition",
      label: "2. 工资构成与季度绩效",
      mode: "source",
      detail:
        "基本30%、岗位50%、绩效20%会校验；绩效当月不发，季度发放月必须使用人力绩效汇总和系数，不从历史金额推算。",
    }),
    Object.freeze({
      id: "attendance",
      label: "3. 考勤扣款",
      mode: "automatic",
      detail:
        "病假三档和事假按需求公式写入考勤扣款；年婚丧陪产假正常发放。整月病假、工伤、产假缺当地规则时停止；事假后社保不足需人工确认补交。",
    }),
    Object.freeze({
      id: "special-deductions",
      label: "4. 专项附加扣除",
      mode: "source",
      detail:
        "只接受税局增减员后重新下载的目标月份数据，按身份证写入累计专项附加字段。",
    }),
    Object.freeze({
      id: "intern-labor",
      label: "5. 实习生津贴与劳务费",
      mode: "source",
      detail:
        "使用人力提供的具体金额；实习生津贴按身份证。劳务费与当前草案金额不同时提示复核，但临时劳务费或本月新增属于正常月度变动，不阻断写入。",
    }),
    Object.freeze({
      id: "income-tax",
      label: "6. 个税",
      mode: "source",
      detail:
        "正式员工和实习生沿用已验证七级累计预扣公式；专项扣除和劳务增值税来自目标月附件。劳务应纳税所得额超过20000元时必须复核30%/40%档。",
    }),
    Object.freeze({
      id: "confidential-allowance",
      label: "7. 保密补贴",
      mode: "manual",
      detail:
        "每季末依据质保部邮件或金额审批表核对；只有人员范围、没有金额时不写入。",
    }),
    Object.freeze({
      id: "sales-commission",
      label: "8. 销售提成",
      mode: "manual",
      detail: "依据当月邮件金额录入；没有提成也要明确确认。",
    }),
    Object.freeze({
      id: "monthly-insurance",
      label: "9. 每月社保明细",
      mode: "source",
      detail:
        "目标月份人力社保明细按身份证写入个人和公司承担字段；加密文件需输入密码后才能核对。",
    }),
    Object.freeze({
      id: "annual-social-base",
      label: "10. 年度社保基数",
      mode: "manual",
      detail:
        "按上年度工资总额月均形成候选；各地时间、上下限、取整及不足12个月处理仍以人力确认和最终附件为准。",
    }),
    Object.freeze({
      id: "severance",
      label: "11. 离职补偿金",
      mode: "manual",
      detail:
        "需要离职前近12个月工资总额月均及明确补偿口径；资料不足时不自动计算，也不把停用人员直接当作补偿金额。",
    }),
  ]);

  function periodMonth(period) {
    return api.periodParts(period)?.month || 0;
  }

  function isPerformancePayoutMonth(period) {
    return PERFORMANCE_PAYOUT_MONTHS.includes(periodMonth(period));
  }

  function isConfidentialReviewMonth(period) {
    return CONFIDENTIAL_REVIEW_MONTHS.includes(periodMonth(period));
  }

  function resolvedResetFields(table) {
    if (!table) {
      return [...MONTHLY_RESET_FIELDS];
    }
    const fields = [];
    for (const canonical of MONTHLY_RESET_FIELDS) {
      const header = api.targetHeader(table, canonical);
      if (header && !fields.includes(header.name)) {
        fields.push(header.name);
      }
    }
    return fields;
  }

  function manualItem(id, label, source, detail, fields = []) {
    return {
      id,
      label,
      source,
      detail,
      fields,
      automatic: false,
      confirmed: false,
    };
  }

  function automaticItem(id, label, source, detail, fields = []) {
    return {
      id,
      label,
      source,
      detail,
      fields,
      automatic: true,
      confirmed: true,
    };
  }

  function monthlyBusinessPlan(period, table = null) {
    const parts = api.periodParts(period);
    if (!parts) {
      return {
        period,
        resetFields: [],
        items: [],
        errors: ["目标月份格式无效"],
      };
    }
    const performancePayout = isPerformancePayoutMonth(period);
    const confidentialReview = isConfidentialReviewMonth(period);
    const items = [
      manualItem(
        "hr-changes",
        "人员与工资变动",
        "人力员工动态及当月金额",
        "按身份证核对入离职日期、转正、转岗、调薪、实习或劳务、离职未发绩效和其他一次性工资；转正统一按生效月份1日使用整月工资，离职补偿缺近12个月口径时必须停止。无变动也需确认。",
        [
          "岗位",
          "部门",
          "基本工资",
          "岗位工资",
          "绩效工资标准",
          "其他工资",
          "年度绩效",
          "代扣借款",
        ],
      ),
      manualItem(
        "attendance",
        "考勤与其他扣款",
        "行政考勤",
        "核对病事假等考勤结果和其他扣款；这些数值已停止沿用上月。",
        ["考勤扣款", "其他扣款"],
      ),
      manualItem(
        "sales-commission",
        "销售提成",
        "当月销售提成邮件",
        "依据当月邮件录入；本月无提成也需确认。",
        ["销售提成"],
      ),
      performancePayout
        ? manualItem(
            "quarterly-performance",
            "季度绩效",
            "人力绩效汇总",
            "本月为绩效发放月，必须依据绩效汇总录入或核对；程序不会从历史月份推算金额。",
            ["月度绩效（季度发放）"],
          )
        : automaticItem(
            "quarterly-performance",
            "季度绩效",
            "季度发放规则",
            "本月不是绩效发放月，已清空。离职人员的未结绩效需在人员与工资变动中单独补录。",
            ["月度绩效（季度发放）"],
          ),
      confidentialReview
        ? manualItem(
            "confidential-allowance",
            "保密补贴",
            "季度末保密补贴邮件",
            "本月为季度末月，需依据邮件核对 BM 津贴；程序保留上月值但不替代本次核对。",
            ["BM津贴"],
          )
        : automaticItem(
            "confidential-allowance",
            "保密补贴",
            "季度核对规则",
            "本月不是季度末月，沿用最近一次已确认的 BM 津贴。",
            ["BM津贴"],
          ),
    ];
    return {
      period,
      performancePayout,
      confidentialReview,
      resetFields: resolvedResetFields(table),
      items,
      errors: [],
    };
  }

  function pendingMonthlyBusiness(plan) {
    return (plan?.items || []).filter(
      (item) => !item.automatic && !item.confirmed,
    );
  }

  Object.assign(api, {
    PERFORMANCE_PAYOUT_MONTHS,
    CONFIDENTIAL_REVIEW_MONTHS,
    MONTHLY_RESET_FIELDS,
    MONTHLY_BUSINESS_RULE_META,
    PAYROLL_REQUIREMENT_COVERAGE,
    isPerformancePayoutMonth,
    isConfidentialReviewMonth,
    monthlyBusinessPlan,
    pendingMonthlyBusiness,
  });
})();
