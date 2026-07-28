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
        "核对入离职、转正、实习或劳务、工资调整、离职结算和其他一次性工资；无变动也需确认。",
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
    isPerformancePayoutMonth,
    isConfidentialReviewMonth,
    monthlyBusinessPlan,
    pendingMonthlyBusiness,
  });
})();
