(() => {
  "use strict";

  const api = window.PayrollLocal.rules;

  const MONTHLY_CHANGE_SOURCE_PROFILES = Object.freeze([
    Object.freeze({
      id: "employee-dynamics",
      label: "员工动态表",
      filenameAny: Object.freeze(["员工动态"]),
      requiredHeaders: Object.freeze([
        Object.freeze(["身份证号", "身份证", "证件号码"]),
        Object.freeze(["姓名"]),
        Object.freeze(["转岗后部门", "转岗后岗位"]),
      ]),
      mappings: Object.freeze([
        Object.freeze({
          source: Object.freeze(["转岗后部门"]),
          target: "部门",
          basis: "目标月份员工动态表明确列示的转岗后部门",
        }),
        Object.freeze({
          source: Object.freeze(["转岗后岗位"]),
          target: "岗位",
          basis: "目标月份员工动态表明确列示的转岗后岗位",
        }),
      ]),
      matchBy: Object.freeze(["idCard"]),
      onlyChanged: true,
      requirePeriod: true,
    }),
    Object.freeze({
      id: "employment-entry-exit",
      label: "员工入离职记录",
      filenameAny: Object.freeze(["员工动态"]),
      requiredHeaders: Object.freeze([
        Object.freeze(["身份证号", "身份证", "证件号码"]),
        Object.freeze(["姓名"]),
        Object.freeze([
          "入职日期",
          "入司日期",
          "入职时间",
          "离职日期",
          "离司日期",
          "离职时间",
        ]),
      ]),
      mappings: Object.freeze([]),
      matchBy: Object.freeze(["idCard"]),
      adapter: "employment-events",
      requirePeriod: true,
    }),
    Object.freeze({
      id: "probation-salary",
      label: "入职转正薪资表",
      filenameAny: Object.freeze(["入职转正薪资", "转正薪资"]),
      requiredHeaders: Object.freeze([
        Object.freeze(["身份证号", "身份证", "证件号码"]),
        Object.freeze(["姓名"]),
        Object.freeze(["转正日期"]),
        Object.freeze(["基本工资", "岗位工资", "绩效工资"]),
      ]),
      mappings: Object.freeze([
        Object.freeze({
          source: Object.freeze(["基本工资"]),
          target: "基本工资",
          basis: "目标月份入职转正薪资表的基本工资字段",
        }),
        Object.freeze({
          source: Object.freeze(["岗位工资"]),
          target: "岗位工资",
          basis: "目标月份入职转正薪资表的岗位工资字段",
        }),
        Object.freeze({
          source: Object.freeze(["绩效工资", "绩效工资标准"]),
          target: "绩效工资标准",
          basis: "目标月份入职转正薪资表的绩效工资字段",
        }),
      ]),
      matchBy: Object.freeze(["idCard"]),
      onlyChanged: true,
      requirePeriod: true,
      adapter: "salary-events",
      eventKind: "regularization",
    }),
    Object.freeze({
      id: "salary-adjustment",
      label: "调薪明细",
      filenameAny: Object.freeze(["入职转正薪资", "转正薪资"]),
      requiredHeaders: Object.freeze([
        Object.freeze(["身份证号", "身份证", "证件号码"]),
        Object.freeze(["姓名"]),
        Object.freeze(["调薪日期"]),
        Object.freeze(["调整后薪酬"]),
      ]),
      mappings: Object.freeze([
        Object.freeze({
          source: Object.freeze(["调整后薪酬"]),
          target: "基本工资",
          basis: "需求确认的工资构成 30%",
        }),
        Object.freeze({
          source: Object.freeze(["调整后薪酬"]),
          target: "岗位工资",
          basis: "需求确认的工资构成 50%",
        }),
        Object.freeze({
          source: Object.freeze(["调整后薪酬"]),
          target: "绩效工资标准",
          basis: "需求确认的工资构成 20%",
        }),
      ]),
      matchBy: Object.freeze(["idCard"]),
      onlyChanged: true,
      requirePeriod: true,
      adapter: "salary-events",
      eventKind: "salary-adjustment",
    }),
    Object.freeze({
      id: "salary-transfer-review",
      label: "入职转正薪资表转岗记录",
      filenameAny: Object.freeze(["入职转正薪资", "转正薪资"]),
      requiredHeaders: Object.freeze([
        Object.freeze(["身份证号", "身份证", "证件号码"]),
        Object.freeze(["姓名"]),
        Object.freeze(["转部门日期"]),
        Object.freeze(["转岗后部门", "转岗后岗位"]),
      ]),
      mappings: Object.freeze([]),
      matchBy: Object.freeze(["idCard"]),
      reviewOnly:
        "已按身份证核对转岗人员与日期；部门和岗位以同月员工动态表为准，避免重复写入。",
      requirePeriod: true,
    }),
    Object.freeze({
      id: "employee-regularization-review",
      label: "员工动态表转正记录",
      filenameAny: Object.freeze(["员工动态"]),
      requiredHeaders: Object.freeze([
        Object.freeze(["身份证号", "身份证", "证件号码"]),
        Object.freeze(["姓名"]),
        Object.freeze(["转正日期"]),
      ]),
      mappings: Object.freeze([]),
      matchBy: Object.freeze(["idCard"]),
      reviewOnly:
        "已按身份证核对转正人员与月份；转正统一按该月1日生效，工资构成以同月入职转正薪资表为准，不做月中折算。",
      requirePeriod: true,
    }),
    Object.freeze({
      id: "employee-salary-adjustment-review",
      label: "员工动态表调薪记录",
      filenameAny: Object.freeze(["员工动态"]),
      requiredHeaders: Object.freeze([
        Object.freeze(["身份证号", "身份证", "证件号码"]),
        Object.freeze(["姓名"]),
        Object.freeze(["调薪日期"]),
      ]),
      mappings: Object.freeze([]),
      matchBy: Object.freeze(["idCard"]),
      reviewOnly:
        "已按身份证核对调薪人员与生效日期；调整后薪酬金额以同月入职转正薪资表为准。",
      requirePeriod: true,
    }),
    Object.freeze({
      id: "attendance-register",
      label: "考勤表",
      filenameAny: Object.freeze(["考勤表"]),
      requiredHeaders: Object.freeze([
        Object.freeze(["身份证号", "身份证", "证件号码"]),
      ]),
      structureMatch: "attendance",
      mappings: Object.freeze([]),
      matchBy: Object.freeze(["idCard"]),
      adapter: "attendance-deductions",
      requirePeriod: true,
    }),
  ]);

  const MONTHLY_CHANGE_SOURCE_RULE_META = Object.freeze({
    id: "monthly-change-source-adapters",
    trigger: "选择目标月份全量工资附件",
    policy:
      "附件数量不限；有身份证字段时只按身份证唯一匹配；同一工作表的全部转正、转岗、调薪分区分别读取；工资构成按30%/50%/20%核对，考勤按需求公式扣款；转正统一按生效月份1日使用整月工资，月中调薪缺少折算口径时停止",
    profiles: Object.freeze(
      MONTHLY_CHANGE_SOURCE_PROFILES.map((profile) => profile.id),
    ),
  });

  Object.assign(api, {
    MONTHLY_CHANGE_SOURCE_PROFILES,
    MONTHLY_CHANGE_SOURCE_RULE_META,
  });
})();
