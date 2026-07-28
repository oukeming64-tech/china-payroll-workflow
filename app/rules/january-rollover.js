(() => {
  "use strict";

  const api = window.PayrollLocal.rules;
  const JANUARY_SCHEMA_PROFILES = Object.freeze({
    2026: Object.freeze({
      id: "2026-tax-deduction-v1",
      evidencePeriod: "2026-01",
      evidenceRef: "verified-structure-profile-2026-01",
      sourceFieldCount: 59,
      targetFieldCount: 63,
      insertAfterField: "累计应计工资",
      addedFields: Object.freeze([
        "扣减税额",
        "之前月份累计扣减税额",
        "累计扣减税额",
        "累计应纳税所得额",
      ]),
      externalSources: Object.freeze([
        "个税工资薪金附件",
        "社保 / 公积金附件",
        "劳务费附件",
      ]),
    }),
  });

  function schemaProfileForPeriod(period) {
    const year = Number(String(period || "").slice(0, 4));
    return JANUARY_SCHEMA_PROFILES[year] || null;
  }

  function buildJanuaryRolloverPlan(
    sourceTable,
    annualAudit,
    targetPeriod,
  ) {
    const profile = schemaProfileForPeriod(targetPeriod);
    const errors = [...(annualAudit?.errors || [])];
    if (!profile) {
      errors.push(`${targetPeriod.slice(0, 4)} 年 schema / 公式版本未验证`);
      return { profile: null, errors };
    }
    if (sourceTable.headers.length !== profile.sourceFieldCount) {
      errors.push(
        `跨年基线应为 ${profile.sourceFieldCount} 字段，当前为 ${sourceTable.headers.length} 字段`,
      );
    }
    const insertAfter = api.targetHeader(sourceTable, profile.insertAfterField);
    if (!insertAfter) {
      errors.push(`缺少跨年插列定位字段“${profile.insertAfterField}”`);
    }
    const resetFields = annualAudit?.resetFields || [];
    const requiredResetFields = sourceTable.headers
      .filter((header) => /^之前月份累计/.test(header.name))
      .map((header) => header.name);
    const unprovedResetFields = requiredResetFields.filter(
      (field) => !resetFields.includes(field),
    );
    if (unprovedResetFields.length) {
      errors.push(
        `一月历史未证明以下累计字段可重置：${unprovedResetFields.join("、")}`,
      );
    }
    const people = api.buildPeople(sourceTable).people;
    const nameOnlyRows = people
      .filter(
        (person) =>
          !api.asText(person.employeeId) &&
          !api.asText(person.idCard) &&
          api.asText(person.name),
      )
      .map((person) => person.rowNumber);
    return {
      profile,
      insertColumn: insertAfter ? insertAfter.column + 1 : null,
      addedFields: [...profile.addedFields],
      resetFields: [...resetFields, "之前月份累计扣减税额"],
      nameOnlyRows,
      errors,
    };
  }

  Object.assign(api, {
    JANUARY_SCHEMA_PROFILES,
    schemaProfileForPeriod,
    buildJanuaryRolloverPlan,
  });
})();
