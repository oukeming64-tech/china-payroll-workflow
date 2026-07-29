(() => {
  "use strict";

  const api = window.PayrollLocal.rules;
  const EVENT_DATES = Object.freeze({
    regularization: Object.freeze(["转正日期"]),
    salaryAdjustment: Object.freeze(["调薪日期", "生效日期"]),
    transfer: Object.freeze(["转部门日期", "转岗日期", "生效日期"]),
    entry: Object.freeze(["入职日期", "入司日期", "入职时间"]),
  });

  function header(table, aliases) {
    return api.headerForAliases(table, aliases);
  }

  function rowsForPeriod(table, aliases, period) {
    const dateHeader = header(table, aliases);
    if (!dateHeader) {
      return [];
    }
    return table.rows.filter((row) => {
      const date = api.parseSalaryEventDate(
        row.get(dateHeader.name),
      );
      return api.salaryEventDatePeriod(date) === period;
    });
  }

  function tableWithRows(table, rows) {
    return {
      ...table,
      rows,
      rowCount: rows.length,
    };
  }

  function profileById(id) {
    return (api.MONTHLY_CHANGE_SOURCE_PROFILES || []).find(
      (profile) => profile.id === id,
    );
  }

  function profileMatchesTable(profile, table) {
    return Boolean(
      profile &&
      profile.requiredHeaders.every((aliases) =>
        header(table, aliases)
      ),
    );
  }

  function structuredWorkbookChangeResults(
    surface,
    table,
    targetTable,
    targetPeriod,
  ) {
    const sourceName = `上月工资表·${surface.name}`;
    const results = [];
    const regularization = profileById("probation-salary");
    if (profileMatchesTable(regularization, table)) {
      const rows = rowsForPeriod(
        table,
        EVENT_DATES.regularization,
        targetPeriod,
      );
      if (rows.length) {
        results.push(
          api.proposalsFromSalaryEventSource(
            tableWithRows(table, rows),
            targetTable,
            targetPeriod,
            sourceName,
            regularization,
            targetPeriod,
          ),
        );
      }
    }
    const salaryAdjustment = profileById("salary-adjustment");
    if (profileMatchesTable(salaryAdjustment, table)) {
      const rows = rowsForPeriod(
        table,
        EVENT_DATES.salaryAdjustment,
        targetPeriod,
      );
      if (rows.length) {
        results.push(
          api.proposalsFromSalaryEventSource(
            tableWithRows(table, rows),
            targetTable,
            targetPeriod,
            sourceName,
            salaryAdjustment,
            targetPeriod,
          ),
        );
      }
    }
    const transfer = profileById("employee-dynamics");
    if (profileMatchesTable(transfer, table)) {
      const rows = rowsForPeriod(
        table,
        EVENT_DATES.transfer,
        targetPeriod,
      );
      if (rows.length) {
        results.push(
          api.proposalsFromBusinessSource(
            tableWithRows(table, rows),
            targetTable,
            targetPeriod,
            sourceName,
            { ...transfer, requirePeriod: false },
          ),
        );
      }
    }
    const employment = profileById("employment-entry-exit");
    if (
      surface.name !== "离职名单" &&
      profileMatchesTable(employment, table)
    ) {
      const rows = rowsForPeriod(
        table,
        EVENT_DATES.entry,
        targetPeriod,
      );
      if (rows.length) {
        results.push(
          api.proposalsFromEmploymentEventSource(
            tableWithRows(table, rows),
            targetTable,
            targetPeriod,
            sourceName,
            employment,
            targetPeriod,
          ),
        );
      }
    }
    return results;
  }

  Object.assign(api, {
    structuredWorkbookChangeResults,
  });
})();
