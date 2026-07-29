(() => {
  "use strict";

  const ui = window.PayrollLocal.ui;
  const rules = window.PayrollEngine;
  const STATUS = Object.freeze({
    automatic: Object.freeze({
      label: "已接入公式",
      className: "",
    }),
    source: Object.freeze({
      label: "按当月资料",
      className: "",
    }),
    manual: Object.freeze({
      label: "不足即停止",
      className: "warning",
    }),
  });

  function renderRequirementCoverage() {
    const container = ui.byId("requirementCoverageList");
    if (!container) {
      return;
    }
    const items = rules.PAYROLL_REQUIREMENT_COVERAGE || [];
    container.innerHTML = items
      .map((item) => {
        const status = STATUS[item.mode] || STATUS.manual;
        return `<article class="diagnostic-card">
          <div class="diagnostic-card-head">
            <strong>${ui.escapeHtml(item.label)}</strong>
            <span class="status-chip ${status.className}">${status.label}</span>
          </div>
          <p>${ui.escapeHtml(item.detail)}</p>
        </article>`;
      })
      .join("");
  }

  Object.assign(ui, {
    renderRequirementCoverage,
  });
})();
