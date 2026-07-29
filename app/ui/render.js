(() => {
  "use strict";

  const ui = window.PayrollLocal.ui;
  const rules = window.PayrollEngine;

  function privateValue(value, kind = "value") {
    if (ui.state.showPrivateData) {
      return ui.escapeHtml(ui.formatValue(value));
    }
    if (kind === "name") {
      return ui.escapeHtml(rules.maskName(value));
    }
    if (kind === "id") {
      return ui.escapeHtml(rules.maskIdentifier(value));
    }
    return value === null || value === undefined || value === "" ? "—" : "••••";
  }

  function renderStats() {
    const table = ui.state.table;
    if (!table) {
      return;
    }
    const people = rules.buildPeople(table).people.filter(
      (person) => !ui.state.disabledRows.has(person.rowNumber),
    );
    ui.byId("mainSheetStat").textContent = table.sheetName;
    ui.byId("headerRowStat").textContent = `第 ${table.headerRow} 行为字段头`;
    ui.byId("peopleStat").textContent = String(people.length);
    ui.byId("fieldStat").textContent = String(table.headers.length);
    ui.byId("formulaStat").textContent = String(
      ui.state.formulaDiagnostics?.totalFormulaNodes ??
        table.formulaCount,
    );
    ui.byId("externalStat").textContent = String(ui.state.externalLinks.length);
    ui.byId("proposalBadge").textContent = String(ui.state.proposals.length);
    ui.byId("historyBadge").textContent = String(ui.state.history.length);
    ui.byId("exportChangeCount").textContent = String(ui.state.history.length);
    ui.byId("exportLinkCount").textContent = String(ui.state.externalLinks.length);
    ui.byId("monthMarkerCount").textContent = String(
      ui.state.monthMarkers.length,
    );
  }

  function peopleColumns(table) {
    const names = [
      "人员编号",
      "姓名",
      "部门",
      "岗位",
      "基本工资",
      "岗位工资",
      "工资合计",
      "应计工资",
      "实发合计",
    ];
    const columns = [];
    for (const name of names) {
      const header = rules.targetHeader(table, name);
      if (header && !columns.some((column) => column.column === header.column)) {
        columns.push(header);
      }
    }
    return columns;
  }

  function renderPeopleTable() {
    const table = ui.state.table;
    if (!table) {
      return;
    }
    const query = rules.normalizeText(ui.byId("peopleSearch").value);
    const identities = rules.identityHeaders(table);
    const columns = peopleColumns(table);
    ui.byId("peopleTable").querySelector("thead").innerHTML = `<tr>${columns
      .map((header) => `<th>${ui.escapeHtml(header.displayName)}</th>`)
      .join("")}</tr>`;
    const people = rules.buildPeople(table).people.filter((person) => {
      if (ui.state.disabledRows.has(person.rowNumber)) {
        return false;
      }
      if (!query) {
        return true;
      }
      return [person.name, person.employeeId, person.department].some((value) =>
        rules.normalizeText(value).includes(query),
      );
    });
    ui.byId("peopleTable").querySelector("tbody").innerHTML = people
      .slice(0, 120)
      .map(
        (person) =>
          `<tr>${columns
            .map((header) => {
              const value = person.row.get(header.name);
              const kind =
                header === identities.name
                  ? "name"
                  : header === identities.employeeId ||
                      header === identities.idCard
                    ? "id"
                    : "value";
              return `<td>${privateValue(value, kind)}</td>`;
            })
            .join("")}</tr>`,
      )
      .join("");
  }

  function proposalStatus(proposal) {
    if (proposal.status === "error") {
      return '<span class="status-dot error">错误</span>';
    }
    if (proposal.warnings?.length) {
      return '<span class="status-dot warning">提醒</span>';
    }
    return '<span class="status-dot">可应用</span>';
  }

  function renderProposals() {
    rules.markConflicts(ui.state.proposals);
    const proposals = ui.state.proposals;
    const ready = proposals.filter((item) => item.status !== "error");
    const errors = proposals.length - ready.length;
    ui.byId("proposalBadge").textContent = String(proposals.length);
    ui.byId("proposalSummary").textContent = proposals.length
      ? `${ready.length} 项可应用，${errors} 项需要处理；错误项不会写入。`
      : "暂无待确认变动。";
    ui.byId("applyProposalsBtn").disabled = !ready.some(
      (item) => item.selected,
    );
    ui.byId("proposalEmpty").hidden = proposals.length > 0;
    ui.byId("proposalTable").querySelector("tbody").innerHTML = proposals
      .map((proposal) => {
        const person = proposal.person
          ? privateValue(proposal.person.name, "name")
          : proposal.kind === "new-person"
            ? privateValue(
                proposal.newPersonValues?.[
                  rules.identityHeaders(ui.state.table).name?.name
                ],
                "name",
              )
            : "未匹配";
        const problem = [...(proposal.errors || []), ...(proposal.warnings || [])]
          .map((message) => `<div>${ui.escapeHtml(message)}</div>`)
          .join("");
        return `<tr>
          <td><input type="checkbox" data-proposal-id="${proposal.id}" ${
            proposal.selected ? "checked" : ""
          } ${proposal.status === "error" ? "disabled" : ""}></td>
          <td>${proposalStatus(proposal)}</td>
          <td>${person}</td>
          <td>${ui.escapeHtml(proposal.field?.displayName || (proposal.kind === "new-person" ? "新增人员行" : "整行"))}</td>
          <td>${ui.escapeHtml(proposal.operation || "—")}</td>
          <td>${privateValue(proposal.currentValue)}</td>
          <td>${privateValue(proposal.inputValue ?? (proposal.kind === "disable-person" ? "隐藏" : "新增"))}</td>
          <td>${ui.escapeHtml(proposal.source || "—")}</td>
          <td class="${proposal.errors?.length ? "cell-error" : "cell-warning"}">${problem || "—"}</td>
        </tr>`;
      })
      .join("");
    ui.byId("proposalTable")
      .querySelectorAll("[data-proposal-id]")
      .forEach((checkbox) => {
        checkbox.addEventListener("change", () => {
          const proposal = proposals.find(
            (item) => item.id === checkbox.dataset.proposalId,
          );
          if (proposal) {
            proposal.selected = checkbox.checked;
            renderProposals();
          }
        });
      });
  }

  function renderSources() {
    ui.byId("sourceSection").hidden = !ui.state.sources.length;
    ui.byId("sourceCountText").textContent =
      `${ui.state.sources.length} 个文件`;
    ui.byId("sourceCards").innerHTML = ui.state.sources.length
      ? ui.state.sources
          .map(
            (source) => `<article class="source-card">
              <div class="source-card-head">
                <strong>${ui.escapeHtml(source.name)}</strong>
                <span class="status-chip ${source.errors?.length ? "warning" : ""}">${ui.escapeHtml(source.category)}</span>
              </div>
              <p>工作表：${ui.escapeHtml(source.sheetName || "未识别")} · 字段匹配 ${source.mappingCount || 0} · 形成预览 ${source.proposalCount || 0}</p>
              ${source.warnings?.length ? `<p class="cell-warning">${ui.escapeHtml(source.warnings.join("；"))}</p>` : ""}
              ${source.errors?.length ? `<p class="cell-error">${ui.escapeHtml(source.errors.join("；"))}</p>` : ""}
            </article>`,
          )
          .join("")
      : '<div class="empty-state">还没有导入来源文件。</div>';
  }

  function renderAttachments() {
    const audit = ui.state.attachments;
    const required = audit.required || [];
    const byCategory = new Map(
      (audit.results || []).map((result) => [
        result.category,
        result,
      ]),
    );
    const ready = required.filter((item) => {
      const result = byCategory.get(item.category);
      return result && !result.errors.length;
    }).length;
    const readFiles = (audit.files || []).filter((file) => file.read).length;
    const changeFiles = (audit.files || []).filter(
      (file) => file.role === "change",
    );
    ui.byId("attachmentBadge").textContent = audit.applied
      ? "✓"
      : `${ready}/${required.length}`;
    ui.byId("attachmentBanner").hidden = audit.applied;
    ui.byId("attachmentSummary").textContent = audit.applied
      ? "目标月份附件已写入，当前草案无外链。"
      : changeFiles.length
        ? `${readFiles}/${audit.files.length} 个文件已读取；${ready}/${required.length} 类工资字段附件通过，${changeFiles.length} 个变动文件已综合匹配。`
        : audit.results.length
        ? `${audit.results.length}/${required.length} 类附件已读取，${ready}/${required.length} 类通过核对。`
        : `${ready}/${required.length} 类附件已通过核对。`;
    const requiredCards = required
      .map((item) => {
        const result = byCategory.get(item.category);
        const status = audit.applied && result
          ? "已写入"
          : result
            ? result.errors.length
              ? "已读取，待处理"
              : "待确认"
            : "待选择";
        const fields = result?.fieldSummaries
          ?.map(
            (field) =>
              `<span>${ui.escapeHtml(field.sourceField)} → ${ui.escapeHtml(field.targetField)} · ${field.matched} 人</span>`,
          )
          .join("") || "<span>尚未核对字段</span>";
        const warnings = result?.warnings?.length
          ? `<p class="cell-warning">${ui.escapeHtml(result.warnings.join("；"))}</p>`
          : "";
        return `<article class="diagnostic-card">
          <div class="diagnostic-card-head">
            <strong>${ui.escapeHtml(item.category)}</strong>
            <span class="status-chip ${status === "已写入" || status === "待确认" ? "" : "warning"}">${status}</span>
          </div>
          <p>${ui.escapeHtml(result ? `${result.sourceName} · ${result.sourceSheet} · 匹配 ${result.matchedPeople}/${result.expectedPeople} 人` : "请选择目标月份对应文件")}</p>
          ${warnings}
          <div class="diagnostic-fields">${fields}</div>
        </article>`;
      })
      .join("");
    const changeCards = changeFiles
      .map((file) => `<article class="diagnostic-card">
        <div class="diagnostic-card-head">
          <strong>${ui.escapeHtml(file.name)}</strong>
          <span class="status-chip ${file.errors?.length ? "warning" : ""}">${ui.escapeHtml(file.category)}</span>
        </div>
        <p>字段匹配 ${file.mappingCount || 0} · 形成变动预览 ${file.proposalCount || 0}</p>
        ${file.warnings?.length ? `<p class="cell-warning">${ui.escapeHtml(file.warnings.join("；"))}</p>` : ""}
        ${file.errors?.length ? `<p class="cell-error">${ui.escapeHtml(file.errors.join("；"))}</p>` : ""}
      </article>`)
      .join("");
    ui.byId("attachmentCards").innerHTML = requiredCards + changeCards;
    ui.byId("attachmentErrors").hidden = !audit.errors.length;
    ui.byId("attachmentErrors").innerHTML = audit.errors.length
      ? audit.errors
          .map((error) => `<div>${ui.escapeHtml(error)}</div>`)
          .join("")
      : "";
    ui.byId("applyAttachmentsBtn").disabled =
      audit.applied ||
      Boolean(audit.errors.length) ||
      !required.length ||
      ready !== required.length;
  }

  function renderFormulaDiagnostics() {
    const diagnostics = ui.state.formulaDiagnostics;
    if (!diagnostics) {
      ui.byId("formulaSummary").textContent = "尚未读取工资表。";
      ui.byId("formulaFieldList").innerHTML = "";
      return;
    }
    ui.byId("formulaSummary").textContent =
      `${diagnostics.totalFormulaNodes} 个公式单元格 · 内部 ${diagnostics.internalFormulaNodes} · 外部 ${diagnostics.externalFormulaNodes} · ${diagnostics.fields.length} 个字段`;
    ui.byId("formulaFieldList").innerHTML = diagnostics.fields
      .map((field) => {
        const sources = field.sourceSheets.length
          ? `来源：${field.sourceSheets.join("、")}`
          : "来源：本表单元格";
        const unresolved = field.unresolvedFormulaNodes
          ? ` · ${field.unresolvedFormulaNodes} 个共享公式未展开`
          : "";
        return `<article class="formula-card">
          <div>
            <strong>${ui.escapeHtml(field.field)}</strong>
            <small>${ui.escapeHtml(sources)} · 单元格 ${ui.escapeHtml(field.sampleCells.join("、"))}${ui.escapeHtml(unresolved)}</small>
          </div>
          <span>${field.formulaNodes} 个</span>
          <small>内部 ${field.internalFormulaNodes} · 外部 ${field.externalFormulaNodes}</small>
        </article>`;
      })
      .join("");
  }

  function renderDiagnostics() {
    const links = ui.state.externalLinks;
    ui.byId("externalDiagnostics").innerHTML = links.length
      ? links
          .map(
            (link) => `<article class="diagnostic-card">
              <div class="diagnostic-card-head">
                <strong>${link.index}. ${ui.escapeHtml(link.filename || "未命名外部工作簿")}</strong>
                <span class="status-chip warning">${ui.escapeHtml(link.category)}</span>
              </div>
              <p>${ui.escapeHtml(link.macStatus)} · ${link.formulaReferenceCount} 个公式引用 · 缓存 ${link.cachedCells} 个单元格</p>
              <div class="diagnostic-fields">${
                link.targetFields.length
                  ? link.targetFields.map((field) => `<span>${ui.escapeHtml(field)}</span>`).join("")
                  : "<span>未发现当前主表引用</span>"
              }</div>
            </article>`,
          )
          .join("")
      : '<div class="empty-state">工作簿未声明外部链接。</div>';
  }

  function renderSourceRouting() {
    const routing = ui.state.sourceRouting;
    if (!routing) {
      ui.byId("sourceRouting").innerHTML =
        '<div class="empty-state">尚未建立月度来源计划。</div>';
      return;
    }
    ui.byId("sourceRouting").innerHTML = routing.plans
      .map(
        (plan) => `<article class="diagnostic-card">
          <div class="diagnostic-card-head">
            <strong>${ui.escapeHtml(plan.category)}</strong>
            <span class="status-chip ${plan.status === "ready" ? "" : "warning"}">${plan.status === "ready" ? "已核对" : "错误"}</span>
          </div>
          <p>${ui.escapeHtml(plan.beforeFilename || "无")} → ${ui.escapeHtml(plan.afterFilename || "无")}</p>
          <div class="diagnostic-fields"><span>${ui.escapeHtml(plan.basis || "—")}</span></div>
        </article>`,
      )
      .join("");
  }

  function renderMonthlyBusiness() {
    const plan = ui.state.monthlyBusiness;
    const items = plan?.items || [];
    const pending = rules.pendingMonthlyBusiness(plan);
    const completed = items.length - pending.length;
    ui.byId("monthlyBusinessSummary").textContent = items.length
      ? `${completed}/${items.length} 项已完成；草案已重置 ${plan.resetFields.length} 个仅属于上月的字段。`
      : "等待建立目标月份草案。";
    ui.byId("monthlyBusinessBadge").textContent = pending.length
      ? `${pending.length} 项待核对`
      : "已完成";
    ui.byId("monthlyBusinessBadge").classList.toggle(
      "warning",
      Boolean(pending.length),
    );
    ui.byId("monthlyBusinessList").innerHTML = items
      .map(
        (item) => `<article class="diagnostic-card">
          <div class="diagnostic-card-head">
            <label class="checkbox-row">
              <input
                type="checkbox"
                data-monthly-business-id="${ui.escapeHtml(item.id)}"
                ${item.confirmed ? "checked" : ""}
                ${item.automatic ? "disabled" : ""}
              />
              <span>
                <strong>${ui.escapeHtml(item.label)}</strong>
                <small>来源：${ui.escapeHtml(item.source)}</small>
              </span>
            </label>
            <span class="status-chip ${item.confirmed ? "" : "warning"}">
              ${item.automatic ? "规则已处理" : item.confirmed ? "已核对" : "待核对"}
            </span>
          </div>
          <p>${ui.escapeHtml(item.detail)}</p>
          <div class="diagnostic-fields">
            ${item.fields.map((field) => `<span>${ui.escapeHtml(field)}</span>`).join("")}
          </div>
        </article>`,
      )
      .join("");
    ui.byId("monthlyBusinessList")
      .querySelectorAll("[data-monthly-business-id]")
      .forEach((checkbox) => {
        checkbox.addEventListener("change", () => {
          const item = items.find(
            (candidate) =>
              candidate.id === checkbox.dataset.monthlyBusinessId,
          );
          if (item && !item.automatic) {
            item.confirmed = checkbox.checked;
            renderMonthlyBusiness();
            ui.updateExportState?.();
          }
        });
      });
  }

  function renderHistory() {
    ui.byId("historyBadge").textContent = String(ui.state.history.length);
    ui.byId("historyList").innerHTML = ui.state.history.length
      ? [...ui.state.history]
          .reverse()
          .map(
            (item) => `<article class="history-item">
              <div class="history-item-head">
                <strong>${ui.escapeHtml(item.label)}</strong><small>${ui.escapeHtml(item.time)}</small>
              </div>
              <p>${ui.escapeHtml(item.detail)}</p>
            </article>`,
          )
          .join("")
      : '<div class="empty-state">还没有已应用变更。</div>';
    renderStats();
  }

  function renderSocialCandidates() {
    const candidates = ui.state.social.adjustedCandidates;
    ui.byId("socialEmpty").hidden = candidates.length > 0;
    ui.byId("socialTable").querySelector("tbody").innerHTML = candidates
      .map(
        (item) => `<tr>
          <td>${privateValue(item.name, "name")}</td>
          <td>${item.coverage}/12</td>
          <td>${privateValue(item.annualTotal)}</td>
          <td>${item.divisor}</td>
          <td>${privateValue(item.candidate)}</td>
          <td>${privateValue(item.adjusted)}</td>
          <td><span class="status-dot ${item.status === "ready" ? "" : "warning"}">${item.status === "ready" ? "齐全" : "待确认"}</span></td>
        </tr>`,
      )
      .join("");
    const incomplete = candidates.filter(
      (item) => item.status !== "ready",
    ).length;
    ui.byId("socialSummary").textContent = candidates.length
      ? `${candidates.length} 人形成候选，${incomplete} 人的历史月份不足 12 个月。`
      : "等待计算。";
    ui.byId("downloadSocialCsvBtn").disabled = !candidates.length;
  }

  function renderAll() {
    renderStats();
    renderPeopleTable();
    renderProposals();
    renderSources();
    renderAttachments();
    renderFormulaDiagnostics();
    renderSourceRouting();
    ui.renderRequirementCoverage?.();
    renderMonthlyBusiness();
    renderDiagnostics();
    renderHistory();
    renderSocialCandidates();
    ui.updateExportState?.();
  }

  Object.assign(ui, {
    privateValue,
    renderStats,
    renderPeopleTable,
    renderProposals,
    renderSources,
    renderAttachments,
    renderFormulaDiagnostics,
    renderSourceRouting,
    renderMonthlyBusiness,
    renderDiagnostics,
    renderHistory,
    renderSocialCandidates,
    renderAll,
  });
})();
