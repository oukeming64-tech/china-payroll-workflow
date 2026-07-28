(() => {
  "use strict";

  const ui = window.PayrollLocal.ui;
  const state = {
    baseFile: null,
    workingSourceFile: null,
    workbook: null,
    mainSheetName: "",
    table: null,
    basePeriod: "",
    workingPeriod: "",
    targetPeriod: "",
    route: null,
    annualHistory: null,
    rolloverPlan: null,
    cumulativeResult: null,
    sourceRouting: null,
    periodEvidence: [],
    monthMarkers: [],
    externalLinks: [],
    formulaDiagnostics: null,
    proposals: [],
    sources: [],
    history: [],
    disabledRows: new Set(),
    personnelSheets: [],
    workbookSync: [],
    monthlyBusiness: {
      period: "",
      resetFields: [],
      items: [],
      errors: [],
    },
    showPrivateData: false,
    baseHash: "",
    attachments: {
      required: [],
      results: [],
      updates: [],
      errors: [],
      applied: false,
      detached: null,
    },
    social: {
      monthlyTables: [],
      targetFile: null,
      targetWorkbook: null,
      targetTable: null,
      targetSheetName: "",
      candidates: [],
      adjustedCandidates: [],
      updatesApplied: false,
    },
  };

  function byId(id) {
    return document.getElementById(id);
  }

  function setLoading(active, text = "正在读取…") {
    byId("loadingText").textContent = text;
    byId("loadingOverlay").hidden = !active;
  }

  function toast(message, type = "success") {
    const node = document.createElement("div");
    node.className = `toast ${type === "error" ? "error" : ""}`;
    node.textContent = message;
    byId("toastRegion").appendChild(node);
    window.setTimeout(() => node.remove(), 4200);
  }

  function openTab(tabName) {
    document.querySelectorAll(".tab").forEach((tab) => {
      tab.classList.toggle("active", tab.dataset.tab === tabName);
    });
    document.querySelectorAll(".tab-panel").forEach((panel) => {
      panel.classList.toggle("active", panel.id === `tab-${tabName}`);
    });
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function bindTabs() {
    document.querySelectorAll(".tab").forEach((tab) => {
      tab.addEventListener("click", () => openTab(tab.dataset.tab));
    });
    document.querySelectorAll("[data-open-tab]").forEach((button) => {
      button.addEventListener("click", () =>
        openTab(button.dataset.openTab),
      );
    });
  }

  function formatBytes(bytes) {
    if (bytes < 1024) {
      return `${bytes} B`;
    }
    if (bytes < 1024 * 1024) {
      return `${(bytes / 1024).toFixed(1)} KB`;
    }
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  function formatValue(value) {
    if (value === null || value === undefined || value === "") {
      return "—";
    }
    if (typeof value === "number") {
      return new Intl.NumberFormat("zh-CN", {
        maximumFractionDigits: 2,
      }).format(value);
    }
    return String(value);
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function money(value) {
    const numeric = Number(value);
    return Number.isFinite(numeric)
      ? new Intl.NumberFormat("zh-CN", {
          style: "currency",
          currency: "CNY",
          maximumFractionDigits: 2,
        }).format(numeric)
      : "—";
  }

  async function sha256File(file) {
    const digest = await crypto.subtle.digest(
      "SHA-256",
      await file.arrayBuffer(),
    );
    return [...new Uint8Array(digest)]
      .map((byte) => byte.toString(16).padStart(2, "0"))
      .join("");
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  function csvCell(value) {
    const text = String(value ?? "");
    return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
  }

  function downloadCsv(rows, filename) {
    const content = rows
      .map((row) => row.map(csvCell).join(","))
      .join("\r\n");
    downloadBlob(
      new Blob([`\uFEFF${content}`], { type: "text/csv;charset=utf-8" }),
      filename,
    );
  }

  function bindDropZone(label, input) {
    for (const eventName of ["dragenter", "dragover"]) {
      label.addEventListener(eventName, (event) => {
        event.preventDefault();
        label.classList.add("dragover");
      });
    }
    for (const eventName of ["dragleave", "drop"]) {
      label.addEventListener(eventName, (event) => {
        event.preventDefault();
        label.classList.remove("dragover");
      });
    }
    label.addEventListener("drop", (event) => {
      const files = event.dataTransfer?.files;
      if (!files?.length) {
        return;
      }
      const transfer = new DataTransfer();
      for (const file of files) {
        transfer.items.add(file);
      }
      input.files = transfer.files;
      input.dispatchEvent(new Event("change", { bubbles: true }));
    });
  }

  Object.assign(ui, {
    state,
    byId,
    setLoading,
    toast,
    openTab,
    bindTabs,
    bindDropZone,
    formatBytes,
    formatValue,
    escapeHtml,
    money,
    sha256File,
    downloadBlob,
    downloadCsv,
  });
})();
