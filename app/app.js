(() => {
  "use strict";

  const ui = window.PayrollLocal?.ui;
  if (!ui?.bindWorkbookFlow || !window.PayrollEngine) {
    throw new Error("应用模块加载不完整");
  }

  function bootstrap() {
    ui.bindTabs();
    ui.bindWorkbookFlow();
    ui.bindAttachmentFlow();
    ui.bindChangeFlow();
    ui.bindSocialBaseFlow();
    ui.bindExportFlow();
    ui.bindDropZone(
      document.querySelector('label[for="baseWorkbookInput"]'),
      ui.byId("baseWorkbookInput"),
    );
    ui.bindDropZone(
      document.querySelector('label[for="sourceFilesInput"]'),
      ui.byId("sourceFilesInput"),
    );
    ui.bindDropZone(
      document.querySelector('label[for="attachmentFilesInput"]'),
      ui.byId("attachmentFilesInput"),
    );
    ui.bindDropZone(
      document.querySelector('label[for="annualPayrollFilesInput"]'),
      ui.byId("annualPayrollFilesInput"),
    );
    ui.bindDropZone(
      document.querySelector('label[for="crossYearFilesInput"]'),
      ui.byId("crossYearFilesInput"),
    );
    window.addEventListener("unhandledrejection", (event) => {
      ui.toast(event.reason?.message || "发生未捕获错误", "error");
    });
    window.addEventListener("error", (event) => {
      if (event.message) {
        ui.toast(event.message, "error");
      }
    });
  }

  bootstrap();
})();
