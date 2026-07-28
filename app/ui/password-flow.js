(() => {
  "use strict";

  const ui = window.PayrollLocal.ui;
  const PASSWORD_CODES = new Set([
    "WORKBOOK_PASSWORD_REQUIRED",
    "WORKBOOK_PASSWORD_INCORRECT",
  ]);
  let passwordDialog = null;
  let activePrompt = null;

  function ensurePasswordDialog() {
    if (passwordDialog) {
      return passwordDialog;
    }
    const dialog = document.createElement("dialog");
    dialog.id = "workbookPasswordDialog";
    dialog.className = "password-dialog";
    dialog.setAttribute("aria-labelledby", "workbookPasswordTitle");
    dialog.innerHTML = `
      <form id="workbookPasswordForm" class="password-dialog-card">
        <div class="password-dialog-head">
          <span class="password-dialog-icon" aria-hidden="true">锁</span>
          <div>
            <h2 id="workbookPasswordTitle">请输入密码</h2>
            <p>
              <strong id="workbookPasswordFile">所选工作簿</strong>
              已加密，请输入打开密码。
            </p>
          </div>
        </div>
        <label class="field" for="workbookPasswordInput">
          <span>工作簿密码</span>
          <input
            id="workbookPasswordInput"
            type="password"
            autocomplete="off"
            maxlength="255"
            spellcheck="false"
          />
        </label>
        <p id="workbookPasswordError" class="password-dialog-error" hidden>
          密码不正确，请重新输入。
        </p>
        <p class="password-dialog-note">
          密码只用于本次本地读取，不会保存到工具或导出文件。
        </p>
        <div class="password-dialog-actions">
          <button
            id="workbookPasswordCancel"
            class="btn btn-ghost"
            type="button"
          >
            取消
          </button>
          <button class="btn btn-primary" type="submit">打开工作簿</button>
        </div>
      </form>
    `;
    document.body.appendChild(dialog);
    const form = dialog.querySelector("#workbookPasswordForm");
    const input = dialog.querySelector("#workbookPasswordInput");
    const errorNode = dialog.querySelector("#workbookPasswordError");

    function finish(value) {
      const resolver = activePrompt?.resolve;
      activePrompt = null;
      input.value = "";
      if (dialog.open) {
        dialog.close();
      }
      resolver?.(value);
    }

    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const password = input.value;
      if (!password) {
        errorNode.textContent = "请输入密码。";
        errorNode.hidden = false;
        input.focus();
        return;
      }
      finish(password);
    });
    dialog
      .querySelector("#workbookPasswordCancel")
      .addEventListener("click", () => finish(null));
    dialog.addEventListener("cancel", (event) => {
      event.preventDefault();
      finish(null);
    });
    passwordDialog = dialog;
    return dialog;
  }

  function requestWorkbookPassword(fileName, incorrect = false) {
    const dialog = ensurePasswordDialog();
    const input = dialog.querySelector("#workbookPasswordInput");
    const errorNode = dialog.querySelector("#workbookPasswordError");
    dialog.querySelector("#workbookPasswordFile").textContent =
      fileName || "所选工作簿";
    errorNode.textContent = incorrect
      ? "密码不正确，请重新输入。"
      : "";
    errorNode.hidden = !incorrect;
    input.value = "";
    return new Promise((resolve) => {
      activePrompt = { resolve };
      dialog.showModal();
      window.requestAnimationFrame(() => input.focus());
    });
  }

  async function loadWorkbookFile(file) {
    let password = "";
    while (true) {
      try {
        return await window.XlsxEngine.SourceWorkbook.load(file, {
          password,
        });
      } catch (error) {
        if (!PASSWORD_CODES.has(error.code)) {
          throw error;
        }
        const nextPassword = await requestWorkbookPassword(
          file.name,
          error.code === "WORKBOOK_PASSWORD_INCORRECT",
        );
        if (nextPassword === null) {
          const cancelled = new Error(`${file.name} 已取消读取`);
          cancelled.code = "WORKBOOK_PASSWORD_CANCELLED";
          throw cancelled;
        }
        password = nextPassword;
      }
    }
  }

  Object.assign(ui, {
    loadWorkbookFile,
    requestWorkbookPassword,
  });
})();
