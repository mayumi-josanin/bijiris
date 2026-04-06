const TICKET_END_FORM_SLUG = "bijiris-ticket-end";
const TICKET_SHEET_FIELD_KEY = "ticket_sheet_number";
const RESPONDENT_NAME_STORAGE_KEY = "bijiris_respondent_name";
const INSTALL_GUIDE_STORAGE_KEY = "bijiris_install_guide_seen";

const formState = {
  slug: getSlugFromPath(),
  form: null,
  ticketSheetNumber: "",
  historyReturnView: "selector",
};

const formEls = {};
const appLifecycle = {
  hiddenAt: 0,
  refreshing: false,
  swRefreshPending: false,
};

const installState = {
  deferredPrompt: null,
  guideOpen: false,
  registration: null,
};

document.addEventListener("DOMContentLoaded", () => {
  bindRespondentElements();
  setupPwaFeatures();
  bootstrap();
});

function getSlugFromPath() {
  const segments = window.location.pathname.split("/").filter(Boolean);
  if (segments[0] === "f" && segments[1]) {
    return segments[1];
  }
  return null;
}

function bindRespondentElements() {
  formEls.selectorPanel = document.getElementById("selectorPanel");
  formEls.selectorList = document.getElementById("selectorList");
  formEls.installCard = document.getElementById("installCard");
  formEls.installStatusBadge = document.getElementById("installStatusBadge");
  formEls.installDescription = document.getElementById("installDescription");
  formEls.installAppButton = document.getElementById("installAppButton");
  formEls.installGuideToggle = document.getElementById("installGuideToggle");
  formEls.installGuide = document.getElementById("installGuide");
  formEls.openHistoryFromSelector = document.getElementById("openHistoryFromSelector");
  formEls.reloadAppFromSelector = document.getElementById("reloadAppFromSelector");
  formEls.backToSelector = document.getElementById("backToSelector");
  formEls.openHistoryFromHeader = document.getElementById("openHistoryFromHeader");
  formEls.respondentHeader = document.getElementById("respondentHeader");
  formEls.reloadAppButton = document.getElementById("reloadAppButton");
  formEls.formTitle = document.getElementById("formTitle");
  formEls.formDescription = document.getElementById("formDescription");
  formEls.formErrorBanner = document.getElementById("formErrorBanner");
  formEls.ticketStepPanel = document.getElementById("ticketStepPanel");
  formEls.ticketSheetInput = document.getElementById("ticketSheetInput");
  formEls.ticketStepContinue = document.getElementById("ticketStepContinue");
  formEls.ticketStepError = document.getElementById("ticketStepError");
  formEls.historyPanel = document.getElementById("historyPanel");
  formEls.historyBackButton = document.getElementById("historyBackButton");
  formEls.historyRespondentName = document.getElementById("historyRespondentName");
  formEls.historySearchButton = document.getElementById("historySearchButton");
  formEls.historySearchError = document.getElementById("historySearchError");
  formEls.historySummary = document.getElementById("historySummary");
  formEls.historyResults = document.getElementById("historyResults");
  formEls.respondentForm = document.getElementById("respondentForm");
  formEls.ticketSummary = document.getElementById("ticketSummary");
  formEls.ticketSummaryValue = document.getElementById("ticketSummaryValue");
  formEls.ticketStepEdit = document.getElementById("ticketStepEdit");
  formEls.categoryLabel = document.getElementById("categoryLabel");
  formEls.categorySelect = document.getElementById("categorySelect");
  formEls.customFields = document.getElementById("customFields");
  formEls.submitError = document.getElementById("submitError");
  formEls.submitButton = document.getElementById("submitButton");
  formEls.successPanel = document.getElementById("successPanel");
  formEls.successMessage = document.getElementById("successMessage");
  formEls.resetButton = document.getElementById("resetButton");
  formEls.openHistoryFromSuccess = document.getElementById("openHistoryFromSuccess");

  formEls.respondentForm.addEventListener("submit", onSubmit);
  formEls.respondentForm.addEventListener("change", onFormInputChange);
  formEls.respondentForm.addEventListener("input", onFormInputChange);
  formEls.ticketStepContinue.addEventListener("click", onTicketStepContinue);
  formEls.ticketSheetInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      onTicketStepContinue();
    }
  });
  formEls.ticketStepEdit.addEventListener("click", () => {
    showTicketStep();
  });
  formEls.openHistoryFromSelector.addEventListener("click", () => {
    openHistoryPanel(getCurrentRespondentName());
  });
  formEls.openHistoryFromHeader.addEventListener("click", () => {
    openHistoryPanel(getCurrentRespondentName());
  });
  formEls.openHistoryFromSuccess.addEventListener("click", () => {
    openHistoryPanel(getCurrentRespondentName());
  });
  formEls.installAppButton.addEventListener("click", onInstallApp);
  formEls.installGuideToggle.addEventListener("click", toggleInstallGuide);
  formEls.historyBackButton.addEventListener("click", restoreFromHistoryPanel);
  formEls.historySearchButton.addEventListener("click", onHistorySearch);
  formEls.historyRespondentName.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      onHistorySearch();
    }
  });
  formEls.backToSelector.addEventListener("click", () => {
    window.location.href = "/f/";
  });
  formEls.reloadAppFromSelector.addEventListener("click", () => {
    forceAppRefresh(true);
  });
  formEls.reloadAppButton.addEventListener("click", () => {
    forceAppRefresh(true);
  });
  formEls.resetButton.addEventListener("click", resetForNextResponse);
  document.addEventListener("visibilitychange", onVisibilityChange);
  window.addEventListener("pageshow", onPageShow);
}

async function bootstrap() {
  if (formState.slug) {
    await loadForm(formState.slug);
    return;
  }
  await loadFormList();
}

function setupPwaFeatures() {
  renderInstallCard();
  registerRespondentServiceWorker();

  window.addEventListener("beforeinstallprompt", (event) => {
    event.preventDefault();
    installState.deferredPrompt = event;
    renderInstallCard();
  });

  window.addEventListener("appinstalled", () => {
    installState.deferredPrompt = null;
    installState.guideOpen = false;
    rememberInstallGuideSeen();
    renderInstallCard();
  });

  const displayMode = window.matchMedia?.("(display-mode: standalone)");
  displayMode?.addEventListener?.("change", () => {
    renderInstallCard();
  });
}

async function registerRespondentServiceWorker() {
  if (!("serviceWorker" in navigator)) {
    renderInstallCard();
    return;
  }
  try {
    const registration = await navigator.serviceWorker.register("/sw.js", { scope: "/" });
    installState.registration = registration;
    navigator.serviceWorker.addEventListener("controllerchange", () => {
      if (appLifecycle.swRefreshPending) {
        return;
      }
      appLifecycle.swRefreshPending = true;
      window.location.reload();
    });
    await registration.update();
  } catch (error) {
    console.warn("service worker registration skipped", error);
  } finally {
    renderInstallCard();
  }
}

async function loadFormList() {
  try {
    const response = await fetch("/api/public/forms");
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "アンケート一覧を読み込めませんでした。");
    }
    renderSelector(data.forms || []);
    showSelector();
  } catch (error) {
    showBanner(error.message);
  }
}

function renderInstallCard() {
  if (!formEls.installCard) {
    return;
  }

  const installed = isStandaloneApp();
  const installable = !!installState.deferredPrompt;
  const iosManual = isIosDevice() && !installed;
  const androidManual = isAndroidDevice() && !installed;
  const supportedContext = installable || iosManual || androidManual || installed;

  formEls.installCard.classList.toggle("hidden", !supportedContext);
  if (!supportedContext) {
    return;
  }

  formEls.installStatusBadge.classList.toggle("hidden", !installed);
  formEls.installStatusBadge.textContent = installed ? "インストール済み" : "";

  if (installed) {
    formEls.installDescription.textContent = "ホーム画面からアプリのように開けます。最新版にしたいときは更新ボタンを押してください。";
    formEls.installAppButton.textContent = "アプリとして利用中";
    formEls.installAppButton.disabled = true;
    formEls.installGuideToggle.classList.add("hidden");
    formEls.installGuide.classList.add("hidden");
    formEls.installGuide.innerHTML = "";
    return;
  }

  formEls.installAppButton.disabled = false;
  if (installable) {
    formEls.installDescription.textContent = "この端末ではそのままインストールできます。追加すると次回からホーム画面からすぐ開けます。";
    formEls.installAppButton.textContent = "ホーム画面に追加";
    formEls.installGuideToggle.classList.add("hidden");
    formEls.installGuide.classList.add("hidden");
    formEls.installGuide.innerHTML = "";
    return;
  }

  formEls.installAppButton.textContent = "追加方法を見る";
  formEls.installGuideToggle.classList.remove("hidden");
  formEls.installGuideToggle.textContent = installState.guideOpen ? "閉じる" : "追加方法を見る";

  if (iosManual) {
    formEls.installDescription.textContent = "iPhone / iPad は Safari の共有メニューからホーム画面に追加してください。";
    formEls.installGuide.innerHTML = buildIosInstallGuide();
  } else {
    formEls.installDescription.textContent = "Android は Chrome のメニューからホーム画面に追加できます。";
    formEls.installGuide.innerHTML = buildGenericInstallGuide();
  }

  const shouldShowGuide = installState.guideOpen || !hasSeenInstallGuide();
  formEls.installGuide.classList.toggle("hidden", !shouldShowGuide);
}

function buildIosInstallGuide() {
  return `
    <strong>iPhoneでの追加方法</strong>
    <ol class="install-guide-list">
      <li>このページを Safari で開きます。</li>
      <li>画面下の共有ボタンを押します。</li>
      <li>「ホーム画面に追加」を選びます。</li>
      <li>右上の「追加」を押すと、ビジリスがホーム画面に入ります。</li>
    </ol>
  `;
}

function buildGenericInstallGuide() {
  return `
    <strong>追加しやすい開き方</strong>
    <ol class="install-guide-list">
      <li>Android は Chrome で開きます。</li>
      <li>iPhone は Safari で開きます。</li>
      <li>表示された案内、またはブラウザメニューの「ホーム画面に追加」を使います。</li>
    </ol>
  `;
}

async function onInstallApp() {
  if (isStandaloneApp()) {
    return;
  }
  if (installState.deferredPrompt) {
    const prompt = installState.deferredPrompt;
    installState.deferredPrompt = null;
    await prompt.prompt();
    try {
      await prompt.userChoice;
    } catch (error) {
      console.warn("install prompt skipped", error);
    }
    renderInstallCard();
    return;
  }
  installState.guideOpen = true;
  rememberInstallGuideSeen();
  renderInstallCard();
}

function toggleInstallGuide() {
  installState.guideOpen = !installState.guideOpen;
  if (installState.guideOpen) {
    rememberInstallGuideSeen();
  }
  renderInstallCard();
}

function isStandaloneApp() {
  return window.matchMedia?.("(display-mode: standalone)").matches || window.navigator.standalone === true;
}

function isIosDevice() {
  const ua = window.navigator.userAgent || "";
  const platform = window.navigator.platform || "";
  const touchMac = platform === "MacIntel" && window.navigator.maxTouchPoints > 1;
  return /iPhone|iPad|iPod/i.test(ua) || touchMac;
}

function isAndroidDevice() {
  return /Android/i.test(window.navigator.userAgent || "");
}

function hasSeenInstallGuide() {
  try {
    return window.localStorage.getItem(INSTALL_GUIDE_STORAGE_KEY) === "1";
  } catch (error) {
    return false;
  }
}

function rememberInstallGuideSeen() {
  try {
    window.localStorage.setItem(INSTALL_GUIDE_STORAGE_KEY, "1");
  } catch (error) {
    console.warn("storage skipped", error);
  }
}

function renderSelector(forms) {
  if (!forms.length) {
    formEls.selectorList.innerHTML = `<div class="empty-state">公開中のアンケートがありません。</div>`;
    return;
  }

  formEls.selectorList.innerHTML = forms
    .map(
      (form) => `
        <button class="selector-card" type="button" data-form-slug="${escapeAttribute(form.slug)}">
          <div>
            <h2>${escapeHtml(form.title)}</h2>
            <p class="muted">${escapeHtml(form.description || "このアンケートに回答します。")}</p>
          </div>
          <span class="pill">${form.questionCount}項目</span>
        </button>
      `
    )
    .join("");

  formEls.selectorList.querySelectorAll("[data-form-slug]").forEach((button) => {
    button.addEventListener("click", () => {
      window.location.href = `/f/${button.dataset.formSlug}`;
    });
  });
}

async function loadForm(slug) {
  try {
    const response = await fetch(`/api/public/forms/${encodeURIComponent(slug)}`);
    const data = await response.json();
    if (!response.ok) {
      if (response.status === 404) {
        formState.slug = null;
        showBanner("このショートカットは古くなっています。最新のアンケート一覧を表示します。");
        await loadFormList();
        return;
      }
      throw new Error(data.error || "フォームを読み込めませんでした。");
    }
    formState.form = data.form;
    renderForm();
  } catch (error) {
    showBanner(error.message);
  }
}

function renderForm() {
  formState.ticketSheetNumber = "";
  formEls.ticketSheetInput.value = "";
  formEls.ticketStepError.textContent = "";
  document.title = `${formState.form.title} | ビジリス`;
  formEls.formTitle.textContent = formState.form.title;
  formEls.formDescription.textContent = formState.form.description || "必要事項を入力して送信してください。";
  formEls.categoryLabel.textContent = formState.form.categoryLabel || "分類";
  formEls.categorySelect.innerHTML = formState.form.categoryOptions
    .map((option) => `<option value="${escapeAttribute(option)}">${escapeHtml(option)}</option>`)
    .join("");
  formEls.customFields.innerHTML = formState.form.fields.map(renderFieldMarkup).join("");
  const nameInput = getRespondentNameInput();
  if (nameInput) {
    nameInput.value = getStoredRespondentName();
  }
  syncTicketSummary();
  applyVisibilityRules();
  if (requiresTicketStep()) {
    showTicketStep();
    return;
  }
  showForm();
}

function getRespondentNameInput() {
  return formEls.respondentForm.querySelector('[name="respondent_name"]');
}

function getStoredRespondentName() {
  try {
    return String(window.localStorage.getItem(RESPONDENT_NAME_STORAGE_KEY) || "");
  } catch (error) {
    return "";
  }
}

function saveStoredRespondentName(value) {
  const name = String(value || "").trim();
  try {
    if (name) {
      window.localStorage.setItem(RESPONDENT_NAME_STORAGE_KEY, name);
    } else {
      window.localStorage.removeItem(RESPONDENT_NAME_STORAGE_KEY);
    }
  } catch (error) {
    console.warn("storage skipped", error);
  }
}

function getCurrentRespondentName() {
  return getRespondentNameInput()?.value.trim() || getStoredRespondentName();
}

function renderFieldMarkup(field) {
  const visibilityValues = (field.visibilityValues || []).join("||");
  const wrapperStart = `
    <div
      class="dynamic-field"
      data-field-card
      data-field-key="${escapeAttribute(field.key)}"
      data-visible-from="${escapeAttribute(field.visibilityFieldKey || "")}"
      data-visible-values="${escapeAttribute(visibilityValues)}"
    >
  `;
  const helpText = field.helpText ? `<p class="field-help">${escapeHtml(field.helpText)}</p>` : "";

  if (field.type === "long_text") {
    return `
      ${wrapperStart}
        <label class="field">
          <span>${escapeHtml(field.label)}${field.required ? " *" : ""}</span>
          <textarea
            name="${escapeAttribute(field.key)}"
            rows="4"
            ${field.required ? "required" : ""}
            placeholder="${escapeAttribute(field.placeholder || "")}"
          ></textarea>
        </label>
        ${helpText}
      </div>
    `;
  }

  if (field.type === "select") {
    return `
      ${wrapperStart}
        <label class="field">
          <span>${escapeHtml(field.label)}${field.required ? " *" : ""}</span>
          <select name="${escapeAttribute(field.key)}" ${field.required ? "required" : ""}>
            <option value="">選択してください</option>
            ${field.options.map((option) => `<option value="${escapeAttribute(option)}">${escapeHtml(option)}</option>`).join("")}
          </select>
        </label>
        ${helpText}
      </div>
    `;
  }

  if (field.type === "radio") {
    return `
      ${wrapperStart}
        <fieldset class="field-set">
          <legend>${escapeHtml(field.label)}${field.required ? " *" : ""}</legend>
          <div class="choice-row">
            ${field.options
              .map(
                (option) => `
                  <label class="choice-pill">
                    <input
                      type="radio"
                      name="${escapeAttribute(field.key)}"
                      value="${escapeAttribute(option)}"
                      ${field.required ? "required" : ""}
                    />
                    <span>${escapeHtml(option)}</span>
                  </label>
                `
              )
              .join("")}
          </div>
        </fieldset>
        ${helpText}
      </div>
    `;
  }

  if (field.type === "checkbox") {
    return `
      ${wrapperStart}
        <fieldset class="field-set">
          <legend>${escapeHtml(field.label)}${field.required ? " *" : ""}</legend>
          <div class="choice-row">
            ${field.options
              .map(
                (option) => `
                  <label class="choice-pill">
                    <input
                      type="checkbox"
                      name="${escapeAttribute(field.key)}"
                      value="${escapeAttribute(option)}"
                    />
                    <span>${escapeHtml(option)}</span>
                  </label>
                `
              )
              .join("")}
            ${
              field.allowOther
                ? `
                  <label class="choice-pill choice-pill-other">
                    <input
                      type="checkbox"
                      name="${escapeAttribute(field.key)}"
                      value="__other__"
                      data-other-toggle="${escapeAttribute(field.key)}"
                    />
                    <span>その他</span>
                  </label>
                  <label class="field other-input-wrap hidden" data-other-input-wrap="${escapeAttribute(field.key)}">
                    <span>その他の内容</span>
                    <input
                      type="text"
                      name="${escapeAttribute(field.key)}__other_text"
                      data-other-input="${escapeAttribute(field.key)}"
                      disabled
                      placeholder="自由にご記入ください"
                    />
                  </label>
                `
                : ""
            }
          </div>
        </fieldset>
        ${helpText}
      </div>
    `;
  }

  if (field.type === "file") {
    return `
      ${wrapperStart}
        <label class="field">
          <span>${escapeHtml(field.label)}${field.required ? " *" : ""}</span>
          <input
            type="file"
            name="${escapeAttribute(field.key)}"
            accept="${escapeAttribute(field.accept || "image/*")}"
            ${field.allowMultiple ? "multiple" : ""}
            ${field.required ? "required" : ""}
          />
        </label>
        ${helpText}
        <div class="image-grid" data-file-preview="${escapeAttribute(field.key)}"></div>
      </div>
    `;
  }

  return `
    ${wrapperStart}
      <label class="field">
        <span>${escapeHtml(field.label)}${field.required ? " *" : ""}</span>
        <input
          type="text"
          name="${escapeAttribute(field.key)}"
          ${field.required ? "required" : ""}
          placeholder="${escapeAttribute(field.placeholder || "")}"
        />
      </label>
      ${helpText}
    </div>
  `;
}

function onFormInputChange(event) {
  if (event.target.matches('input[type="file"]')) {
    renderFilePreviewForInput(event.target);
  }
  if (event.target.matches("[data-other-toggle]")) {
    syncOtherInputState(event.target.dataset.otherToggle);
  }
  applyVisibilityRules();
}

function renderFilePreviewForInput(input) {
  const preview = formEls.customFields.querySelector(`[data-file-preview="${CSS.escape(input.name)}"]`);
  if (!preview) {
    return;
  }
  const files = Array.from(input.files || []);
  if (!files.length) {
    preview.innerHTML = "";
    return;
  }
  preview.innerHTML = files
    .map((file) => {
      const url = URL.createObjectURL(file);
      return `
        <figure class="preview-tile">
          <img src="${escapeAttribute(url)}" alt="${escapeAttribute(file.name)}" />
          <figcaption>${escapeHtml(file.name)}</figcaption>
        </figure>
      `;
    })
    .join("");
}

function applyVisibilityRules() {
  formEls.customFields.querySelectorAll("[data-field-card]").forEach((card) => {
    const dependsOn = card.dataset.visibleFrom;
    if (!dependsOn) {
      enableFieldCard(card, true);
      return;
    }
    const requiredValues = (card.dataset.visibleValues || "")
      .split("||")
      .map((value) => value.trim())
      .filter(Boolean);
    const currentValues = getFieldValues(dependsOn);
    const visible = requiredValues.some((value) => currentValues.includes(value));
    enableFieldCard(card, visible);
  });
  formEls.customFields.querySelectorAll("[data-other-toggle]").forEach((checkbox) => {
    syncOtherInputState(checkbox.dataset.otherToggle);
  });
}

function getFieldValues(fieldKey) {
  const nodes = Array.from(formEls.respondentForm.querySelectorAll(`[name="${CSS.escape(fieldKey)}"]`));
  if (!nodes.length) {
    return [];
  }
  if (nodes[0].type === "radio") {
    const selected = nodes.find((node) => node.checked);
    return selected ? [selected.value] : [];
  }
  if (nodes[0].type === "checkbox") {
    return nodes.filter((node) => node.checked).map((node) => node.value);
  }
  return nodes[0].value ? [nodes[0].value] : [];
}

function enableFieldCard(card, visible) {
  card.classList.toggle("hidden", !visible);
  card.querySelectorAll("input, select, textarea").forEach((input) => {
    input.disabled = !visible;
  });
}

function syncOtherInputState(fieldKey) {
  const toggle = formEls.respondentForm.querySelector(`[data-other-toggle="${CSS.escape(fieldKey)}"]`);
  const wrapper = formEls.respondentForm.querySelector(`[data-other-input-wrap="${CSS.escape(fieldKey)}"]`);
  const input = formEls.respondentForm.querySelector(`[data-other-input="${CSS.escape(fieldKey)}"]`);
  if (!toggle || !wrapper || !input) {
    return;
  }
  const enabled = toggle.checked && !toggle.disabled;
  wrapper.classList.toggle("hidden", !enabled);
  input.disabled = !enabled;
  if (!enabled) {
    input.value = "";
  }
}

async function onSubmit(event) {
  event.preventDefault();
  formEls.submitError.textContent = "";
  formEls.submitButton.disabled = true;
  formEls.submitButton.textContent = "送信中...";

  try {
    const formData = new FormData(formEls.respondentForm);
    const respondentName = String(formData.get("respondent_name") || "").trim();
    if (requiresTicketStep()) {
      if (!formState.ticketSheetNumber) {
        showTicketStep();
        throw new Error("回数券が何枚目かを先に入力してください。");
      }
      formData.set(TICKET_SHEET_FIELD_KEY, formState.ticketSheetNumber);
    }
    const response = await fetch(`/api/public/forms/${encodeURIComponent(formState.slug)}/submit`, {
      method: "POST",
      body: formData,
    });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "送信に失敗しました。");
    }
    saveStoredRespondentName(respondentName);
    formEls.successMessage.textContent = data.message || "送信ありがとうございました。";
    showSuccessPanel();
  } catch (error) {
    formEls.submitError.textContent = error.message;
  } finally {
    formEls.submitButton.disabled = false;
    formEls.submitButton.textContent = "送信する";
  }
}

function resetForNextResponse() {
  formEls.respondentForm.reset();
  const nameInput = getRespondentNameInput();
  if (nameInput) {
    nameInput.value = getStoredRespondentName();
  }
  formState.ticketSheetNumber = "";
  formEls.ticketSheetInput.value = "";
  formEls.ticketStepError.textContent = "";
  formEls.customFields.querySelectorAll("[data-file-preview]").forEach((preview) => {
    preview.innerHTML = "";
  });
  formEls.submitError.textContent = "";
  formEls.successPanel.classList.add("hidden");
  syncTicketSummary();
  applyVisibilityRules();
  if (requiresTicketStep()) {
    showTicketStep();
    return;
  }
  formEls.respondentForm.classList.remove("hidden");
}

async function onHistorySearch() {
  const respondentName = formEls.historyRespondentName.value.trim();
  formEls.historySearchError.textContent = "";
  formEls.historySearchButton.disabled = true;
  formEls.historySearchButton.textContent = "読み込み中...";
  try {
    if (!respondentName) {
      throw new Error("お名前を入力してください。");
    }
    const response = await fetch(`/api/public/respondents/history?name=${encodeURIComponent(respondentName)}`);
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(data.error || "回答履歴を読み込めませんでした。");
    }
    saveStoredRespondentName(data.respondent?.respondentName || respondentName);
    renderRespondentHistoryResults(data.respondent || null, data.history || []);
  } catch (error) {
    formEls.historySummary.classList.add("hidden");
    formEls.historySummary.textContent = "";
    formEls.historyResults.innerHTML = "";
    formEls.historySearchError.textContent = error.message;
  } finally {
    formEls.historySearchButton.disabled = false;
    formEls.historySearchButton.textContent = "履歴を表示";
  }
}

function renderRespondentHistoryResults(respondent, history) {
  const orderedHistory = [...history].sort(compareHistoryItemsByDateAsc);
  const responseCount = Number(respondent?.responseCount || history.length || 0);
  const lastResponseAt = respondent?.lastResponseAt
    ? formatDateTime(respondent.lastResponseAt)
    : orderedHistory[orderedHistory.length - 1]?.createdAt
      ? formatDateTime(orderedHistory[orderedHistory.length - 1].createdAt)
      : "まだ回答がありません。";
  formEls.historySummary.innerHTML = `
    <strong>${escapeHtml(respondent?.respondentName || formEls.historyRespondentName.value.trim())}</strong>
    <div>回答 ${responseCount}件 / 最新 ${escapeHtml(lastResponseAt)} / 古い日付順</div>
  `;
  formEls.historySummary.classList.remove("hidden");
  if (!orderedHistory.length) {
    formEls.historyResults.innerHTML = `<div class="empty-state">まだ回答履歴がありません。</div>`;
    return;
  }
  formEls.historyResults.innerHTML = orderedHistory.map(renderPublicHistoryCard).join("");
  bindPublicHistoryToggles();
}

function renderPublicHistoryCard(item) {
  const detailId = `history-detail-${escapeAttribute(String(item.id || ""))}`;
  const answerMarkup = (item.answers || []).length
    ? (item.answers || [])
        .map(
          (answer) => `
            <div class="respondent-history-answer-row">
              <dt>${escapeHtml(answer.label || "-")}</dt>
              <dd>${escapeHtml(answer.value || "-")}</dd>
            </div>
          `
        )
        .join("")
    : `<div class="empty-state">回答内容はありません。</div>`;
  return `
    <article class="respondent-history-card stack-gap compact-gap">
      <div class="spread respondent-history-head">
        <div>
          <h3>${escapeHtml(item.formTitle || "アンケート")}</h3>
          <p class="muted">${escapeHtml(formatDateTime(item.createdAt || ""))}</p>
        </div>
        ${item.category && item.category !== item.formTitle ? `<span class="pill respondent-history-category">${escapeHtml(item.category)}</span>` : ""}
      </div>
      <div class="button-row respondent-history-actions">
        <button
          class="secondary-button compact-button respondent-history-toggle"
          type="button"
          data-history-toggle
          data-detail-id="${detailId}"
        >
          詳細を見る
        </button>
      </div>
      <div id="${detailId}" class="respondent-history-detail hidden">
        <div class="respondent-history-answer-list">
          ${answerMarkup}
        </div>
        ${
          (item.files || []).length
            ? `
              <div class="stack-gap compact-gap">
                <strong class="respondent-history-file-title">提出ファイル</strong>
                <div class="respondent-history-file-list">
                  ${item.files
                    .map(
                      (fileItem) => `
                        <span class="respondent-history-file-chip">
                          ${escapeHtml(fileItem.label || "画像")} / ${escapeHtml(fileItem.originalName || "ファイル")}
                        </span>
                      `
                    )
                    .join("")}
                </div>
              </div>
            `
            : ""
        }
      </div>
    </article>
  `;
}

function bindPublicHistoryToggles() {
  formEls.historyResults.querySelectorAll("[data-history-toggle]").forEach((button) => {
    button.addEventListener("click", () => {
      const detail = formEls.historyResults.querySelector(`#${CSS.escape(button.dataset.detailId || "")}`);
      if (!detail) {
        return;
      }
      const isHidden = detail.classList.contains("hidden");
      detail.classList.toggle("hidden", !isHidden);
      button.textContent = isHidden ? "詳細を閉じる" : "詳細を見る";
    });
  });
}

function compareHistoryItemsByDateAsc(a, b) {
  const aTime = historyItemTimestamp(a);
  const bTime = historyItemTimestamp(b);
  return aTime - bTime;
}

function historyItemTimestamp(item) {
  const value = item?.createdAt ? Date.parse(String(item.createdAt).replace(" ", "T")) : Number.NaN;
  return Number.isFinite(value) ? value : 0;
}

function currentView() {
  if (!formEls.historyPanel.classList.contains("hidden")) {
    return "history";
  }
  if (!formEls.selectorPanel.classList.contains("hidden")) {
    return "selector";
  }
  if (!formEls.successPanel.classList.contains("hidden")) {
    return "success";
  }
  if (!formEls.ticketStepPanel.classList.contains("hidden")) {
    return "ticket";
  }
  if (!formEls.respondentForm.classList.contains("hidden")) {
    return "form";
  }
  return formState.slug ? "form" : "selector";
}

function openHistoryPanel(preferredName = "") {
  const nextReturnView = currentView();
  formState.historyReturnView = nextReturnView === "history" ? formState.historyReturnView : nextReturnView;
  formEls.selectorPanel.classList.add("hidden");
  formEls.respondentHeader.classList.add("hidden");
  formEls.ticketStepPanel.classList.add("hidden");
  formEls.respondentForm.classList.add("hidden");
  formEls.successPanel.classList.add("hidden");
  formEls.historyPanel.classList.remove("hidden");
  formEls.historySearchError.textContent = "";
  const name = preferredName || getStoredRespondentName();
  formEls.historyRespondentName.value = name;
  if (name) {
    onHistorySearch();
    return;
  }
  formEls.historySummary.classList.add("hidden");
  formEls.historySummary.textContent = "";
  formEls.historyResults.innerHTML = `<div class="empty-state">お名前を入力して履歴を表示してください。</div>`;
}

function restoreFromHistoryPanel() {
  const targetView = formState.historyReturnView || (formState.slug ? "form" : "selector");
  formEls.historyPanel.classList.add("hidden");
  if (targetView === "selector") {
    showSelector();
    return;
  }
  if (targetView === "success") {
    showSuccessPanel();
    return;
  }
  if (targetView === "ticket") {
    showTicketStep();
    return;
  }
  showForm();
}

function onVisibilityChange() {
  if (document.visibilityState === "hidden") {
    appLifecycle.hiddenAt = Date.now();
    return;
  }
  if (!appLifecycle.hiddenAt || appLifecycle.refreshing) {
    return;
  }
  appLifecycle.hiddenAt = 0;
  if (hasPendingInput()) {
    return;
  }
  forceAppRefresh(false);
}

function onPageShow(event) {
  if (appLifecycle.refreshing) {
    return;
  }
  if (event.persisted && !hasPendingInput()) {
    forceAppRefresh(false);
  }
}

function showSelector() {
  formEls.selectorPanel.classList.remove("hidden");
  formEls.respondentHeader.classList.add("hidden");
  formEls.ticketStepPanel.classList.add("hidden");
  formEls.historyPanel.classList.add("hidden");
  formEls.respondentForm.classList.add("hidden");
  formEls.successPanel.classList.add("hidden");
}

function showForm() {
  if (requiresTicketStep() && !formState.ticketSheetNumber) {
    showTicketStep();
    return;
  }
  formEls.selectorPanel.classList.add("hidden");
  formEls.respondentHeader.classList.remove("hidden");
  formEls.backToSelector.classList.remove("hidden");
  formEls.openHistoryFromHeader.classList.remove("hidden");
  formEls.ticketStepPanel.classList.add("hidden");
  formEls.historyPanel.classList.add("hidden");
  formEls.respondentForm.classList.remove("hidden");
  formEls.successPanel.classList.add("hidden");
  const hasCategory = (formState.form.categoryOptions || []).length > 0;
  formEls.categorySelect.closest(".field").classList.toggle("hidden", !hasCategory);
  formEls.categorySelect.disabled = !hasCategory;
  formEls.categorySelect.required = hasCategory;
  syncTicketSummary();
}

function showTicketStep() {
  formEls.selectorPanel.classList.add("hidden");
  formEls.respondentHeader.classList.remove("hidden");
  formEls.backToSelector.classList.remove("hidden");
  formEls.openHistoryFromHeader.classList.remove("hidden");
  formEls.ticketStepPanel.classList.remove("hidden");
  formEls.historyPanel.classList.add("hidden");
  formEls.respondentForm.classList.add("hidden");
  formEls.successPanel.classList.add("hidden");
  syncTicketSummary();
}

function showSuccessPanel() {
  formEls.selectorPanel.classList.add("hidden");
  formEls.respondentHeader.classList.remove("hidden");
  formEls.backToSelector.classList.remove("hidden");
  formEls.openHistoryFromHeader.classList.remove("hidden");
  formEls.ticketStepPanel.classList.add("hidden");
  formEls.historyPanel.classList.add("hidden");
  formEls.respondentForm.classList.add("hidden");
  formEls.successPanel.classList.remove("hidden");
}

function formatDateTime(value) {
  if (!value) {
    return "-";
  }
  const normalized = String(value).replace(" ", "T");
  const date = new Date(normalized);
  if (Number.isNaN(date.getTime())) {
    return String(value);
  }
  return new Intl.DateTimeFormat("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function requiresTicketStep() {
  return formState.form?.slug === TICKET_END_FORM_SLUG;
}

function onTicketStepContinue() {
  try {
    const normalized = normalizeTicketSheetNumber(formEls.ticketSheetInput.value);
    formState.ticketSheetNumber = normalized.raw;
    formEls.ticketSheetInput.value = normalized.raw;
    formEls.ticketStepError.textContent = "";
    syncTicketSummary(normalized.label);
    showForm();
  } catch (error) {
    formEls.ticketStepError.textContent = error.message;
  }
}

function syncTicketSummary(forcedLabel = "") {
  const visible = requiresTicketStep() && !!formState.ticketSheetNumber;
  formEls.ticketSummary.classList.toggle("hidden", !visible);
  if (!visible) {
    formEls.ticketSummaryValue.textContent = "";
    return;
  }
  formEls.ticketSummaryValue.textContent = forcedLabel || `${formState.ticketSheetNumber}枚目`;
}

function normalizeTicketSheetNumber(value) {
  const normalized = String(value ?? "")
    .trim()
    .replace(/[０-９]/g, (digit) => String.fromCharCode(digit.charCodeAt(0) - 0xfee0));
  const match = normalized.match(/^(\d{1,3})(?:枚目)?$/);
  if (!match) {
    throw new Error("回数券が何枚目かを数字で入力してください。");
  }
  const number = Number.parseInt(match[1], 10);
  if (!Number.isFinite(number) || number <= 0) {
    throw new Error("回数券が何枚目かは1以上で入力してください。");
  }
  return {
    raw: String(number),
    label: `${number}枚目`,
  };
}

function showBanner(message) {
  formEls.formErrorBanner.textContent = message;
  formEls.formErrorBanner.classList.remove("hidden");
}

function hasPendingInput() {
  if (!formEls.successPanel.classList.contains("hidden")) {
    return false;
  }
  if (formState.ticketSheetNumber) {
    return true;
  }
  const formData = new FormData(formEls.respondentForm);
  for (const [key, value] of formData.entries()) {
    if (key === "category" && !String(value || "").trim()) {
      continue;
    }
    if (value instanceof File) {
      if (value.name) {
        return true;
      }
      continue;
    }
    if (String(value || "").trim()) {
      return true;
    }
  }
  return false;
}

async function forceAppRefresh(confirmIfDirty) {
  if (appLifecycle.refreshing) {
    return;
  }
  if (confirmIfDirty && hasPendingInput()) {
    const shouldReload = window.confirm("入力中の内容は失われます。最新版に更新しますか？");
    if (!shouldReload) {
      return;
    }
  }
  appLifecycle.refreshing = true;
  try {
    if (installState.registration) {
      await installState.registration.update();
      if (installState.registration.waiting) {
        installState.registration.waiting.postMessage({ type: "SKIP_WAITING" });
      }
    }
    if ("caches" in window) {
      const keys = await caches.keys();
      await Promise.all(keys.map((key) => caches.delete(key)));
    }
  } catch (error) {
    console.warn("cache clear skipped", error);
  }
  const url = new URL(window.location.href);
  url.searchParams.set("_refresh", String(Date.now()));
  window.location.replace(url.toString());
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replaceAll("`", "&#96;");
}
