const MEASUREMENT_CATEGORIES = [
  "モニター",
  "回数券",
  "トライアル",
  "単発",
  "初回お試し",
  "乗り放題キャンペーン",
  "その他",
];

const runtime = window.BijirisRuntime;

const state = {
  forms: [],
  stats: { formCount: 0, responseCount: 0, respondentCount: 0 },
  recentResponses: [],
  selectedFormId: null,
  selectedRespondentFormId: "",
  activeRespondentId: null,
  editingFormId: null,
  responseCategory: "",
  responseSearch: "",
  respondentSearch: "",
  measurementImportFeedback: null,
  respondentDirectory: [],
  measurementSearchQuery: "",
  measurementRespondentFilter: "",
  publicBaseUrl: "",
  settings: {
    configuredPublicBaseUrl: "",
    publicBaseUrl: "",
    publicBaseUrlSource: "auto",
    defaultPasswordInUse: false,
  },
  operationsStatus: null,
  backups: [],
};

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => Array.from(document.querySelectorAll(selector));

const els = {};

document.addEventListener("DOMContentLoaded", () => {
  bindElements();
  bindEvents();
  $("#loginHelp").textContent = "管理者パスワードを入力してください。初回は bijiris-admin です。";
  bootstrap();
});

function bindElements() {
  els.loginPanel = $("#loginPanel");
  els.adminApp = $("#adminApp");
  els.loginForm = $("#loginForm");
  els.loginPassword = $("#loginPassword");
  els.loginError = $("#loginError");
  els.logoutButton = $("#logoutButton");
  els.tabLinks = $$(".tab-link");
  els.panels = $$(".panel");
  els.statFormCount = $("#statFormCount");
  els.statResponseCount = $("#statResponseCount");
  els.statRespondentCount = $("#statRespondentCount");
  els.overviewForms = $("#overviewForms");
  els.recentResponses = $("#recentResponses");
  els.newFormButton = $("#newFormButton");
  els.formEditor = $("#formEditor");
  els.editorFormId = $("#editorFormId");
  els.editorTitle = $("#editorTitle");
  els.editorSlug = $("#editorSlug");
  els.editorDescription = $("#editorDescription");
  els.editorSuccessMessage = $("#editorSuccessMessage");
  els.editorCategoryLabel = $("#editorCategoryLabel");
  els.editorCategoryOptions = $("#editorCategoryOptions");
  els.editorIsActive = $("#editorIsActive");
  els.builderFields = $("#builderFields");
  els.formCards = $("#formCards");
  els.formError = $("#formError");
  els.responseFormFilter = $("#responseFormFilter");
  els.responseCategoryFilter = $("#responseCategoryFilter");
  els.responseSearch = $("#responseSearch");
  els.responseSearchButton = $("#responseSearchButton");
  els.categorySummary = $("#categorySummary");
  els.responseList = $("#responseList");
  els.responseDetail = $("#responseDetail");
  els.respondentFormFilter = $("#respondentFormFilter");
  els.respondentSearch = $("#respondentSearch");
  els.respondentSearchButton = $("#respondentSearchButton");
  els.newRespondentButton = $("#newRespondentButton");
  els.respondentList = $("#respondentList");
  els.respondentHistory = $("#respondentHistory");
  els.respondentDirectoryOptions = $("#respondentDirectoryOptions");
  els.measurementRespondentFilter = $("#measurementRespondentFilter");
  els.measurementSearch = $("#measurementSearch");
  els.measurementSearchButton = $("#measurementSearchButton");
  els.measurementManagerForm = $("#measurementManagerForm");
  els.measurementRespondentInput = $("#measurementRespondentInput");
  els.measurementEntryDate = $("#measurementEntryDate");
  els.measurementEntryCategory = $("#measurementEntryCategory");
  els.measurementEntryWaist = $("#measurementEntryWaist");
  els.measurementEntryHip = $("#measurementEntryHip");
  els.measurementEntryThigh = $("#measurementEntryThigh");
  els.measurementManagerError = $("#measurementManagerError");
  els.measurementManagerFeedback = $("#measurementManagerFeedback");
  els.measurementManagerList = $("#measurementManagerList");
  els.refreshOperationsButton = $("#refreshOperationsButton");
  els.operationsStatus = $("#operationsStatus");
  els.publicBaseUrlForm = $("#publicBaseUrlForm");
  els.settingsPublicBaseUrl = $("#settingsPublicBaseUrl");
  els.settingsPublicBaseUrlClear = $("#settingsPublicBaseUrlClear");
  els.publicBaseUrlHelp = $("#publicBaseUrlHelp");
  els.publicBaseUrlError = $("#publicBaseUrlError");
  els.passwordChangeForm = $("#passwordChangeForm");
  els.currentPassword = $("#currentPassword");
  els.newPassword = $("#newPassword");
  els.confirmPassword = $("#confirmPassword");
  els.passwordChangeHelp = $("#passwordChangeHelp");
  els.passwordChangeError = $("#passwordChangeError");
  els.createBackupButton = $("#createBackupButton");
  els.backupFeedback = $("#backupFeedback");
  els.backupList = $("#backupList");
  els.imageModal = $("#imageModal");
  els.imageModalClose = $("#imageModalClose");
  els.imageModalImg = $("#imageModalImg");
  els.imageModalTitle = $("#imageModalTitle");
  els.imageModalMeta = $("#imageModalMeta");
}

function bindEvents() {
  els.loginForm.addEventListener("submit", onLogin);
  els.logoutButton.addEventListener("click", onLogout);
  els.newFormButton.addEventListener("click", () => {
    state.editingFormId = null;
    renderEditor();
    setActivePanel("forms");
  });
  els.formEditor.addEventListener("submit", onSaveForm);
  els.responseSearchButton.addEventListener("click", loadResponses);
  els.respondentSearchButton.addEventListener("click", loadRespondents);
  els.newRespondentButton.addEventListener("click", renderRespondentCreateForm);
  els.measurementSearchButton.addEventListener("click", loadMeasurementManager);
  els.measurementManagerForm.addEventListener("submit", onMeasurementManagerSubmit);
  els.publicBaseUrlForm?.addEventListener("submit", onPublicBaseUrlSubmit);
  els.settingsPublicBaseUrlClear?.addEventListener("click", onPublicBaseUrlClear);
  els.passwordChangeForm?.addEventListener("submit", onPasswordChangeSubmit);
  els.createBackupButton?.addEventListener("click", onCreateBackup);
  els.refreshOperationsButton?.addEventListener("click", refreshOperationsStatus);
  els.imageModalClose?.addEventListener("click", closeImageModal);
  $$("[data-image-modal-close]").forEach((node) => {
    node.addEventListener("click", closeImageModal);
  });
  els.imageModalImg?.addEventListener("error", () => {
    if (els.imageModalImg.dataset.fallbackApplied === "1") {
      return;
    }
    const fallback = String(els.imageModalImg.dataset.fallbackSrc || "").trim();
    if (!fallback || fallback === els.imageModalImg.currentSrc || fallback === els.imageModalImg.src) {
      return;
    }
    els.imageModalImg.dataset.fallbackApplied = "1";
    els.imageModalImg.src = fallback;
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && els.imageModal && !els.imageModal.classList.contains("hidden")) {
      closeImageModal();
    }
  });
  els.responseFormFilter.addEventListener("change", () => {
    state.selectedFormId = Number(els.responseFormFilter.value || 0) || null;
    state.responseCategory = "";
    els.responseCategoryFilter.value = "";
    loadResponses();
  });
  els.responseCategoryFilter.addEventListener("change", () => {
    state.responseCategory = els.responseCategoryFilter.value;
    loadResponses();
  });
  els.respondentFormFilter.addEventListener("change", async () => {
    state.selectedRespondentFormId = els.respondentFormFilter.value;
    await loadRespondents();
    if (state.activeRespondentId) {
      await loadRespondentHistory(state.activeRespondentId, false);
    } else {
      els.respondentHistory.innerHTML = "回答者を選択してください。";
    }
  });
  els.responseSearch.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      loadResponses();
    }
  });
  els.respondentSearch.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      loadRespondents();
    }
  });
  els.measurementSearch.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      loadMeasurementManager();
    }
  });
  els.measurementRespondentFilter.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      loadMeasurementManager();
    }
  });

  $$("[data-add-field]").forEach((button) => {
    button.addEventListener("click", () => {
      addBuilderField({ type: button.dataset.addField });
    });
  });

  els.tabLinks.forEach((button) => {
    button.addEventListener("click", () => setActivePanel(button.dataset.panelTarget));
  });
}

async function bootstrap() {
  try {
    const data = await api("/api/admin/bootstrap");
    state.forms = data.forms || [];
    state.stats = data.stats || state.stats;
    state.recentResponses = data.recentResponses || [];
    state.publicBaseUrl = data.publicBaseUrl || runtime.respondentHomeUrl();
    state.settings = data.settings || state.settings;
    state.operationsStatus = data.operationsStatus || null;
    state.backups = data.backups || [];

    if (!state.selectedFormId || !state.forms.some((form) => form.id === state.selectedFormId)) {
      state.selectedFormId = state.forms[0]?.id ?? null;
    }
    if (
      state.selectedRespondentFormId &&
      !state.forms.some((form) => String(form.id) === String(state.selectedRespondentFormId))
    ) {
      state.selectedRespondentFormId = "";
    }
    if (state.editingFormId && !state.forms.some((form) => form.id === state.editingFormId)) {
      state.editingFormId = null;
    }

    showAdmin();
    renderOverview();
    renderForms();
    renderSettingsPanel();
    populateFilterSelects();
    renderEditor();
    await loadRespondentDirectory();
    await loadResponses();
    await loadRespondents();
    await loadMeasurementManager();
  } catch (error) {
    showLogin(error.message === "認証が必要です。" ? "" : error.message);
  }
}

function showAdmin() {
  els.loginPanel.classList.add("hidden");
  els.adminApp.classList.remove("hidden");
}

function showLogin(message = "") {
  els.loginPanel.classList.remove("hidden");
  els.adminApp.classList.add("hidden");
  els.loginError.textContent = message;
}

async function onLogin(event) {
  event.preventDefault();
  els.loginError.textContent = "";
  try {
    await api("/api/admin/login", {
      method: "POST",
      body: { password: els.loginPassword.value },
    });
    els.loginPassword.value = "";
    await bootstrap();
  } catch (error) {
    els.loginError.textContent = error.message;
  }
}

async function onLogout() {
  await api("/api/admin/logout", { method: "POST", body: {} }).catch(() => null);
  showLogin("");
}

async function refreshOperationsStatus() {
  try {
    const data = await api("/api/admin/operations/status");
    state.operationsStatus = data.status || null;
    state.backups = data.backups || [];
    if (state.operationsStatus) {
      state.settings = {
        ...state.settings,
        configuredPublicBaseUrl: state.operationsStatus.configuredPublicBaseUrl || "",
        publicBaseUrl: state.operationsStatus.publicUrl || state.settings.publicBaseUrl || "",
        publicBaseUrlSource: state.operationsStatus.publicBaseUrlSource || "auto",
        defaultPasswordInUse: !!state.operationsStatus.defaultPasswordInUse,
      };
      state.publicBaseUrl = state.operationsStatus.publicUrl || state.publicBaseUrl;
    }
    renderSettingsPanel();
  } catch (error) {
    if (els.backupFeedback) {
      els.backupFeedback.textContent = error.message;
    }
  }
}

async function onPublicBaseUrlSubmit(event) {
  event.preventDefault();
  els.publicBaseUrlError.textContent = "";
  const submitButton = els.publicBaseUrlForm.querySelector('button[type="submit"]');
  submitButton.disabled = true;
  submitButton.textContent = "保存中...";
  try {
    const data = await api("/api/admin/settings/public-base-url", {
      method: "POST",
      body: {
        publicBaseUrl: els.settingsPublicBaseUrl.value.trim(),
      },
    });
    state.settings.configuredPublicBaseUrl = data.configuredPublicBaseUrl || "";
    state.settings.publicBaseUrl = data.publicBaseUrl || "";
    state.settings.publicBaseUrlSource = data.publicBaseUrlSource || "auto";
    state.publicBaseUrl = data.publicBaseUrl || state.publicBaseUrl;
    await refreshOperationsStatus();
    renderForms();
    renderSettingsPanel();
  } catch (error) {
    els.publicBaseUrlError.textContent = error.message;
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "公開URLを保存";
  }
}

async function onPublicBaseUrlClear() {
  if (!els.settingsPublicBaseUrl) {
    return;
  }
  els.settingsPublicBaseUrl.value = "";
  els.publicBaseUrlError.textContent = "";
  try {
    const data = await api("/api/admin/settings/public-base-url", {
      method: "POST",
      body: { publicBaseUrl: "" },
    });
    state.settings.configuredPublicBaseUrl = data.configuredPublicBaseUrl || "";
    state.settings.publicBaseUrl = data.publicBaseUrl || "";
    state.settings.publicBaseUrlSource = data.publicBaseUrlSource || "auto";
    state.publicBaseUrl = data.publicBaseUrl || state.publicBaseUrl;
    await refreshOperationsStatus();
    renderForms();
    renderSettingsPanel();
  } catch (error) {
    els.publicBaseUrlError.textContent = error.message;
  }
}

async function onPasswordChangeSubmit(event) {
  event.preventDefault();
  els.passwordChangeError.textContent = "";
  const submitButton = els.passwordChangeForm.querySelector('button[type="submit"]');
  submitButton.disabled = true;
  submitButton.textContent = "変更中...";
  try {
    await api("/api/admin/settings/password", {
      method: "POST",
      body: {
        currentPassword: els.currentPassword.value,
        newPassword: els.newPassword.value,
        confirmPassword: els.confirmPassword.value,
      },
    });
    els.currentPassword.value = "";
    els.newPassword.value = "";
    els.confirmPassword.value = "";
    state.settings.defaultPasswordInUse = false;
    await refreshOperationsStatus();
    renderSettingsPanel();
  } catch (error) {
    els.passwordChangeError.textContent = error.message;
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "パスワードを変更";
  }
}

async function onCreateBackup() {
  if (!els.createBackupButton) {
    return;
  }
  els.backupFeedback.textContent = "";
  els.createBackupButton.disabled = true;
  els.createBackupButton.textContent = "作成中...";
  try {
    const data = await api("/api/admin/backups/create", { method: "POST", body: {} });
    state.backups = data.backups || [];
    await refreshOperationsStatus();
    renderSettingsPanel();
    if (data.backup?.downloadUrl) {
      els.backupFeedback.innerHTML = `バックアップを作成しました。<a href="${escapeAttribute(data.backup.downloadUrl)}">${escapeHtml(data.backup.name || "ダウンロード")}</a>`;
    } else {
      els.backupFeedback.textContent = "バックアップを作成しました。";
    }
  } catch (error) {
    els.backupFeedback.textContent = error.message;
  } finally {
    els.createBackupButton.disabled = false;
    els.createBackupButton.textContent = "バックアップを作成";
  }
}

function renderOverview() {
  els.statFormCount.textContent = String(state.stats.formCount || 0);
  els.statResponseCount.textContent = String(state.stats.responseCount || 0);
  els.statRespondentCount.textContent = String(state.stats.respondentCount || 0);

  if (!state.forms.length) {
    els.overviewForms.innerHTML = `<div class="empty-state">まだフォームがありません。</div>`;
  } else {
    els.overviewForms.innerHTML = state.forms
      .map(
        (form) => `
          <article class="summary-card">
            <div class="spread">
              <div>
                <h3>${escapeHtml(form.title)}</h3>
                <p class="muted">/${escapeHtml(form.slug)}</p>
              </div>
              <span class="status-badge ${form.isActive ? "active" : "inactive"}">
                ${form.isActive ? "公開中" : "非公開"}
              </span>
            </div>
            <div class="mini-meta">
              <span>回答 ${form.responseCount}</span>
              <span>回答者 ${form.respondentCount}</span>
            </div>
          </article>
        `
      )
      .join("");
  }

  if (!state.recentResponses.length) {
    els.recentResponses.innerHTML = `<div class="empty-state">まだ回答がありません。</div>`;
  } else {
    els.recentResponses.innerHTML = state.recentResponses.map((item) => responseCardMarkup(item, true)).join("");
    bindResponseCards(els.recentResponses);
    attachPreviewFallbackHandlers(els.recentResponses);
  }
}

function renderForms() {
  if (!state.forms.length) {
    els.formCards.innerHTML = `<div class="empty-state">フォームを作成してください。</div>`;
    return;
  }

  const hubUrl = publicHubUrl();
  const qrAccessNote =
    state.publicBaseUrl && state.publicBaseUrl !== runtime.respondentHomeUrl()
      ? `このQRコードはスマホ用に ${state.publicBaseUrl} を使っています。`
      : "お客さんはこのQRコードからアンケートタイトルを選択します。";
  const hubCard = `
    <article class="form-card hub-card">
      <div class="spread">
        <div>
          <h3>共通QRコード</h3>
          <p class="muted">${escapeHtml(hubUrl)}</p>
        </div>
        <span class="status-badge active">共通導線</span>
      </div>
      <p class="muted">${escapeHtml(qrAccessNote)}</p>
      <div class="qr-block">
        <img
          alt="QR code"
          class="qr-image"
          src="${qrCodeUrl(hubUrl)}"
          data-qr-fallback-src="${escapeAttribute(qrCodeFallbackUrl(hubUrl))}"
          data-qr-text="${escapeAttribute(hubUrl)}"
        />
        <p class="muted qr-error-message hidden" data-qr-error-message>QRコードを表示できません。下のURLをコピーしてご利用ください。</p>
        <div class="button-row">
          <button class="ghost-button" type="button" data-copy-url="${escapeAttribute(hubUrl)}">共通URLコピー</button>
        </div>
      </div>
    </article>
  `;

  els.formCards.innerHTML =
    hubCard +
    state.forms
      .map((form) => {
        const publicUrl = publicFormUrl(form.slug);
        return `
        <article class="form-card">
          <div class="spread">
            <div>
              <h3>${escapeHtml(form.title)}</h3>
              <p class="muted">${escapeHtml(publicUrl)}</p>
            </div>
            <span class="status-badge ${form.isActive ? "active" : "inactive"}">
              ${form.isActive ? "公開中" : "非公開"}
            </span>
          </div>
          <div class="mini-meta">
            <span>回答 ${form.responseCount}</span>
            <span>回答者 ${form.respondentCount}</span>
          </div>
          <div class="button-row">
            <button class="ghost-button" type="button" data-edit-form="${form.id}">編集</button>
            <button class="ghost-button" type="button" data-copy-url="${escapeAttribute(publicUrl)}">直リンクコピー</button>
            <button class="ghost-button" type="button" data-toggle-form="${form.id}">
              ${form.isActive ? "非公開にする" : "公開する"}
            </button>
          </div>
        </article>
      `;
      })
      .join("");

  $$("[data-edit-form]").forEach((button) => {
    button.addEventListener("click", () => {
      state.editingFormId = Number(button.dataset.editForm);
      renderEditor();
      setActivePanel("forms");
    });
  });

  $$("[data-copy-url]").forEach((button) => {
    button.addEventListener("click", async () => {
      const url = button.dataset.copyUrl;
      try {
        await navigator.clipboard.writeText(url);
        button.textContent = "コピー済み";
        setTimeout(() => {
          button.textContent = "URLコピー";
        }, 1200);
      } catch (error) {
        window.prompt("URLをコピーしてください", url);
      }
    });
  });

  $$("[data-toggle-form]").forEach((button) => {
    button.addEventListener("click", async () => {
      try {
        await api(`/api/admin/forms/${button.dataset.toggleForm}/toggle`, { method: "POST", body: {} });
        await bootstrap();
      } catch (error) {
        alert(error.message);
      }
    });
  });
  bindQrImages(els.formCards);
}

function formatBytes(bytes) {
  const value = Number(bytes || 0);
  if (!Number.isFinite(value) || value <= 0) {
    return "0 B";
  }
  const units = ["B", "KB", "MB", "GB", "TB"];
  let size = value;
  let index = 0;
  while (size >= 1024 && index < units.length - 1) {
    size /= 1024;
    index += 1;
  }
  return `${size >= 10 || index === 0 ? size.toFixed(0) : size.toFixed(1)} ${units[index]}`;
}

function operationsStatusCards(status) {
  if (!status) {
    return `<div class="empty-state">運用状態を読み込めませんでした。</div>`;
  }
  const cards = [
    { label: "公開URL", value: status.publicUrl || "-", note: status.publicBaseUrlSource === "env" ? "環境変数を使用中" : status.publicBaseUrlSource === "config" ? "設定保存値を使用中" : "自動判定" },
    { label: "ローカルURL", value: status.localUrl || "-", note: "このMac上での確認用" },
    { label: "DBサイズ", value: formatBytes(status.databaseSizeBytes), note: status.databasePath || "" },
    { label: "アップロード画像", value: `${status.uploadFileCount || 0}件`, note: `ローカル ${status.localImageCount || 0} / 外部参照 ${status.externalImageCount || 0}` },
    { label: "バックアップ", value: `${status.backupCount || 0}件`, note: status.latestBackup ? `最新 ${formatDate(status.latestBackup.createdAt)}` : "まだ未作成" },
    { label: "パスワード", value: status.defaultPasswordInUse ? "初期設定のまま" : "変更済み", note: status.defaultPasswordInUse ? "変更を推奨" : "運用可能" },
  ];
  return cards
    .map(
      (card) => `
        <article class="note-box compact-note operations-card">
          <strong>${escapeHtml(card.label)}</strong>
          <div class="operations-card-value">${escapeHtml(card.value)}</div>
          <p class="muted operations-card-note">${escapeHtml(card.note)}</p>
        </article>
      `
    )
    .join("");
}

function renderBackupList(backups) {
  if (!els.backupList) {
    return;
  }
  if (!backups.length) {
    els.backupList.innerHTML = `<div class="empty-state">バックアップはまだありません。</div>`;
    return;
  }
  els.backupList.innerHTML = backups
    .map(
      (backup) => `
        <article class="note-box compact-note backup-card">
          <div class="spread">
            <div>
              <strong>${escapeHtml(backup.name || "-")}</strong>
              <p class="muted">${escapeHtml(formatDate(backup.createdAt))} / ${escapeHtml(formatBytes(backup.sizeBytes))}</p>
            </div>
            <a class="ghost-button compact-button" href="${escapeAttribute(backup.downloadUrl || "#")}">ダウンロード</a>
          </div>
        </article>
      `
    )
    .join("");
}

function renderSettingsPanel() {
  if (els.operationsStatus) {
    els.operationsStatus.innerHTML = operationsStatusCards(state.operationsStatus);
  }
  if (els.settingsPublicBaseUrl) {
    els.settingsPublicBaseUrl.value = state.settings.configuredPublicBaseUrl || "";
  }
  if (els.publicBaseUrlHelp) {
    const sourceLabel =
      state.settings.publicBaseUrlSource === "env"
        ? "現在は環境変数が優先されています。設定保存値は次回以降の起動で使えます。"
        : state.settings.publicBaseUrlSource === "config"
          ? "保存済みの公開URLを使用中です。"
          : "現在は自動判定のURLを使用中です。";
    const temporaryNote = state.operationsStatus?.publicUrlIsTemporary
      ? "現在の公開URLは一時URLです。固定ドメインへ切り替えることを推奨します。"
      : "";
    els.publicBaseUrlHelp.textContent = [sourceLabel, temporaryNote].filter(Boolean).join(" ");
  }
  if (els.passwordChangeHelp) {
    els.passwordChangeHelp.textContent = state.settings.defaultPasswordInUse
      ? "初期パスワードのままです。早めに変更してください。"
      : "現在のパスワードを入力して変更できます。";
  }
  renderBackupList(state.backups || []);
}

function populateFilterSelects() {
  const formOptions = state.forms
    .map((form) => `<option value="${form.id}">${escapeHtml(form.title)}</option>`)
    .join("");
  els.responseFormFilter.innerHTML = formOptions || `<option value="">フォームなし</option>`;
  if (state.selectedFormId) {
    els.responseFormFilter.value = String(state.selectedFormId);
  }

  els.respondentFormFilter.innerHTML =
    `<option value="">全フォーム</option>` +
    state.forms
      .map((form) => `<option value="${form.id}">${escapeHtml(form.title)}</option>`)
      .join("");
  els.respondentFormFilter.value = state.selectedRespondentFormId || "";
}

function renderEditor() {
  const target = state.forms.find((form) => form.id === state.editingFormId) || null;
  els.formError.textContent = "";
  els.builderFields.innerHTML = "";

  if (!target) {
    els.editorFormId.value = "";
    els.editorTitle.value = "";
    els.editorSlug.value = "";
    els.editorDescription.value = "";
    els.editorSuccessMessage.value = "送信ありがとうございました。";
    els.editorCategoryLabel.value = "分類";
    els.editorCategoryOptions.value = "";
    els.editorIsActive.value = "true";
    return;
  }

  els.editorFormId.value = String(target.id);
  els.editorTitle.value = target.title;
  els.editorSlug.value = target.slug;
  els.editorDescription.value = target.description || "";
  els.editorSuccessMessage.value = target.successMessage || "";
  els.editorCategoryLabel.value = target.categoryLabel || "分類";
  els.editorCategoryOptions.value = (target.categoryOptions || []).join("\n");
  els.editorIsActive.value = String(Boolean(target.isActive));
  (target.fields || []).forEach((field) => addBuilderField(field));
}

function addBuilderField(field = {}) {
  const card = document.createElement("div");
  card.className = "builder-card";
  card.innerHTML = `
    <div class="grid-two">
      <label class="field">
        <span>ラベル</span>
        <input class="builder-label" type="text" value="${escapeAttribute(field.label || "")}" />
      </label>
      <label class="field">
        <span>キー</span>
        <input class="builder-key" type="text" value="${escapeAttribute(field.key || "")}" placeholder="site_name" />
      </label>
    </div>
    <div class="grid-two">
      <label class="field">
        <span>タイプ</span>
        <select class="builder-type">
          <option value="short_text">短文</option>
          <option value="long_text">長文</option>
          <option value="select">プルダウン</option>
          <option value="radio">ラジオ</option>
          <option value="checkbox">複数選択</option>
          <option value="file">画像アップロード</option>
        </select>
      </label>
      <label class="field">
        <span>必須設定</span>
        <select class="builder-required">
          <option value="true">必須</option>
          <option value="false">任意</option>
        </select>
      </label>
    </div>
    <label class="field builder-options-wrap">
      <span>選択肢</span>
      <textarea class="builder-options" rows="3" placeholder="高&#10;中&#10;低">${escapeHtml(
        (field.options || []).join("\n")
      )}</textarea>
    </label>
    <label class="field">
      <span>補足文</span>
      <textarea class="builder-help" rows="2" placeholder="回答者向けの説明">${escapeHtml(field.helpText || "")}</textarea>
    </label>
    <label class="field">
      <span>プレースホルダー</span>
      <input class="builder-placeholder" type="text" value="${escapeAttribute(field.placeholder || "")}" />
    </label>
    <div class="builder-other-wrap">
      <label class="field">
        <span>その他入力</span>
        <select class="builder-allow-other">
          <option value="true">有効</option>
          <option value="false">無効</option>
        </select>
      </label>
    </div>
    <div class="grid-two builder-file-wrap">
      <label class="field">
        <span>画像受け付け形式</span>
        <input class="builder-accept" type="text" value="${escapeAttribute(field.accept || "image/*")}" />
      </label>
      <label class="field">
        <span>複数枚</span>
        <select class="builder-allow-multiple">
          <option value="true">許可する</option>
          <option value="false">1枚のみ</option>
        </select>
      </label>
    </div>
    <div class="grid-two builder-visibility-wrap">
      <label class="field">
        <span>表示条件の元キー</span>
        <input class="builder-visibility-key" type="text" value="${escapeAttribute(field.visibilityFieldKey || "")}" placeholder="future_plan" />
      </label>
      <label class="field">
        <span>表示条件の値</span>
        <input class="builder-visibility-values" type="text" value="${escapeAttribute((field.visibilityValues || []).join(', '))}" placeholder="継続していきたい" />
      </label>
    </div>
    <div class="button-row end">
      <button class="danger-button" type="button">削除</button>
    </div>
  `;
  card.querySelector(".builder-type").value = field.type || "short_text";
  card.querySelector(".builder-required").value = String(Boolean(field.required));
  card.querySelector(".builder-allow-multiple").value = String(Boolean(field.allowMultiple));
  card.querySelector(".builder-allow-other").value = String(Boolean(field.allowOther));
  card.querySelector(".danger-button").addEventListener("click", () => card.remove());
  card.querySelector(".builder-type").addEventListener("change", () => syncBuilderField(card));
  syncBuilderField(card);
  els.builderFields.appendChild(card);
}

function syncBuilderField(card) {
  const type = card.querySelector(".builder-type").value;
  const optionsWrap = card.querySelector(".builder-options-wrap");
  const fileWrap = card.querySelector(".builder-file-wrap");
  const otherWrap = card.querySelector(".builder-other-wrap");
  optionsWrap.classList.toggle("hidden", !["select", "radio", "checkbox"].includes(type));
  fileWrap.classList.toggle("hidden", type !== "file");
  otherWrap.classList.toggle("hidden", type !== "checkbox");
}

async function onSaveForm(event) {
  event.preventDefault();
  els.formError.textContent = "";
  try {
    const payload = serializeEditor();
    if (els.editorFormId.value) {
      await api(`/api/admin/forms/${els.editorFormId.value}`, { method: "PUT", body: payload });
      state.editingFormId = Number(els.editorFormId.value);
    } else {
      const result = await api("/api/admin/forms", { method: "POST", body: payload });
      state.editingFormId = result.form.id;
    }
    await bootstrap();
  } catch (error) {
    els.formError.textContent = error.message;
  }
}

function serializeEditor() {
  const fields = $$(".builder-card").map((card) => {
    const type = card.querySelector(".builder-type").value;
    const optionLines = card.querySelector(".builder-options").value
      .split("\n")
      .map((item) => item.trim())
      .filter(Boolean);
    return {
      label: card.querySelector(".builder-label").value.trim(),
      key: card.querySelector(".builder-key").value.trim(),
      type,
      required: card.querySelector(".builder-required").value === "true",
      options: optionLines,
      helpText: card.querySelector(".builder-help").value.trim(),
      placeholder: card.querySelector(".builder-placeholder").value.trim(),
      accept: card.querySelector(".builder-accept").value.trim(),
      allowMultiple: card.querySelector(".builder-allow-multiple").value === "true",
      allowOther: card.querySelector(".builder-allow-other").value === "true",
      visibilityFieldKey: card.querySelector(".builder-visibility-key").value.trim(),
      visibilityValues: card.querySelector(".builder-visibility-values").value
        .split(",")
        .map((item) => item.trim())
        .filter(Boolean),
    };
  });

  return {
    title: els.editorTitle.value.trim(),
    slug: els.editorSlug.value.trim(),
    description: els.editorDescription.value.trim(),
    successMessage: els.editorSuccessMessage.value.trim(),
    categoryLabel: els.editorCategoryLabel.value.trim(),
    categoryOptions: els.editorCategoryOptions.value
      .split("\n")
      .map((item) => item.trim())
      .filter(Boolean),
    isActive: els.editorIsActive.value === "true",
    fields,
  };
}

async function loadResponses() {
  if (!state.selectedFormId) {
    els.responseList.innerHTML = `<div class="empty-state">フォームを選択してください。</div>`;
    els.responseDetail.innerHTML = "回答を選択してください。";
    els.categorySummary.innerHTML = "";
    return;
  }

  state.responseSearch = els.responseSearch.value.trim();
  const query = new URLSearchParams();
  if (state.responseCategory) {
    query.set("category", state.responseCategory);
  }
  if (state.responseSearch) {
    query.set("respondent", state.responseSearch);
  }

  try {
    const data = await api(`/api/admin/forms/${state.selectedFormId}/responses?${query.toString()}`);
    renderCategorySummary(data.categorySummary || []);
    renderResponseList(data.responses || []);
  } catch (error) {
    els.responseList.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
  }
}

function renderCategorySummary(summary) {
  const currentValue = state.responseCategory;
  const options = [`<option value="">すべて</option>`]
    .concat(summary.map((item) => `<option value="${escapeAttribute(item.category)}">${escapeHtml(item.category)}</option>`))
    .join("");
  els.responseCategoryFilter.innerHTML = options;
  els.responseCategoryFilter.value = currentValue;

  if (!summary.length) {
    els.categorySummary.innerHTML = `<div class="empty-state">分類別の集計はまだありません。</div>`;
    return;
  }

  const max = Math.max(...summary.map((item) => item.count), 1);
  els.categorySummary.innerHTML = summary
    .map(
      (item) => `
        <div class="summary-row">
          <div class="spread">
            <span>${escapeHtml(item.category)}</span>
            <strong>${item.count}</strong>
          </div>
          <div class="summary-track">
            <div class="summary-fill" style="width:${Math.max((item.count / max) * 100, 8)}%"></div>
          </div>
        </div>
      `
    )
    .join("");
}

function renderResponseList(items) {
  if (!items.length) {
    els.responseList.innerHTML = `<div class="empty-state">条件に合う回答はありません。</div>`;
    els.responseDetail.innerHTML = "回答を選択してください。";
    return;
  }
  els.responseList.innerHTML = items.map((item) => responseCardMarkup(item, false)).join("");
  bindResponseCards(els.responseList);
  attachPreviewFallbackHandlers(els.responseList);
}

function respondentShortKey(respondentId) {
  const compact = String(respondentId || "")
    .replace(/[^A-Za-z0-9]/g, "")
    .toUpperCase();
  return compact.slice(-6) || "ANON";
}

function respondentLookupKey(value) {
  return String(value || "")
    .normalize("NFKC")
    .trim()
    .replace(/\s+/g, "")
    .toLowerCase();
}

function findRespondentByName(name) {
  const key = respondentLookupKey(name);
  if (!key) {
    return null;
  }
  return state.respondentDirectory.find((item) => respondentLookupKey(item.respondentName) === key) || null;
}

function respondentPrimaryLabel(item) {
  if (item.respondentName) {
    return item.respondentName;
  }
  if (item.respondentEmail) {
    return item.respondentEmail;
  }
  return `匿名回答 ${respondentShortKey(item.respondentId)}`;
}

function respondentSecondaryLabel(item) {
  if (item.respondentName) {
    return "お名前で履歴管理";
  }
  const shortKey = respondentShortKey(item.respondentId);
  return `自動識別 ${shortKey}`;
}

function respondentProfileDateLabel(item) {
  return item.profileDate ? formatDateOnly(item.profileDate) : "記録未登録";
}

function respondentProfileTitleLabel(item) {
  return item.profileTitle || "画像記録なし";
}

function respondentTicketSheetLabel(item) {
  const parts = [];
  if (item.latestTicketSheet) {
    parts.push(`回数券 ${item.latestTicketSheet}`);
  }
  if (item.currentTicketBookType) {
    parts.push(item.currentTicketBookType);
  }
  if (item.currentTicketStampMax) {
    parts.push(`${item.currentTicketStampCount || 0}/${item.currentTicketStampMax}`);
  }
  return parts.join(" / ");
}

function respondentTicketSheetMeta(item) {
  const label = respondentTicketSheetLabel(item);
  if (!label) {
    return "まだ記録がありません。";
  }
  return label;
}

function respondentTicketSheetInputValue(item) {
  const match = String(item.latestTicketSheetManualValue || "").match(/(\d{1,3})/);
  return match ? match[1] : "";
}

function ticketBookTypeMax(ticketBookType) {
  if (ticketBookType === "6回券") {
    return 5;
  }
  if (ticketBookType === "10回券") {
    return 9;
  }
  return 0;
}

function ticketStampCountValue(item) {
  return Number(item.currentTicketStampCount || 0);
}

function ticketStampAutoValue(item) {
  return Number(item.currentTicketStampAutoValue || 0);
}

function ticketStampManualValue(item) {
  return item.currentTicketStampManualEnabled ? Number(item.currentTicketStampManualValue || 0) : "";
}

function ticketStampManualEnabled(item) {
  return !!item.currentTicketStampManualEnabled;
}

function ticketStampSummary(ticketBookType, stampCount, manualEnabled = false, autoCount = 0) {
  if (!ticketBookType) {
    return "回数券種別を選択してください。";
  }
  const suffix = manualEnabled ? "管理者上書き" : `アンケート連動 ${autoCount}/${ticketBookTypeMax(ticketBookType)}`;
  return `${stampCount}/${ticketBookTypeMax(ticketBookType)} / ${suffix}`;
}

function ticketStampCardButtons(ticketBookType, stampCount) {
  const max = ticketBookTypeMax(ticketBookType);
  if (!max) {
    return "";
  }
  return Array.from({ length: max }, (_, index) => {
    const value = index + 1;
    const active = value <= stampCount ? " active" : "";
    return `
      <button
        class="ticket-stamp-dot${active}"
        type="button"
        data-ticket-stamp-choice="${value}"
        aria-pressed="${value <= stampCount ? "true" : "false"}"
      >
        ${value}
      </button>
    `;
  }).join("");
}

function respondentMeasurementMeta(item) {
  const latest = item.latestMeasurements;
  if (!latest) {
    return "まだ記録がありません。";
  }
  const dateLabel = formatDateOnly(item.latestMeasurementDate || "");
  const categoryLabel = latest.category ? `${latest.category} / ` : "";
  return `${dateLabel || "最新"} / ${categoryLabel}W ${latest.waistLabel} / H ${latest.hipLabel} / T ${latest.thighLabel}`;
}

function measurementValueLabel(value) {
  const amount = Number.parseFloat(String(value ?? ""));
  if (!Number.isFinite(amount)) {
    return "-";
  }
  return Number.isInteger(amount) ? `${amount}` : amount.toFixed(1);
}

function renderMeasurementCategoryOptions(selected = "") {
  return [
    `<option value="">未選択</option>`,
    ...MEASUREMENT_CATEGORIES.map(
      (category) =>
        `<option value="${escapeAttribute(category)}"${selected === category ? " selected" : ""}>${escapeHtml(category)}</option>`
    ),
  ].join("");
}

function adminImageProxyUrl(src) {
  const raw = String(src || "").trim();
  if (!raw) {
    return "";
  }
  if (runtime.isGasMode()) {
    return raw;
  }
  return `/api/admin/image-proxy?src=${encodeURIComponent(raw)}`;
}

function directImageSource(image) {
  return String(image?.previewUrl || image?.url || "").trim();
}

function directOpenImageSource(image) {
  return String(image?.url || image?.previewUrl || "").trim();
}

function imagePreviewSource(image) {
  return adminImageProxyUrl(directImageSource(image));
}

function imageOpenSource(image) {
  return adminImageProxyUrl(directOpenImageSource(image));
}

function attachPreviewFallbackHandlers(container) {
  if (!container) {
    return;
  }
  container.querySelectorAll("img[data-fallback-src]").forEach((img) => {
    img.addEventListener("error", () => {
      if (img.dataset.fallbackApplied === "1") {
        return;
      }
      const fallback = String(img.dataset.fallbackSrc || "").trim();
      if (!fallback || fallback === img.currentSrc || fallback === img.src) {
        return;
      }
      img.dataset.fallbackApplied = "1";
      img.src = fallback;
    });
  });
}

function respondentProfileImageMarkup(item, compact = false, link = false) {
  const image = item?.profileImage;
  const previewUrl = imagePreviewSource(image);
  if (previewUrl) {
    const className = `respondent-avatar ${compact ? "compact" : ""}`;
    const imageMarkup = `<img src="${escapeAttribute(previewUrl)}" alt="${escapeAttribute(respondentPrimaryLabel(item))}" loading="lazy" referrerpolicy="no-referrer" />`;
    if (link && image?.url) {
      return `
        <a
          class="${className}"
          href="${escapeAttribute(image.url)}"
          target="_blank"
          rel="noreferrer"
        >
          ${imageMarkup}
        </a>
      `;
    }
    return `<div class="${className}">${imageMarkup}</div>`;
  }
  return `<div class="respondent-avatar placeholder ${compact ? "compact" : ""}">NO IMAGE</div>`;
}

function responseCardMarkup(item, compact = false) {
  const responsePreviewUrl = item.files?.[0] ? imagePreviewSource(item.files[0]) : "";
  const responseFallbackUrl = item.files?.[0] ? directImageSource(item.files[0]) : "";
  const thumb = responsePreviewUrl
    ? `<img class="response-thumb" src="${escapeAttribute(responsePreviewUrl)}" data-fallback-src="${escapeAttribute(responseFallbackUrl)}" alt="uploaded image" loading="lazy" referrerpolicy="no-referrer" />`
    : `<div class="response-thumb placeholder">NO IMAGE</div>`;
  return `
    <article class="response-card ${compact ? "compact" : ""}" data-response-id="${item.id}">
      ${thumb}
      <div class="response-copy">
        <div class="spread">
          <div>
            <h3>${escapeHtml(respondentPrimaryLabel(item))}</h3>
            <p class="muted">${escapeHtml(respondentSecondaryLabel(item))}</p>
          </div>
          <span class="category-chip">${escapeHtml(item.category)}</span>
        </div>
        <div class="mini-meta">
          <span>${escapeHtml(item.formTitle)}</span>
          <span>${formatDate(item.createdAt)}</span>
        </div>
      </div>
    </article>
  `;
}

function bindResponseCards(container) {
  container.querySelectorAll("[data-response-id]").forEach((card) => {
    card.addEventListener("click", async () => {
      const responseId = card.dataset.responseId;
      await openResponseDetail(responseId);
      setActivePanel("responses");
    });
  });
}

async function openResponseDetail(responseId) {
  try {
    const detail = await api(`/api/admin/responses/${responseId}`);
    const response = detail.response;
    const answerRows = detail.answers.length
      ? detail.answers
          .map(
            (answer) => `
              <tr>
                <th>${escapeHtml(answer.label)}</th>
                <td>${escapeHtml(answer.value || "-")}</td>
              </tr>
            `
          )
          .join("")
      : `<tr><th>追加質問</th><td>-</td></tr>`;

    const fileGallery = renderFileGroups(detail.files);

    els.responseDetail.innerHTML = `
      <div class="stack-gap compact-gap">
        <div class="detail-header">
          <div>
            <h3>${escapeHtml(respondentPrimaryLabel(response))}</h3>
            <p class="muted">${escapeHtml(respondentSecondaryLabel(response))} / ${escapeHtml(response.formTitle)}</p>
          </div>
          <button class="ghost-button" type="button" data-history-id="${escapeAttribute(response.respondentId)}">
            この履歴をまとめて見る
          </button>
        </div>
        <div class="meta-grid">
          <div><span>分類</span><strong>${escapeHtml(response.category)}</strong></div>
          <div><span>日時</span><strong>${formatDate(response.createdAt)}</strong></div>
          <div><span>お名前</span><strong>${escapeHtml(response.respondentName || "-")}</strong></div>
        </div>
        <div>
          <h4>追加回答</h4>
          <table class="answer-table">${answerRows}</table>
        </div>
        <div>
          <h4>アップロード画像</h4>
          ${fileGallery}
        </div>
      </div>
    `;
    const historyButton = els.responseDetail.querySelector("[data-history-id]");
    historyButton?.addEventListener("click", () => loadRespondentHistory(response.respondentId, true));
    bindImageOpenButtons(els.responseDetail);
    attachPreviewFallbackHandlers(els.responseDetail);
  } catch (error) {
    els.responseDetail.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
  }
}

async function loadRespondents() {
  const query = new URLSearchParams();
  if (state.selectedRespondentFormId) {
    query.set("form_id", state.selectedRespondentFormId);
  }
  state.respondentSearch = els.respondentSearch.value.trim();
  if (state.respondentSearch) {
    query.set("q", state.respondentSearch);
  }

  try {
    const data = await api(`/api/admin/respondents?${query.toString()}`);
    renderRespondentList(data.respondents || []);
  } catch (error) {
    els.respondentList.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
  }
}

async function loadRespondentDirectory() {
  try {
    const data = await api("/api/admin/respondents?limit=500");
    state.respondentDirectory = data.respondents || [];
    renderRespondentDirectoryOptions();
  } catch (error) {
    state.respondentDirectory = [];
    renderRespondentDirectoryOptions();
  }
}

function renderRespondentDirectoryOptions() {
  if (!els.respondentDirectoryOptions) {
    return;
  }
  els.respondentDirectoryOptions.innerHTML = state.respondentDirectory
    .map((item) => `<option value="${escapeAttribute(item.respondentName || "")}"></option>`)
    .join("");
}

async function loadMeasurementManager() {
  const query = new URLSearchParams();
  state.measurementSearchQuery = els.measurementSearch.value.trim();
  state.measurementRespondentFilter = els.measurementRespondentFilter.value.trim();
  if (state.measurementSearchQuery) {
    query.set("q", state.measurementSearchQuery);
  }
  if (state.measurementRespondentFilter) {
    query.set("respondent_name", state.measurementRespondentFilter);
  }
  const respondent = findRespondentByName(state.measurementRespondentFilter);
  if (respondent && !state.measurementSearchQuery) {
    query.set("respondent_id", respondent.respondentId);
    query.delete("respondent_name");
  }

  try {
    const data = await api(`/api/admin/measurements?${query.toString()}`);
    renderMeasurementManagerList(data.records || []);
  } catch (error) {
    els.measurementManagerFeedback.innerHTML = "";
    els.measurementManagerList.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
  }
}

function renderRespondentList(items) {
  if (!items.length) {
    state.activeRespondentId = null;
    els.respondentList.innerHTML = `<div class="empty-state">履歴候補が見つかりません。</div>`;
    renderRespondentCreateForm();
    return;
  }
  const activeExists = items.some((item) => item.respondentId === state.activeRespondentId);
  if (!activeExists) {
    state.activeRespondentId = null;
    els.respondentHistory.innerHTML = "回答者を選択してください。";
  }
  els.respondentList.innerHTML = items
    .map(
      (item) => `
        <article
          class="respondent-card-row ${state.activeRespondentId === item.respondentId ? "active" : ""}"
          data-respondent-id="${escapeAttribute(item.respondentId)}"
        >
          <div class="respondent-card-main">
            <h3>${escapeHtml(respondentPrimaryLabel(item))}</h3>
            <p class="muted respondent-card-subline">${escapeHtml(respondentSecondaryLabel(item))}</p>
            <p class="muted respondent-card-subline">${escapeHtml(respondentProfileTitleLabel(item))}</p>
            <p class="muted respondent-card-subline">最新日付 ${escapeHtml(respondentProfileDateLabel(item))}</p>
            ${
              respondentTicketSheetLabel(item)
                ? `<p class="muted respondent-card-subline respondent-ticket-line">${escapeHtml(respondentTicketSheetLabel(item))}</p>`
                : ""
            }
          </div>
          <div class="respondent-card-side">
            <div class="mini-meta vertical respondent-card-metrics">
              <span>回答 ${item.responseCount}</span>
              <span>画像記録 ${item.profileRecordCount || 0}</span>
              <span>計測 ${item.measurementCount || 0}</span>
              <span>${formatDate(item.lastResponseAt)}</span>
            </div>
            <div class="button-row respondent-card-actions">
              <button
                class="ghost-button compact-button"
                type="button"
                data-respondent-edit="${escapeAttribute(item.respondentId)}"
                data-respondent-name="${escapeAttribute(item.respondentName || respondentPrimaryLabel(item))}"
              >
                編集
              </button>
              <button
                class="danger-button compact-button"
                type="button"
                data-respondent-delete="${escapeAttribute(item.respondentId)}"
                data-respondent-name="${escapeAttribute(item.respondentName || respondentPrimaryLabel(item))}"
              >
                削除
              </button>
            </div>
          </div>
        </article>
      `
    )
    .join("");
  els.respondentList.querySelectorAll("[data-respondent-id]").forEach((card) => {
    card.addEventListener("click", () => loadRespondentHistory(card.dataset.respondentId, false));
  });
  bindRespondentActionButtons(els.respondentList);
}

async function loadRespondentHistory(respondentId, switchPanel) {
  state.activeRespondentId = respondentId;
  if (switchPanel) {
    setActivePanel("respondents");
  }
  const query = new URLSearchParams();
  if (state.selectedRespondentFormId) {
    query.set("form_id", state.selectedRespondentFormId);
  }

  try {
    const data = await api(`/api/admin/respondents/${encodeURIComponent(respondentId)}/history?${query.toString()}`);
    renderRespondentHistory(
      respondentId,
      data.respondent || null,
      data.history || [],
      data.imageRecords || data.profileRecords || [],
      data.measurementRecords || []
    );
  } catch (error) {
    els.respondentHistory.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
  }
}

function renderRespondentHistory(respondentId, respondent, items, imageRecords, measurementRecords) {
  if (!respondent && !items.length) {
    els.respondentHistory.innerHTML = `<div class="empty-state">履歴がありません。</div>`;
    return;
  }
  const base = respondent || items[0];
  const primary = respondentPrimaryLabel(base);
  const sheets = buildRespondentHistorySheets(items);
  const current = respondent || items[0];
  const measurementImportFeedback =
    state.measurementImportFeedback && state.measurementImportFeedback.respondentId === respondentId
      ? `
        <div class="success-panel compact-success-panel">
          <strong>${escapeHtml(state.measurementImportFeedback.title)}</strong>
          <p class="muted">${escapeHtml(state.measurementImportFeedback.message)}</p>
        </div>
      `
      : "";
  els.respondentHistory.innerHTML = `
    <div class="stack-gap">
      <div class="detail-header">
        <div>
          <h3>${escapeHtml(primary)}</h3>
          <p class="muted">${escapeHtml(respondentSecondaryLabel(base))}</p>
          <p class="muted">全回答履歴</p>
        </div>
        <div class="button-row respondent-card-actions">
          <button
            class="ghost-button compact-button"
            type="button"
            data-respondent-edit="${escapeAttribute(respondentId)}"
            data-respondent-name="${escapeAttribute(current.respondentName || primary)}"
          >
            お名前編集
          </button>
          <button
            class="danger-button compact-button"
            type="button"
            data-respondent-delete="${escapeAttribute(respondentId)}"
            data-respondent-name="${escapeAttribute(current.respondentName || primary)}"
          >
            回答者削除
          </button>
        </div>
      </div>
      <section class="history-sheet stack-gap compact-gap">
        <div class="spread history-sheet-head">
          <div>
            <h4>回答者情報</h4>
            <p class="muted">お名前はここで編集し、画像記録と計測記録は下で管理します。</p>
          </div>
        </div>
        <form class="stack-gap compact-gap" data-respondent-profile-form="${escapeAttribute(respondentId)}">
          <div class="respondent-profile-layout single-column">
            <div class="stack-gap compact-gap">
              <div class="grid-two">
                <div class="stack-gap compact-gap">
                  <div class="grid-two">
                    <label class="field">
                      <span>お名前</span>
                      <input
                        type="text"
                        name="respondent_name"
                        value="${escapeAttribute(current.respondentName || primary)}"
                        required
                      />
                    </label>
                    <label class="field">
                      <span>最新の回数券</span>
                      <input
                        type="text"
                        name="ticket_sheet_manual_value"
                        inputmode="numeric"
                        placeholder="例: 2 または 2枚目"
                        value="${escapeAttribute(respondentTicketSheetInputValue(current))}"
                      />
                    </label>
                  </div>
                  <div class="grid-two">
                    <label class="field">
                      <span>現在の回数券種別</span>
                      <select name="current_ticket_book_type">
                        <option value="">未設定</option>
                        <option value="6回券" ${current.currentTicketBookType === "6回券" ? "selected" : ""}>6回券</option>
                        <option value="10回券" ${current.currentTicketBookType === "10回券" ? "selected" : ""}>10回券</option>
                      </select>
                    </label>
                    <div class="field">
                      <span>回数券スタンプカード</span>
                      <input type="hidden" name="current_ticket_stamp_count" value="${escapeAttribute(String(ticketStampManualValue(current)))}" />
                      <input type="hidden" name="current_ticket_stamp_manual_enabled" value="${ticketStampManualEnabled(current) ? "1" : "0"}" />
                      <div class="ticket-stamp-card" data-ticket-stamp-card></div>
                      <div class="spread ticket-stamp-meta">
                        <p
                          class="muted"
                          data-ticket-stamp-summary
                          data-ticket-stamp-auto-value="${escapeAttribute(String(ticketStampAutoValue(current)))}"
                        ></p>
                        <button class="ghost-button compact-button" type="button" data-ticket-stamp-reset>リセット</button>
                      </div>
                    </div>
                  </div>
                  <p class="muted profile-inline-note">スタンプは施術後アンケートの回数に連動します。管理者が押した場合はその値を優先します。</p>
                </div>
                <div class="stack-gap compact-gap">
                  <div class="note-box compact-note">
                    <strong>画像記録</strong>
                    <div>${imageRecords.length}件</div>
                  </div>
                  <div class="note-box compact-note">
                    <strong>最新の回数券</strong>
                    <div>${escapeHtml(respondentTicketSheetMeta(current))}</div>
                  </div>
                  <div class="note-box compact-note">
                    <strong>最新の計測</strong>
                    <div>${escapeHtml(respondentMeasurementMeta(current))}</div>
                  </div>
                </div>
              </div>
            </div>
          </div>
          <div class="button-row">
            <button class="primary-button compact-button" type="submit">回答者情報を保存</button>
            <p class="muted profile-inline-note">保存すると回答履歴側のお名前も更新されます。</p>
          </div>
          <p class="error-text" data-respondent-profile-error></p>
        </form>
      </section>
      <section class="history-sheet stack-gap compact-gap">
        <div class="spread history-sheet-head">
          <div>
            <h4>画像記録を追加</h4>
            <p class="muted">タイトル・日付・メモ・画像を追加できます。アンケート画像も一覧に統合されます。</p>
          </div>
        </div>
        <form class="stack-gap compact-gap" data-respondent-record-form="${escapeAttribute(respondentId)}">
          <div class="grid-two">
            <label class="field">
              <span>タイトル</span>
              <input type="text" name="title" maxlength="120" required />
            </label>
            <label class="field">
              <span>日付</span>
              <input type="date" name="entry_date" required />
            </label>
          </div>
          <label class="field">
            <span>メモ</span>
            <textarea name="memo" rows="3" placeholder="気づいたことや補足を入力してください"></textarea>
          </label>
          <label class="field">
            <span>画像をアップロード</span>
            <input type="file" name="profile_image" accept="image/*" required />
          </label>
          <div class="button-row">
            <button class="primary-button compact-button" type="submit">画像記録を追加</button>
            <p class="muted profile-inline-note">一覧は日付の新しい順で表示します。</p>
          </div>
          <p class="error-text" data-respondent-record-error></p>
        </form>
      </section>
      <section class="history-sheet stack-gap compact-gap">
        <div class="spread history-sheet-head">
          <div>
            <h4>画像一覧</h4>
            <p class="muted">お客さんがアンケートでアップロードした画像と、管理者追加の画像をまとめて日付順に表示します。</p>
          </div>
          <button class="secondary-button compact-button" type="button" data-image-records-toggle="closed">
            一覧を見る
          </button>
        </div>
        <div class="hidden" data-image-records-panel>
          ${renderRespondentProfileRecords(imageRecords)}
        </div>
      </section>
      <section class="history-sheet stack-gap compact-gap">
        <div class="spread history-sheet-head">
          <div>
            <h4>計測記録</h4>
            <p class="muted">管理者側で日付ごとのウエスト・ヒップ・太ももを入力できます。変化グラフも自動で更新します。</p>
          </div>
        </div>
        ${measurementImportFeedback}
        <form class="stack-gap compact-gap" data-measurement-record-form="${escapeAttribute(respondentId)}">
          <div class="measurement-input-grid">
            <label class="field">
              <span>計測日</span>
              <input type="date" name="entry_date" required />
            </label>
            <label class="field">
              <span>カテゴリ</span>
              <select name="category">${renderMeasurementCategoryOptions("")}</select>
            </label>
            <label class="field">
              <span>ウエスト(cm)</span>
              <input type="number" name="waist" min="0" step="0.1" required />
            </label>
            <label class="field">
              <span>ヒップ(cm)</span>
              <input type="number" name="hip" min="0" step="0.1" required />
            </label>
            <label class="field">
              <span>太もも(cm)</span>
              <input type="number" name="thigh" min="0" step="0.1" required />
            </label>
          </div>
          <div class="button-row">
            <button class="primary-button compact-button" type="submit">計測記録を追加</button>
            <p class="muted profile-inline-note">日付順に一覧化し、グラフは古い日付から線でつなぎます。</p>
          </div>
          <p class="error-text" data-measurement-record-error></p>
        </form>
        <form class="stack-gap compact-gap measurement-import-form" data-measurement-import-form="${escapeAttribute(respondentId)}">
          <label class="field">
            <span>GoogleスプレッドシートURLから取込</span>
            <input
              type="url"
              name="sheet_url"
              inputmode="url"
              placeholder="https://docs.google.com/spreadsheets/d/..."
              required
            />
          </label>
          <div class="button-row">
            <button class="secondary-button compact-button" type="submit">計測記録を取り込む</button>
            <p class="muted profile-inline-note">
              閲覧できるシートURLを貼ると、お名前が一致する行だけを計測記録として取り込みます。
            </p>
          </div>
          <p class="error-text" data-measurement-import-error></p>
        </form>
        ${renderMeasurementSection(measurementRecords)}
      </section>
      ${sheets || `<section class="history-sheet"><div class="empty-state">アンケート回答はまだありません。</div></section>`}
    </div>
  `;
  bindRespondentProfileForm(els.respondentHistory);
  bindRespondentRecordForm(els.respondentHistory);
  bindMeasurementRecordForm(els.respondentHistory);
  bindRespondentRecordDeleteButtons(els.respondentHistory);
  bindRespondentRecordEditButtons(els.respondentHistory);
  bindRespondentRecordEditForms(els.respondentHistory);
  bindMeasurementEditButtons(els.respondentHistory);
  bindMeasurementEditForms(els.respondentHistory);
  bindMeasurementDeleteButtons(els.respondentHistory);
  bindMeasurementImportForm(els.respondentHistory);
  bindImageRecordToggle(els.respondentHistory);
  bindRespondentActionButtons(els.respondentHistory);
  bindImageOpenButtons(els.respondentHistory);
  attachPreviewFallbackHandlers(els.respondentHistory);
}

function renderRespondentCreateForm() {
  state.activeRespondentId = null;
  els.respondentHistory.innerHTML = `
    <section class="history-sheet stack-gap compact-gap">
      <div class="spread history-sheet-head">
        <div>
          <h4>回答者を登録</h4>
          <p class="muted">管理者側で先に回答者を作成できます。あとから同じお名前でアンケート回答が来ると、この回答者に履歴が追加されます。</p>
        </div>
      </div>
      <form class="stack-gap compact-gap" data-respondent-create-form>
        <label class="field">
          <span>お名前</span>
          <input type="text" name="respondent_name" placeholder="例: 山田 花子" required />
        </label>
        <div class="button-row">
          <button class="primary-button compact-button" type="submit">回答者を作成</button>
        </div>
        <p class="error-text" data-respondent-create-error></p>
      </form>
    </section>
  `;
  bindRespondentCreateForm(els.respondentHistory);
}

function renderRespondentProfileRecords(records) {
  if (!records.length) {
    return `<div class="empty-state">画像記録はまだありません。</div>`;
  }
  return `
    <div class="profile-record-list">
      ${records
        .map(
          (record) => `
            <article class="profile-record-card">
              ${
                record.image?.url
                  ? `
                    <button
                      class="profile-record-image profile-record-image-button"
                      type="button"
                      data-image-open="${escapeAttribute(imageOpenSource(record.image))}"
                      data-image-fallback="${escapeAttribute(directOpenImageSource(record.image))}"
                      data-image-title="${escapeAttribute(record.title || "画像記録")}"
                      data-image-meta="${escapeAttribute(profileRecordSourceLabel(record))}"
                    >
                      <img src="${escapeAttribute(imagePreviewSource(record.image))}" data-fallback-src="${escapeAttribute(directImageSource(record.image))}" alt="${escapeAttribute(record.title || "画像記録")}" loading="lazy" referrerpolicy="no-referrer" />
                    </button>
                  `
                  : `
                    <div class="profile-record-image placeholder">NO IMAGE</div>
                  `
              }
              <div class="profile-record-body">
                <div class="spread profile-record-head">
                  <div>
                    <h5>${escapeHtml(record.title || "無題")}</h5>
                    <p class="muted">${escapeHtml(formatDateOnly(record.date || ""))}</p>
                    <p class="muted">${escapeHtml(profileRecordSourceLabel(record))}</p>
                  </div>
                  <div class="button-row">
                    <button
                      class="ghost-button compact-button"
                      type="button"
                      data-profile-record-edit-toggle="${escapeAttribute(record.sourceType)}:${escapeAttribute(String(record.recordId))}"
                    >
                      編集
                    </button>
                    ${
                      record.deletable
                        ? `
                          <button
                            class="danger-button compact-button"
                            type="button"
                            data-profile-record-delete="${escapeAttribute(record.sourceType)}:${escapeAttribute(String(record.recordId))}"
                          >
                            削除
                          </button>
                        `
                        : ""
                    }
                  </div>
                </div>
                <p class="profile-record-memo">${escapeHtml(record.memo || "メモなし")}</p>
                ${
                  record.image?.originalName
                    ? `
                      <div class="history-file-links profile-record-link-list">
                        <button
                          type="button"
                          data-image-open="${escapeAttribute(imageOpenSource(record.image))}"
                          data-image-fallback="${escapeAttribute(directOpenImageSource(record.image))}"
                          data-image-title="${escapeAttribute(record.title || record.image.originalName || "画像記録")}"
                          data-image-meta="${escapeAttribute(profileRecordSourceLabel(record))}"
                        >
                          ${escapeHtml(record.image.originalName)}
                        </button>
                      </div>
                    `
                    : ""
                }
                <form
                  class="profile-record-edit-form hidden"
                  data-profile-record-edit-form="${escapeAttribute(record.sourceType)}:${escapeAttribute(String(record.recordId))}"
                >
                  <div class="grid-two">
                    <label class="field">
                      <span>タイトル</span>
                      <input type="text" name="title" value="${escapeAttribute(record.title || "")}" required />
                    </label>
                    <label class="field">
                      <span>日付</span>
                      <input type="date" name="entry_date" value="${escapeAttribute(record.date || "")}" required />
                    </label>
                  </div>
                  <label class="field">
                    <span>メモ</span>
                    <textarea name="memo" rows="3" placeholder="管理者メモを入力してください">${escapeHtml(record.memo || "")}</textarea>
                  </label>
                  <div class="button-row">
                    <button class="primary-button compact-button" type="submit">保存</button>
                    <button
                      class="secondary-button compact-button"
                      type="button"
                      data-profile-record-edit-cancel="${escapeAttribute(record.sourceType)}:${escapeAttribute(String(record.recordId))}"
                    >
                      閉じる
                    </button>
                  </div>
                  <p class="error-text" data-profile-record-edit-error></p>
                </form>
              </div>
            </article>
          `
        )
        .join("")}
    </div>
  `;
}

function profileRecordSourceLabel(record) {
  if (record.sourceType === "response") {
    return [record.formTitle, record.sourceLabel].filter(Boolean).join(" / ") || "アンケート画像";
  }
  return record.sourceLabel || "管理者追加";
}

function renderMeasurementSection(records) {
  return `
    <div class="measurement-dashboard">
      <div class="measurement-chart-card">
        <div class="spread">
          <h5>変化グラフ</h5>
          <div class="measurement-legend">
            <span class="waist">ウエスト</span>
            <span class="hip">ヒップ</span>
            <span class="thigh">太もも</span>
          </div>
        </div>
        ${renderMeasurementGraph(records)}
      </div>
      <div class="measurement-table-card">
        <div class="spread">
          <h5>計測一覧</h5>
          <p class="muted">${records.length}件</p>
        </div>
        ${renderMeasurementTable(records)}
      </div>
    </div>
  `;
}

function renderMeasurementGraph(records) {
  if (!records.length) {
    return `<div class="empty-state measurement-empty">計測記録を追加すると変化グラフが表示されます。</div>`;
  }
  const ordered = [...records].sort((left, right) => String(left.date || "").localeCompare(String(right.date || "")));
  const values = ordered.flatMap((record) => [record.waist, record.hip, record.thigh]).map((value) => Number(value));
  const min = Math.min(...values);
  const max = Math.max(...values);
  const yMin = Math.floor((min - 1) / 5) * 5;
  const yMax = Math.ceil((max + 1) / 5) * 5;
  const width = 720;
  const height = 260;
  const margin = { top: 18, right: 22, bottom: 36, left: 40 };
  const innerWidth = width - margin.left - margin.right;
  const innerHeight = height - margin.top - margin.bottom;
  const hasSinglePoint = ordered.length === 1;
  const divisor = Math.max(1, ordered.length - 1);
  const yRange = Math.max(1, yMax - yMin);
  const xFor = (index) => (hasSinglePoint ? margin.left + innerWidth / 2 : margin.left + (innerWidth * index) / divisor);
  const yFor = (value) => margin.top + ((yMax - Number(value)) / yRange) * innerHeight;
  const makePoints = (key) =>
    ordered
      .map((record, index) => `${xFor(index).toFixed(1)},${yFor(record[key]).toFixed(1)}`)
      .join(" ");
  const guides = Array.from({ length: 4 }, (_, index) => {
    const ratio = index / 3;
    const y = margin.top + innerHeight * ratio;
    const label = (yMax - yRange * ratio).toFixed(0);
    return `
      <g>
        <line x1="${margin.left}" y1="${y.toFixed(1)}" x2="${(width - margin.right).toFixed(1)}" y2="${y.toFixed(1)}" />
        <text x="${margin.left - 10}" y="${(y + 4).toFixed(1)}">${escapeHtml(label)}</text>
      </g>
    `;
  }).join("");
  const xLabels = ordered
    .map((record, index) => {
      const x = xFor(index);
      return `<text x="${x.toFixed(1)}" y="${height - 10}" text-anchor="middle">${escapeHtml(formatDateOnly(record.date || ""))}</text>`;
    })
    .join("");
  return `
    <div class="measurement-chart-wrap">
      <svg class="measurement-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="計測値の変化グラフ">
        <g class="measurement-guides">${guides}</g>
        <polyline class="measurement-line waist" points="${makePoints("waist")}" />
        <polyline class="measurement-line hip" points="${makePoints("hip")}" />
        <polyline class="measurement-line thigh" points="${makePoints("thigh")}" />
        ${ordered
          .flatMap((record, index) => [
            `<circle class="measurement-point waist" cx="${xFor(index).toFixed(1)}" cy="${yFor(record.waist).toFixed(1)}" r="4" />`,
            `<circle class="measurement-point hip" cx="${xFor(index).toFixed(1)}" cy="${yFor(record.hip).toFixed(1)}" r="4" />`,
            `<circle class="measurement-point thigh" cx="${xFor(index).toFixed(1)}" cy="${yFor(record.thigh).toFixed(1)}" r="4" />`,
          ])
          .join("")}
        <g class="measurement-xlabels">${xLabels}</g>
      </svg>
    </div>
  `;
}

function renderMeasurementTable(records) {
  if (!records.length) {
    return `<div class="empty-state measurement-empty">計測記録はまだありません。</div>`;
  }
  return `
    <div class="history-table-wrap">
      <table class="answer-table measurement-table">
        <thead>
          <tr>
            <th>日付</th>
            <th>カテゴリ</th>
            <th>ウエスト</th>
            <th>ヒップ</th>
            <th>太もも</th>
            <th>画像</th>
            <th>操作</th>
          </tr>
        </thead>
        <tbody>
          ${records
            .map(
              (record) => `
                <tr>
                  <td>${escapeHtml(formatDateOnly(record.date || ""))}</td>
                  <td>${escapeHtml(record.category || "-")}</td>
                  <td>${escapeHtml(measurementValueLabel(record.waist))} cm</td>
                  <td>${escapeHtml(measurementValueLabel(record.hip))} cm</td>
                  <td>${escapeHtml(measurementValueLabel(record.thigh))} cm</td>
                  <td>${renderMeasurementImageLinks(record.imageLinks || [])}</td>
                  <td>
                    <div class="button-row measurement-actions">
                      <button
                        class="ghost-button compact-button measurement-action-button"
                        type="button"
                        data-measurement-edit-toggle="${escapeAttribute(String(record.recordId))}"
                      >
                        編集
                      </button>
                      <button
                        class="danger-button compact-button measurement-action-button"
                        type="button"
                        data-measurement-delete="${escapeAttribute(String(record.recordId))}"
                      >
                        削除
                      </button>
                    </div>
                  </td>
                </tr>
                <tr class="measurement-edit-row hidden" data-measurement-edit-row="${escapeAttribute(String(record.recordId))}">
                  <td colspan="7">
                    <form class="measurement-edit-form" data-measurement-edit-form="${escapeAttribute(String(record.recordId))}">
                      <div class="measurement-input-grid">
                        <label class="field">
                          <span>計測日</span>
                          <input type="date" name="entry_date" value="${escapeAttribute(record.date || "")}" required />
                        </label>
                        <label class="field">
                          <span>カテゴリ</span>
                          <select name="category">${renderMeasurementCategoryOptions(record.category || "")}</select>
                        </label>
                        <label class="field">
                          <span>ウエスト(cm)</span>
                          <input type="number" name="waist" min="0" step="0.1" value="${escapeAttribute(measurementValueLabel(record.waist))}" required />
                        </label>
                        <label class="field">
                          <span>ヒップ(cm)</span>
                          <input type="number" name="hip" min="0" step="0.1" value="${escapeAttribute(measurementValueLabel(record.hip))}" required />
                        </label>
                        <label class="field">
                          <span>太もも(cm)</span>
                          <input type="number" name="thigh" min="0" step="0.1" value="${escapeAttribute(measurementValueLabel(record.thigh))}" required />
                        </label>
                      </div>
                      <div class="grid-two measurement-upload-grid">
                        <label class="field">
                          <span>画像を追加(任意)</span>
                          <input type="file" name="measurement_image" accept="image/*" multiple />
                        </label>
                        <label class="field">
                          <span>画像タイトル(任意)</span>
                          <input type="text" name="measurement_image_title" maxlength="120" placeholder="未入力なら計測画像の日付名" />
                        </label>
                      </div>
                      <p class="muted measurement-upload-note">画像は複数枚追加できます。同じ日付の画像リンクとして一覧に表示します。</p>
                      <div class="button-row">
                        <button class="primary-button compact-button" type="submit">保存</button>
                        <button
                          class="secondary-button compact-button"
                          type="button"
                          data-measurement-edit-cancel="${escapeAttribute(String(record.recordId))}"
                        >
                          閉じる
                        </button>
                      </div>
                      <p class="error-text" data-measurement-edit-error></p>
                    </form>
                  </td>
                </tr>
              `
            )
            .join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderMeasurementValueInline(record, compact = false) {
  const className = compact ? "measurement-inline-values compact" : "measurement-inline-values";
  return `
    <div class="${className}">
      <span>W ${escapeHtml(measurementValueLabel(record.waist))}</span>
      <span>H ${escapeHtml(measurementValueLabel(record.hip))}</span>
      <span>T ${escapeHtml(measurementValueLabel(record.thigh))}</span>
    </div>
  `;
}

function renderMeasurementManagerList(records) {
  if (!records.length) {
    els.measurementManagerFeedback.innerHTML = "";
    els.measurementManagerList.innerHTML = `<div class="empty-state">計測記録はまだありません。</div>`;
    return;
  }
  els.measurementManagerFeedback.innerHTML = `
    <div class="note-box compact-note measurement-manager-summary">
      <strong>表示件数</strong>
      <div>${records.length}件の計測記録を表示しています。</div>
    </div>
  `;
  els.measurementManagerList.innerHTML = `
    <section class="history-sheet stack-gap compact-gap">
      <div class="spread history-sheet-head">
        <div>
          <h4>全体の計測一覧</h4>
          <p class="muted">ここで編集した内容は、各回答者の計測記録にも反映されます。</p>
        </div>
      </div>
      <div class="history-table-wrap">
        <table class="answer-table measurement-table measurement-manager-table">
          <thead>
            <tr>
              <th>日付</th>
              <th>カテゴリ</th>
              <th>お名前</th>
              <th>計測値</th>
              <th>画像</th>
              <th>操作</th>
            </tr>
          </thead>
          <tbody>
            ${records
              .map(
                (record) => `
                  <tr>
                    <td>${escapeHtml(formatDateOnly(record.date || ""))}</td>
                    <td>${escapeHtml(record.category || "-")}</td>
                    <td>
                      <button
                        class="inline-link-button"
                        type="button"
                        data-measurement-open-respondent="${escapeAttribute(record.respondentId || "")}"
                      >
                        ${escapeHtml(record.respondentName || "-")}
                      </button>
                    </td>
                    <td>${renderMeasurementValueInline(record, true)}</td>
                    <td>${renderMeasurementImageLinks(record.imageLinks || [])}</td>
                    <td>
                      <div class="button-row measurement-actions">
                        <button
                          class="ghost-button compact-button measurement-action-button"
                          type="button"
                          data-measurement-edit-toggle="${escapeAttribute(String(record.recordId))}"
                        >
                          編集
                        </button>
                        <button
                          class="danger-button compact-button measurement-action-button"
                          type="button"
                          data-measurement-delete="${escapeAttribute(String(record.recordId))}"
                          data-measurement-respondent-id="${escapeAttribute(record.respondentId || "")}"
                        >
                          削除
                        </button>
                      </div>
                    </td>
                  </tr>
                  <tr class="measurement-edit-row hidden" data-measurement-edit-row="${escapeAttribute(String(record.recordId))}">
                    <td colspan="6">
                      <form
                        class="measurement-edit-form"
                        data-measurement-edit-form="${escapeAttribute(String(record.recordId))}"
                        data-measurement-respondent-id="${escapeAttribute(record.respondentId || "")}"
                      >
                        <div class="measurement-input-grid">
                          <label class="field">
                            <span>計測日</span>
                            <input type="date" name="entry_date" value="${escapeAttribute(record.date || "")}" required />
                          </label>
                          <label class="field">
                            <span>カテゴリ</span>
                            <select name="category">${renderMeasurementCategoryOptions(record.category || "")}</select>
                          </label>
                          <label class="field">
                            <span>ウエスト(cm)</span>
                            <input type="number" name="waist" min="0" step="0.1" value="${escapeAttribute(measurementValueLabel(record.waist))}" required />
                          </label>
                          <label class="field">
                            <span>ヒップ(cm)</span>
                            <input type="number" name="hip" min="0" step="0.1" value="${escapeAttribute(measurementValueLabel(record.hip))}" required />
                          </label>
                          <label class="field">
                            <span>太もも(cm)</span>
                            <input type="number" name="thigh" min="0" step="0.1" value="${escapeAttribute(measurementValueLabel(record.thigh))}" required />
                          </label>
                        </div>
                        <div class="grid-two measurement-upload-grid">
                          <label class="field">
                            <span>画像を追加(任意)</span>
                            <input type="file" name="measurement_image" accept="image/*" multiple />
                          </label>
                          <label class="field">
                            <span>画像タイトル(任意)</span>
                            <input type="text" name="measurement_image_title" maxlength="120" placeholder="未入力なら計測画像の日付名" />
                          </label>
                        </div>
                        <p class="muted measurement-upload-note">画像は複数枚追加できます。同じ日付の画像リンクとして一覧に表示します。</p>
                        <div class="button-row">
                          <button class="primary-button compact-button" type="submit">保存</button>
                          <button
                            class="secondary-button compact-button"
                            type="button"
                            data-measurement-edit-cancel="${escapeAttribute(String(record.recordId))}"
                          >
                            閉じる
                          </button>
                        </div>
                        <p class="error-text" data-measurement-edit-error></p>
                      </form>
                    </td>
                  </tr>
                `
              )
              .join("")}
          </tbody>
        </table>
      </div>
    </section>
  `;
  bindMeasurementEditButtons(els.measurementManagerList);
  bindMeasurementEditForms(els.measurementManagerList);
  bindMeasurementDeleteButtons(els.measurementManagerList);
  bindMeasurementOpenRespondentButtons(els.measurementManagerList);
  bindImageOpenButtons(els.measurementManagerList);
  attachPreviewFallbackHandlers(els.measurementManagerList);
}

function renderMeasurementImageLinks(links) {
  if (!links.length) {
    return `<span class="muted">-</span>`;
  }
  return `
    <div class="history-file-links compact measurement-image-links">
      ${links
        .map(
          (link, index) => `
            <button
              type="button"
              data-image-open="${escapeAttribute(adminImageProxyUrl(link.url || link.previewUrl || ""))}"
              data-image-fallback="${escapeAttribute(link.url || link.previewUrl || "")}"
              data-image-title="${escapeAttribute(link.originalName || link.label || `画像 ${index + 1}`)}"
              data-image-meta="${escapeAttribute(link.label || `画像 ${index + 1}`)}"
              title="${escapeAttribute(link.originalName || link.label || `画像 ${index + 1}`)}"
            >
              ${escapeHtml(link.label || `画像 ${index + 1}`)}
            </button>
          `
        )
        .join("")}
    </div>
  `;
}

function buildRespondentHistorySheets(items) {
  const groups = new Map();
  items.forEach((item) => {
    const key = String(item.formId);
    if (!groups.has(key)) {
      groups.set(key, {
        formId: item.formId,
        formTitle: item.formTitle,
        items: [],
      });
    }
    groups.get(key).items.push(item);
  });

  return Array.from(groups.values())
    .map((group) => {
      const answerLabels = [];
      const seenLabels = new Set();
      let hasFiles = false;

      group.items.forEach((item) => {
        if (item.files?.length) {
          hasFiles = true;
        }
        item.answers.forEach((answer) => {
          if (!seenLabels.has(answer.label)) {
            seenLabels.add(answer.label);
            answerLabels.push(answer.label);
          }
        });
      });

      const headerCells = [
        "<th>タイムスタンプ</th>",
        "<th>分類</th>",
        ...answerLabels.map((label) => `<th>${escapeHtml(label)}</th>`),
      ];
      if (hasFiles) {
        headerCells.push("<th>画像</th>");
      }

      const bodyRows = group.items
        .map((item) => {
          const answersByLabel = new Map(item.answers.map((answer) => [answer.label, answer.value || "-"]));
          const cells = [
            `<td>${escapeHtml(formatDate(item.createdAt))}</td>`,
            `<td>${escapeHtml(item.category || "-")}</td>`,
            ...answerLabels.map((label) => `<td>${escapeHtml(answersByLabel.get(label) || "-")}</td>`),
          ];
          if (hasFiles) {
            cells.push(`<td>${renderResponseFileLinks(item.files || [])}</td>`);
          }
          return `<tr>${cells.join("")}</tr>`;
        })
        .join("");

      return `
        <section class="history-sheet stack-gap compact-gap">
          <div class="spread history-sheet-head">
            <div>
              <h4>${escapeHtml(group.formTitle)}</h4>
              <p class="muted">回答 ${group.items.length} 件を横長の表で表示しています。</p>
            </div>
          </div>
          <div class="history-table-wrap">
            <table class="answer-table history-answer-table sheet-table">
              <thead>
                <tr>${headerCells.join("")}</tr>
              </thead>
              <tbody>${bodyRows}</tbody>
            </table>
          </div>
        </section>
      `;
    })
    .join("");
}

function renderResponseFileLinks(files) {
  if (!files.length) {
    return "-";
  }
  const groups = new Map();
  files.forEach((file) => {
    const key = file.label || "画像";
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key).push(file);
  });
  const links = Array.from(groups.entries()).flatMap(([label, items]) =>
    items.map(
      (file, index) => `
        <button
          type="button"
          data-image-open="${escapeAttribute(adminImageProxyUrl(file.url || file.previewUrl || ""))}"
          data-image-fallback="${escapeAttribute(file.url || file.previewUrl || "")}"
          data-image-title="${escapeAttribute(file.originalName || compactFileLinkLabel(label, index, items.length))}"
          data-image-meta="${escapeAttribute(label || "画像")}"
          title="${escapeAttribute(file.originalName || compactFileLinkLabel(label, index, items.length))}"
        >
          ${escapeHtml(compactFileLinkLabel(label, index, items.length))}
        </button>
      `
    )
  );
  return `<div class="history-file-links compact">${links.join("")}</div>`;
}

function compactFileLinkLabel(label, index, count) {
  const rawLabel = String(label || "画像").trim();
  const match = rawLabel.match(/^計測写真[\(（](.+?)[\)）]$/);
  const baseLabel = match ? match[1] : rawLabel.replace(/^計測写真/, "").trim() || "画像";
  return count > 1 ? `${baseLabel} ${index + 1}` : baseLabel;
}

function openImageModal(src, title = "", meta = "", fallbackSrc = "") {
  const resolved = String(src || "").trim();
  if (!resolved || !els.imageModal || !els.imageModalImg) {
    return;
  }
  els.imageModalImg.dataset.fallbackSrc = String(fallbackSrc || "").trim();
  els.imageModalImg.dataset.fallbackApplied = "";
  els.imageModalImg.src = resolved;
  els.imageModalImg.alt = title || "画像プレビュー";
  if (els.imageModalTitle) {
    els.imageModalTitle.textContent = title || "画像プレビュー";
  }
  if (els.imageModalMeta) {
    els.imageModalMeta.textContent = meta || "";
  }
  els.imageModal.classList.remove("hidden");
  els.imageModal.setAttribute("aria-hidden", "false");
}

function closeImageModal() {
  if (!els.imageModal || !els.imageModalImg) {
    return;
  }
  els.imageModal.classList.add("hidden");
  els.imageModal.setAttribute("aria-hidden", "true");
  els.imageModalImg.removeAttribute("src");
  delete els.imageModalImg.dataset.fallbackSrc;
  delete els.imageModalImg.dataset.fallbackApplied;
}

function bindImageOpenButtons(container) {
  if (!container) {
    return;
  }
  container.querySelectorAll("[data-image-open]").forEach((button) => {
    button.addEventListener("click", (event) => {
      event.preventDefault();
      openImageModal(
        button.dataset.imageOpen,
        button.dataset.imageTitle || "",
        button.dataset.imageMeta || "",
        button.dataset.imageFallback || ""
      );
    });
  });
}

function setActivePanel(name) {
  els.tabLinks.forEach((button) => {
    button.classList.toggle("active", button.dataset.panelTarget === name);
  });
  els.panels.forEach((panel) => {
    panel.classList.toggle("active", panel.id === `panel-${name}`);
  });
}

async function api(url, options = {}) {
  return runtime.request(url, options);
}

function bindRespondentProfileForm(container) {
  const form = container.querySelector("[data-respondent-profile-form]");
  if (!form) {
    return;
  }
  bindTicketStampEditor(form);
  form.addEventListener("submit", onRespondentProfileSubmit);
}

function bindTicketStampEditor(form) {
  const typeSelect = form.querySelector('[name="current_ticket_book_type"]');
  const countInput = form.querySelector('[name="current_ticket_stamp_count"]');
  const manualEnabledInput = form.querySelector('[name="current_ticket_stamp_manual_enabled"]');
  const card = form.querySelector("[data-ticket-stamp-card]");
  const summary = form.querySelector("[data-ticket-stamp-summary]");
  const resetButton = form.querySelector("[data-ticket-stamp-reset]");
  if (!typeSelect || !countInput || !manualEnabledInput || !card || !summary) {
    return;
  }

  const autoValue = Number.parseInt(summary.dataset.ticketStampAutoValue || "0", 10) || 0;

  const sync = () => {
    const ticketBookType = typeSelect.value.trim();
    const max = ticketBookTypeMax(ticketBookType);
    const manualEnabled = manualEnabledInput.value === "1";
    let count = Number.parseInt(countInput.value || "0", 10);
    if (!Number.isFinite(count) || count < 0) {
      count = 0;
    }
    const autoCount = max ? Math.min(autoValue, max) : 0;
    if (max && count > max) {
      count = max;
    }
    const displayCount = manualEnabled ? count : autoCount;
    countInput.value = manualEnabled ? String(count) : "0";
    card.innerHTML = ticketStampCardButtons(ticketBookType, displayCount);
    summary.textContent = ticketStampSummary(ticketBookType, displayCount, manualEnabled, autoCount);
  };

  typeSelect.addEventListener("change", sync);
  card.addEventListener("click", (event) => {
    const button = event.target.closest("[data-ticket-stamp-choice]");
    if (!button) {
      return;
    }
    countInput.value = button.dataset.ticketStampChoice || "0";
    manualEnabledInput.value = "1";
    sync();
  });
  resetButton?.addEventListener("click", () => {
    manualEnabledInput.value = "0";
    countInput.value = "0";
    sync();
  });
  sync();
}

function bindRespondentCreateForm(container) {
  const form = container.querySelector("[data-respondent-create-form]");
  if (!form) {
    return;
  }
  form.addEventListener("submit", onRespondentCreateSubmit);
}

async function onRespondentCreateSubmit(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const submitButton = form.querySelector('button[type="submit"]');
  const errorNode = form.querySelector("[data-respondent-create-error]");
  const respondentName = form.querySelector('[name="respondent_name"]').value.trim();

  if (errorNode) {
    errorNode.textContent = "";
  }
  submitButton.disabled = true;
  submitButton.textContent = "作成中...";

  try {
    const result = await api("/api/admin/respondents/create", {
      method: "POST",
      body: { name: respondentName },
    });
    state.activeRespondentId = result.respondent.respondentId;
    await bootstrap();
    setActivePanel("respondents");
    await loadRespondentHistory(result.respondent.respondentId, false);
  } catch (error) {
    if (errorNode) {
      errorNode.textContent = error.message;
    }
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "回答者を作成";
  }
}

async function onRespondentProfileSubmit(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const respondentId = form.dataset.respondentProfileForm;
  const submitButton = form.querySelector('button[type="submit"]');
  const errorNode = form.querySelector("[data-respondent-profile-error]");
  const respondentName = form.querySelector('[name="respondent_name"]').value.trim();
  const ticketSheetManualValue = form.querySelector('[name="ticket_sheet_manual_value"]')?.value.trim() || "";
  const ticketBookType = form.querySelector('[name="current_ticket_book_type"]')?.value.trim() || "";
  const ticketStampCount = form.querySelector('[name="current_ticket_stamp_count"]')?.value.trim() || "0";
  const ticketStampManualEnabled = form.querySelector('[name="current_ticket_stamp_manual_enabled"]')?.value.trim() || "0";

  if (errorNode) {
    errorNode.textContent = "";
  }
  submitButton.disabled = true;
  submitButton.textContent = "保存中...";

  try {
    const result = await api(`/api/admin/respondents/${encodeURIComponent(respondentId)}/profile`, {
      method: "POST",
      body: {
        name: respondentName,
        ticketSheet: ticketSheetManualValue,
        ticketBookType,
        ticketStampCount,
        ticketStampManualEnabled,
        formId: state.selectedRespondentFormId || "",
      },
    });
    state.activeRespondentId = result.respondentId;
    await bootstrap();
    setActivePanel("respondents");
    await loadRespondentHistory(result.respondentId, false);
  } catch (error) {
    if (errorNode) {
      errorNode.textContent = error.message;
    }
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "回答者情報を保存";
  }
}

function bindRespondentRecordForm(container) {
  const form = container.querySelector("[data-respondent-record-form]");
  if (!form) {
    return;
  }
  form.addEventListener("submit", onRespondentRecordSubmit);
}

function bindMeasurementRecordForm(container) {
  const form = container.querySelector("[data-measurement-record-form]");
  if (!form) {
    return;
  }
  form.addEventListener("submit", onMeasurementRecordSubmit);
}

function bindMeasurementImportForm(container) {
  const form = container.querySelector("[data-measurement-import-form]");
  if (!form) {
    return;
  }
  form.addEventListener("submit", onMeasurementImportSubmit);
}

async function onRespondentRecordSubmit(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const respondentId = form.dataset.respondentRecordForm;
  const submitButton = form.querySelector('button[type="submit"]');
  const errorNode = form.querySelector("[data-respondent-record-error]");
  const formData = new FormData(form);
  formData.append("form_id", state.selectedRespondentFormId || "");

  if (errorNode) {
    errorNode.textContent = "";
  }
  submitButton.disabled = true;
  submitButton.textContent = "追加中...";

  try {
    await multipartApi(`/api/admin/respondents/${encodeURIComponent(respondentId)}/profile-records`, formData);
    await bootstrap();
    setActivePanel("respondents");
    await loadRespondentHistory(state.activeRespondentId || respondentId, false);
  } catch (error) {
    if (errorNode) {
      errorNode.textContent = error.message;
    }
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "画像記録を追加";
  }
}

async function onMeasurementRecordSubmit(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const respondentId = form.dataset.measurementRecordForm;
  const submitButton = form.querySelector('button[type="submit"]');
  const errorNode = form.querySelector("[data-measurement-record-error]");

  if (errorNode) {
    errorNode.textContent = "";
  }
  submitButton.disabled = true;
  submitButton.textContent = "追加中...";

  try {
    await api(`/api/admin/respondents/${encodeURIComponent(respondentId)}/measurements`, {
      method: "POST",
      body: {
        formId: state.selectedRespondentFormId || "",
        entryDate: form.querySelector('[name="entry_date"]').value,
        category: form.querySelector('[name="category"]').value,
        waist: form.querySelector('[name="waist"]').value,
        hip: form.querySelector('[name="hip"]').value,
        thigh: form.querySelector('[name="thigh"]').value,
      },
    });
    await bootstrap();
    setActivePanel("respondents");
    await loadRespondentHistory(state.activeRespondentId || respondentId, false);
  } catch (error) {
    if (errorNode) {
      errorNode.textContent = error.message;
    }
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "計測記録を追加";
  }
}

async function onMeasurementImportSubmit(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const respondentId = form.dataset.measurementImportForm;
  const submitButton = form.querySelector('button[type="submit"]');
  const errorNode = form.querySelector("[data-measurement-import-error]");
  const sheetUrl = form.querySelector('[name="sheet_url"]').value.trim();

  if (errorNode) {
    errorNode.textContent = "";
  }
  state.measurementImportFeedback = null;
  submitButton.disabled = true;
  submitButton.textContent = "取込中...";

  try {
    const result = await api(`/api/admin/respondents/${encodeURIComponent(respondentId)}/measurements/import-sheet`, {
      method: "POST",
      body: {
        sheetUrl,
      },
    });
    state.measurementImportFeedback = {
      respondentId,
      title: "取込が完了しました",
      message: `追加 ${result.importedCount || 0}件 / 更新 ${result.updatedCount || 0}件 / 既存のまま ${result.skippedCount || 0}件`,
    };
    await bootstrap();
    setActivePanel("respondents");
    await loadRespondentHistory(state.activeRespondentId || respondentId, false);
  } catch (error) {
    if (errorNode) {
      errorNode.textContent = error.message;
    }
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "計測記録を取り込む";
  }
}

async function onMeasurementManagerSubmit(event) {
  event.preventDefault();
  const respondent = findRespondentByName(els.measurementRespondentInput.value.trim());
  if (els.measurementManagerError) {
    els.measurementManagerError.textContent = "";
  }
  if (!respondent) {
    els.measurementManagerError.textContent = "既存の回答者名を選択してください。";
    return;
  }
  const submitButton = els.measurementManagerForm.querySelector('button[type="submit"]');
  submitButton.disabled = true;
  submitButton.textContent = "追加中...";
  try {
    await api(`/api/admin/respondents/${encodeURIComponent(respondent.respondentId)}/measurements`, {
      method: "POST",
      body: {
        entryDate: els.measurementEntryDate.value,
        category: els.measurementEntryCategory.value,
        waist: els.measurementEntryWaist.value,
        hip: els.measurementEntryHip.value,
        thigh: els.measurementEntryThigh.value,
      },
    });
    els.measurementEntryDate.value = "";
    els.measurementEntryCategory.value = "";
    els.measurementEntryWaist.value = "";
    els.measurementEntryHip.value = "";
    els.measurementEntryThigh.value = "";
    await bootstrap();
    setActivePanel("measurements");
  } catch (error) {
    els.measurementManagerError.textContent = error.message;
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "計測記録を追加";
  }
}

function bindRespondentActionButtons(container) {
  container.querySelectorAll("[data-respondent-edit]").forEach((button) => {
    button.addEventListener("click", async (event) => {
      event.stopPropagation();
      await editRespondent(button.dataset.respondentEdit);
    });
  });
  container.querySelectorAll("[data-respondent-delete]").forEach((button) => {
    button.addEventListener("click", async (event) => {
      event.stopPropagation();
      await deleteRespondent(button.dataset.respondentDelete, button.dataset.respondentName || "");
    });
  });
}

function bindRespondentRecordDeleteButtons(container) {
  container.querySelectorAll("[data-profile-record-delete]").forEach((button) => {
    button.addEventListener("click", async (event) => {
      event.preventDefault();
      const target = button.dataset.profileRecordDelete;
      if (!target || !state.activeRespondentId) {
        return;
      }
      const [sourceType, recordId] = target.split(":");
      const confirmed = window.confirm("この画像記録を削除します。");
      if (!confirmed) {
        return;
      }
      await deleteRespondentProfileRecord(state.activeRespondentId, sourceType, recordId);
    });
  });
}

async function deleteRespondentProfileRecord(respondentId, sourceType, recordId) {
  try {
    await api(
      `/api/admin/respondents/${encodeURIComponent(respondentId)}/profile-records/${encodeURIComponent(sourceType)}/${recordId}/delete`,
      {
      method: "POST",
      body: {},
      }
    );
    await bootstrap();
    setActivePanel("respondents");
    await loadRespondentHistory(respondentId, false);
  } catch (error) {
    window.alert(error.message);
  }
}

function bindRespondentRecordEditButtons(container) {
  container.querySelectorAll("[data-profile-record-edit-toggle]").forEach((button) => {
    button.addEventListener("click", () => {
      const target = button.dataset.profileRecordEditToggle;
      const form = container.querySelector(`[data-profile-record-edit-form="${cssEscape(target)}"]`);
      form?.classList.toggle("hidden");
    });
  });
  container.querySelectorAll("[data-profile-record-edit-cancel]").forEach((button) => {
    button.addEventListener("click", () => {
      const target = button.dataset.profileRecordEditCancel;
      const form = container.querySelector(`[data-profile-record-edit-form="${cssEscape(target)}"]`);
      form?.classList.add("hidden");
    });
  });
}

function bindRespondentRecordEditForms(container) {
  container.querySelectorAll("[data-profile-record-edit-form]").forEach((form) => {
    form.addEventListener("submit", onRespondentRecordEditSubmit);
  });
}

async function onRespondentRecordEditSubmit(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const target = form.dataset.profileRecordEditForm;
  const [sourceType, recordId] = target.split(":");
  const submitButton = form.querySelector('button[type="submit"]');
  const errorNode = form.querySelector("[data-profile-record-edit-error]");
  if (!state.activeRespondentId) {
    return;
  }

  if (errorNode) {
    errorNode.textContent = "";
  }
  submitButton.disabled = true;
  submitButton.textContent = "保存中...";

  try {
    await api(
      `/api/admin/respondents/${encodeURIComponent(state.activeRespondentId)}/profile-records/${encodeURIComponent(sourceType)}/${recordId}/update`,
      {
        method: "POST",
        body: {
          title: form.querySelector('[name="title"]').value.trim(),
          entryDate: form.querySelector('[name="entry_date"]').value,
          memo: form.querySelector('[name="memo"]').value,
        },
      }
    );
    await bootstrap();
    setActivePanel("respondents");
    await loadRespondentHistory(state.activeRespondentId, false);
  } catch (error) {
    if (errorNode) {
      errorNode.textContent = error.message;
    }
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "保存";
  }
}

function bindImageRecordToggle(container) {
  const button = container.querySelector("[data-image-records-toggle]");
  const panel = container.querySelector("[data-image-records-panel]");
  if (!button || !panel) {
    return;
  }
  button.addEventListener("click", () => {
    const isHidden = panel.classList.contains("hidden");
    panel.classList.toggle("hidden", !isHidden);
    button.dataset.imageRecordsToggle = isHidden ? "open" : "closed";
    button.textContent = isHidden ? "一覧を閉じる" : "一覧を見る";
  });
}

function bindMeasurementEditButtons(container) {
  container.querySelectorAll("[data-measurement-edit-toggle]").forEach((button) => {
    button.addEventListener("click", () => {
      const target = button.dataset.measurementEditToggle;
      const row = container.querySelector(`[data-measurement-edit-row="${cssEscape(target)}"]`);
      row?.classList.toggle("hidden");
    });
  });
  container.querySelectorAll("[data-measurement-edit-cancel]").forEach((button) => {
    button.addEventListener("click", () => {
      const target = button.dataset.measurementEditCancel;
      const row = container.querySelector(`[data-measurement-edit-row="${cssEscape(target)}"]`);
      row?.classList.add("hidden");
    });
  });
}

function bindMeasurementEditForms(container) {
  container.querySelectorAll("[data-measurement-edit-form]").forEach((form) => {
    form.addEventListener("submit", onMeasurementEditSubmit);
  });
}

function buildMeasurementImageTitle(entryDate, explicitTitle = "") {
  const title = explicitTitle.trim();
  if (title) {
    return title;
  }
  if (entryDate) {
    return `計測画像 ${entryDate.replace(/-/g, "/")}`;
  }
  return "計測画像";
}

function buildMeasurementImageTitleForIndex(entryDate, explicitTitle, index, total) {
  const baseTitle = buildMeasurementImageTitle(entryDate, explicitTitle);
  if (total <= 1) {
    return baseTitle;
  }
  return `${baseTitle} ${index + 1}`;
}

async function maybeUploadMeasurementImageRecord(respondentId, form) {
  const fileInput = form.querySelector('[name="measurement_image"]');
  const selectedFiles = Array.from(fileInput?.files || []);
  if (!selectedFiles.length) {
    return 0;
  }
  const entryDate = form.querySelector('[name="entry_date"]').value;
  const titleInput = form.querySelector('[name="measurement_image_title"]');
  for (const [index, selectedFile] of selectedFiles.entries()) {
    const formData = new FormData();
    formData.append(
      "title",
      buildMeasurementImageTitleForIndex(entryDate, titleInput?.value || "", index, selectedFiles.length)
    );
    formData.append("entry_date", entryDate);
    formData.append("memo", "");
    formData.append("profile_image", selectedFile);
    await multipartApi(`/api/admin/respondents/${encodeURIComponent(respondentId)}/profile-records`, formData);
  }
  return selectedFiles.length;
}

async function onMeasurementEditSubmit(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const respondentId = form.dataset.measurementRespondentId || state.activeRespondentId;
  if (!respondentId) {
    return;
  }
  const recordId = form.dataset.measurementEditForm;
  const submitButton = form.querySelector('button[type="submit"]');
  const errorNode = form.querySelector("[data-measurement-edit-error]");
  if (errorNode) {
    errorNode.textContent = "";
  }
  submitButton.disabled = true;
  submitButton.textContent = "保存中...";
  let imageUploadError = null;

  try {
    await api(`/api/admin/respondents/${encodeURIComponent(respondentId)}/measurements/${recordId}/update`, {
      method: "POST",
      body: {
        entryDate: form.querySelector('[name="entry_date"]').value,
        category: form.querySelector('[name="category"]').value,
        waist: form.querySelector('[name="waist"]').value,
        hip: form.querySelector('[name="hip"]').value,
        thigh: form.querySelector('[name="thigh"]').value,
      },
    });
    try {
      await maybeUploadMeasurementImageRecord(respondentId, form);
    } catch (uploadError) {
      imageUploadError = uploadError;
    }
    await bootstrap();
    if (form.closest("#panel-measurements")) {
      setActivePanel("measurements");
    } else {
      setActivePanel("respondents");
      await loadRespondentHistory(respondentId, false);
    }
    if (imageUploadError) {
      window.alert(`計測記録は保存しましたが、画像の追加に失敗しました。${imageUploadError.message}`);
    }
  } catch (error) {
    if (errorNode) {
      errorNode.textContent = error.message;
    }
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "保存";
  }
}

function bindMeasurementDeleteButtons(container) {
  container.querySelectorAll("[data-measurement-delete]").forEach((button) => {
    button.addEventListener("click", async () => {
      const respondentId = button.dataset.measurementRespondentId || state.activeRespondentId;
      if (!respondentId) {
        return;
      }
      const recordId = button.dataset.measurementDelete;
      const confirmed = window.confirm("この計測記録を削除します。");
      if (!confirmed) {
        return;
      }
      try {
        await api(`/api/admin/respondents/${encodeURIComponent(respondentId)}/measurements/${recordId}/delete`, {
          method: "POST",
          body: {},
        });
        await bootstrap();
        if (button.closest("#panel-measurements")) {
          setActivePanel("measurements");
        } else {
          setActivePanel("respondents");
          await loadRespondentHistory(respondentId, false);
        }
      } catch (error) {
        window.alert(error.message);
      }
    });
  });
}

function bindMeasurementOpenRespondentButtons(container) {
  container.querySelectorAll("[data-measurement-open-respondent]").forEach((button) => {
    button.addEventListener("click", async () => {
      const respondentId = button.dataset.measurementOpenRespondent;
      if (!respondentId) {
        return;
      }
      await loadRespondentHistory(respondentId, true);
    });
  });
}

async function editRespondent(respondentId) {
  state.activeRespondentId = respondentId;
  await loadRespondentHistory(respondentId, true);
  const target = els.respondentHistory.querySelector('[name="respondent_name"]');
  target?.focus();
  target?.select();
}

async function deleteRespondent(respondentId, currentName) {
  const label = currentName || respondentId;
  const confirmed = window.confirm(
    `${respondentScopeLabel()}の「${label}」の回答履歴を削除します。この操作は元に戻せません。`
  );
  if (!confirmed) {
    return;
  }
  try {
    await api(`/api/admin/respondents/${encodeURIComponent(respondentId)}/delete`, {
      method: "POST",
      body: {
        formId: state.selectedRespondentFormId || "",
      },
    });
    state.activeRespondentId = null;
    await bootstrap();
    setActivePanel("respondents");
    els.respondentHistory.innerHTML = "回答者を選択してください。";
  } catch (error) {
    window.alert(error.message);
  }
}

function respondentScopeLabel() {
  return state.selectedRespondentFormId ? "このフォーム内" : "全フォーム";
}

async function multipartApi(url, formData) {
  return runtime.requestMultipart(url, formData);
}

function publicFormUrl(slug) {
  const rawBase = String(state.publicBaseUrl || "").trim();
  if (rawBase) {
    const url = new URL(rawBase);
    if (slug) {
      url.searchParams.set("form", slug);
    } else {
      url.searchParams.delete("form");
    }
    return url.toString();
  }
  return runtime.respondentFormUrl(slug);
}

function publicHubUrl() {
  const rawBase = String(state.publicBaseUrl || "").trim();
  if (rawBase) {
    const url = new URL(rawBase);
    url.searchParams.delete("form");
    return url.toString();
  }
  return runtime.respondentFormUrl("");
}

function resolvedPublicBaseUrl() {
  const raw = String(state.publicBaseUrl || "").trim();
  if (raw) {
    return raw;
  }
  return runtime.respondentHomeUrl();
}

function qrCodeUrl(text) {
  return `https://quickchart.io/qr?size=220&margin=1&text=${encodeURIComponent(text)}`;
}

function qrCodeFallbackUrl(text) {
  return `https://api.qrserver.com/v1/create-qr-code/?size=220x220&data=${encodeURIComponent(text)}`;
}

function bindQrImages(container) {
  if (!container) {
    return;
  }
  container.querySelectorAll(".qr-image[data-qr-fallback-src]").forEach((img) => {
    img.addEventListener("error", () => {
      if (img.dataset.qrFallbackApplied === "1") {
        const message = img.closest(".qr-block")?.querySelector("[data-qr-error-message]");
        message?.classList.remove("hidden");
        img.classList.add("hidden");
        return;
      }
      const fallback = String(img.dataset.qrFallbackSrc || "").trim();
      if (!fallback || fallback === img.currentSrc || fallback === img.src) {
        return;
      }
      img.dataset.qrFallbackApplied = "1";
      img.src = fallback;
    });
  });
}

function renderFileGroups(files, small = false) {
  if (!files.length) {
    return `<div class="empty-state">画像はありません。</div>`;
  }
  const groups = new Map();
  files.forEach((file) => {
    const key = file.label || "画像";
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key).push(file);
  });
  return Array.from(groups.entries())
    .map(
      ([label, items]) => `
        <div class="stack-gap compact-gap">
          <h5>${escapeHtml(label)}</h5>
          <div class="image-grid admin-gallery">
            ${items
              .map(
                (file) => `
                  <button
                    class="image-tile ${small ? "small" : ""}"
                    type="button"
                    data-image-open="${escapeAttribute(adminImageProxyUrl(file.url || file.previewUrl || ""))}"
                    data-image-fallback="${escapeAttribute(file.url || file.previewUrl || "")}"
                    data-image-title="${escapeAttribute(file.originalName || label || "画像")}"
                    data-image-meta="${escapeAttribute(label || "画像")}"
                  >
                    <img src="${escapeAttribute(adminImageProxyUrl(file.previewUrl || file.url || ""))}" data-fallback-src="${escapeAttribute(file.previewUrl || file.url || "")}" alt="${escapeAttribute(file.originalName)}" loading="lazy" referrerpolicy="no-referrer" />
                    <span>${escapeHtml(file.originalName)}</span>
                  </button>
                `
              )
              .join("")}
          </div>
        </div>
      `
    )
    .join("");
}

function formatDate(value) {
  if (!value) {
    return "-";
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  return new Intl.DateTimeFormat("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function formatDateOnly(value) {
  if (!value) {
    return "-";
  }
  const date = new Date(`${value}T00:00:00`);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  return new Intl.DateTimeFormat("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(date);
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

function cssEscape(value) {
  if (window.CSS?.escape) {
    return window.CSS.escape(String(value ?? ""));
  }
  return String(value ?? "").replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}
