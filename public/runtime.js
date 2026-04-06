(() => {
  const GLOBAL_CONFIG = window.BIJIRIS_CONFIG || {};
  const GAS_URL_STORAGE_KEY = "BIJIRIS_GAS_URL";
  const ADMIN_TOKEN_STORAGE_KEY = "BIJIRIS_ADMIN_TOKEN";

  function normalizeApiMode(value) {
    const raw = String(value || "")
      .trim()
      .toLowerCase();
    return raw === "gas" ? "gas" : "rest";
  }

  function configuredValue(key) {
    return String(GLOBAL_CONFIG?.[key] || "").trim();
  }

  function configuredGasUrl() {
    const override = safeStorageGet(GAS_URL_STORAGE_KEY);
    if (override) {
      return override;
    }
    return configuredValue("gasUrl");
  }

  function apiMode() {
    const explicit = normalizeApiMode(configuredValue("apiMode") || configuredValue("mode"));
    if (explicit === "gas" && configuredGasUrl()) {
      return "gas";
    }
    return explicit;
  }

  function isGasMode() {
    return apiMode() === "gas" && !!configuredGasUrl();
  }

  function safeStorageGet(key) {
    try {
      return String(window.localStorage?.getItem(key) || "").trim();
    } catch (_error) {
      return "";
    }
  }

  function safeStorageSet(key, value) {
    try {
      if (!value) {
        window.localStorage?.removeItem(key);
        return;
      }
      window.localStorage?.setItem(key, String(value));
    } catch (_error) {
      // noop
    }
  }

  function siteRootUrl() {
    const configured = configuredValue("siteRootUrl");
    if (configured) {
      return new URL(configured, window.location.href).toString();
    }
    if (window.location.pathname.includes("/admin/")) {
      return new URL("../", window.location.href).toString();
    }
    return new URL("./", window.location.href).toString();
  }

  function respondentHomeUrl() {
    const configured = configuredValue("respondentUrl") || configuredValue("appUrl");
    if (configured) {
      return new URL(configured, window.location.href).toString();
    }
    return siteRootUrl();
  }

  function adminHomeUrl() {
    const configured = configuredValue("adminUrl");
    if (configured) {
      return new URL(configured, window.location.href).toString();
    }
    return new URL("admin/", siteRootUrl()).toString();
  }

  function respondentFormUrl(slug = "") {
    const url = new URL(respondentHomeUrl());
    const normalized = String(slug || "").trim();
    if (normalized) {
      url.searchParams.set("form", normalized);
    } else {
      url.searchParams.delete("form");
    }
    return url.toString();
  }

  function getAdminToken() {
    return safeStorageGet(ADMIN_TOKEN_STORAGE_KEY);
  }

  function setAdminToken(value) {
    safeStorageSet(ADMIN_TOKEN_STORAGE_KEY, value);
  }

  function clearAdminToken() {
    safeStorageSet(ADMIN_TOKEN_STORAGE_KEY, "");
  }

  function setGasUrl(value) {
    safeStorageSet(GAS_URL_STORAGE_KEY, value ? String(value).trim() : "");
  }

  function appendFormField(target, key, value) {
    if (Object.prototype.hasOwnProperty.call(target, key)) {
      if (Array.isArray(target[key])) {
        target[key].push(value);
      } else {
        target[key] = [target[key], value];
      }
      return;
    }
    target[key] = value;
  }

  function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error("ファイルを読み込めませんでした。"));
      reader.onload = () => {
        const raw = String(reader.result || "");
        const marker = raw.indexOf("base64,");
        resolve(marker >= 0 ? raw.slice(marker + 7) : raw);
      };
      reader.readAsDataURL(file);
    });
  }

  async function serializeFormData(formData) {
    const fields = {};
    const files = [];
    for (const [key, value] of formData.entries()) {
      if (value instanceof File) {
        if (!value.name) {
          continue;
        }
        files.push({
          fieldName: key,
          name: value.name,
          type: value.type || "application/octet-stream",
          size: Number(value.size || 0),
          base64: await fileToBase64(value),
        });
        continue;
      }
      appendFormField(fields, key, String(value ?? ""));
    }
    return { fields, files };
  }

  async function gasRequest(path, options = {}) {
    const gasUrl = configuredGasUrl();
    if (!gasUrl) {
      throw new Error("Google Apps Script URL が設定されていません。");
    }

    const payload = {
      action: "api",
      method: String(options.method || "GET").toUpperCase(),
      path: path,
      body: options.body === undefined ? null : options.body,
      authToken: getAdminToken(),
      requestedAt: Date.now(),
    };

    if (options.formData instanceof FormData) {
      payload.formData = await serializeFormData(options.formData);
    }

    const response = await fetch(gasUrl, {
      method: "POST",
      headers: {
        "Content-Type": "text/plain;charset=utf-8",
      },
      body: JSON.stringify(payload),
    });
    const data = await response.json().catch(() => ({}));
    const statusCode = Number(data.statusCode || (response.ok ? 200 : response.status) || 200);
    const result = data.data !== undefined ? data.data : data;
    const errorMessage = data.error || result?.error || "通信に失敗しました。";

    if (path === "/api/admin/login" && result?.authToken) {
      setAdminToken(result.authToken);
    }
    if (path === "/api/admin/logout") {
      clearAdminToken();
    }

    if (statusCode === 401) {
      if (path !== "/api/admin/login") {
        clearAdminToken();
      }
      throw new Error(errorMessage);
    }
    if (statusCode < 200 || statusCode >= 300 || data.error) {
      throw new Error(errorMessage);
    }
    return result;
  }

  async function request(path, options = {}) {
    if (isGasMode()) {
      return gasRequest(path, options);
    }
    const settings = {
      method: String(options.method || "GET").toUpperCase(),
      credentials: "same-origin",
      headers: {},
    };
    if (options.body !== undefined) {
      settings.headers["Content-Type"] = "application/json";
      settings.body = JSON.stringify(options.body);
    }
    const response = await fetch(path, settings);
    const data = await response.json().catch(() => ({}));
    if (response.status === 401) {
      throw new Error(data.error || "認証が必要です。");
    }
    if (!response.ok) {
      throw new Error(data.error || "通信に失敗しました。");
    }
    return data;
  }

  async function requestMultipart(path, formData) {
    if (isGasMode()) {
      return gasRequest(path, { method: "POST", formData });
    }
    const response = await fetch(path, {
      method: "POST",
      credentials: "same-origin",
      body: formData,
    });
    const data = await response.json().catch(() => ({}));
    if (response.status === 401) {
      throw new Error(data.error || "認証が必要です。");
    }
    if (!response.ok) {
      throw new Error(data.error || "通信に失敗しました。");
    }
    return data;
  }

  window.BijirisRuntime = {
    apiMode,
    isGasMode,
    getGasUrl: configuredGasUrl,
    setGasUrl,
    respondentHomeUrl,
    respondentFormUrl,
    adminHomeUrl,
    siteRootUrl,
    getAdminToken,
    setAdminToken,
    clearAdminToken,
    request,
    requestMultipart,
  };
})();
