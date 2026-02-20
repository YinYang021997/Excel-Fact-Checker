/* global document, Office, OfficeRuntime */

Office.onReady(() => {
  // ── Element refs ──────────────────────────────────────────────────────────
  const apiKeyInput   = document.getElementById("apiKey")       as HTMLInputElement;
  const saveBtn       = document.getElementById("saveBtn")      as HTMLButtonElement;
  const showBtn       = document.getElementById("showBtn")      as HTMLButtonElement;
  const clearKeyBtn   = document.getElementById("clearKeyBtn")  as HTMLButtonElement;
  const clearCacheBtn = document.getElementById("clearCacheBtn")as HTMLButtonElement;
  const countCacheBtn = document.getElementById("countCacheBtn")as HTMLButtonElement;
  const keyStatus     = document.getElementById("keyStatus")    as HTMLDivElement;
  const cacheMsg      = document.getElementById("cacheMsg")     as HTMLDivElement;
  const keyBadge      = document.getElementById("keyBadge")     as HTMLSpanElement;

  // ── On load: check whether a key is already saved ────────────────────────
  OfficeRuntime.storage.getItem("anthropic_api_key").then((key) => {
    if (key) {
      setBadgeSaved(key);
    }
  }).catch(() => { /* storage not available yet — no-op */ });

  // ── Save key ──────────────────────────────────────────────────────────────
  saveBtn.onclick = async () => {
    const key = apiKeyInput.value.trim();
    if (!key) {
      flash(keyStatus, "Please enter an API key.", "danger");
      return;
    }
    if (!key.startsWith("sk-ant-")) {
      flash(keyStatus, "Warning: key doesn't look like sk-ant-… — saved anyway.", "warn");
    }
    await OfficeRuntime.storage.setItem("anthropic_api_key", key);
    apiKeyInput.value = "";
    apiKeyInput.type  = "password";
    showBtn.textContent = "👁";
    setBadgeSaved(key);
    if (key.startsWith("sk-ant-")) {
      flash(keyStatus, "API key saved successfully.", "ok");
    }
  };

  // ── Toggle show/hide ──────────────────────────────────────────────────────
  showBtn.onclick = () => {
    if (apiKeyInput.type === "password") {
      apiKeyInput.type = "text";
      showBtn.textContent = "🙈";
    } else {
      apiKeyInput.type = "password";
      showBtn.textContent = "👁";
    }
  };

  // ── Clear key ─────────────────────────────────────────────────────────────
  clearKeyBtn.onclick = async () => {
    await OfficeRuntime.storage.removeItem("anthropic_api_key");
    apiKeyInput.value   = "";
    apiKeyInput.type    = "password";
    showBtn.textContent = "👁";
    keyBadge.textContent = "No Key";
    keyBadge.className   = "badge badge-none";
    flash(keyStatus, "API key cleared.", "warn");
  };

  // ── Count cache entries ───────────────────────────────────────────────────
  countCacheBtn.onclick = async () => {
    const keys = await OfficeRuntime.storage.getKeys();
    const n = keys.filter((k) => k.startsWith("fc_")).length;
    cacheMsg.textContent = `${n} cached fact-check result${n !== 1 ? "s" : ""} in storage.`;
  };

  // ── Clear cache ───────────────────────────────────────────────────────────
  clearCacheBtn.onclick = async () => {
    const keys = await OfficeRuntime.storage.getKeys();
    const cacheKeys = keys.filter((k) => k.startsWith("fc_"));
    await Promise.all(cacheKeys.map((k) => OfficeRuntime.storage.removeItem(k)));
    cacheMsg.textContent = `Cleared ${cacheKeys.length} cached result${cacheKeys.length !== 1 ? "s" : ""}.`;
  };

  // ── Helpers ───────────────────────────────────────────────────────────────

  function setBadgeSaved(key: string) {
    keyBadge.textContent = `Saved (…${key.slice(-4)})`;
    keyBadge.className   = "badge badge-saved";
  }

  function flash(el: HTMLElement, msg: string, type: "ok" | "warn" | "danger") {
    const colors = { ok: "#107c10", warn: "#ca5010", danger: "#a4262c" };
    el.textContent  = msg;
    el.style.color  = colors[type];
    setTimeout(() => { el.textContent = ""; }, 4500);
  }
});
