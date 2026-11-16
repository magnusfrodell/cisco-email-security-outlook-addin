/**
 * @license
 * Copyright (c) 2026 Cisco and/or its affiliates.
 *
 * This software is licensed to you under the terms of the Cisco Sample
 * Code License, Version 1.1 (the "License"). You may obtain a copy of the
 * License at
 *
 *                https://developer.cisco.com/docs/licenses
 *
 * All use of the material herein must be in accordance with the terms of
 * the License. All rights not expressly granted by the License are
 * reserved. Unless required by applicable law or agreed to separately in
 * writing, software distributed under the License is distributed on an "AS
 * IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
 * or implied.
 */
/* global Office, msal */
/*
 * Cisco Email Security Reporter – task pane logic.
 *
 * Two send modes (SEND_MODE in config.js, default "graph"):
 *   graph   : The mail is fetched as EML (getAsFileAsync) and sent through Microsoft Graph
 *             /me/sendMail with Nested App Authentication (NAA). Needs CLIENT_ID in
 *             config.js, Mailbox 1.14 and NestedAppAuth 1.1 in the Outlook client.
 *   compose : A new message to Cisco's address is opened with the selected mail attached
 *             as an item ("forward as attachment"). The user presses Send. Works without
 *             an app registration (Mailbox 1.6+).
 *
 * With SEND_MODE "graph" and COMPOSE_FALLBACK true, compose is used only as a reserve:
 * when graph's preconditions are missing (a warning is shown in the pane), when a mail is
 * too large for Graph, or on demand ("send manually instead") after a failed send.
 * With COMPOSE_FALLBACK false the pane shows an error instead of opening a new message.
 *
 * Shared mailboxes and shared folders: the manifest declares <SupportsSharedFolders>, so the
 * pane also opens on items the user reads in a mailbox that is not their own. Such an item is
 * recognised with item.getSharedPropertiesAsync(); it does not exist under /me, so reading
 * its MIME content and moving it go through /users/{mailbox} with a *.Shared Graph scope.
 * The report itself is still sent from the user's own mailbox (/me/sendMail).
 *
 * Languages: every user-facing text comes from window.REPORTER_LOCALES[<lang>]
 * (src/locales/*.js). The active language is resolved in resolveLanguage(). This file
 * contains no language-dependent text; log messages are English by design.
 */
(function () {
  "use strict";

  var VERSION = "1.4.0";
  var CFG = window.REPORTER_CONFIG || {};
  var LOCALES = window.REPORTER_LOCALES || {};
  var LOG_PREFIX = "[CiscoReporter]";
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var AUTO = "auto";
  var SHARED_LOOKUP_TIMEOUT_MS = 4000;

  var L = {};   // active locale object
  var S = {};   // active locale strings

  var state = {
    mode: "compose",      // "graph" | "compose" | "none"
    modeDetected: false,
    notice: null,         // { kind, key, reasonKey } shown persistently under the header
    fallbackOk: false,    // compose may be used as a fallback for graph
    language: null,
    msalApp: null,
    busy: false,
    item: null,
    shared: null,         // { owner, targetMailbox, permissions } when the item is in a shared mailbox or folder
    sharedReady: null,    // promise that settles when the shared-properties lookup for the item is done
    sharedSeq: 0          // discards a late lookup answer for a previously selected item
  };

  /* ------------------------------------------------------------------ */
  /* Logging                                                            */
  /* ------------------------------------------------------------------ */

  function log(level, msg, extra) {
    var line = new Date().toISOString().substr(11, 8) + " " + level.toUpperCase() + " " + msg;
    if (extra !== undefined && extra !== null) {
      try {
        line += " " + (typeof extra === "string" ? extra : JSON.stringify(extra));
      } catch (e) {
        line += " " + String(extra);
      }
    }
    var fn = console[level] || console.log;
    try { fn.call(console, LOG_PREFIX, line); } catch (e) { /* ignore */ }
    var el = document.getElementById("log");
    if (el) {
      el.textContent += line + "\n";
      el.scrollTop = el.scrollHeight;
    }
  }

  function errText(e) {
    if (!e) { return "unknown error"; }
    if (typeof e === "string") { return e; }
    var parts = [];
    if (e.name) { parts.push(e.name); }
    if (e.errorCode) { parts.push(e.errorCode); }
    if (e.message) { parts.push(e.message); }
    return parts.length ? parts.join(": ") : String(e);
  }

  /* ------------------------------------------------------------------ */
  /* Languages                                                          */
  /* ------------------------------------------------------------------ */

  function languages() { return Object.keys(LOCALES); }

  /** Map "sv-SE", "SV", "da-dk" … to a loaded locale key, or null. */
  function matchLanguage(tag) {
    var t = String(tag || "").toLowerCase();
    if (!t) { return null; }
    var keys = languages(), i;
    for (i = 0; i < keys.length; i++) { if (keys[i].toLowerCase() === t) { return keys[i]; } }
    var base = t.split(/[-_]/)[0];
    for (i = 0; i < keys.length; i++) { if (keys[i].toLowerCase() === base) { return keys[i]; } }
    return null;
  }

  function savedLanguage() {
    try {
      var v = Office.context.roamingSettings.get("language");
      if (typeof v === "string" && v) { return v; }
    } catch (e) { /* ignore */ }
    return null;
  }

  function displayLanguage() {
    try { return Office.context.displayLanguage || null; } catch (e) { return null; }
  }

  /**
   * Precedence: the user's own choice (if the selector is enabled) → LANGUAGE in config
   * (unless "auto") → the Outlook display language → DEFAULT_LANGUAGE → first locale loaded.
   */
  function resolveLanguage() {
    var keys = languages();
    if (!keys.length) { return null; }
    var chosen = null, source = "";

    var saved = CFG.SHOW_LANGUAGE_SELECTOR !== false ? savedLanguage() : null;
    if (saved && saved !== AUTO && matchLanguage(saved)) {
      chosen = matchLanguage(saved); source = "user setting";
    } else if (CFG.LANGUAGE && CFG.LANGUAGE !== AUTO && matchLanguage(CFG.LANGUAGE)) {
      chosen = matchLanguage(CFG.LANGUAGE); source = "config LANGUAGE";
    } else {
      var disp = displayLanguage();
      if (disp && matchLanguage(disp)) {
        chosen = matchLanguage(disp); source = "Outlook display language " + disp;
      }
    }
    if (!chosen) {
      chosen = matchLanguage(CFG.DEFAULT_LANGUAGE) || keys[0];
      source = "default";
    }
    log("info", "Language: " + chosen, { source: source, available: keys });
    return chosen;
  }

  /** Text from the active locale with {name} placeholders filled in. Unknown key → the key. */
  function t(key, params) {
    var text = Object.prototype.hasOwnProperty.call(S, key) ? String(S[key]) : key;
    if (params) {
      Object.keys(params).forEach(function (k) {
        text = text.split("{" + k + "}").join(String(params[k]));
      });
    }
    return text;
  }

  function labelOf(cat) {
    var c = (L.categories || {})[cat.id];
    return (c && c.label) || cat.id;
  }

  function hintOf(cat) {
    var c = (L.categories || {})[cat.id];
    return (c && c.hint) || "";
  }

  function applyStrings() {
    var nodes = document.querySelectorAll("[data-str]");
    for (var i = 0; i < nodes.length; i++) {
      var key = nodes[i].getAttribute("data-str");
      if (Object.prototype.hasOwnProperty.call(S, key)) {
        nodes[i].textContent = S[key];
      }
    }
    if (S.title) { document.title = S.title; }
  }

  /** Switch the whole pane to a locale: strings, group titles, buttons, chip, selector. */
  function applyLocale(code) {
    L = LOCALES[code] || {};
    S = L.strings || {};
    state.language = code;
    if (code) { document.documentElement.setAttribute("lang", code); }
    applyStrings();
    renderCategories();
    updateModeChip();
    renderNotice();
    renderLanguageSelector();
    clearStatus();
  }

  function renderLanguageSelector() {
    var row = document.getElementById("language-row");
    var sel = document.getElementById("lang-select");
    if (!row || !sel) { return; }
    var enabled = CFG.SHOW_LANGUAGE_SELECTOR !== false && languages().length > 1;
    row.hidden = !enabled;
    if (!enabled) { return; }

    var saved = savedLanguage();
    var current = (saved && saved !== AUTO && matchLanguage(saved)) ? matchLanguage(saved) : AUTO;
    sel.innerHTML = "";
    var optAuto = document.createElement("option");
    optAuto.value = AUTO;
    optAuto.textContent = t("languageAuto");
    sel.appendChild(optAuto);
    languages().forEach(function (code) {
      var o = document.createElement("option");
      o.value = code;
      o.textContent = (LOCALES[code] && LOCALES[code].name) || code;
      sel.appendChild(o);
    });
    sel.value = current;
  }

  function initLanguageSelector() {
    var sel = document.getElementById("lang-select");
    if (!sel) { return; }
    sel.addEventListener("change", function () {
      var value = sel.value;
      try {
        Office.context.roamingSettings.set("language", value);
        Office.context.roamingSettings.saveAsync(function (r) {
          if (r.status !== Office.AsyncResultStatus.Succeeded) {
            log("warn", "Could not save language setting", r.error && r.error.message);
          }
        });
      } catch (e) {
        log("warn", "roamingSettings unavailable", errText(e));
      }
      log("info", "Language changed by user", { value: value });
      applyLocale(resolveLanguage());
      refreshItem();
    });
  }

  /* ------------------------------------------------------------------ */
  /* Helpers                                                            */
  /* ------------------------------------------------------------------ */

  function $(id) { return document.getElementById(id); }

  function isSet(name, version) {
    try {
      return Office.context.requirements.isSetSupported(name, version);
    } catch (e) {
      return false;
    }
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }

  function reporterAddress() {
    try { return (Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress) || ""; }
    catch (e) { return ""; }
  }

  function senderOf(item) {
    var from = (item && item.from) || {};
    if (!from.emailAddress) { return ""; }
    return from.displayName ? from.displayName + " <" + from.emailAddress + ">" : from.emailAddress;
  }

  function fill(template, params) {
    var text = String(template);
    Object.keys(params).forEach(function (k) {
      text = text.split("{" + k + "}").join(String(params[k] === undefined || params[k] === null ? "" : params[k]));
    });
    return text;
  }

  function subjectFor(cat) {
    var item = state.item || {};
    var from = item.from || {};
    var subject = String(item.subject || t("noSubject")).replace(/\s+/g, " ").trim();
    if (subject.length > 120) { subject = subject.slice(0, 117) + "…"; }
    return fill(CFG.SUBJECT || "Cisco Secure Email submission – {category}", {
      category: labelOf(cat),
      subject: subject,
      sender: from.emailAddress || "",
      reporter: reporterAddress()
    }).trim();
  }

  function bodyText() {
    return String(L.bodyText || "");
  }

  /** Administrator-configured SOC recipients (SOC_ADDRESSES; CC_ADDRESSES kept as an alias). */
  function socAddresses() {
    var seen = {}, out = [];
    [].concat(CFG.SOC_ADDRESSES || [], CFG.CC_ADDRESSES || []).forEach(function (a) {
      var v = String(a || "").trim().toLowerCase();
      if (v && !seen[v]) { seen[v] = true; out.push(v); }
    });
    return out;
  }

  function socMode() {
    return String(CFG.SOC_COPY_MODE || "cc").toLowerCase() === "bcc" ? "bcc" : "cc";
  }

  function sendAs() {
    return String(CFG.SEND_AS || "").trim();
  }

  /** Machine-readable details appended to the report body (for the SOC copy). */
  function reportDetails(cat) {
    var item = state.item || {};
    var lines = [
      ["Category", labelOf(cat) + " (" + cat.id + ")"],
      ["Submitted to", cat.address],
      ["Reported by", reporterAddress()],
      ["Sent as", (state.mode === "graph" && sendAs()) ? sendAs() : (reporterAddress() || "-")]
    ];
    if (state.shared) {
      // Only present for items reported from a mailbox that is not the reporting user's own.
      lines.push(["Shared mailbox", sharedLabel(state.shared)]);
    }
    lines = lines.concat([
      ["Original subject", item.subject || ""],
      ["Original sender", senderOf(item)],
      ["Original date", item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : ""],
      ["Message-ID", item.internetMessageId || ""],
      ["Client", clientInfo()],
      ["Add-in", "Cisco Email Security Reporter " + VERSION + " (" + state.mode + ")"]
    ]);
    return lines.map(function (l) { return l[0] + ": " + l[1]; }).join("\n");
  }

  function clientInfo() {
    try {
      var d = Office.context.diagnostics;
      return d.host + " " + d.platform + " " + d.version;
    } catch (e) { return ""; }
  }

  function bodyPlain(cat) {
    var text = bodyText();
    if (CFG.REPORT_DETAILS !== false) {
      text += "\n\n--- Report details ---\n" + reportDetails(cat);
    }
    return text;
  }

  function bodyHtml(cat) {
    var html = "<p>" + escapeHtml(bodyText()) + "</p>";
    if (CFG.REPORT_DETAILS !== false) {
      html += "<pre>--- Report details ---\n" + escapeHtml(reportDetails(cat)) + "</pre>";
    }
    return html;
  }

  function socStatusSuffix() {
    var list = socAddresses();
    return (list.length && CFG.SOC_SHOW_IN_STATUS) ? t("copySentTo", { address: list.join(", ") }) : "";
  }

  function folderName(id) {
    if (id === "junkemail") { return t("folderJunk"); }
    if (id === "inbox") { return t("folderInbox"); }
    return String(id);
  }

  function randomString() {
    return Math.random().toString(36).slice(2, 10);
  }

  function enabledCategories() {
    return (CFG.CATEGORIES || []).filter(function (c) { return c.enabled !== false; });
  }

  /* ------------------------------------------------------------------ */
  /* UI                                                                 */
  /* ------------------------------------------------------------------ */

  function setModeChip(text, cls) {
    var chip = $("mode-chip");
    chip.textContent = text;
    chip.className = "chip " + cls;
  }

  function updateModeChip() {
    if (!state.modeDetected) {
      setModeChip(t("chipStarting"), "chip-neutral");
    } else if (state.mode === "graph") {
      setModeChip(t("chipAuto"), "chip-ok");
    } else if (state.mode === "compose") {
      setModeChip(t("chipCompose"), "chip-neutral");
    } else {
      setModeChip(t("chipUnavailable"), "chip-err");
    }
  }

  /** Persistent notice under the header (e.g. "automatic send is not available"). */
  function renderNotice() {
    var box = $("notice");
    if (!box) { return; }
    var n = state.notice;
    if (!n) { box.hidden = true; return; }
    box.hidden = false;
    box.setAttribute("data-kind", n.kind);
    var chip = $("notice-chip");
    chip.textContent = t(n.kind === "err" ? "chipErr" : "chipWarn");
    chip.className = "chip chip-" + n.kind;
    $("notice-text").textContent = t(n.key, { reason: n.reasonKey ? t(n.reasonKey) : "" });
  }

  function setStatus(kind, text, opts) {
    var box = $("status");
    var chip = $("status-chip");
    var actions = $("status-actions");
    var chipKeys = { ok: "chipOk", warn: "chipWarn", err: "chipErr", wait: "chipWait", neutral: "chipInfo" };
    var chipText = (opts && opts.chip) || t(chipKeys[kind] || "chipInfo");

    box.hidden = false;
    box.setAttribute("data-kind", kind);
    chip.textContent = chipText;
    chip.className = "chip chip-" + kind;
    $("status-text").textContent = text;
    actions.innerHTML = "";

    if (opts && opts.retryCompose) {
      var b = document.createElement("button");
      b.type = "button";
      b.className = "btn-secondary";
      b.textContent = t("retryCompose");
      b.addEventListener("click", function () {
        runGuarded(function () { return reportViaCompose(opts.retryCompose, t("manualChosen")); });
      });
      actions.appendChild(b);
    }

    try {
      var reduce = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
      box.scrollIntoView({ behavior: reduce ? "auto" : "smooth", block: "start" });
    } catch (e) { /* ignore */ }
  }

  function clearStatus() {
    $("status").hidden = true;
    $("status-actions").innerHTML = "";
  }

  function setButtonsEnabled(on) {
    var enabled = on && state.mode !== "none";
    var btns = document.querySelectorAll(".btn-cat");
    for (var i = 0; i < btns.length; i++) { btns[i].disabled = !enabled; }
  }

  function renderCategories() {
    var groups = L.groups || {};
    ["missed", "false_positive"].forEach(function (g) {
      var title = $("group-" + g + "-title");
      if (title) { title.textContent = groups[g] || g; }
      var container = $("group-" + g);
      if (container) { container.innerHTML = ""; }
    });

    enabledCategories().forEach(function (cat) {
      var container = $("group-" + cat.group) || $("group-missed");
      var b = document.createElement("button");
      b.type = "button";
      b.className = "btn-cat";
      b.setAttribute("data-id", cat.id);
      b.setAttribute("title", t("sentTo", { address: cat.address }));
      b.disabled = state.busy || !state.item || state.mode === "none";
      var l = document.createElement("span");
      l.className = "btn-cat-label";
      l.textContent = labelOf(cat);
      var h = document.createElement("span");
      h.className = "btn-cat-hint";
      h.textContent = hintOf(cat);
      b.appendChild(l);
      b.appendChild(h);
      b.addEventListener("click", function () { report(cat); });
      container.appendChild(b);
    });
  }

  function refreshItem() {
    var item = null;
    try { item = Office.context.mailbox.item; } catch (e) { item = null; }
    state.item = item;

    var isMessage = !!item && item.itemType === Office.MailboxEnums.ItemType.Message;
    if (!isMessage) {
      state.item = null;
      loadSharedProperties(null);
      $("mail-subject").textContent = t("noMailSelected");
      $("mail-from").textContent = t("selectMailHint");
      setButtonsEnabled(false);
      log("info", "No mail selected");
      return;
    }

    $("mail-subject").textContent = item.subject || t("noSubject");
    var from = item.from || {};
    $("mail-from").textContent = from.emailAddress
      ? t("fromPrefix") + (from.displayName ? from.displayName + " <" + from.emailAddress + ">" : from.emailAddress)
      : "";
    setButtonsEnabled(!state.busy);
    clearStatus();
    log("info", "Mail selected", { subject: item.subject, from: from.emailAddress, itemId: (item.itemId || "").slice(0, 24) + "…" });
    loadSharedProperties(item);
  }

  /* ------------------------------------------------------------------ */
  /* Shared mailboxes and shared folders                                */
  /* ------------------------------------------------------------------ */

  /** The address shown and logged for a shared item: the owner, or the client's location value. */
  function sharedLabel(shared) {
    return shared.owner || shared.targetMailbox || "";
  }

  /**
   * Graph user segments to try for a shared item. The owner's address comes first; the
   * client-specific targetMailbox value is the second attempt when it differs.
   */
  function sharedMailboxIds(shared) {
    var seen = {}, out = [];
    [shared.owner, shared.targetMailbox].forEach(function (v) {
      var key = String(v || "").toLowerCase();
      if (key && !seen[key]) { seen[key] = true; out.push(String(v)); }
    });
    return out;
  }

  function renderShared() {
    var el = $("mail-shared");
    if (!el) { return; }
    el.hidden = !state.shared;
    el.textContent = state.shared ? t("sharedMailbox", { address: sharedLabel(state.shared) }) : "";
  }

  /**
   * Find out whether the selected item lives in a shared mailbox or in a folder shared with
   * the user. Office.js answers "failed" for an item in the user's own mailbox, so anything
   * but a successful answer with an owner means "not shared". Sets state.shared and
   * state.sharedReady; the promise always resolves and never rejects.
   */
  function loadSharedProperties(item) {
    var seq = ++state.sharedSeq;
    state.shared = null;
    renderShared();

    if (!item || typeof item.getSharedPropertiesAsync !== "function") {
      state.sharedReady = Promise.resolve(null);
      return state.sharedReady;
    }

    state.sharedReady = new Promise(function (resolve) {
      var settled = false;
      var timer = null;
      var finish = function (value) {
        if (settled) { return; }
        settled = true;
        if (timer) { clearTimeout(timer); }
        if (seq !== state.sharedSeq) { resolve(null); return; }   // another item was selected meanwhile
        state.shared = value;
        renderShared();
        resolve(value);
      };
      timer = setTimeout(function () {
        log("warn", "getSharedPropertiesAsync did not answer – treating the item as not shared");
        finish(null);
      }, SHARED_LOOKUP_TIMEOUT_MS);

      try {
        item.getSharedPropertiesAsync(function (r) {
          if (!r || r.status !== Office.AsyncResultStatus.Succeeded || !r.value) { finish(null); return; }
          var owner = String(r.value.owner || "").trim();
          var target = String(r.value.targetMailbox || "").trim();
          if (!owner && !target) { finish(null); return; }
          var shared = { owner: owner, targetMailbox: target, permissions: r.value.delegatePermissions };
          if (seq === state.sharedSeq) {
            log("info", "Item is in a shared mailbox or folder", {
              owner: owner, targetMailbox: target, delegatePermissions: shared.permissions, user: reporterAddress()
            });
          }
          finish(shared);
        });
      } catch (e) {
        log("warn", "getSharedPropertiesAsync failed", errText(e));
        finish(null);
      }
    });
    return state.sharedReady;
  }

  /** Wait until the shared-properties lookup for the selected item has finished. */
  function sharedLookupDone() {
    return state.sharedReady || Promise.resolve(null);
  }

  /* ------------------------------------------------------------------ */
  /* Settings (roamingSettings)                                         */
  /* ------------------------------------------------------------------ */

  function keepCopy() {
    try {
      var v = Office.context.roamingSettings.get("keepCopy");
      if (v === true || v === false) { return v; }
    } catch (e) { /* ignore */ }
    return !!CFG.SAVE_TO_SENT_DEFAULT;
  }

  function initSettings() {
    var cb = $("keep-copy");
    cb.checked = keepCopy();
    cb.addEventListener("change", function () {
      try {
        Office.context.roamingSettings.set("keepCopy", cb.checked);
        Office.context.roamingSettings.saveAsync(function (r) {
          if (r.status !== Office.AsyncResultStatus.Succeeded) {
            log("warn", "Could not save setting", r.error && r.error.message);
          } else {
            log("info", "Setting saved", { keepCopy: cb.checked });
          }
        });
      } catch (e) {
        log("warn", "roamingSettings unavailable", errText(e));
      }
    });
    initLanguageSelector();
  }

  /** The settings panel is shown if any of its rows is; the keep-copy row only in graph mode. */
  function updateSettingsVisibility() {
    var keepRow = $("keep-copy-row");
    var langRow = $("language-row");
    keepRow.hidden = (state.mode !== "graph");
    $("settings").hidden = keepRow.hidden && langRow.hidden;
  }

  /* ------------------------------------------------------------------ */
  /* Send: compose (new message with the item attached)                 */
  /* ------------------------------------------------------------------ */

  function reportViaCompose(cat, reason) {
    var item = state.item;
    if (!item || !item.itemId) {
      throw new Error(t("selectFirst"));
    }
    var soc = socAddresses();
    var params = {
      toRecipients: [cat.address],
      ccRecipients: socMode() === "cc" ? soc : [],
      bccRecipients: socMode() === "bcc" ? soc : [],
      subject: subjectFor(cat),
      htmlBody: bodyHtml(cat),
      attachments: [{
        type: "item",
        name: (item.subject || t("reportedMailName")).slice(0, 120),
        itemId: item.itemId
      }]
    };
    if (sendAs()) { log("warn", "SEND_AS is ignored in compose mode – the report is sent from the user's mailbox"); }
    if (state.shared) {
      log("warn", "Compose mode with an item from a shared mailbox – whether Outlook can attach it by item id depends on the client", sharedLabel(state.shared));
    }
    log("info", "Opening new message (compose mode)", { to: cat.address, soc: soc, socMode: socMode(), reason: reason || null });

    var done = function () {
      setStatus("ok",
        t("composeOpened", { address: cat.address }) + (reason ? " (" + reason + ")" : "") + socStatusSuffix(),
        { chip: t("chipPressSend") });
    };

    if (isSet("Mailbox", "1.9") && typeof Office.context.mailbox.displayNewMessageFormAsync === "function") {
      return new Promise(function (resolve, reject) {
        Office.context.mailbox.displayNewMessageFormAsync(params, function (r) {
          if (r.status === Office.AsyncResultStatus.Succeeded) {
            done();
            resolve();
          } else {
            reject(new Error("displayNewMessageFormAsync: " + (r.error && r.error.message)));
          }
        });
      });
    }
    Office.context.mailbox.displayNewMessageForm(params);
    done();
    return Promise.resolve();
  }

  /* ------------------------------------------------------------------ */
  /* Send: Microsoft Graph via NAA                                      */
  /* ------------------------------------------------------------------ */

  function getEmlBase64() {
    return new Promise(function (resolve, reject) {
      Office.context.mailbox.item.getAsFileAsync(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
          resolve(r.value);
        } else {
          reject(new Error("getAsFileAsync: " + (r.error && r.error.message)));
        }
      });
    });
  }

  /**
   * Acquire a Graph token for the scopes. silentOnly = true never opens a sign-in or consent
   * window: if interaction would be needed, an error with consentRequired = true is thrown.
   */
  async function getToken(scopes, silentOnly) {
    if (!state.msalApp) {
      log("info", "Initialising MSAL (NAA)", { authority: CFG.AUTHORITY });
      state.msalApp = await msal.createNestablePublicClientApplication({
        auth: { clientId: CFG.CLIENT_ID, authority: CFG.AUTHORITY },
        cache: { cacheLocation: "localStorage" }
      });
    }
    var request = { scopes: scopes };
    try {
      var silent = await state.msalApp.acquireTokenSilent(request);
      log("info", "Token acquired (silent)", { scopes: scopes });
      return silent.accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) {
        if (silentOnly) {
          log("warn", "Silent token failed and no sign-in window is opened for this step", { scopes: scopes, error: errText(e) });
          var denied = new Error("Consent or sign-in required for " + scopes.join(", ") + " (" + errText(e) + ")");
          denied.consentRequired = true;
          denied.scopes = scopes;
          throw denied;
        }
        log("warn", "Silent token failed – trying interactive", errText(e));
        var popup = await state.msalApp.acquireTokenPopup(request);
        log("info", "Token acquired (interactive)", { scopes: scopes });
        return popup.accessToken;
      }
      throw e;
    }
  }

  async function graphFailure(path, res) {
    var text = "";
    try { text = await res.text(); } catch (e) { /* ignore */ }
    var err = new Error("Graph " + path + " returned " + res.status + " " + text.slice(0, 300));
    err.status = res.status;
    return err;
  }

  async function graphPost(path, token, body) {
    var res = await fetch(GRAPH + path, {
      method: "POST",
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    });
    if (!res.ok) { throw await graphFailure(path, res); }
    return res;
  }

  function bytesToBase64(bytes) {
    var binary = "", chunk = 0x8000;
    for (var i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
  }

  /** GET a Graph resource as raw bytes, Base64-encoded – used for the MIME content of a message. */
  async function graphGetBase64(path, token) {
    var res = await fetch(GRAPH + path, {
      method: "GET",
      headers: { Authorization: "Bearer " + token }
    });
    if (!res.ok) { throw await graphFailure(path, res); }
    return bytesToBase64(new Uint8Array(await res.arrayBuffer()));
  }

  function restItemId() {
    return Office.context.mailbox.convertToRestId(
      state.item.itemId, Office.MailboxEnums.RestVersion.v2_0);
  }

  /**
   * Run a Graph call against the mailbox that holds the selected item: /me for the user's own
   * mailbox, /users/{mailbox} for a shared one. For a shared item the owner's address is
   * tried first and the client's targetMailbox value second. A 401/403 is a permission
   * problem that another identifier cannot fix, so it ends the attempts.
   */
  async function withItemMailbox(shared, call) {
    if (!shared) { return call("/me"); }
    var ids = sharedMailboxIds(shared);
    var lastError = null;
    for (var i = 0; i < ids.length; i++) {
      try {
        return await call("/users/" + encodeURIComponent(ids[i]));
      } catch (e) {
        lastError = e;
        if (e && (e.status === 401 || e.status === 403)) { break; }
        if (i < ids.length - 1) {
          log("warn", "Graph call failed for mailbox " + ids[i] + " – trying " + ids[i + 1], errText(e));
        }
      }
    }
    throw lastError || new Error("No mailbox identifier for the shared item");
  }

  async function moveItem(token, folder, shared) {
    var id = encodeURIComponent(restItemId());
    await withItemMailbox(shared, function (base) {
      return graphPost(base + "/messages/" + id + "/move", token, { destinationId: folder });
    });
  }

  /**
   * The selected message as Base64 EML. Office.js (getAsFileAsync) comes first because it
   * needs no Graph permission. If it fails, the MIME content is read through Graph – from
   * /me for the user's own mailbox, from /users/{mailbox} for a shared one.
   * mailboxToken() supplies the token for that read (acquired silently, on first use).
   */
  async function fetchEml(shared, mailboxToken) {
    var officeError;
    try {
      return await getEmlBase64();
    } catch (e) {
      officeError = e;
      log("warn", "getAsFileAsync failed – reading the MIME content through Graph instead", errText(e));
    }
    try {
      var token = await mailboxToken();
      var id = encodeURIComponent(restItemId());
      var eml = await withItemMailbox(shared, function (base) {
        return graphGetBase64(base + "/messages/" + id + "/$value", token);
      });
      log("info", "MIME content read through Graph", { mailbox: shared ? sharedLabel(shared) : "me" });
      return eml;
    } catch (e2) {
      log("error", "Reading the MIME content through Graph failed", errText(e2));
      if (!shared) { throw officeError; }   // own mailbox: report the original, well-known error
      e2.sharedMailbox = sharedLabel(shared);
      throw e2;
    }
  }

  async function reportViaGraph(cat) {
    var shared = state.shared;
    var asMailbox = sendAs();
    var wantMove = !!(CFG.MOVE_AFTER_REPORT && cat.moveTo) &&
      (!shared || CFG.MOVE_IN_SHARED_MAILBOX !== false);
    var scopes = [asMailbox ? "Mail.Send.Shared" : "Mail.Send"];
    if (wantMove && !shared) { scopes.push("Mail.ReadWrite"); }

    var token = await getToken(scopes);

    // Token for reading or moving the item itself. In the user's own mailbox the send token
    // already carries Mail.ReadWrite when a move is wanted. In every other case – above all
    // for a shared mailbox, which needs a *.Shared scope – it is requested on first use and
    // silently, so that a missing consent can never block a report with a sign-in window.
    var itemScope = shared ? (wantMove ? "Mail.ReadWrite.Shared" : "Mail.Read.Shared") : "Mail.ReadWrite";
    var itemToken = (!shared && wantMove) ? token : null;
    var mailboxToken = async function () {
      if (!itemToken) {
        try {
          itemToken = await getToken([itemScope], true);
        } catch (e) {
          e.permission = itemScope;
          throw e;
        }
      }
      return itemToken;
    };

    var eml = await fetchEml(shared, mailboxToken);
    var bytes = Math.floor(eml.length * 3 / 4);
    log("info", "EML fetched", { bytes: bytes });
    if (bytes > (CFG.MAX_GRAPH_ATTACHMENT_BYTES || 3 * 1024 * 1024)) {
      var mb = (bytes / (1024 * 1024)).toFixed(1);
      log("warn", "Message too large for Graph sendMail – falling back to compose", { mb: mb });
      return { fallback: true, reason: t("tooLarge", { mb: mb }) };
    }

    var recipients = function (list) {
      return (list || []).map(function (a) { return { emailAddress: { address: a } }; });
    };

    var soc = socAddresses();
    var message = {
      subject: subjectFor(cat),
      body: { contentType: "Text", content: bodyPlain(cat) },
      toRecipients: recipients([cat.address]),
      internetMessageHeaders: [{ name: "X-MS-AddIn-Base64Encode", value: "true" }],
      attachments: [{
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: "orig_msg_" + randomString() + ".raw.eml",
        contentType: "message/rfc822",
        contentBytes: eml
      }]
    };
    if (soc.length) {
      if (socMode() === "bcc") { message.bccRecipients = recipients(soc); }
      else { message.ccRecipients = recipients(soc); }
    }
    if (asMailbox) {
      // Send from the shared mailbox: /me/sendMail with "from" set needs only Send As (or
      // Send on Behalf) on that mailbox plus Mail.Send.Shared – not Full Access.
      message.from = { emailAddress: { address: asMailbox } };
    }

    await graphPost("/me/sendMail", token, {
      message: message,
      saveToSentItems: keepCopy()
    });
    log("info", "Report sent via Graph", {
      to: cat.address, category: cat.id, as: asMailbox || "me", soc: soc, socMode: socMode(),
      sharedMailbox: shared ? sharedLabel(shared) : undefined
    });

    var moved = false;
    if (wantMove) {
      try {
        await moveItem(await mailboxToken(), cat.moveTo, shared);
        moved = true;
        log("info", "Message moved", { to: cat.moveTo, mailbox: shared ? sharedLabel(shared) : "me" });
      } catch (e) {
        log("warn", "Report sent, but the message could not be moved", errText(e));
      }
    } else if (shared && CFG.MOVE_AFTER_REPORT && cat.moveTo) {
      log("info", "Message left in place – MOVE_IN_SHARED_MAILBOX is false", { mailbox: sharedLabel(shared) });
    }

    setStatus("ok",
      t("reported", { category: labelOf(cat).toLowerCase(), address: cat.address }) +
      (moved ? t("movedTo", { folder: folderName(cat.moveTo) }) : "") + "." + socStatusSuffix());
    return { fallback: false };
  }

  /* ------------------------------------------------------------------ */
  /* Reporting                                                          */
  /* ------------------------------------------------------------------ */

  function friendlyError(e) {
    var text = errText(e);
    if (e && e.sharedMailbox && e.consentRequired) {
      return t("errSharedConsent", { permission: e.permission || "Mail.ReadWrite.Shared" });
    }
    if (e && e.sharedMailbox && e.status) {
      return t("errSharedAccess", { mailbox: e.sharedMailbox, status: e.status });
    }
    if (/InteractionRequired|consent|AADSTS65001/i.test(text)) {
      return t("errConsent");
    }
    if (/ErrorSendAsDenied|SendAsDenied/i.test(text)) {
      return t("errGraphDenied", { status: "Send As" });
    }
    if (e && (e.status === 401 || e.status === 403)) {
      return t("errGraphDenied", { status: e.status });
    }
    if (/getAsFileAsync/i.test(text)) {
      return t("errGetFile");
    }
    return text;
  }

  async function runGuarded(fn) {
    if (state.busy) { return; }
    state.busy = true;
    setButtonsEnabled(false);
    try {
      await fn();
    } catch (e) {
      log("error", "Error", errText(e));
      setStatus("err", t("notSent") + friendlyError(e));
    } finally {
      state.busy = false;
      setButtonsEnabled(!!state.item);
    }
  }

  async function report(cat) {
    if (state.busy) { return; }
    if (state.mode === "none") {
      renderNotice();
      return;
    }
    if (!state.item) {
      setStatus("warn", t("selectFirst"));
      return;
    }
    state.busy = true;
    setButtonsEnabled(false);
    setStatus("wait", t("sending", { category: labelOf(cat).toLowerCase() }));

    try {
      await sharedLookupDone();
      log("info", "Report started", {
        category: cat.id, mode: state.mode, language: state.language,
        sharedMailbox: state.shared ? sharedLabel(state.shared) : undefined
      });
      if (state.mode === "graph") {
        var r;
        try {
          r = await reportViaGraph(cat);
        } catch (e) {
          log("error", "Graph send failed", errText(e));
          setStatus("err", t("notSentAuto") + friendlyError(e),
            state.fallbackOk ? { retryCompose: cat } : undefined);
          return;
        }
        if (r.fallback) {
          if (state.fallbackOk) {
            await reportViaCompose(cat, r.reason);
          } else {
            log("warn", "Compose fallback disabled – report not sent", r.reason);
            setStatus("err", t("notSentAuto") + r.reason);
          }
        }
      } else {
        await reportViaCompose(cat);
      }
    } catch (e) {
      log("error", "Report failed", errText(e));
      setStatus("err", t("notSent") + friendlyError(e));
    } finally {
      state.busy = false;
      setButtonsEnabled(!!state.item);
    }
  }

  /* ------------------------------------------------------------------ */
  /* Startup                                                            */
  /* ------------------------------------------------------------------ */

  /**
   * Decide the send mode.
   *   SEND_MODE "graph" (default): graph when configured and supported; otherwise compose
   *   as a fallback (with a visible warning) if COMPOSE_FALLBACK is not false, else "none".
   *   SEND_MODE "compose": compose, no warning.
   */
  function detectMode() {
    var wantGraph = String(CFG.SEND_MODE || "graph").toLowerCase() !== "compose";
    var reasons = [];
    var reasonKey = null;

    if (wantGraph) {
      if (!CFG.CLIENT_ID) {
        reasons.push("CLIENT_ID is empty in config.js");
        reasonKey = reasonKey || "reasonNotConfigured";
      }
      if (!isSet("Mailbox", "1.14")) {
        reasons.push("Mailbox 1.14 (getAsFileAsync) not supported by this client");
        reasonKey = reasonKey || "reasonClientUnsupported";
      }
      if (!isSet("NestedAppAuth", "1.1")) {
        reasons.push("NestedAppAuth 1.1 not supported by this client");
        reasonKey = reasonKey || "reasonClientUnsupported";
      }
      if (!window.msal || typeof msal.createNestablePublicClientApplication !== "function") {
        reasons.push("MSAL (lib/msal-browser.min.js) not loaded");
        reasonKey = reasonKey || "reasonMsalMissing";
      }
    }

    var composeSupported = isSet("Mailbox", "1.6");
    var fallbackAllowed = CFG.COMPOSE_FALLBACK !== false;
    state.fallbackOk = composeSupported && fallbackAllowed;

    if (wantGraph && reasons.length === 0) {
      state.mode = "graph";
      state.notice = null;
    } else if (!wantGraph && composeSupported) {
      state.mode = "compose";
      state.notice = null;
    } else if (wantGraph && state.fallbackOk) {
      state.mode = "compose";
      state.notice = { kind: "warn", key: "autoUnavailable", reasonKey: reasonKey };
    } else if (wantGraph) {
      state.mode = "none";
      state.notice = { kind: "err", key: "autoUnavailableNoFallback", reasonKey: reasonKey };
      if (!composeSupported) { reasons.push("Mailbox 1.6 (displayNewMessageForm) not supported by this client"); }
    } else {
      state.mode = "none";
      state.notice = { kind: "err", key: "composeUnavailable", reasonKey: null };
      reasons.push("Mailbox 1.6 (displayNewMessageForm) not supported by this client");
    }

    state.modeDetected = true;
    updateModeChip();
    renderNotice();
    log("info", "Send mode: " + state.mode, {
      configured: wantGraph ? "graph" : "compose",
      fallback: state.fallbackOk,
      reasons: reasons.length ? reasons : undefined
    });
    updateSettingsVisibility();
  }

  function init() {
    var diag = { addin: VERSION };
    try {
      diag.host = Office.context.diagnostics.host;
      diag.platform = Office.context.diagnostics.platform;
      diag.version = Office.context.diagnostics.version;
      diag.displayLanguage = displayLanguage();
    } catch (e) { /* ignore */ }
    log("info", "Start", diag);

    if (!languages().length) {
      setStatus("err", "No locale files loaded (src/locales/*.js)");
      return;
    }

    applyLocale(resolveLanguage());
    initSettings();
    detectMode();
    refreshItem();

    if (isSet("Mailbox", "1.5")) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () {
        refreshItem();
      }, function (r) {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          log("warn", "Could not register ItemChanged handler", r.error && r.error.message);
        }
      });
    }
  }

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Outlook) {
      if (languages().length) { applyLocale(resolveLanguage()); }
      setStatus("err", t("errNotOutlook"));
      return;
    }
    try {
      init();
    } catch (e) {
      log("error", "Startup failed", errText(e));
      setStatus("err", t("errStart") + errText(e));
    }
  });

})();
