/* global Office, msal */
/*
 * Rapportér mail til Cisco – task pane-logik.
 *
 * To afsendelsesmåder:
 *   graph   : Mailen hentes som EML (getAsFileAsync) og sendes via Microsoft Graph
 *             /me/sendMail med Nested App Authentication (NAA). Kræver CLIENT_ID i
 *             config.js, Mailbox 1.14 og NestedAppAuth 1.1 i Outlook-klienten.
 *   compose : Der åbnes en ny mail til Ciscos adresse med den valgte besked
 *             vedhæftet som element ("videresend som vedhæftning"). Brugeren
 *             trykker selv Send. Virker uden appregistrering (Mailbox 1.6+).
 *
 * graph falder automatisk tilbage til compose hvis forudsætningerne mangler,
 * og tilbyder compose som "Send manuelt i stedet" hvis afsendelsen fejler.
 *
 * Alle tekster hentes fra RAPPORT_CONFIG.STRINGS (config.js). Denne fil
 * indeholder ingen sprogafhængige tekster.
 */
(function () {
  "use strict";

  var VERSION = "1.0.0";
  var CFG = window.RAPPORT_CONFIG || {};
  var S = CFG.STRINGS || {};
  var LOG_PREFIX = "[CiscoRapport]";
  var GRAPH = "https://graph.microsoft.com/v1.0";

  var state = {
    mode: "compose",
    msalApp: null,
    busy: false,
    item: null
  };

  /* ------------------------------------------------------------------ */
  /* Tekster                                                            */
  /* ------------------------------------------------------------------ */

  /** Hent tekst fra STRINGS og erstat {navn}-pladsholdere. Ukendt nøgle -> nøglen selv. */
  function t(key, params) {
    var text = Object.prototype.hasOwnProperty.call(S, key) ? String(S[key]) : key;
    if (params) {
      Object.keys(params).forEach(function (k) {
        text = text.split("{" + k + "}").join(String(params[k]));
      });
    }
    return text;
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
    if (CFG.LANG) { document.documentElement.setAttribute("lang", CFG.LANG); }
  }

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
    try { fn.call(console, LOG_PREFIX, line); } catch (e) { /* ignorer */ }
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
  /* Hjælpere                                                           */
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

  function subjectFor(cat) {
    return String(CFG.SUBJECT || "Cisco Secure Email submission").split("{category}").join(cat.label);
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
    } catch (e) { /* ignorer */ }
  }

  function clearStatus() {
    $("status").hidden = true;
    $("status-actions").innerHTML = "";
  }

  function setButtonsEnabled(on) {
    var btns = document.querySelectorAll(".btn-cat");
    for (var i = 0; i < btns.length; i++) { btns[i].disabled = !on; }
  }

  function renderCategories() {
    var groups = CFG.GROUPS || {};
    Object.keys(groups).forEach(function (g) {
      var el = $("group-" + g + "-title");
      if (el) { el.textContent = groups[g]; }
    });

    enabledCategories().forEach(function (cat) {
      var container = $("group-" + cat.group) || $("group-missed");
      var b = document.createElement("button");
      b.type = "button";
      b.className = "btn-cat";
      b.setAttribute("data-id", cat.id);
      b.setAttribute("title", t("sentTo", { address: cat.address }));
      var l = document.createElement("span");
      l.className = "btn-cat-label";
      l.textContent = cat.label;
      var h = document.createElement("span");
      h.className = "btn-cat-hint";
      h.textContent = cat.hint || "";
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
  }

  /* ------------------------------------------------------------------ */
  /* Indstillinger (roamingSettings)                                    */
  /* ------------------------------------------------------------------ */

  function keepCopy() {
    try {
      var v = Office.context.roamingSettings.get("keepCopy");
      if (v === true || v === false) { return v; }
    } catch (e) { /* ignorer */ }
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
  }

  /* ------------------------------------------------------------------ */
  /* Afsendelse: compose (ny mail med vedhæftning)                      */
  /* ------------------------------------------------------------------ */

  function reportViaCompose(cat, reason) {
    var item = state.item;
    if (!item || !item.itemId) {
      throw new Error(t("selectFirst"));
    }
    var params = {
      toRecipients: [cat.address],
      ccRecipients: CFG.CC_ADDRESSES || [],
      subject: subjectFor(cat),
      htmlBody: "<p>" + escapeHtml(CFG.BODY_TEXT || "") + "</p>",
      attachments: [{
        type: "item",
        name: (item.subject || t("reportedMailName")).slice(0, 120),
        itemId: item.itemId
      }]
    };
    log("info", "Opening new message (compose mode)", { to: cat.address, reason: reason || null });

    var done = function () {
      setStatus("ok",
        t("composeOpened", { address: cat.address }) + (reason ? " (" + reason + ")" : ""),
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
  /* Afsendelse: Microsoft Graph via NAA                                */
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

  async function getToken(scopes) {
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
      log("info", "Token acquired (silent)");
      return silent.accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) {
        log("warn", "Silent token failed – trying interactive", errText(e));
        var popup = await state.msalApp.acquireTokenPopup(request);
        log("info", "Token acquired (interactive)");
        return popup.accessToken;
      }
      throw e;
    }
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
    if (!res.ok) {
      var text = "";
      try { text = await res.text(); } catch (e) { /* ignorer */ }
      var err = new Error("Graph " + path + " returned " + res.status + " " + text.slice(0, 300));
      err.status = res.status;
      throw err;
    }
    return res;
  }

  async function moveItem(token, folder) {
    var restId = Office.context.mailbox.convertToRestId(
      state.item.itemId, Office.MailboxEnums.RestVersion.v2_0);
    await graphPost("/me/messages/" + encodeURIComponent(restId) + "/move", token, {
      destinationId: folder
    });
  }

  async function reportViaGraph(cat) {
    var scopes = ["Mail.Send"];
    if (CFG.MOVE_AFTER_REPORT && cat.moveTo) { scopes.push("Mail.ReadWrite"); }

    var token = await getToken(scopes);

    var eml = await getEmlBase64();
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

    var message = {
      subject: subjectFor(cat),
      body: { contentType: "Text", content: CFG.BODY_TEXT || "" },
      toRecipients: recipients([cat.address]),
      internetMessageHeaders: [{ name: "X-MS-AddIn-Base64Encode", value: "true" }],
      attachments: [{
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: "orig_msg_" + randomString() + ".raw.eml",
        contentType: "message/rfc822",
        contentBytes: eml
      }]
    };
    if (CFG.CC_ADDRESSES && CFG.CC_ADDRESSES.length) {
      message.ccRecipients = recipients(CFG.CC_ADDRESSES);
    }

    await graphPost("/me/sendMail", token, {
      message: message,
      saveToSentItems: keepCopy()
    });
    log("info", "Report sent via Graph", { to: cat.address, category: cat.id });

    var moved = false;
    if (CFG.MOVE_AFTER_REPORT && cat.moveTo) {
      try {
        await moveItem(token, cat.moveTo);
        moved = true;
        log("info", "Message moved", { to: cat.moveTo });
      } catch (e) {
        log("warn", "Report sent, but the message could not be moved", errText(e));
      }
    }

    setStatus("ok",
      t("reported", { category: cat.label.toLowerCase(), address: cat.address }) +
      (moved ? t("movedTo", { folder: folderName(cat.moveTo) }) : "") + ".");
    return { fallback: false };
  }

  /* ------------------------------------------------------------------ */
  /* Rapportering                                                       */
  /* ------------------------------------------------------------------ */

  function friendlyError(e) {
    var text = errText(e);
    if (/InteractionRequired|consent|AADSTS65001/i.test(text)) {
      return t("errConsent");
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
    if (!state.item) {
      setStatus("warn", t("selectFirst"));
      return;
    }
    state.busy = true;
    setButtonsEnabled(false);
    setStatus("wait", t("sending", { category: cat.label.toLowerCase() }));
    log("info", "Report started", { category: cat.id, mode: state.mode });

    try {
      if (state.mode === "graph") {
        var r;
        try {
          r = await reportViaGraph(cat);
        } catch (e) {
          log("error", "Graph send failed", errText(e));
          setStatus("err", t("notSentAuto") + friendlyError(e), { retryCompose: cat });
          return;
        }
        if (r.fallback) {
          await reportViaCompose(cat, r.reason);
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
  /* Opstart                                                            */
  /* ------------------------------------------------------------------ */

  function detectMode() {
    var reasons = [];
    if (!CFG.CLIENT_ID) { reasons.push("CLIENT_ID is empty in config.js"); }
    if (!isSet("Mailbox", "1.14")) { reasons.push("Mailbox 1.14 (getAsFileAsync) not supported by this client"); }
    if (!isSet("NestedAppAuth", "1.1")) { reasons.push("NestedAppAuth 1.1 not supported by this client"); }
    if (!window.msal || typeof msal.createNestablePublicClientApplication !== "function") {
      reasons.push("MSAL (lib/msal-browser.min.js) not loaded");
    }
    if (!isSet("Mailbox", "1.6")) {
      reasons.push("Mailbox 1.6 (displayNewMessageForm) not supported by this client");
    }

    if (reasons.length === 0) {
      state.mode = "graph";
      setModeChip(t("chipAuto"), "chip-ok");
    } else {
      state.mode = "compose";
      setModeChip(t("chipCompose"), "chip-neutral");
    }
    log("info", "Send mode: " + state.mode, reasons.length ? reasons : undefined);
    $("settings").hidden = (state.mode !== "graph");
  }

  function init() {
    var diag = { addin: VERSION };
    try {
      diag.host = Office.context.diagnostics.host;
      diag.platform = Office.context.diagnostics.platform;
      diag.version = Office.context.diagnostics.version;
    } catch (e) { /* ignorer */ }
    log("info", "Start", diag);

    applyStrings();
    renderCategories();
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
      applyStrings();
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
