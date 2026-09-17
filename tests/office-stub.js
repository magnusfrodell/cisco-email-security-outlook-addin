/*
 * Minimal Office.js stub used by tests/smoke_test.py so the task pane can be
 * rendered and exercised in a plain browser, outside Outlook.
 * It is never shipped – src/ must not reference this file.
 */
(function () {
  function versionAtLeast(have, want) {
    var a = String(have).split(".").map(Number), b = String(want).split(".").map(Number);
    for (var i = 0; i < Math.max(a.length, b.length); i++) {
      var x = a[i] || 0, y = b[i] || 0;
      if (x !== y) { return x > y; }
    }
    return true;
  }
  var supported = { Mailbox: "1.15" }; // NestedAppAuth deliberately absent -> compose mode

  window.__stubCalls = { displayNewMessageForm: [], displayNewMessageFormAsync: [] };

  window.Office = {
    HostType: { Outlook: "Outlook" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    MailboxEnums: { ItemType: { Message: "message" }, RestVersion: { v2_0: "v2.0" } },
    EventType: { ItemChanged: "itemChanged" },
    context: {
      requirements: {
        isSetSupported: function (name, version) {
          return !!supported[name] && versionAtLeast(supported[name], version || "1.1");
        }
      },
      diagnostics: { host: "Outlook", platform: "Stub", version: "0.0" },
      roamingSettings: {
        _s: {},
        get: function (k) { return this._s[k]; },
        set: function (k, v) { this._s[k] = v; },
        saveAsync: function (cb) { cb({ status: "succeeded" }); }
      },
      mailbox: {
        item: {
          itemType: "message",
          itemId: "AAMkAGI2THVSAAA=",
          subject: "Din konto er spærret – bekræft dine oplysninger nu",
          from: { displayName: "Netbank Sikkerhed", emailAddress: "sikkerhed@netbank-login.example" }
        },
        addHandlerAsync: function (type, handler, cb) { window.__stubItemChanged = handler; cb({ status: "succeeded" }); },
        displayNewMessageForm: function (p) { window.__stubCalls.displayNewMessageForm.push(p); },
        displayNewMessageFormAsync: function (p, cb) { window.__stubCalls.displayNewMessageFormAsync.push(p); cb({ status: "succeeded" }); },
        convertToRestId: function (id) { return id; }
      }
    },
    onReady: function (cb) { setTimeout(function () { cb({ host: "Outlook" }); }, 0); }
  };
})();
