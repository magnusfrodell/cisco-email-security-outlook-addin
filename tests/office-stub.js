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
/*
 * Minimal Office.js stub used by tests/smoke_test.py so the task pane can be
 * rendered and exercised in a plain browser, outside Outlook.
 * displayLanguage is da-DK so LANGUAGE: "auto" resolves to Danish.
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
      displayLanguage: "da-DK",   // drives LANGUAGE: "auto" in the tests
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
        userProfile: { emailAddress: "bruger@firma.example", displayName: "Bruger Brugersen" },
        item: {
          itemType: "message",
          itemId: "AAMkAGI2THVSAAA=",
          subject: "Din konto er spærret – bekræft dine oplysninger nu",
          from: { displayName: "Netbank Sikkerhed", emailAddress: "sikkerhed@netbank-login.example" },
          dateTimeCreated: new Date("2026-09-17T08:15:00Z"),
          internetMessageId: "<20260917081500.ABC123@netbank-login.example>"
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
