#!/usr/bin/env python3
# Copyright (c) 2026 Cisco and/or its affiliates.
#
# This software is licensed to you under the terms of the Cisco Sample
# Code License, Version 1.1 (the "License"). You may obtain a copy of the
# License at
#
#                https://developer.cisco.com/docs/licenses
#
# All use of the material herein must be in accordance with the terms of
# the License. All rights not expressly granted by the License are
# reserved. Unless required by applicable law or agreed to separately in
# writing, software distributed under the License is distributed on an "AS
# IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
# or implied.
"""
Smoke test for the task pane, run in headless Chromium with a stubbed Office.js.

It verifies that the pane renders from config.js and the locale files, that all enabled
categories become buttons, that the compose-mode flow opens a new message addressed to the
right Cisco address with the item attached, that the language selector switches every text,
and that no JavaScript errors occur.

Automatic send is exercised too, against a mocked MSAL and a mocked Microsoft Graph: which
tokens are requested, which Graph calls are made and what they carry – for a mail in the
user's own mailbox and for a mail in a shared mailbox, with and without a working
getAsFileAsync, and with the *.Shared permission missing. No network access is needed.

Usage:
    pip install playwright && playwright install chromium
    python3 tests/smoke_test.py [--screenshots DIR]
"""
import argparse
import asyncio
import pathlib
import re
import shutil
import sys
import tempfile

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC = ROOT / "src"
STUB = ROOT / "tests" / "office-stub.js"

# The stub's item id as it appears in a Graph URL, and the shared-mailbox contexts used below.
ITEM = "AAMkAGI2THVSAAA%3D"
SHARED = {"owner": "support@firma.example", "targetMailbox": "support@firma.example", "delegatePermissions": 63}
SHARED_PATH = "/users/support%40firma.example"
# Some clients report a location value instead of the address in targetMailbox.
SHARED_GUID = {
    "owner": "support@firma.example",
    "targetMailbox": "0f3b6c1e-aaaa-4bbb-8ccc-123456789abc@7d1f0c2a-dddd-4eee-9fff-abcdef012345",
    "delegatePermissions": 63,
}
SHARED_GUID_PATH = "/users/0f3b6c1e-aaaa-4bbb-8ccc-123456789abc%407d1f0c2a-dddd-4eee-9fff-abcdef012345"

# Replaces MSAL's nested-app client and window.fetch inside the page. Tokens are named after
# their scopes so a test can see which token a Graph call was made with.
GRAPH_MOCK_JS = r"""
(() => {
  window.__graph = {
    tokens: [], calls: [], deny: [], notFound: [],
    mime: "From: graph@example.com\r\nSubject: MIME content read through Graph\r\n\r\nbody\r\n"
  };
  msal.createNestablePublicClientApplication = async () => ({
    acquireTokenSilent: async (req) => {
      window.__graph.tokens.push({ scopes: req.scopes, how: "silent" });
      const denied = req.scopes.find((s) => window.__graph.deny.includes(s));
      if (denied) {
        throw new msal.InteractionRequiredAuthError("consent_required", "AADSTS65001: consent missing for " + denied);
      }
      return { accessToken: "tok[" + req.scopes.join(" ") + "]" };
    },
    acquireTokenPopup: async (req) => {
      window.__graph.tokens.push({ scopes: req.scopes, how: "popup" });
      throw new Error("a sign-in window must not be opened in this test");
    }
  });
  window.fetch = async (url, init) => {
    init = init || {};
    const path = String(url).replace("https://graph.microsoft.com/v1.0", "");
    const call = {
      path: path,
      method: init.method || "GET",
      auth: (init.headers || {}).Authorization || "",
      body: init.body ? JSON.parse(init.body) : null
    };
    window.__graph.calls.push(call);
    if (window.__graph.notFound.some((fragment) => path.includes(fragment))) {
      return new Response(JSON.stringify({ error: { code: "ErrorItemNotFound",
        message: "The specified object was not found in the store." } }), { status: 404 });
    }
    if (call.method === "GET" && path.endsWith("/$value")) { return new Response(window.__graph.mime, { status: 200 }); }
    if (call.method === "POST" && path.endsWith("/sendMail")) { return new Response(null, { status: 202 }); }
    if (call.method === "POST" && path.endsWith("/move")) { return new Response(JSON.stringify({ id: "moved" }), { status: 201 }); }
    return new Response("{}", { status: 400 });
  };
})()
"""


async def graph_report(page, category: str, *, shared=None, file_fails: bool = False,
                       deny=(), not_found=(), config=None) -> dict:
    """Put the stub's selected item into the given context, report it and return what happened."""
    await page.evaluate(
        """(o) => {
            window.__graph.tokens = []; window.__graph.calls = [];
            window.__graph.deny = o.deny; window.__graph.notFound = o.notFound;
            window.__stub.shared = o.shared; window.__stub.getAsFileFails = o.fileFails;
            Object.assign(window.REPORTER_CONFIG, o.config);
            window.__stubItemChanged({});
        }""",
        {"shared": shared, "fileFails": file_fails, "deny": list(deny), "notFound": list(not_found),
         "config": config or {}},
    )
    await page.wait_for_timeout(100)
    await page.click(f'button[data-id="{category}"]')
    await page.wait_for_function(
        "() => { const s = document.getElementById('status'); return !s.hidden && s.getAttribute('data-kind') !== 'wait'; }",
        timeout=5000,
    )
    return await page.evaluate(
        """() => ({
            tokens: window.__graph.tokens,
            calls: window.__graph.calls,
            kind: document.getElementById('status').getAttribute('data-kind'),
            text: document.getElementById('status-text').textContent,
            retry: !!document.querySelector('#status-actions button'),
            sharedLine: document.getElementById('mail-shared').hidden
                ? null : document.getElementById('mail-shared').textContent
        })"""
    )


def build_preview(tmp: pathlib.Path, graph_ready: bool = False) -> pathlib.Path:
    """Copy src/ to a temp dir and swap the hosted office.js for the stub.

    graph_ready=True simulates a fully configured deployment (CLIENT_ID set, client with
    NestedAppAuth) so detectMode() picks graph. Sending in that variant runs against the
    mocked MSAL and Graph installed by GRAPH_MOCK_JS.
    """
    dest = tmp / ("graph" if graph_ready else "compose")
    shutil.copytree(SRC, dest)
    stub = STUB.read_text(encoding="utf-8")
    if graph_ready:
        stub, n = re.subn(r'var supported = \{ Mailbox: "1\.15" \};', 'var supported = { Mailbox: "1.15", NestedAppAuth: "1.1" };', stub)
        assert n == 1, "supported sets not found in stub"
        cfg = (dest / "config.js").read_text(encoding="utf-8")
        cfg, n = re.subn(r'CLIENT_ID: "",', 'CLIENT_ID: "00000000-0000-0000-0000-000000000000",', cfg)
        assert n == 1, "CLIENT_ID not found in config.js"
        (dest / "config.js").write_text(cfg, encoding="utf-8")
    else:
        # compose scenario: an administrator-configured SOC copy
        cfg = (dest / "config.js").read_text(encoding="utf-8")
        cfg, n = re.subn(r'SOC_ADDRESSES: \[\],', 'SOC_ADDRESSES: ["soc@example.com"],', cfg)
        assert n == 1, "SOC_ADDRESSES not found in config.js"
        (dest / "config.js").write_text(cfg, encoding="utf-8")
    (dest / "office-stub.js").write_text(stub, encoding="utf-8")
    html = (dest / "taskpane.html").read_text(encoding="utf-8")
    html, n = re.subn(r'src="https://appsforoffice\.microsoft\.com/[^"]+"', 'src="office-stub.js"', html)
    assert n == 1, "office.js script tag not found in taskpane.html"
    (dest / "taskpane.html").write_text(html, encoding="utf-8")
    return dest / "taskpane.html"


def read_config() -> dict:
    """Extract a few facts from config.js and the locale files without a JS engine."""
    js = (SRC / "config.js").read_text(encoding="utf-8")
    categories = re.findall(r'\{\s*id:\s*"([a-z_]+)",\s*group:\s*"([a-z_]+)",\s*enabled:\s*(true|false),\s*address:\s*"([^"]+)"', js)
    titles = {}
    for f in sorted((SRC / "locales").glob("*.js")):
        titles[f.stem] = re.search(r'\btitle:\s*"([^"]+)"', f.read_text(encoding="utf-8")).group(1)
    return {"categories": categories, "addresses": {c[0]: c[3] for c in categories}, "titles": titles}


async def run(screenshots: pathlib.Path | None) -> int:
    from playwright.async_api import async_playwright

    cfg = read_config()
    enabled = [c for c in cfg["categories"] if c[2] == "true"]
    failures: list[str] = []

    def check(cond: bool, msg: str) -> None:
        print(("PASS  " if cond else "FAIL  ") + msg)
        if not cond:
            failures.append(msg)

    with tempfile.TemporaryDirectory() as tmp:
        page_path = build_preview(pathlib.Path(tmp))
        async with async_playwright() as p:
            browser = await p.chromium.launch()
            page = await browser.new_page(viewport={"width": 360, "height": 900}, device_scale_factor=2)
            errors: list[str] = []
            page.on("pageerror", lambda e: errors.append(str(e)))
            page.on("console", lambda m: errors.append(m.text) if m.type == "error" else None)

            await page.goto(page_path.as_uri())
            await page.wait_for_timeout(500)

            # --- rendering from config + locales -------------------------------------
            check(len(cfg["titles"]) >= 3 and {"da", "en", "sv"} <= set(cfg["titles"]), "locale files da, en, sv present")
            check(await page.text_content(".head-title") == cfg["titles"]["da"], "LANGUAGE auto follows Outlook display language (da-DK → da)")
            check(await page.get_attribute("html", "lang") == "da", "<html lang> set from the active locale")
            buttons = await page.query_selector_all(".btn-cat")
            check(len(buttons) == len(enabled), f"{len(enabled)} enabled categories rendered as buttons ({len(buttons)} found)")
            chip = (await page.text_content("#mode-chip") or "").strip()
            check(chip != "" and "Starter" not in chip, f"mode chip set ({chip!r})")
            check(await page.is_visible("#notice") and await page.get_attribute("#notice", "data-kind") == "warn",
                  "graph is the default: warning notice shown when CLIENT_ID is empty")
            notice = await page.text_content("#notice-text") or ""
            check("CLIENT_ID" in notice and "ny mail" in notice, f"notice explains the fallback in Danish ({notice[:60]!r}…)")
            check(await page.is_enabled('button[data-id="spam"]'), "buttons enabled – compose fallback active")
            check(await page.is_visible("#settings"), "settings panel visible (language selector)")
            check(await page.is_hidden("#keep-copy-row"), "keep-copy row hidden in compose mode")
            check((await page.text_content("#mail-subject") or "").startswith("Din konto"), "selected mail subject shown")
            options = await page.eval_on_selector_all("#lang-select option", "els => els.map(e => e.value)")
            check(options == ["auto", "da", "en", "sv"], f"language selector lists auto + locales ({options})")
            check(await page.input_value("#lang-select") == "auto", "selector starts on automatic")

            if screenshots:
                screenshots.mkdir(parents=True, exist_ok=True)
                await page.screenshot(path=str(screenshots / "taskpane-fallback.png"), full_page=True)

            # --- language switching ----------------------------------------------------
            await page.select_option("#lang-select", "sv")
            await page.wait_for_timeout(300)
            check(await page.text_content(".head-title") == cfg["titles"]["sv"], "switching to Swedish changes the title")
            check(await page.get_attribute("html", "lang") == "sv", "<html lang> follows the switch")
            first_label = await page.text_content(".btn-cat .btn-cat-label")
            check(first_label == "Skräppost", f"category buttons re-rendered in Swedish ({first_label!r})")
            check("CLIENT_ID saknas" in (await page.text_content("#notice-text") or ""), "notice re-rendered in Swedish")
            saved = await page.evaluate("Office.context.roamingSettings.get('language')")
            check(saved == "sv", "language choice saved in roamingSettings")

            await page.select_option("#lang-select", "en")
            await page.wait_for_timeout(300)
            check(await page.text_content(".head-title") == cfg["titles"]["en"], "switching to English changes the title")

            await page.select_option("#lang-select", "auto")
            await page.wait_for_timeout(300)
            check(await page.text_content(".head-title") == cfg["titles"]["da"], "automatic returns to the Outlook language")

            # --- compose flow ----------------------------------------------------------
            await page.click('button[data-id="phish"]')
            await page.wait_for_timeout(400)
            calls = await page.evaluate("window.__stubCalls")
            forms = calls["displayNewMessageFormAsync"] + calls["displayNewMessageForm"]
            check(len(forms) == 1, "one new-message form opened")
            if forms:
                f = forms[0]
                check(f["toRecipients"] == [cfg["addresses"]["phish"]], f"addressed to {cfg['addresses']['phish']}")
                att = f.get("attachments") or []
                check(len(att) == 1 and att[0]["type"] == "item" and att[0]["itemId"], "reported mail attached as item")
                check("Phishing" in (f.get("subject") or ""), "subject carries the localized category label")
                check("Din konto er spærret" in (f.get("subject") or ""), "subject carries the original subject ({subject})")
                check("Rapport sendt" in (f.get("htmlBody") or ""), "body text comes from the active locale")
                check(f.get("ccRecipients") == ["soc@example.com"] and not f.get("bccRecipients"),
                      "SOC copy added as Cc (SOC_ADDRESSES, SOC_COPY_MODE cc)")
                body = f.get("htmlBody") or ""
                check("Reported by: bruger@firma.example" in body and "Submitted to: phish@access.ironport.com" in body
                      and "Message-ID: &lt;20260917081500.ABC123@netbank-login.example&gt;" in body,
                      "report details block present in the body")
            check(await page.get_attribute("#status", "data-kind") == "ok", "status box shows ok state")
            status_text = await page.text_content("#status-text") or ""
            check(cfg["addresses"]["phish"] in status_text, "status text names the Cisco address")
            check(await page.is_enabled('button[data-id="spam"]'), "buttons re-enabled after report")

            if screenshots:
                await page.screenshot(path=str(screenshots / "taskpane-status.png"), full_page=False)

            # --- item in a shared mailbox: the pane names it, the report body carries it -----
            check(await page.is_hidden("#mail-shared"), "no shared-mailbox line for a mail in the own mailbox")
            await page.evaluate("(s) => { window.__stub.shared = s; window.__stubItemChanged({}); }", SHARED)
            await page.wait_for_timeout(150)
            shared_line = (await page.text_content("#mail-shared") or "") if await page.is_visible("#mail-shared") else ""
            check(shared_line == "Delt postkasse: support@firma.example",
                  f"shared mailbox named under the selected mail ({shared_line!r})")
            await page.click('button[data-id="spam"]')
            await page.wait_for_timeout(400)
            calls = await page.evaluate("window.__stubCalls")
            forms = calls["displayNewMessageFormAsync"] + calls["displayNewMessageForm"]
            check(len(forms) == 2 and "Shared mailbox: support@firma.example" in (forms[-1].get("htmlBody") or ""),
                  "report details name the shared mailbox")
            check("Shared mailbox:" not in (forms[0].get("htmlBody") or ""),
                  "no shared-mailbox line in the details of an own-mailbox report")
            await page.evaluate("() => { window.__stub.shared = null; window.__stubItemChanged({}); }")
            await page.wait_for_timeout(150)
            check(await page.is_hidden("#mail-shared"), "shared-mailbox line gone again for an own-mailbox mail")

            # --- selection change with no item -------------------------------------------
            await page.evaluate("Office.context.mailbox.item = null; window.__stubItemChanged && window.__stubItemChanged({})")
            await page.wait_for_timeout(200)
            check(not await page.is_enabled('button[data-id="spam"]'), "buttons disabled when no mail is selected")

            check(not errors, "no JavaScript errors" + (f": {errors}" if errors else ""))

            # --- fully configured deployment: graph must be chosen, no warning ---------------
            graph_path = build_preview(pathlib.Path(tmp), graph_ready=True)
            gpage = await browser.new_page(viewport={"width": 360, "height": 900}, device_scale_factor=2)
            gerrors: list[str] = []
            gpage.on("pageerror", lambda e: gerrors.append(str(e)))
            await gpage.goto(graph_path.as_uri())
            await gpage.wait_for_timeout(500)
            chip = (await gpage.text_content("#mode-chip") or "").strip()
            check(chip == "Automatisk afsendelse", f"configured deployment selects graph mode ({chip!r})")
            check(await gpage.is_hidden("#notice"), "no warning notice in graph mode")
            check(await gpage.is_visible("#keep-copy-row"), "keep-copy setting visible in graph mode")
            check(await gpage.is_enabled('button[data-id="spam"]'), "buttons enabled in graph mode")
            if screenshots:
                await gpage.screenshot(path=str(screenshots / "taskpane.png"), full_page=True)
                await gpage.select_option("#lang-select", "en")
                await gpage.wait_for_timeout(300)
                await gpage.screenshot(path=str(screenshots / "taskpane-en.png"), full_page=True)
                await gpage.select_option("#lang-select", "sv")
                await gpage.wait_for_timeout(300)
                await gpage.screenshot(path=str(screenshots / "taskpane-sv.png"), full_page=True)
                await gpage.select_option("#lang-select", "auto")
                await gpage.wait_for_timeout(300)

            # --- automatic send against a mocked MSAL and Graph ------------------------------
            await gpage.evaluate(GRAPH_MOCK_JS)
            stub_eml = await gpage.evaluate("window.__stub.eml")
            graph_mime = await gpage.evaluate("btoa(window.__graph.mime)")
            phish = cfg["addresses"]["phish"]

            def paths(r: dict) -> list[str]:
                return [c["method"] + " " + c["path"] for c in r["calls"]]

            def scopes(r: dict) -> list[list[str]]:
                return [t["scopes"] for t in r["tokens"]]

            # Helpers return empty stand-ins instead of raising, so that a missing Graph call
            # shows up as a failed check rather than as a traceback.
            no_send = {"auth": "", "body": {"message": {
                "toRecipients": [], "internetMessageHeaders": [], "body": {"content": ""},
                "attachments": [{"contentBytes": "", "contentType": "", "name": ""}]}}}

            def send_call(r: dict) -> dict:
                return next((c for c in r["calls"] if c["path"].endswith("/sendMail")), no_send)

            def call(r: dict, i: int) -> dict:
                return r["calls"][i] if len(r["calls"]) > i else {"path": "", "method": "", "auth": "", "body": None}

            def no_popup(r: dict) -> bool:
                return all(t["how"] == "silent" for t in r["tokens"])

            # own mailbox, getAsFileAsync works – must behave exactly as before 1.4.0
            r = await graph_report(gpage, "phish")
            check(scopes(r) == [["Mail.Send", "Mail.ReadWrite"]], f"own mailbox: one token with Mail.Send + Mail.ReadWrite ({scopes(r)})")
            check(paths(r) == ["POST /me/sendMail", f"POST /me/messages/{ITEM}/move"], f"own mailbox: sendMail, then move under /me ({paths(r)})")
            msg = send_call(r)["body"]["message"]
            att = msg["attachments"][0]
            check(msg["toRecipients"] == [{"emailAddress": {"address": phish}}], f"own mailbox: report addressed to {phish}")
            check(att["contentBytes"] == stub_eml and att["contentType"] == "message/rfc822"
                  and re.fullmatch(r"orig_msg_[a-z0-9]+\.raw\.eml", att["name"]) is not None,
                  "own mailbox: EML from getAsFileAsync attached as message/rfc822 (orig_msg_*.raw.eml)")
            check(msg["internetMessageHeaders"] == [{"name": "X-MS-AddIn-Base64Encode", "value": "true"}], "own mailbox: Cisco's custom header set")
            check("Shared mailbox:" not in msg["body"]["content"] and "Reported by: bruger@firma.example" in msg["body"]["content"],
                  "own mailbox: details block unchanged (no shared-mailbox line)")
            check(call(r, 1)["body"] == {"destinationId": "junkemail"}, "own mailbox: phishing moved to junkemail")
            check(r["kind"] == "ok" and "flyttet til" in r["text"] and r["sharedLine"] is None, "own mailbox: status says sent and moved")

            # own mailbox, getAsFileAsync fails – MIME content read from /me
            r = await graph_report(gpage, "phish", file_fails=True)
            check(paths(r) == [f"GET /me/messages/{ITEM}/$value", "POST /me/sendMail", f"POST /me/messages/{ITEM}/move"],
                  f"own mailbox, no getAsFileAsync: MIME read from /me, then sendMail and move ({paths(r)})")
            check(send_call(r)["body"]["message"]["attachments"][0]["contentBytes"] == graph_mime,
                  "own mailbox, no getAsFileAsync: the MIME bytes from Graph are what gets attached")
            check(scopes(r) == [["Mail.Send", "Mail.ReadWrite"]], "own mailbox, no getAsFileAsync: no extra token needed")

            # shared mailbox, getAsFileAsync works – report from the user, move in the shared mailbox
            r = await graph_report(gpage, "ham", shared=SHARED)
            check(r["sharedLine"] == "Delt postkasse: support@firma.example", "shared mailbox: named in the pane")
            check(scopes(r) == [["Mail.Send"], ["Mail.ReadWrite.Shared"]] and no_popup(r),
                  f"shared mailbox: Mail.Send for the report, Mail.ReadWrite.Shared (silent) for the move ({scopes(r)})")
            check(paths(r) == ["POST /me/sendMail", f"POST {SHARED_PATH}/messages/{ITEM}/move"],
                  f"shared mailbox: sent from /me, moved under /users/{{mailbox}} ({paths(r)})")
            check(call(r, 1)["auth"] == "Bearer tok[Mail.ReadWrite.Shared]" and call(r, 1)["body"] == {"destinationId": "inbox"},
                  "shared mailbox: the move uses the *.Shared token and the category's folder")
            check("Shared mailbox: support@firma.example" in send_call(r)["body"]["message"]["body"]["content"],
                  "shared mailbox: details block names the shared mailbox")
            check(r["kind"] == "ok" and "flyttet til" in r["text"], "shared mailbox: status says sent and moved")

            # shared mailbox, getAsFileAsync fails – the case in which Cisco's add-in gives up
            r = await graph_report(gpage, "spam", shared=SHARED, file_fails=True)
            check(paths(r) == [f"GET {SHARED_PATH}/messages/{ITEM}/$value", "POST /me/sendMail", f"POST {SHARED_PATH}/messages/{ITEM}/move"],
                  f"shared mailbox, no getAsFileAsync: MIME read from /users/{{mailbox}}, never from /me ({paths(r)})")
            check(scopes(r) == [["Mail.Send"], ["Mail.ReadWrite.Shared"]], "shared mailbox, no getAsFileAsync: one *.Shared token serves read and move")
            check(send_call(r)["body"]["message"]["attachments"][0]["contentBytes"] == graph_mime and r["kind"] == "ok",
                  "shared mailbox, no getAsFileAsync: report sent with the MIME bytes from the shared mailbox")

            # shared mailbox, Mail.ReadWrite.Shared not consented
            r = await graph_report(gpage, "spam", shared=SHARED, deny=["Mail.ReadWrite.Shared"])
            check(paths(r) == ["POST /me/sendMail"] and r["kind"] == "ok" and "flyttet til" not in r["text"] and no_popup(r),
                  f"shared mailbox, *.Shared not consented: report still sent, move skipped, no sign-in window ({paths(r)})")
            r = await graph_report(gpage, "spam", shared=SHARED, file_fails=True, deny=["Mail.ReadWrite.Shared"])
            check(r["calls"] == [] and r["kind"] == "err" and "Mail.ReadWrite.Shared" in r["text"] and r["retry"] and no_popup(r),
                  f"shared mailbox, *.Shared not consented and no getAsFileAsync: clear error naming the permission ({r['text'][:70]!r}…)")

            # shared mailbox whose targetMailbox is a location value; the owner address is refused
            r = await graph_report(gpage, "spam", shared=SHARED_GUID, file_fails=True, not_found=[SHARED_PATH + "/"])
            check(paths(r) == [f"GET {SHARED_PATH}/messages/{ITEM}/$value", f"GET {SHARED_GUID_PATH}/messages/{ITEM}/$value",
                               "POST /me/sendMail",
                               f"POST {SHARED_PATH}/messages/{ITEM}/move", f"POST {SHARED_GUID_PATH}/messages/{ITEM}/move"]
                  and r["kind"] == "ok" and "flyttet til" in r["text"],
                  f"shared mailbox: targetMailbox is tried when the owner address is not found ({len(r['calls'])} calls)")

            # shared mailbox the user cannot read
            r = await graph_report(gpage, "spam", shared=SHARED, file_fails=True, not_found=["/users/"])
            check(paths(r) == [f"GET {SHARED_PATH}/messages/{ITEM}/$value"] and r["kind"] == "err"
                  and "support@firma.example" in r["text"] and "404" in r["text"] and r["retry"],
                  f"shared mailbox not readable: nothing sent, error names mailbox and status ({r['text'][:70]!r}…)")

            # MOVE_IN_SHARED_MAILBOX: false – report, leave the mail where it is
            r = await graph_report(gpage, "spam", shared=SHARED, config={"MOVE_IN_SHARED_MAILBOX": False})
            check(paths(r) == ["POST /me/sendMail"] and scopes(r) == [["Mail.Send"]] and "flyttet til" not in r["text"],
                  f"MOVE_IN_SHARED_MAILBOX false: no move and no *.Shared token ({scopes(r)})")
            r = await graph_report(gpage, "spam", shared=SHARED, file_fails=True, config={"MOVE_IN_SHARED_MAILBOX": False})
            check(scopes(r) == [["Mail.Send"], ["Mail.Read.Shared"]] and paths(r) == [f"GET {SHARED_PATH}/messages/{ITEM}/$value", "POST /me/sendMail"],
                  f"MOVE_IN_SHARED_MAILBOX false, no getAsFileAsync: read-only Mail.Read.Shared is enough ({scopes(r)})")

            # back to a mail in the own mailbox – nothing of the shared context may linger
            r = await graph_report(gpage, "ads", config={"MOVE_IN_SHARED_MAILBOX": True})
            check(paths(r) == ["POST /me/sendMail", f"POST /me/messages/{ITEM}/move"] and scopes(r) == [["Mail.Send", "Mail.ReadWrite"]]
                  and r["sharedLine"] is None and "Shared mailbox:" not in send_call(r)["body"]["message"]["body"]["content"],
                  "own mailbox again after a shared one: /me and the plain scopes")

            check(not gerrors, "no JavaScript errors in graph mode" + (f": {gerrors}" if gerrors else ""))
            await browser.close()

    print(f"\n{len(failures)} failure(s)")
    return 1 if failures else 0


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--screenshots", type=pathlib.Path, help="directory to write the documentation screenshots")
    args = ap.parse_args()
    sys.exit(asyncio.run(run(args.screenshots)))


if __name__ == "__main__":
    main()
