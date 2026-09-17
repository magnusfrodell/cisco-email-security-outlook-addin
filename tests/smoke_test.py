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


def build_preview(tmp: pathlib.Path) -> pathlib.Path:
    """Copy src/ to a temp dir and swap the hosted office.js for the stub."""
    dest = tmp / "src"
    shutil.copytree(SRC, dest)
    shutil.copy(STUB, dest / "office-stub.js")
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
            check(await page.is_visible("#settings"), "settings panel visible (language selector)")
            check(await page.is_hidden("#keep-copy-row"), "keep-copy row hidden in compose mode")
            check((await page.text_content("#mail-subject") or "").startswith("Din konto"), "selected mail subject shown")
            options = await page.eval_on_selector_all("#lang-select option", "els => els.map(e => e.value)")
            check(options == ["auto", "da", "en", "sv"], f"language selector lists auto + locales ({options})")
            check(await page.input_value("#lang-select") == "auto", "selector starts on automatic")

            if screenshots:
                screenshots.mkdir(parents=True, exist_ok=True)
                await page.screenshot(path=str(screenshots / "taskpane.png"), full_page=True)

            # --- language switching ----------------------------------------------------
            await page.select_option("#lang-select", "sv")
            await page.wait_for_timeout(300)
            check(await page.text_content(".head-title") == cfg["titles"]["sv"], "switching to Swedish changes the title")
            check(await page.get_attribute("html", "lang") == "sv", "<html lang> follows the switch")
            first_label = await page.text_content(".btn-cat .btn-cat-label")
            check(first_label == "Skräppost", f"category buttons re-rendered in Swedish ({first_label!r})")
            saved = await page.evaluate("Office.context.roamingSettings.get('language')")
            check(saved == "sv", "language choice saved in roamingSettings")
            if screenshots:
                await page.screenshot(path=str(screenshots / "taskpane-sv.png"), full_page=True)

            await page.select_option("#lang-select", "en")
            await page.wait_for_timeout(300)
            check(await page.text_content(".head-title") == cfg["titles"]["en"], "switching to English changes the title")
            if screenshots:
                await page.screenshot(path=str(screenshots / "taskpane-en.png"), full_page=True)

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
                check("Rapport sendt" in (f.get("htmlBody") or ""), "body text comes from the active locale")
            check(await page.get_attribute("#status", "data-kind") == "ok", "status box shows ok state")
            status_text = await page.text_content("#status-text") or ""
            check(cfg["addresses"]["phish"] in status_text, "status text names the Cisco address")
            check(await page.is_enabled('button[data-id="spam"]'), "buttons re-enabled after report")

            if screenshots:
                await page.screenshot(path=str(screenshots / "taskpane-status.png"), full_page=False)

            # --- selection change with no item -------------------------------------------
            await page.evaluate("Office.context.mailbox.item = null; window.__stubItemChanged && window.__stubItemChanged({})")
            await page.wait_for_timeout(200)
            check(not await page.is_enabled('button[data-id="spam"]'), "buttons disabled when no mail is selected")

            check(not errors, "no JavaScript errors" + (f": {errors}" if errors else ""))
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
