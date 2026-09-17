#!/usr/bin/env python3
"""
Smoke test for the task pane, run in headless Chromium with a stubbed Office.js.

It verifies that the pane renders from config.js, that all enabled categories
become buttons, that the compose-mode flow opens a new message addressed to the
right Cisco address with the item attached, and that no JavaScript errors occur.

Usage:
    pip install playwright && playwright install chromium
    python3 tests/smoke_test.py [--screenshots DIR]
"""
import argparse
import asyncio
import json
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
    """Extract a few facts from config.js without a JS engine."""
    js = (SRC / "config.js").read_text(encoding="utf-8")
    categories = re.findall(r'id:\s*"([a-z_]+)",\s*group:\s*"([a-z_]+)",\s*enabled:\s*(true|false)', js)
    addresses = dict(re.findall(r'id:\s*"([a-z_]+)"[\s\S]*?address:\s*"([^"]+)"', js))
    title = re.search(r'title:\s*"([^"]+)"', js).group(1)
    return {"categories": categories, "addresses": addresses, "title": title}


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

            check(await page.text_content(".head-title") == cfg["title"], "title comes from config.js STRINGS")
            buttons = await page.query_selector_all(".btn-cat")
            check(len(buttons) == len(enabled), f"{len(enabled)} enabled categories rendered as buttons ({len(buttons)} found)")
            chip = (await page.text_content("#mode-chip") or "").strip()
            check(chip != "" and "Starter" not in chip, f"mode chip set ({chip!r})")
            check(await page.is_hidden("#settings"), "settings panel hidden in compose mode")
            check((await page.text_content("#mail-subject") or "").startswith("Din konto"), "selected mail subject shown")

            if screenshots:
                screenshots.mkdir(parents=True, exist_ok=True)
                await page.screenshot(path=str(screenshots / "taskpane.png"), full_page=True)

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
                check(bool(f.get("subject")), "subject set")
            check(await page.get_attribute("#status", "data-kind") == "ok", "status box shows ok state")
            status_text = await page.text_content("#status-text") or ""
            check(cfg["addresses"]["phish"] in status_text, "status text names the Cisco address")
            check(await page.is_enabled('button[data-id="spam"]'), "buttons re-enabled after report")

            if screenshots:
                await page.screenshot(path=str(screenshots / "taskpane-status.png"), full_page=False)

            # simulate selection change with no item
            await page.evaluate("Office.context.mailbox.item = null; window.__stubItemChanged && window.__stubItemChanged({})")
            await page.wait_for_timeout(200)
            check(not await page.is_enabled('button[data-id="spam"]'), "buttons disabled when no mail is selected")

            check(not errors, "no JavaScript errors" + (f": {errors}" if errors else ""))
            await browser.close()

    print(f"\n{len(failures)} failure(s)")
    return 1 if failures else 0


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--screenshots", type=pathlib.Path, help="directory to write taskpane.png and taskpane-status.png")
    args = ap.parse_args()
    sys.exit(asyncio.run(run(args.screenshots)))


if __name__ == "__main__":
    main()
