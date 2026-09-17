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
Build docs/index.html – one self-contained HTML file with all documentation.

  python3 scripts/build-docs.py            # writes docs/index.html
  python3 scripts/build-docs.py -o out.html

Requires: pip install markdown pygments
Mermaid diagrams are rendered in the browser by mermaid.js (loaded from jsdelivr); everything
else (CSS, Pygments styles, images) is embedded, so the file works offline apart from diagrams.
"""
import argparse
import base64
import datetime as dt
import html
import json
import pathlib
import re

import markdown
from pygments.formatters import HtmlFormatter

ROOT = pathlib.Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"

# (chapter slug, sidebar title, source file)
CHAPTERS = [
    ("overview", "Overview", ROOT / "README.md"),
    ("deployment", "Deployment guide", DOCS / "DEPLOYMENT.md"),
    ("installation-da", "Installationsvejledning (dansk)", DOCS / "INSTALLATION.da.md"),
    ("configuration", "Configuration", DOCS / "CONFIGURATION.md"),
    ("architecture", "Architecture & code reference", DOCS / "ARCHITECTURE.md"),
    ("localization", "Localization", DOCS / "LOCALIZATION.md"),
    ("troubleshooting", "Troubleshooting", DOCS / "TROUBLESHOOTING.md"),
    ("security", "Security", ROOT / "SECURITY.md"),
    ("changelog", "Changelog", ROOT / "CHANGELOG.md"),
    ("notice", "Notice", ROOT / "NOTICE"),
]

# links inside the Markdown that should point at chapters of this file
LINK_MAP = {
    "docs/DEPLOYMENT.md": "deployment",
    "docs/INSTALLATION.da.md": "installation-da",
    "docs/CONFIGURATION.md": "configuration",
    "docs/ARCHITECTURE.md": "architecture",
    "docs/LOCALIZATION.md": "localization",
    "docs/TROUBLESHOOTING.md": "troubleshooting",
    "docs/index.html": "overview",
    "SECURITY.md": "security",
    "CHANGELOG.md": "changelog",
    "NOTICE": "notice",
    "CONTRIBUTING.md": None,
    "CODE_OF_CONDUCT.md": None,
    "LICENSE": None,
}

MERMAID_RE = re.compile(r"```mermaid\n(.*?)\n```", re.S)
IMG_RE = re.compile(r'(src=|\]\()["\']?((?:docs/)?images/[^)"\'\s]+)["\']?')


def version() -> str:
    return json.loads((ROOT / "package.json").read_text(encoding="utf-8"))["version"]


def embed_images(text: str, base: pathlib.Path) -> str:
    """Replace images/… and docs/images/… references with data URIs."""
    def repl(m: re.Match) -> str:
        rel = m.group(2)
        path = (ROOT / rel) if rel.startswith("docs/") else (base / rel)
        if not path.exists():
            return m.group(0)
        mime = "image/png" if path.suffix == ".png" else "image/svg+xml"
        data = base64.b64encode(path.read_bytes()).decode("ascii")
        return f'{m.group(1)}"data:{mime};base64,{data}"' if m.group(1) == "src=" else f"]({'data:' + mime + ';base64,' + data}"
    return IMG_RE.sub(repl, text)


def mermaid_blocks(text: str) -> str:
    """Turn ```mermaid fences into <pre class="mermaid"> raw HTML blocks."""
    return MERMAID_RE.sub(lambda m: '\n<pre class="mermaid">\n' + html.escape(m.group(1)) + "\n</pre>\n", text)


def rewrite_links(body: str, slug: str) -> str:
    """Prefix heading ids with the chapter slug and map cross-file links to chapter anchors."""
    body = re.sub(r'id="([^"]+)"', lambda m: f'id="{slug}--{m.group(1)}"', body)
    body = re.sub(r'href="#([^"]+)"', lambda m: f'href="#{slug}--{m.group(1)}"', body)

    def link(m: re.Match) -> str:
        target, anchor = m.group(1), m.group(2)
        chapter = LINK_MAP.get(target)
        if chapter is None:
            return m.group(0)
        return f'href="#{chapter}--{anchor}"' if anchor else f'href="#{chapter}"'

    body = re.sub(r'href="(?:\./)?((?:docs/)?[A-Za-z_.-]+\.(?:md|html)|LICENSE|NOTICE)(?:#([^"]+))?"', link, body)
    return body


def render_chapter(slug: str, title: str, path: pathlib.Path) -> tuple[str, list[tuple[str, str]]]:
    text = path.read_text(encoding="utf-8")
    if path.suffix == "":  # plain-text files such as NOTICE
        text = f"# {title}\n\n```text\n{text.rstrip()}\n```\n"
    text = embed_images(text, path.parent)
    text = mermaid_blocks(text)
    md = markdown.Markdown(
        extensions=["fenced_code", "codehilite", "tables", "toc", "sane_lists", "attr_list"],
        extension_configs={
            "codehilite": {"guess_lang": False, "css_class": "hl"},
            "toc": {"toc_depth": "2-3"},
        },
    )
    body = md.convert(text)
    body = rewrite_links(body, slug)
    # h2 entries for the sidebar
    subs = [(f"{slug}--{t['id']}", t["name"]) for t in md.toc_tokens[0]["children"]] if md.toc_tokens else []
    if md.toc_tokens and md.toc_tokens[0]["level"] != 1:
        subs = [(f"{slug}--{t['id']}", t["name"]) for t in md.toc_tokens]
    return f'<section class="chapter" id="{slug}">\n{body}\n</section>', subs


CSS = """
:root{--bg:#f4f6f9;--accent:#1f77b4;--accent-dark:#16609a;--ink:#1b2733;--muted:#5b6b7b;--line:#d9e1ea;
--dark-a:#0f2a43;--dark-b:#0a1f33;--side:240px;--shadow:0 2px 10px rgba(15,42,67,.08)}
*{box-sizing:border-box}
html{scroll-behavior:smooth;scroll-padding-top:16px}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif}
a{color:var(--accent-dark)}
.side{position:fixed;top:0;left:0;bottom:0;width:var(--side);background:linear-gradient(180deg,var(--dark-a),var(--dark-b));color:#c9d6e3;overflow-y:auto;padding:20px 0 30px;z-index:20;transition:transform .2s ease}
.side .brand{padding:0 18px 14px;border-bottom:1px solid rgba(255,255,255,.08);margin-bottom:10px}
.side .brand strong{display:block;color:#fff;font-size:15px;letter-spacing:.01em}
.side .brand span{display:block;color:#8ea6bd;font-size:12px;margin-top:4px}
.side .grp{padding:12px 18px 4px;font-size:10.5px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;color:#7f99b3}
.side a{display:block;padding:5px 18px;color:#c9d6e3;text-decoration:none;font-size:13px;border-left:3px solid transparent}
.side a.sub{padding-left:30px;font-size:12.5px;color:#a9bccd}
.side a:hover{background:rgba(255,255,255,.06);color:#fff}
.side a.active{border-left-color:var(--accent);background:rgba(31,119,180,.18);color:#fff}
.main{margin-left:var(--side);padding:28px 40px 60px;max-width:calc(var(--side) + 960px)}
.chapter{background:#fff;border-radius:12px;box-shadow:var(--shadow);padding:28px 34px;margin-bottom:26px}
.chapter h1{margin:0 0 14px;font-size:24px;color:var(--dark-a);letter-spacing:-.01em}
.chapter h2{margin:30px 0 10px;font-size:18px;color:var(--dark-a);padding-bottom:6px;border-bottom:1px solid var(--line)}
.chapter h3{margin:22px 0 8px;font-size:15px;color:var(--dark-a)}
.chapter p,.chapter li{max-width:78ch}
.chapter img{max-width:100%;border-radius:8px;box-shadow:var(--shadow)}
.chapter code{font:12.5px/1.5 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;background:#eef2f6;padding:1px 5px;border-radius:5px}
.chapter pre{overflow-x:auto;border-radius:8px;border:1px solid var(--line);background:#f8fafc;padding:12px 14px}
.chapter pre code{background:none;padding:0}
.chapter .hl{margin:12px 0}
.chapter table{border-collapse:collapse;width:100%;margin:12px 0;font-size:13px;display:block;overflow-x:auto}
.chapter th,.chapter td{border:1px solid var(--line);padding:6px 9px;text-align:left;vertical-align:top}
.chapter th{background:#eef3f8;color:var(--dark-a);font-weight:600}
.chapter tr:nth-child(even) td{background:#fafcfe}
.chapter blockquote{margin:12px 0;padding:8px 14px;border-left:4px solid var(--accent);background:#f3f8fc;border-radius:0 8px 8px 0}
.chapter .mermaid{background:#fff;border:1px solid var(--line);border-radius:8px;padding:12px;text-align:center;overflow-x:auto}
.chapter p[align=center]{text-align:center}
.chip{display:inline-block;padding:2px 10px;border-radius:999px;font-size:12px;font-weight:600;background:#e4ebf3;color:var(--dark-a)}
.toggle{display:none;position:fixed;top:12px;left:12px;z-index:30;border:0;border-radius:8px;background:var(--dark-a);color:#fff;padding:8px 12px;font:inherit;font-weight:600;box-shadow:var(--shadow)}
.foot{color:var(--muted);font-size:12px;padding:6px 4px}
@media (max-width:900px){
 .side{transform:translateX(-100%)}
 .side.open{transform:none}
 .toggle{display:block}
 .main{margin-left:0;padding:60px 14px 40px}
 .chapter{padding:18px 16px}
 .side .brand{padding-top:44px}
}
@media (prefers-reduced-motion:reduce){html{scroll-behavior:auto}.side{transition:none}}
"""

JS = """
(function(){
 var side=document.getElementById('side');
 document.getElementById('toggle').addEventListener('click',function(){side.classList.toggle('open')});
 side.addEventListener('click',function(e){if(e.target.tagName==='A'&&window.innerWidth<=900){side.classList.remove('open')}});
 var links=[].slice.call(side.querySelectorAll('a[href^="#"]'));
 var targets=links.map(function(a){return document.getElementById(a.getAttribute('href').slice(1))}).filter(Boolean);
 function update(){
  var y=window.scrollY+40,cur=null;
  for(var i=0;i<targets.length;i++){if(targets[i].offsetTop<=y){cur=targets[i].id}}
  links.forEach(function(a){a.classList.toggle('active',a.getAttribute('href')==='#'+cur)});
 }
 window.addEventListener('scroll',update,{passive:true});update();
 if(window.mermaid){mermaid.initialize({startOnLoad:true,theme:'neutral',securityLevel:'loose'});}
})();
"""


def build(out: pathlib.Path) -> None:
    sections, nav = [], []
    for slug, title, path in CHAPTERS:
        section, subs = render_chapter(slug, title, path)
        sections.append(section)
        nav.append(f'<a href="#{slug}">{html.escape(title)}</a>')
        nav.extend(f'<a class="sub" href="#{sid}">{html.escape(name)}</a>' for sid, name in subs)

    pyg = HtmlFormatter(style="friendly").get_style_defs(".hl")
    built = dt.date.today().isoformat()
    ver = version()
    page = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Cisco Email Security Reporter – documentation v{ver}</title>
<style>{CSS}{pyg}</style>
<script src="https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js"></script>
</head>
<body>
<button class="toggle" id="toggle" aria-label="Menu">☰ Menu</button>
<nav class="side" id="side">
  <div class="brand"><strong>Cisco Email Security Reporter</strong><span>Outlook add-in · v{ver} · built {built}</span></div>
  <div class="grp">Contents</div>
  {''.join(nav)}
</nav>
<main class="main">
{''.join(sections)}
<div class="foot">Generated by scripts/build-docs.py from the Markdown files in the repository. Diagrams need internet access (mermaid.js from jsdelivr).</div>
</main>
<script>{JS}</script>
</body>
</html>
"""
    out.write_text(page, encoding="utf-8")
    print(f"wrote {out} ({out.stat().st_size // 1024} kB, {len(CHAPTERS)} chapters)")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-o", "--out", type=pathlib.Path, default=DOCS / "index.html")
    build(ap.parse_args().out)
