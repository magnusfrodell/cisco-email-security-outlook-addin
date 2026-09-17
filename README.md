# Rapportér mail til Cisco

[![CI](https://github.com/OWNER/outlook-cisco-submission-addin/actions/workflows/ci.yml/badge.svg)](https://github.com/OWNER/outlook-cisco-submission-addin/actions/workflows/ci.yml)

An Outlook add-in, in Danish, that lets users report spam, phishing, virus and marketing mail to
Cisco Talos – and report legitimate mail that was wrongly stopped. Reports go to Cisco's
documented submission addresses in the same format Cisco's own *Secure Email Submission* add-in
uses, so they show up in the Talos Email Status Portal like any other submission.

It exists because Cisco's add-in only has an English task pane and cannot be localized. All
strings live in one file, so the same code serves Swedish, Norwegian or any other language
(`examples/config.sv-SE.js`).

<p align="center"><img src="docs/images/taskpane.png" width="340" alt="The task pane in Danish"></p>

## Features

- Six categories: spam, phishing, virus, marketing (→ Junk) and legitimate, not-marketing (→ Inbox);
  any of them can be hidden.
- **Automatic mode** – one click: the message is fetched as EML, sent via Microsoft Graph on the
  user's behalf (Nested App Authentication, no secrets) and moved to Junk or Inbox.
- **Compose mode** – zero infrastructure: a new message to the right Cisco address opens with the
  reported mail attached; the user presses Send. Also the fallback when a message is larger than
  Graph accepts or when anything fails.
- Optional Cc to an internal mailbox (SOC), per-user "keep a copy in Sent Items".
- Technical log inside the pane for support; no backend, no telemetry.
- Static files only – host them anywhere with HTTPS.

## Quick start

```bash
git clone https://github.com/OWNER/outlook-cisco-submission-addin.git
cd outlook-cisco-submission-addin
scripts/set-host.sh addin.firma.dk      # put your HTTPS host into src/manifest.xml
# upload src/ to https://addin.firma.dk/ , then sideload src/manifest.xml via https://aka.ms/olksideload
```

That gives you compose mode. For automatic mode, register an Entra app with the SPA redirect
`brk-multihub://addin.firma.dk` and delegated `Mail.Send` + `Mail.ReadWrite`, and put its client
ID into `src/config.js`. The full procedure is in [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md)
(English) and [docs/INSTALLATION.da.md](docs/INSTALLATION.da.md) (Danish, for the customer).

## Documentation

| Document                                          | Contents                                                        |
| ------------------------------------------------- | --------------------------------------------------------------- |
| [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md)           | Hosting, Entra app registration, sideloading, central deployment |
| [docs/INSTALLATION.da.md](docs/INSTALLATION.da.md) | The same, in Danish, written for the customer's IT department   |
| [docs/CONFIGURATION.md](docs/CONFIGURATION.md)     | Every key in `config.js`, with recipes                          |
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)       | How it works, sequence diagrams, function-by-function reference |
| [docs/LOCALIZATION.md](docs/LOCALIZATION.md)       | Translating the pane and the manifest                           |
| [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | Reading the log, error codes, fixes                             |
| [docs/index.html](docs/index.html)                 | All of the above as one styled, offline HTML file (`npm run docs` rebuilds it) |
| [SECURITY.md](SECURITY.md)                         | What data goes where, permissions                               |

## Repository layout

```text
src/            the deployable add-in (manifest, task pane, config, MSAL, icons)
docs/           documentation (Markdown + generated docs/index.html)
examples/       config.sv-SE.js – Swedish strings as a translation template
scripts/        set-host.sh / set-host.ps1, build-docs.py
tests/          Office.js stub and headless smoke test
.github/        CI (validate + test) and optional GitHub Pages deployment
```

## Development

There is no build step; `src/` runs as-is.

```bash
npm install                 # validator, dev certs, static server
npm run check               # node --check on the JavaScript
npm run validate            # Microsoft's manifest validator
pip install playwright && playwright install chromium
npm test                    # headless smoke test with the Office.js stub
npm run dev-certs && npm run serve   # https://localhost:3000 for sideloading a dev manifest
npm run docs                # regenerate docs/index.html
```

`scripts/set-host.sh localhost:3000 manifest-dev.xml` produces a manifest for the local server.

Client support: Outlook on the web, new Outlook on Windows, classic Outlook on Windows and
Outlook on Mac (Microsoft 365). Not Outlook mobile. Details in
[docs/ARCHITECTURE.md § 7](docs/ARCHITECTURE.md#7-client-support).

## Relationship to Cisco

This project is not provided or supported by Cisco. It sends reports to the public submission
addresses Cisco documents for customers (`*@access.ironport.com`) using the same message shape
as Cisco's add-in. Cisco Secure Awareness simulated-phishing handling is not replicated.

## License

MIT – see [LICENSE](LICENSE). Bundles MSAL.js (`src/lib/msal-browser.min.js`), MIT, © Microsoft.
