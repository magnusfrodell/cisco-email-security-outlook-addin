# Guidance on how to contribute

Contributions to this code are welcome and appreciated.
Please adhere to our [Code of Conduct](./CODE_OF_CONDUCT.md) at all times.

> All contributions to this code will be released under the terms of the [LICENSE](./LICENSE)
> of this code. By submitting a pull request or filing a bug, issue, or feature request, you are
> agreeing to comply with this waiver of copyright interest. Details can be found in our
> [LICENSE](./LICENSE).

There are two primary ways to contribute:

1. Using the issue tracker
2. Changing the codebase

## Using the issue tracker

Use the issue tracker to suggest feature requests, report bugs, and ask questions. This is also
a great way to connect with the developers of the project as well as others who are interested
in this solution.

When reporting a bug, include the pane's **Technical log** (expand it at the bottom of the
pane), the Outlook client and version from its `Start` line, and – for automatic send – the
`AADSTS` or Graph error code. `docs/TROUBLESHOOTING.md` explains how to read the log.

Use the issue tracker to find ways to contribute. Find a bug or a feature, mention in the issue
that you will take on that effort, then follow the _Changing the codebase_ guidance below.

## Changing the codebase

Generally speaking, you should fork this repository, make changes in your own fork, and then
submit a pull request. All new code should have associated tests (if applicable) that validate
implemented features and the presence or lack of defects – for this project that means
extending `tests/smoke_test.py`.

Additionally, the code should follow the guidelines below. In the absence of a guideline,
mimic the styles and patterns in the existing codebase.

### Project guidelines

- No build step. `src/` is deployed as-is, so keep the JavaScript free of syntax that older
  Office webviews cannot parse (no optional chaining `?.`, no nullish coalescing `??`). `async`/`await`
  is fine.
- No user-facing text in `taskpane.js` or `config.js`. Every string the user sees comes from a
  locale file in `src/locales/` (see `docs/LOCALIZATION.md`). A new key must be added to **all**
  locale files; CI checks that the key sets match. Log messages are in English and stay in the code.
- Never change the Cisco submission addresses in `CATEGORIES`. They are Cisco's documented
  addresses; changing them silently breaks reporting.
- Keep the send-mode logic intact: graph by default, compose only as the fallback or when
  configured. A change that makes reporting *possible* in fewer situations is a regression.
- Every new source file (JS, CSS, HTML, XML, Python, shell, PowerShell) starts with the Cisco
  Sample Code License header – copy it from an existing file of the same type.

### Before you open a pull request

```bash
npm run check      # node --check on the JS files
npm run validate   # Microsoft's manifest validator (needs network)
npm test           # headless smoke test (pip install playwright && playwright install chromium)
npm run docs       # regenerate docs/index.html if you changed anything under docs/
```

Bump the version in `package.json`, `VERSION` in `src/taskpane.js` and `<Version>` in
`src/manifest.xml`, and add a line to `CHANGELOG.md`.

### Testing in a real Outlook

The smoke test proves the pane renders, the language handling works and the compose flow is
wired correctly. Anything that touches Office.js or Microsoft Graph for real must be tested by
sideloading (`docs/DEPLOYMENT.md`, step 4) in at least Outlook on the web and one desktop client.
