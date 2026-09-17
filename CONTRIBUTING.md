# Contributing

## Ground rules

- No build step. `src/` is deployed as-is, so keep the JavaScript free of syntax that older
  Office webviews cannot parse (no optional chaining `?.`, no nullish coalescing `??`). `async`/`await`
  is fine.
- No user-facing text in `taskpane.js` or `config.js`. Every string the user sees comes from a
  locale file in `src/locales/` (see `docs/LOCALIZATION.md`). A new key must be added to **all**
  locale files. Log messages are in English and stay in the code.
- Never change the Cisco submission addresses in `CATEGORIES`. They are Cisco's documented
  addresses; changing them silently breaks reporting.
- Keep the existing fallback chain intact: graph → compose on preconditions, size or failure.
  A change that makes reporting *possible* in fewer situations is a regression.

## Before you open a pull request

```bash
npm run check      # node --check on the JS files
npm run validate   # Microsoft's manifest validator (needs network)
npm test           # headless smoke test (pip install playwright && playwright install chromium)
npm run docs       # regenerate docs/index.html if you changed anything under docs/
```

Bump the version in `package.json`, `VERSION` in `src/taskpane.js` and `<Version>` in
`src/manifest.xml`, and add a line to `CHANGELOG.md`.

## Testing in a real Outlook

The smoke test only proves the pane renders and the compose flow is wired correctly. Anything
that touches Office.js or Graph for real must be tested by sideloading
(`docs/DEPLOYMENT.md`, step 4) in at least Outlook on the web and one desktop client.
