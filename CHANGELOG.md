# Changelog

All notable changes to this project are documented here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses
[Semantic Versioning](https://semver.org/). The add-in version is kept in three places that must
match: `package.json`, `VERSION` in `src/taskpane.js`, and `<Version>` in `src/manifest.xml`
(four-part form, e.g. `1.0.0.0`).

## [1.3.0] – 2026-09-17

### Added
- SOC visibility, configured by the administrator only: `SOC_ADDRESSES` copies every report to
  the SOC (Cc, or Bcc with `SOC_COPY_MODE: "bcc"`) in both send modes; `SOC_SHOW_IN_STATUS`
  optionally tells the user.
- `SEND_AS`: automatic reports can be sent from a shared (SOC) mailbox (`from` set on
  `/me/sendMail` with `Mail.Send.Shared`), so the Talos Email Status Portal tracks all
  submissions under one address. Requires Send As (or Send on Behalf) rights; the compose
  fallback still sends from the user.
- `REPORT_DETAILS`: a `--- Report details ---` block in the report body (category, Cisco
  address, reporter, sent-as mailbox, original subject/sender/date, Message-ID, client,
  add-in version) for the SOC copy and SIEM parsing.
- `SUBJECT` placeholders `{subject}`, `{sender}`, `{reporter}` in addition to `{category}`;
  the default subject now includes the original subject.
- Locale string `copySentTo`; smoke-test coverage for the SOC copy and the details block.

### Changed
- `CC_ADDRESSES` renamed to `SOC_ADDRESSES`; the old key is still honoured as an alias.

## [1.2.0] – 2026-09-17

### Changed
- Automatic send through Microsoft Graph is now the configured default (`SEND_MODE: "graph"`).
  Opening a new message is a fallback, not a mode of its own: it is used when automatic send
  is unavailable or fails, and the pane then shows a persistent warning explaining why
  (missing `CLIENT_ID`, unsupported client, component not loaded).
- `COMPOSE_FALLBACK` (default `true`) controls whether the fallback is allowed at all; with
  `false` the pane shows an error and never opens a new message.
- `SEND_MODE: "compose"` restores the previous always-new-message behaviour without warnings.
- Deployment guides now treat the Entra app registration as part of a standard installation
  and describe the shared multi-tenant registration model (one app, admin consent per tenant).

### Added
- Locale strings `chipUnavailable`, `autoUnavailable`, `autoUnavailableNoFallback`,
  `composeUnavailable`, `reasonNotConfigured`, `reasonClientUnsupported`, `reasonMsalMissing`.
- Smoke test coverage for the fallback warning.
- Repository documentation restructured to the Cisco DevNet Code Exchange template: README
  with use case, installation, configuration, usage, known issues, getting help/involved,
  credits and licensing sections; `CODE_OF_CONDUCT.md`; `NOTICE` (replaces
  `THIRD_PARTY_NOTICES.md`); `CONTRIBUTING.md` in the template's structure.

## [1.1.0] – 2026-09-17

### Changed
- Renamed to **Cisco Email Security Reporter** (repository
  `cisco-email-security-outlook-addin`). Manifest `DisplayName`, ribbon group
  (*Email security*) and button (*Report to Cisco*) renamed; config global is now
  `window.REPORTER_CONFIG`, log prefix `[CiscoReporter]`. The manifest `Id` is unchanged, so
  the update installs over 1.0.0.
- All user-facing text moved out of `config.js` into one locale file per language in
  `src/locales/` (`window.REPORTER_LOCALES.<code>`), including category labels/hints, group
  headings and the report body text. `config.js` now holds structure only.
- Manifest `DefaultLocale` is `en-US` with `da-DK` and `sv-SE` overrides for description, group
  label, button label and tooltip.
- `examples/config.sv-SE.js` removed – Swedish is built in.
- License changed from MIT to the Cisco Sample Code License, Version 1.1 (`LICENSE`); the
  license header added to every source file; bundled MSAL.js listed in
  `THIRD_PARTY_NOTICES.md`.

### Added
- English and Swedish locales alongside Danish.
- `LANGUAGE` (`"auto"` follows the Outlook display language, or a fixed code),
  `DEFAULT_LANGUAGE` and `SHOW_LANGUAGE_SELECTOR` in `config.js`.
- Language selector in the pane; the user's choice is stored in `roamingSettings` and
  overrides the configured language for that user. *Automatic* resets it.
- Trilingual `support.html`.
- Smoke test coverage for automatic language resolution and switching.

## [1.0.0] – 2026-09-17

### Added
- Danish task pane with six report categories mapped to Cisco's documented submission addresses
  (spam, phish, virus, ads, ham, not_ads @access.ironport.com).
- Automatic send mode: EML via `getAsFileAsync`, Microsoft Graph `sendMail` with Nested App
  Authentication, message moved to Junk/Inbox afterwards. Payload mirrors Cisco's own add-in
  (`orig_msg_*.raw.eml`, `message/rfc822`, `X-MS-AddIn-Base64Encode: true`).
- Compose fallback mode: pre-addressed new message with the reported item attached, no app
  registration required.
- Automatic fallback for messages larger than the Graph inline attachment limit and a
  "send manually instead" action on any failure.
- All user-facing strings in `config.js` (`STRINGS`, `CATEGORIES`, `GROUPS`); Swedish example
  in `examples/config.sv-SE.js`.
- Per-user "keep a copy in Sent Items" setting stored in `roamingSettings`.
- In-pane technical log and console logging with the `[CiscoRapport]` prefix.
- Headless smoke test (`tests/smoke_test.py`) with an Office.js stub; GitHub Actions CI
  (manifest validation, syntax check, smoke test) and optional GitHub Pages deployment.
- Documentation: architecture and code reference, configuration, deployment (EN), installation
  guide (DA), localization, troubleshooting; single-file HTML build in `docs/index.html`.
