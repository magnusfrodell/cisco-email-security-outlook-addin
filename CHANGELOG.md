# Changelog

All notable changes to this project are documented here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses
[Semantic Versioning](https://semver.org/). The add-in version is kept in three places that must
match: `package.json`, `VERSION` in `src/taskpane.js`, and `<Version>` in `src/manifest.xml`
(four-part form, e.g. `1.0.0.0`).

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
