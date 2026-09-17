# Localization

The add-in ships with three languages – **English (`en`)**, **Danish (`da`)** and
**Swedish (`sv`)** – as one file each in `src/locales/`. The JavaScript in `taskpane.js`
contains no user-facing text.

## How the language is chosen

`resolveLanguage()` in `taskpane.js` picks a locale in this order and logs the result
(`Language: sv {"source":"…"}`):

1. **The user's own choice** in the pane's language selector, stored in the user's mailbox
   (`roamingSettings`, key `language`). Only if `SHOW_LANGUAGE_SELECTOR` is not `false`.
2. **`LANGUAGE` in `config.js`**, if it is not `"auto"`.
3. **The Outlook display language** (`Office.context.displayLanguage`, e.g. `sv-SE`), matched
   first exactly and then by its base language (`sv`).
4. **`DEFAULT_LANGUAGE`**, and failing that the first locale file loaded.

Selecting *Automatic* in the pane stores `"auto"` and returns the user to steps 2–4.

```javascript
LANGUAGE: "auto",              // or "da" | "en" | "sv"
DEFAULT_LANGUAGE: "en",
SHOW_LANGUAGE_SELECTOR: true,
```

The ribbon button, group label and store description are localized separately by Outlook, from
the `Override` elements in the manifest (section 3). Both mechanisms key off the same Outlook
display language, so with `LANGUAGE: "auto"` the ribbon and the pane agree.

## 1. Locale files

Each file assigns one property on `window.REPORTER_LOCALES`:

```javascript
window.REPORTER_LOCALES = window.REPORTER_LOCALES || {};
window.REPORTER_LOCALES.sv = {
  name: "Svenska",                     // shown in the language selector
  bodyText: "…",                       // body of the report mail
  groups: { missed: "…", false_positive: "…" },
  categories: {                        // keyed by the ids in config.js CATEGORIES
    spam:    { label: "Skräppost", hint: "…" },
    …
  },
  strings: { title: "…", … }           // everything else (table below)
};
```

Missing pieces degrade gracefully: a missing `strings` key shows the key itself
(e.g. `chipAuto`), a missing category shows the category id, a missing group shows the group
key. The CI syntax-checks every locale file; the smoke test checks that all three switch the
pane correctly.

### Adding a language

1. Copy `src/locales/en.js` to `src/locales/<code>.js` (`nb`, `fi`, `de`, …) and translate
   everything except the keys.
2. Change `REPORTER_LOCALES.en` to `REPORTER_LOCALES.<code>` at the top of the new file.
3. Add `<script src="locales/<code>.js"></script>` next to the other locale scripts at the bottom
   of `src/taskpane.html`.
4. Optionally add `Override` elements for the new locale in `src/manifest.xml` (section 3).
5. Run `node --check src/locales/<code>.js` and `npm test`.

The new language appears in the selector automatically and is matched by `auto` when the
user's Outlook language starts with `<code>`.

### Removing a language

Delete the `<script>` line from `taskpane.html` (and the file). If `DEFAULT_LANGUAGE` pointed at
it, change it.

### `strings` keys

Placeholders in braces are replaced at runtime; keep them in the translation.

| Key                  | Placeholders            | Where it appears                                          |
| -------------------- | ----------------------- | --------------------------------------------------------- |
| `title`              |                         | Pane header and browser title                             |
| `chipStarting`       |                         | Header chip before Office.js is ready                     |
| `chipAuto`           |                         | Header chip in graph mode                                 |
| `chipCompose`        |                         | Header chip in compose mode                               |
| `selectedMail`       |                         | Heading of the selected-mail panel                        |
| `noMailSelected`     |                         | Subject line when nothing is selected                     |
| `selectMailHint`     |                         | Sender line when nothing is selected                      |
| `noSubject`          |                         | Subject line for a mail without subject                   |
| `fromPrefix`         |                         | Prefix before the sender ("From: ")                       |
| `missedHelp`         |                         | Help text under the "missed" group heading                |
| `falsePositiveHelp`  |                         | Help text under the "false positive" group heading        |
| `sentTo`             | `{address}`             | Tooltip on each category button                           |
| `settings`           |                         | Settings panel heading                                    |
| `keepCopy`           |                         | Checkbox label                                            |
| `language`           |                         | Label of the language selector                            |
| `languageAuto`       |                         | First option of the language selector                     |
| `logSummary`         |                         | Collapsed log heading                                     |
| `footer`             |                         | Footer text before the portal link                        |
| `footerLink`         |                         | Portal link text                                          |
| `chipOk` `chipWarn` `chipErr` `chipWait` `chipInfo` | | Status chips                                  |
| `chipPressSend`      |                         | Status chip after a compose-mode report                   |
| `copySentTo`         | `{address}`             | Appended to the status when `SOC_SHOW_IN_STATUS` is on     |
| `chipUnavailable`    |                         | Header chip when neither send mode is available            |
| `autoUnavailable`    | `{reason}`              | Persistent notice: automatic send unavailable, fallback active |
| `autoUnavailableNoFallback` | `{reason}`       | Persistent notice: automatic send unavailable, no fallback |
| `composeUnavailable` |                         | Persistent notice: `SEND_MODE: "compose"` on a client without Mailbox 1.6 |
| `reasonNotConfigured` `reasonClientUnsupported` `reasonMsalMissing` | | Text substituted into `{reason}` |
| `sending`            | `{category}`            | Status while a report is in progress                      |
| `composeOpened`      | `{address}`             | Status after the new-message form opened                  |
| `reported`           | `{category}` `{address}` | Status after a graph-mode report                         |
| `movedTo`            | `{folder}`              | Appended to `reported` when the message was moved         |
| `folderJunk` `folderInbox` |                   | Folder names used in `{folder}`                           |
| `selectFirst`        |                         | Warning when a button is pressed with nothing selected    |
| `notSentAuto`        |                         | Prefix of the error status in graph mode                  |
| `notSent`            |                         | Prefix of the error status in compose mode                |
| `retryCompose`       |                         | Label of the "send manually instead" button               |
| `tooLarge`           | `{mb}`                  | Reason appended when the EML exceeds the Graph limit      |
| `manualChosen`       |                         | Reason appended after "send manually instead"             |
| `reportedMailName`   |                         | Attachment name when the mail has no subject              |
| `errConsent`         |                         | Consent / AADSTS65001 errors                               |
| `errGraphDenied`     | `{status}`              | Graph 401 / 403                                            |
| `errGetFile`         |                         | `getAsFileAsync` failure                                   |
| `errNotOutlook`      |                         | Pane opened outside Outlook                                |
| `errStart`           |                         | Prefix for unexpected startup errors                      |

`SUBJECT` in `config.js` is deliberately not per language: it is read by Cisco, not by the user.
`{category}` in it is replaced with the label in the active language.

## 2. Ribbon and store card (`src/manifest.xml`)

Outlook localizes these itself from the manifest. `DefaultLocale` is `en-US`; each string has
`Override` children for `da-DK` and `sv-SE`, and Outlook picks the one matching the user's
Office display language:

| Resource                    | en-US (default)                                  | da-DK                                    | sv-SE                                      |
| --------------------------- | ------------------------------------------------ | ---------------------------------------- | ------------------------------------------ |
| `DisplayName`               | Cisco Email Security Reporter                    | (same)                                   | (same)                                     |
| `Description`               | Report spam, phishing, virus and marketing mail … | Rapportér spam, phishing, virus og … | Rapportera skräppost, nätfiske, virus … |
| `GroupLabel`                | Email security                                   | Mailsikkerhed                            | E-postsäkerhet                             |
| `TaskpaneButton.Label`      | Report to Cisco                                  | Rapportér til Cisco                      | Rapportera till Cisco                      |
| `TaskpaneButton.Tooltip`    | Report spam, phishing, virus or marketing mail … | Rapportér spam, phishing, virus eller … | Rapportera skräppost, nätfiske, virus … |

To add a locale, add an `Override` to each of the four localized strings:

```xml
<bt:String id="TaskpaneButton.Label" DefaultValue="Report to Cisco">
  <bt:Override Locale="da-DK" Value="Rapportér til Cisco"/>
  <bt:Override Locale="sv-SE" Value="Rapportera till Cisco"/>
  <bt:Override Locale="nb-NO" Value="Rapporter til Cisco"/>
</bt:String>
```

Locale codes must be full `language-REGION` tags. Keep `ShortStrings` under 125 characters and
`LongStrings` under 250; the validator enforces this. `DisplayName` is kept identical in all
languages on purpose (it is the product name), but it accepts overrides in the same way.

Manifest strings only change when the manifest is re-uploaded (bump `<Version>`); pane strings
change as soon as the locale file on the web host changes.

## 3. `src/support.html`

A static page with one section per language. Add a section for a new language if you want.

## What not to translate

- Category `id`s, `address`es, `moveTo` values and the `groups` keys (`missed`,
  `false_positive`).
- Log messages in `taskpane.js` – they are for support staff and are grepped in troubleshooting.
- `SUBJECT` in `config.js` (read by Cisco).
