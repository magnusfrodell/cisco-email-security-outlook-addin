# Localization

The add-in ships in Danish. Translating it to another language touches two files and nothing in
`taskpane.js`.

## 1. `src/config.js`

Three parts carry user-facing text:

| Part          | What it holds                                              |
| ------------- | ---------------------------------------------------------- |
| `CATEGORIES`  | `label` and `hint` for each button (keep `id`, `address`, `moveTo`) |
| `GROUPS`      | The two panel headings                                     |
| `LANG`        | Language tag written to `<html lang>` (`da`, `sv`, `nb`, `fi`, …) |
| `STRINGS`     | Everything else the user can see                           |
| `BODY_TEXT`   | Body of the report mail                                    |

A complete Swedish example is in `examples/config.sv-SE.js`. Copy it over `src/config.js`, then
re-apply your settings (`CLIENT_ID`, `CC_ADDRESSES`, …).

### `STRINGS` keys

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
| `fromPrefix`         |                         | Prefix before the sender ("Fra: ")                        |
| `missedHelp`         |                         | Help text under the "missed" group heading                |
| `falsePositiveHelp`  |                         | Help text under the "false positive" group heading        |
| `sentTo`             | `{address}`             | Tooltip on each category button                           |
| `settings`           |                         | Settings panel heading                                    |
| `keepCopy`           |                         | Checkbox label                                            |
| `logSummary`         |                         | Collapsed log heading                                     |
| `footer`             |                         | Footer text before the portal link                        |
| `footerLink`         |                         | Portal link text                                          |
| `chipOk` `chipWarn` `chipErr` `chipWait` `chipInfo` | | Status chips                                  |
| `chipPressSend`      |                         | Status chip after a compose-mode report                   |
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

`t(key)` returns the key itself when a translation is missing, so a forgotten string shows up as
e.g. `chipAuto` in the UI rather than breaking anything.

## 2. `src/manifest.xml`

Five strings, all plain `DefaultValue` attributes:

| Element / resource id           | Danish value                                              |
| ------------------------------- | --------------------------------------------------------- |
| `DisplayName`                   | Rapportér mail til Cisco                                  |
| `Description`                   | Rapportér spam, phishing, virus og marketing til Cisco …  |
| `GroupLabel`                    | Mailsikkerhed                                             |
| `TaskpaneButton.Label`          | Rapportér mail                                            |
| `TaskpaneButton.Tooltip`        | Rapportér spam, phishing, virus eller marketing til Cisco … |

Also change `DefaultLocale` (e.g. `sv-SE`). Keep `ShortStrings` under 125 characters and
`LongStrings` under 250, which the validator enforces.

### One manifest, several languages

If one tenant has users with different Office display languages, keep `DefaultLocale` and add
`<Override Locale="…" Value="…"/>` children to each string and URL resource. Outlook picks the
override that matches the user's Office language and falls back to the default:

```xml
<bt:String id="TaskpaneButton.Label" DefaultValue="Rapportér mail">
  <bt:Override Locale="sv-SE" Value="Rapportera mejl"/>
  <bt:Override Locale="en-US" Value="Report mail"/>
</bt:String>
```

The pane itself does not switch language per user – it shows whatever `config.js` contains. To
serve different languages from one host, publish one copy of `src/` per language (e.g.
`/da/`, `/sv/`) and use `<bt:Url>` overrides in the same way:

```xml
<bt:Url id="Taskpane.Url" DefaultValue="https://addin.firma.dk/da/taskpane.html">
  <bt:Override Locale="sv-SE" Value="https://addin.firma.dk/sv/taskpane.html"/>
</bt:Url>
```

## 3. `src/support.html`

A static page; translate the text in place. It is only reached through the *Support* link in the
add-in store card.

## What not to translate

- Category `id`s, `address`es, `moveTo` values and `GROUPS` keys (`missed`, `false_positive`).
- Log messages in `taskpane.js` – they are for support staff and are grepped in troubleshooting.
- `SUBJECT` may be translated, but keep `{category}` if you want the category in the subject.
