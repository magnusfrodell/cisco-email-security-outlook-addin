# Configuration reference

All configuration is in `src/config.js`, a plain JavaScript file that assigns
`window.RAPPORT_CONFIG`. It is read when the pane opens; changes on the web host take effect
the next time a user opens the pane, with no manifest update.

## Sending

| Key                          | Type      | Default                                            | Description                                                                                                                                                  |
| ---------------------------- | --------- | -------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `CLIENT_ID`                  | string    | `""`                                               | Entra *Application (client) ID*. Empty disables graph mode; the add-in then always opens a new message (compose mode).                                        |
| `AUTHORITY`                  | string    | `https://login.microsoftonline.com/organizations`  | MSAL authority. Use `https://login.microsoftonline.com/<tenant-id>` for a single-tenant registration.                                                        |
| `MOVE_AFTER_REPORT`          | boolean   | `true`                                             | After a successful graph-mode report, move the message to the category's `moveTo` folder. Requires the delegated `Mail.ReadWrite` permission.                 |
| `SAVE_TO_SENT_DEFAULT`       | boolean   | `false`                                            | Default for the per-user "keep a copy in Sent Items" checkbox (graph mode). The user's choice is stored in `roamingSettings` and overrides this.              |
| `CC_ADDRESSES`               | string[]  | `[]`                                               | Extra recipients on every report, e.g. an internal SOC mailbox. Applied in both modes.                                                                       |
| `SUBJECT`                    | string    | `Cisco Secure Email submission – {category}`       | Subject of the report mail. `{category}` is replaced with the category label.                                                                                |
| `BODY_TEXT`                  | string    | Danish sentence                                    | Plain-text body of the report mail (rendered as a paragraph in compose mode).                                                                                |
| `MAX_GRAPH_ATTACHMENT_BYTES` | number    | `3145728` (3 MB)                                   | Above this EML size graph mode hands over to compose mode. Microsoft Graph rejects larger inline attachments in `sendMail`; Cisco accepts at most 10 MB.      |

## Categories

`CATEGORIES` is an array; order is display order within each group.

```javascript
{
  id: "phish",                 // unique key; used in logs and data-id attributes
  group: "missed",             // key into GROUPS → which panel the button is in
  enabled: true,               // false hides the button without deleting the entry
  label: "Phishing",           // button text; also substituted into SUBJECT
  hint: "Forsøg på at …",      // one line under the label
  address: "phish@access.ironport.com",  // Cisco address – do not change
  moveTo: "junkemail"          // "junkemail" | "inbox" | null (graph mode only)
}
```

`GROUPS` maps group keys to the panel headings:

```javascript
GROUPS: {
  missed: "Fejlagtigt leveret til indbakken",
  false_positive: "Fejlagtigt stoppet eller markeret"
}
```

The pane has two group containers (`missed`, `false_positive`). A category with an unknown
group lands in `missed`.

### Cisco addresses

These are Cisco's documented direct-submission addresses (Cisco doc 214133). Do not change
them; Talos will not see reports sent anywhere else.

| Purpose                                  | Address                        |
| ---------------------------------------- | ------------------------------ |
| Missed spam                              | `spam@access.ironport.com`     |
| Missed phishing                          | `phish@access.ironport.com`    |
| Missed virus / malicious attachment      | `virus@access.ironport.com`    |
| Missed marketing (graymail)              | `ads@access.ironport.com`      |
| False positive – legitimate mail         | `ham@access.ironport.com`      |
| False positive – not marketing           | `not_ads@access.ironport.com`  |

Cisco's own add-in uses three internal addresses (`addin_spam@`, `addin_ham@`, `addin_mktg@`).
Those are reserved for Cisco's add-in; this project uses the public ones.

## Strings

`LANG` (BCP-47 language tag, used for `<html lang>`) and `STRINGS` (every text in the pane).
See `docs/LOCALIZATION.md` for the full key list and the placeholders each string supports.

## Recipes

**Only spam, phishing and "not spam" (three buttons like Cisco's add-in):** set `enabled: false`
on `virus`, `ads` and `not_ads`.

**Copy every report to the SOC:**

```javascript
CC_ADDRESSES: ["soc@firma.dk"],
```

**Automatic mode, single tenant:**

```javascript
CLIENT_ID: "3f1c1d6e-…",
AUTHORITY: "https://login.microsoftonline.com/7a2b…-tenant-id",
MOVE_AFTER_REPORT: true,
```

**Automatic mode without moving messages (only `Mail.Send` granted):**

```javascript
MOVE_AFTER_REPORT: false,
```

**Keep a copy of reports by default:** `SAVE_TO_SENT_DEFAULT: true` – users can still turn it
off in the pane.
