# Configuration reference

All configuration is in `src/config.js`, a plain JavaScript file that assigns
`window.REPORTER_CONFIG`. It is read when the pane opens; changes on the web host take effect
the next time a user opens the pane, with no manifest update. User-facing texts are not in this
file – they live in `src/locales/<lang>.js` (see `docs/LOCALIZATION.md`).

## Language

| Key                       | Type    | Default  | Description                                                                                                                                   |
| ------------------------- | ------- | -------- | --------------------------------------------------------------------------------------------------------------------------------------------- |
| `LANGUAGE`                | string  | `"auto"` | `"auto"` follows the user's Outlook display language. Any locale code with a file in `src/locales/` (`"da"`, `"en"`, `"sv"`) forces that language. |
| `DEFAULT_LANGUAGE`        | string  | `"en"`   | Used when `auto` cannot match the Outlook language to a locale file.                                                                          |
| `SHOW_LANGUAGE_SELECTOR`  | boolean | `true`   | Shows the language selector in the pane. A user's choice is stored in their mailbox and overrides `LANGUAGE` for that user.                    |

## Sending

| Key                          | Type      | Default                                            | Description                                                                                                                                                  |
| ---------------------------- | --------- | -------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `SEND_MODE`                  | string    | `"graph"`                                          | `"graph"`: send automatically through Microsoft Graph (needs `CLIENT_ID`). `"compose"`: always open a new message with the mail attached; no app registration, no warnings. |
| `COMPOSE_FALLBACK`           | boolean   | `true`                                             | In graph mode: when automatic send is unavailable (missing `CLIENT_ID`, unsupported client) or fails, fall back to a new message and show a warning. `false`: show an error and never open a new message. |
| `CLIENT_ID`                  | string    | `""`                                               | Entra *Application (client) ID* used by graph mode. Empty with `SEND_MODE: "graph"` → persistent warning in the pane and the compose fallback (if allowed).    |
| `AUTHORITY`                  | string    | `https://login.microsoftonline.com/organizations`  | MSAL authority. Use `https://login.microsoftonline.com/<tenant-id>` for a single-tenant registration.                                                        |
| `MOVE_AFTER_REPORT`          | boolean   | `true`                                             | After a successful graph-mode report, move the message to the category's `moveTo` folder. Requires the delegated `Mail.ReadWrite` permission.                 |
| `SAVE_TO_SENT_DEFAULT`       | boolean   | `false`                                            | Default for the per-user "keep a copy in Sent Items" checkbox (graph mode). The user's choice is stored in `roamingSettings` and overrides this.              |
| `SUBJECT`                    | string    | `Cisco Secure Email submission – {category}: {subject}` | Subject of the report mail. Placeholders: `{category}` (label in the active language), `{subject}` (original subject, shortened to 120 characters), `{sender}` (original sender address), `{reporter}` (reporting user). |
| `REPORT_DETAILS`             | boolean   | `true`                                             | Append a `--- Report details ---` block to the body: category, Cisco address, reporter, sent-as mailbox, original subject/sender/date, Message-ID, client, add-in version. Cisco ignores the body; the SOC copy and SIEMs can parse it. |
| `MAX_GRAPH_ATTACHMENT_BYTES` | number    | `3145728` (3 MB)                                   | Above this EML size graph mode hands over to compose mode. Microsoft Graph rejects larger inline attachments in `sendMail`; Cisco accepts at most 10 MB.      |

The body text of the report mail is per language (`bodyText` in each locale file); the details
block that follows it is always English so it stays machine-readable.

## SOC visibility

Set by the administrator in `config.js`; nothing here is visible or changeable in the pane.

| Key                  | Type     | Default | Description                                                                                                                                                                                                 |
| -------------------- | -------- | ------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `SOC_ADDRESSES`      | string[] | `[]`    | Every report is also delivered to these addresses – a SOC mailbox, a ticketing address, a SIEM ingestion address. Applies in both send modes. (`CC_ADDRESSES` from 1.0–1.2 is still honoured as an alias.) |
| `SOC_COPY_MODE`      | string   | `"cc"`  | `"cc"`: the SOC address is visible to Cisco on the report. `"bcc"`: the copy is invisible to Cisco.                                                                                                         |
| `SOC_SHOW_IN_STATUS` | boolean  | `false` | Append "A copy was sent to …" to the status text so users know the SOC sees their reports.                                                                                                                  |
| `SEND_AS`            | string   | `""`    | Automatic send only: send every report **from** this shared mailbox (e.g. the SOC mailbox) instead of the user's own, so all submissions are tracked under one address in the Talos Email Status Portal. Needs *Send As* rights for each user on the mailbox (Send on Behalf also works but the mail is marked "on behalf of") and the delegated Graph permission `Mail.Send.Shared` with admin consent; no Full Access is needed. The compose fallback always sends from the user. |

What the SOC receives is the same mail Cisco receives: subject with category and original
subject, the details block, and the original message as a `.raw.eml` attachment – enough to
open the original, correlate it with mail-flow logs by Message-ID, and follow the case in Talos.


## Categories

`CATEGORIES` is an array; order is display order within each group. Labels and hints come from
the locale files, keyed by `id`.

```javascript
{
  id: "phish",                 // key into <locale>.categories; also used in logs
  group: "missed",             // "missed" | "false_positive" → which panel the button is in
  enabled: true,               // false hides the button without deleting the entry
  address: "phish@access.ironport.com",  // Cisco address – do not change
  moveTo: "junkemail"          // "junkemail" | "inbox" | null (graph mode only)
}
```

The pane has two group containers (`missed`, `false_positive`) whose headings come from
`<locale>.groups`. A category with an unknown group lands in `missed`.

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

## Recipes

**Danish for everyone, no selector:**

```javascript
LANGUAGE: "da",
SHOW_LANGUAGE_SELECTOR: false,
```

**Follow Outlook, fall back to Swedish:**

```javascript
LANGUAGE: "auto",
DEFAULT_LANGUAGE: "sv",
```

**Only spam, phishing and "not spam" (three buttons like Cisco's add-in):** set `enabled: false`
on `virus`, `ads` and `not_ads`.

**Copy every report to the SOC, invisibly to Cisco, and tell the user:**

```javascript
SOC_ADDRESSES: ["soc@contoso.com"],
SOC_COPY_MODE: "bcc",
SOC_SHOW_IN_STATUS: true,
```

**Submit everything from the SOC mailbox so Talos shows one submitter for the organisation:**

```javascript
SEND_AS: "soc@contoso.com",        // shared mailbox; users need Send As; app needs Mail.Send.Shared
SOC_ADDRESSES: ["soc@contoso.com"],  // optional – the SOC also gets its own copy
```

**Shared multi-tenant app (one registration for every customer, as Cisco's add-in does):**

```javascript
SEND_MODE: "graph",
CLIENT_ID: "<client ID of the shared app>",
AUTHORITY: "https://login.microsoftonline.com/organizations",
```

Each customer tenant then only grants admin consent (`docs/DEPLOYMENT.md`, step 2).

**Zero infrastructure (no Entra app), no warnings:**

```javascript
SEND_MODE: "compose",
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

**Add a category** (e.g. a second phishing address for an internal team): add an entry to
`CATEGORIES` with a new `id`, then add `categories.<id>` with `label` and `hint` to every locale
file.
