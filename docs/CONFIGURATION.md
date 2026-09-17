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
| `CC_ADDRESSES`               | string[]  | `[]`                                               | Extra recipients on every report, e.g. an internal SOC mailbox. Applied in both modes.                                                                       |
| `SUBJECT`                    | string    | `Cisco Secure Email submission – {category}`       | Subject of the report mail. `{category}` is replaced with the category label in the active language.                                                         |
| `MAX_GRAPH_ATTACHMENT_BYTES` | number    | `3145728` (3 MB)                                   | Above this EML size graph mode hands over to compose mode. Microsoft Graph rejects larger inline attachments in `sendMail`; Cisco accepts at most 10 MB.      |

The body text of the report mail is per language (`bodyText` in each locale file).

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

**Copy every report to the SOC:**

```javascript
CC_ADDRESSES: ["soc@contoso.com"],
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
