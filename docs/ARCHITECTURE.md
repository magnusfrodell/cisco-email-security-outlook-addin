# Architecture and code reference

This document describes how the add-in is put together, how a report travels from Outlook to
Cisco Talos, and what every file and function is for. It is written from the source in `src/`;
if the two disagree, the source wins and this document needs an update.

## 1. Overview

The add-in is a classic Office web add-in: an XML manifest that adds a button to the message
read ribbon, and a static web page (HTML/CSS/JS) that Outlook loads in a task pane. There is no
backend and no build step. Everything under `src/` is deployed unchanged to any HTTPS web host.

```mermaid
flowchart LR
    subgraph Outlook["Outlook (Windows / Mac / web)"]
        R[Ribbon button<br/>Rapportér mail]
        TP[Task pane<br/>taskpane.html + js]
        OJS[Office.js]
    end
    subgraph Host["Static HTTPS host"]
        M[manifest.xml]
        F[taskpane.html/css/js<br/>config.js, lib/, assets/]
    end
    subgraph MS["Microsoft 365"]
        EX[Exchange Online]
        G[Microsoft Graph]
        ID[Entra ID]
    end
    T[Cisco Talos<br/>*@access.ironport.com]

    M -. deployed via<br/>Integrated apps .-> Outlook
    R --> TP
    TP --> OJS
    F --> TP
    OJS -- getAsFileAsync<br/>displayNewMessageForm --> EX
    TP -- NAA token --> ID
    TP -- sendMail / move --> G
    G --> EX
    EX -- report mail with .eml attached --> T
```

A report is always an **e-mail with the original message attached as `message/rfc822`**, sent
from the user's own mailbox to one of Cisco's submission addresses. That is exactly what Cisco's
own *Secure Email Submission* add-in does under the hood, and what Cisco's documentation asks
for when submitting manually. Talos therefore processes reports from this add-in like any other.

## 2. Send modes

`detectMode()` picks one of two modes when the pane opens and shows it in the chip in the header.

```mermaid
flowchart TD
    A[Pane opens] --> B{CLIENT_ID set?}
    B -- no --> C[compose]
    B -- yes --> D{Mailbox 1.14<br/>getAsFileAsync?}
    D -- no --> C
    D -- yes --> E{NestedAppAuth 1.1?}
    E -- no --> C
    E -- yes --> F{MSAL loaded?}
    F -- no --> C
    F -- yes --> G[graph]
    C --> H{Mailbox 1.6?}
    H -- no --> X[Pane shows error]
    H -- yes --> C2[compose ready]
```

| Mode        | What happens                                                                                                                                                    | Needs                                                                     |
| ----------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------- |
| **graph**   | EML fetched with `getAsFileAsync`, sent with Graph `POST /me/sendMail`, message moved with `POST /me/messages/{id}/move`. No user interaction beyond the click. | `CLIENT_ID`, Entra app registration, Mailbox 1.14, NestedAppAuth 1.1      |
| **compose** | `displayNewMessageForm(Async)` opens a new message to Cisco's address with the selected item attached; the user presses Send.                                   | Mailbox 1.6 (1.9 for the async variant with error reporting)               |

Runtime fallbacks from graph to compose:

- the EML is larger than `MAX_GRAPH_ATTACHMENT_BYTES` (default 3 MB, the Graph `sendMail`
  limit for inline attachments) → compose is used automatically and the status explains why;
- token acquisition, `getAsFileAsync` or the Graph call fails → the status shows the error and a
  **Send manually instead** button that runs the compose flow for the same category.

The move step never blocks a report: if `move` fails after `sendMail` succeeded, the report
counts as sent and the status omits the "moved to …" part.

### 2.1 Sequence: graph mode

```mermaid
sequenceDiagram
    actor U as User
    participant P as Task pane
    participant O as Office.js
    participant MS as MSAL (NAA)
    participant ID as Entra ID
    participant G as Graph
    participant T as Talos

    U->>P: click category (e.g. Phishing)
    P->>MS: acquireTokenSilent(Mail.Send, Mail.ReadWrite)
    MS->>ID: brokered by Outlook (brk-multihub)
    ID-->>MS: access token
    alt interaction required
        P->>MS: acquireTokenPopup(...)
        MS-->>P: token
    end
    P->>O: item.getAsFileAsync()
    O-->>P: base64 EML
    alt EML > 3 MB
        P->>O: displayNewMessageForm(...)  (compose fallback)
    else
        P->>G: POST /me/sendMail {attachment orig_msg_*.raw.eml}
        G-->>T: report mail delivered
        G-->>P: 202 Accepted
        P->>G: POST /me/messages/{restId}/move {junkemail | inbox}
        G-->>P: 201 Created
    end
    P-->>U: status "Rapporteret som phishing …"
```

### 2.2 Sequence: compose mode

```mermaid
sequenceDiagram
    actor U as User
    participant P as Task pane
    participant O as Office.js
    participant OL as Outlook
    participant T as Talos

    U->>P: click category
    P->>O: displayNewMessageFormAsync({to, cc, subject, htmlBody, attachments:[{type:item, itemId}]})
    O->>OL: open compose window
    OL-->>P: callback (Mailbox 1.9+) or nothing (1.6)
    P-->>U: status "Tryk Send …"
    U->>OL: Send
    OL-->>T: mail with original as message/rfc822
```

## 3. The report message

Both modes produce the same kind of message. In graph mode the payload is built explicitly:

```json
{
  "message": {
    "subject": "Cisco Secure Email submission – Phishing",
    "body": { "contentType": "Text", "content": "…BODY_TEXT…" },
    "toRecipients": [ { "emailAddress": { "address": "phish@access.ironport.com" } } ],
    "ccRecipients": [ … CC_ADDRESSES … ],
    "internetMessageHeaders": [ { "name": "X-MS-AddIn-Base64Encode", "value": "true" } ],
    "attachments": [ {
      "@odata.type": "#microsoft.graph.fileAttachment",
      "name": "orig_msg_k3f9a2qz.raw.eml",
      "contentType": "message/rfc822",
      "contentBytes": "<base64 EML from getAsFileAsync>"
    } ]
  },
  "saveToSentItems": false
}
```

The attachment name pattern, content type and the `X-MS-AddIn-Base64Encode` header are the
ones Cisco's add-in uses, so the message arriving at Talos is indistinguishable in shape.
`saveToSentItems` follows the user's "keep a copy" setting.

In compose mode Outlook builds the message itself from an *item attachment*; when the message
leaves Exchange the attached item is emitted as a `message/rfc822` MIME part, which is the format
Cisco documents for manual submissions ("forward as attachment").

| Category id | Address                       | Move to      |
| ----------- | ----------------------------- | ------------ |
| `spam`      | spam@access.ironport.com      | `junkemail`  |
| `phish`     | phish@access.ironport.com     | `junkemail`  |
| `virus`     | virus@access.ironport.com     | `junkemail`  |
| `ads`       | ads@access.ironport.com       | `junkemail`  |
| `ham`       | ham@access.ironport.com       | `inbox`      |
| `not_ads`   | not_ads@access.ironport.com   | `inbox`      |

## 4. Files

```text
src/
├── manifest.xml          Office add-in manifest (da-DK). Placeholder host addin.example.com.
├── taskpane.html         Pane markup. Every visible text node has data-str="<key>".
├── taskpane.css          Styling; no framework.
├── taskpane.js           All logic (this document, section 5).
├── config.js             Settings, categories and every user-facing string.
├── support.html          Static support page referenced by <SupportUrl>.
├── assets/icon-*.png     Ribbon and store icons, 16/32/64/80/128 px.
└── lib/msal-browser.min.js  MSAL.js 4.30.0 (LTS), self-hosted. Loaded always, used in graph mode.
tests/
├── office-stub.js        Fake Office object for browser-only testing.
└── smoke_test.py         Headless Chromium test (Playwright).
scripts/
├── set-host.sh / .ps1    Replace the placeholder host in the manifest.
└── build-docs.py         Build docs/index.html from the Markdown files.
```

Load order in `taskpane.html` matters: `office.js` in `<head>`, then `config.js` (defines
`window.RAPPORT_CONFIG`), then `lib/msal-browser.min.js` (defines the `msal` global), then
`taskpane.js`. `taskpane.js` reads the config once at load; changing `config.js` on the host
takes effect the next time a pane opens.

## 5. `taskpane.js` reference

The file is a single IIFE. Nothing is exported; everything is reachable only through the DOM
events it attaches. Global state lives in one object:

```javascript
var state = {
  mode: "compose",   // "graph" | "compose", decided by detectMode()
  msalApp: null,     // lazily created NestablePublicClientApplication
  busy: false,       // a report is in progress; buttons are disabled
  item: null         // Office.context.mailbox.item at last refresh, or null
};
```

### 5.1 Strings

| Function            | Purpose                                                                                                                                  |
| ------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| `t(key, params)`    | Returns `STRINGS[key]` with `{name}` placeholders replaced from `params`. Unknown keys return the key itself, so a missing translation is visible rather than fatal. |
| `applyStrings()`    | Fills every `[data-str]` element from `STRINGS`, sets `document.title` and `<html lang>` from `LANG`. Runs once in `init()`.               |

### 5.2 Logging

| Function              | Purpose                                                                                                              |
| --------------------- | -------------------------------------------------------------------------------------------------------------------- |
| `log(level, msg, extra)` | Writes `HH:MM:SS LEVEL msg extra` to the in-pane `<pre id="log">` and to the console with the `[CiscoRapport]` prefix. `extra` is JSON-serialised. Levels: `info`, `warn`, `error`. |
| `errText(e)`          | Normalises an `Error`, MSAL error or string to `Name: code: message`.                                                |

Log messages are English and technical; they are meant for support staff, not end users.
User-facing status text always goes through `t()`.

### 5.3 UI

| Function                 | Purpose                                                                                                              |
| ------------------------ | -------------------------------------------------------------------------------------------------------------------- |
| `setModeChip(text, cls)` | Header chip showing the active send mode.                                                                            |
| `setStatus(kind, text, opts)` | Shows the status box (`ok`, `warn`, `err`, `wait`, `neutral`), optional chip override (`opts.chip`) and optional **Send manually instead** button (`opts.retryCompose = category`). Scrolls the box into view (respects `prefers-reduced-motion`). |
| `clearStatus()`          | Hides the status box; called when the selected item changes.                                                         |
| `setButtonsEnabled(on)`  | Enables/disables all category buttons.                                                                               |
| `renderCategories()`     | Builds one `<button class="btn-cat" data-id=…>` per enabled category into its group container and sets group titles. |
| `refreshItem()`          | Reads `Office.context.mailbox.item`, shows subject/sender or the "no mail selected" state. Bound to `ItemChanged` for pinned panes. |

### 5.4 Settings

`keepCopy()` reads the per-user `keepCopy` value from `Office.context.roamingSettings`, falling
back to `SAVE_TO_SENT_DEFAULT`. `initSettings()` wires the checkbox and persists changes with
`roamingSettings.saveAsync`. The settings panel is hidden in compose mode because Outlook's own
Send decides what goes to Sent Items there.

### 5.5 Sending

| Function                     | Purpose                                                                                                                                                                                                                                                              |
| ---------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `reportViaCompose(cat, reason)` | Builds the `displayNewMessageForm` parameters (to, cc, subject, HTML body, one `item` attachment with the current `itemId`). Uses the async variant when Mailbox 1.9 is available so failures surface; otherwise the sync call. Returns a promise.                       |
| `getEmlBase64()`             | Promise wrapper around `item.getAsFileAsync`. Resolves with the base64 EML.                                                                                                                                                                                          |
| `getToken(scopes)`           | Creates the MSAL nestable public client on first use (`clientId`, `authority`, `localStorage` cache), then `acquireTokenSilent`; on `InteractionRequiredAuthError` falls back to `acquireTokenPopup`. Any other error propagates.                                    |
| `graphPost(path, token, body)` | `fetch` POST to `https://graph.microsoft.com/v1.0` with bearer token. Non-2xx responses become an `Error` with `.status` and the first 300 characters of the body.                                                                                                     |
| `moveItem(token, folder)`    | Converts the EWS item id with `convertToRestId(…, v2_0)` and posts to `/me/messages/{id}/move` with `destinationId` `junkemail` or `inbox`.                                                                                                                           |
| `reportViaGraph(cat)`        | Orchestrates graph mode: scopes → token → EML → size check (returns `{fallback:true, reason}` if too large) → `sendMail` → optional move → success status. Returns `{fallback:false}` on success.                                                                     |

### 5.6 Orchestration and errors

| Function            | Purpose                                                                                                                                                                                                                                                   |
| ------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `report(cat)`       | Entry point from a category button. Guards against re-entry (`state.busy`), shows the "sending" status, dispatches to the active mode, handles the graph→compose fallback, and always re-enables the buttons in `finally`.                                  |
| `runGuarded(fn)`    | Same busy/finally discipline for secondary actions (the **Send manually instead** button).                                                                                                                                                                 |
| `friendlyError(e)`  | Maps common failures to user text: consent/`AADSTS65001` → `errConsent`; Graph 401/403 → `errGraphDenied`; `getAsFileAsync` failures → `errGetFile`; anything else → the raw `errText`.                                                                    |

Error strategy in one sentence: every failure ends in a visible status with a way forward,
and the technical detail is in the log, never lost.

### 5.7 Startup

`Office.onReady` → host check → `init()`:

1. log diagnostics (add-in version, host, platform, Office version);
2. `applyStrings()`;
3. `renderCategories()`;
4. `initSettings()`;
5. `detectMode()` (logs the reasons if graph mode is not available);
6. `refreshItem()`;
7. register the `ItemChanged` handler (Mailbox 1.5+) so a pinned pane follows the selection.

## 6. `manifest.xml` reference

| Element                                   | Value / purpose                                                                                                  |
| ----------------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| `Id`                                      | `8cec0077-122f-4b5b-a7ec-b779f92668c7`. Generate a new GUID if you fork the add-in for another organisation.     |
| `Version`                                 | Four-part, e.g. `1.0.0.0`. Must be bumped for the admin center to accept an update.                              |
| `DefaultLocale`                           | `da-DK`. All strings are plain `DefaultValue`s; see `docs/LOCALIZATION.md` for per-locale overrides.               |
| `AppDomains`                              | The add-in host and `https://login.microsoftonline.com` (interactive MSAL fallback).                             |
| `Requirements` / `bt:Sets`                | Mailbox 1.6 minimum – `displayNewMessageForm` is the floor. Graph mode is detected at runtime, not in the manifest. |
| `FormSettings` → `SourceLocation`         | `taskpane.html`, used by clients that ignore `VersionOverrides` (none in practice, but required by the schema).  |
| `Permissions`                             | `ReadWriteItem`.                                                                                                 |
| `Rule`                                    | `ItemIs Message Read` – the button only appears when reading a message.                                          |
| `VersionOverrides` 1.1 → `MessageReadCommandSurface` | One group (`Mailsikkerhed`) with one `ShowTaskpane` button (`Rapportér mail`), `SupportsPinning` true.  |
| `Resources`                               | Icons 16/32/80, task pane URL, group label, button label, tooltip.                                               |

The manifest is validated in CI with Microsoft's `office-addin-manifest validate`, which also
reports the supported platforms (Outlook on Windows, Mac and web).

## 7. Client support

| Client                                | compose mode | graph mode                                     |
| ------------------------------------- | ------------ | ---------------------------------------------- |
| Outlook on the web                    | yes          | yes                                            |
| New Outlook on Windows                | yes          | yes                                            |
| Classic Outlook on Windows (M365)     | yes          | yes on current builds (Mailbox 1.14 + NAA)     |
| Outlook on Mac (M365)                 | yes          | yes on current builds                          |
| Outlook on iOS / Android              | no – `displayNewMessageForm` and `getAsFileAsync` are not available on mobile | no |
| Perpetual / volume-licensed Outlook   | compose only, if Mailbox 1.6 is met            | no                                             |

Requirement sets used: Mailbox 1.6 (`displayNewMessageForm`), 1.5 (`ItemChanged`, pinning),
1.9 (`displayNewMessageFormAsync`), 1.14 (`getAsFileAsync`), NestedAppAuth 1.1 (MSAL brokering).

## 8. Testing

`tests/office-stub.js` provides just enough of the `Office` global to render the pane in a
normal browser: requirement-set answers (Mailbox 1.15, no NestedAppAuth → compose mode), a fake
selected message, `roamingSettings`, `displayNewMessageForm(Async)` that records its arguments,
and an `ItemChanged` hook.

`tests/smoke_test.py` copies `src/` to a temp directory, swaps the hosted `office.js` for the
stub, and drives the pane with Playwright: strings from config, one button per enabled category,
compose flow addressed to the right Cisco address with an item attachment, status rendering,
button enable/disable on selection change, and zero JavaScript errors. It can also write the
screenshots used in the documentation (`--screenshots docs/images`).

CI (`.github/workflows/ci.yml`) runs the syntax check, the manifest validator, a version
consistency check and the smoke test on every push and pull request.

Graph mode cannot be exercised without a real Outlook and tenant; see
`docs/DEPLOYMENT.md` step 4 for the sideload procedure.

## 9. Design decisions

- **No backend.** Cisco's add-in also sends from the user's mailbox; adding a relay would add a
  system to secure and a place for reports to get stuck, with no gain.
- **Mirror Cisco's payload.** Same attachment name pattern, content type and custom header, so
  Talos sees the format it already accepts.
- **Compose mode as a first-class citizen.** It needs no app registration and no consent, and it
  works on every supported client. Organisations that want one-click reporting add the Entra app
  later without redeploying the manifest.
- **Self-hosted MSAL.** No CDN dependency, no CSP surprises, and the version is pinned by
  what is in the repository.
- **Strings in `config.js`.** One file to translate; the JavaScript stays language-neutral.
