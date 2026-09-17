# Troubleshooting

## Start with the log

Every pane has a collapsed **Teknisk log** at the bottom. It is the first thing to ask a user
for. The same lines go to the browser console with the prefix `[CiscoRapport]`
(Outlook desktop: right-click in the pane → *Inspect*; Outlook on the web: F12).

A healthy graph-mode report looks like this:

```text
13:26:06 INFO Start {"addin":"1.0.0","host":"Outlook","platform":"PC","version":"16.0.18…"}
13:26:06 INFO Send mode: graph
13:26:06 INFO Mail selected {"subject":"…","from":"…","itemId":"AAMkAGI2THVSAAA…"}
13:26:10 INFO Report started {"category":"phish","mode":"graph"}
13:26:10 INFO Initialising MSAL (NAA) {"authority":"https://login.microsoftonline.com/…"}
13:26:10 INFO Token acquired (silent)
13:26:11 INFO EML fetched {"bytes":48213}
13:26:12 INFO Report sent via Graph {"to":"phish@access.ironport.com","category":"phish"}
13:26:12 INFO Message moved {"to":"junkemail"}
```

The `Send mode` line lists the reasons whenever graph mode is not active, for example
`["CLIENT_ID is empty in config.js"]` or `["NestedAppAuth 1.1 not supported by this client"]`.

## Symptoms

| Symptom                                                                 | Cause and fix                                                                                                                                                                                                                              |
| ----------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Chip says *Åbner ny mail til afsendelse* although `CLIENT_ID` is set     | Read the `Send mode` reasons. Old Outlook build without Mailbox 1.14 / NestedAppAuth → update the client. `MSAL … not loaded` → `lib/msal-browser.min.js` returns 404 or is blocked by CSP.                                                    |
| `AADSTS65001` / *mangler samtykke*                                       | Admin consent not granted for `Mail.Send` (and `Mail.ReadWrite`). Grant it in the app registration, or let the user consent if the tenant allows user consent.                                                                             |
| `AADSTS50011` redirect URI mismatch                                      | The SPA redirect URI `brk-multihub://<origin>` does not match the origin serving `taskpane.html`. Compare scheme-less host (and port) exactly. A path in the redirect URI is invalid.                                                          |
| `AADSTS700016` application not found                                     | `CLIENT_ID` is wrong or the registration is in another tenant than `AUTHORITY` points to.                                                                                                                                                  |
| `BrowserAuthError: popup_window_error` / nothing happens after the click | Interactive fallback needed but pop-ups are blocked. Grant admin consent so the silent path succeeds, or allow pop-ups for the add-in origin.                                                                                                |
| *Rapporten er sendt, men mailen kunne ikke flyttes* in the log           | `Mail.ReadWrite` missing or not consented. Grant it or set `MOVE_AFTER_REPORT: false`. The report itself was delivered.                                                                                                                       |
| Graph `403 ErrorAccessDenied` on `sendMail`                              | `Mail.Send` not consented, or a mailbox policy blocks sending (shared mailbox without Send As, disabled mailbox). The user can still use *Send manuelt i stedet*.                                                                             |
| Graph `413` or `ErrorMessageSizeExceeded`                                | EML larger than Graph accepts inline. Lower `MAX_GRAPH_ATTACHMENT_BYTES` so the automatic fallback triggers earlier.                                                                                                                       |
| Large messages always open a new mail                                    | Expected above `MAX_GRAPH_ATTACHMENT_BYTES` (3 MB). Compose mode has no such limit; Cisco accepts up to 10 MB in total.                                                                                                                   |
| `getAsFileAsync` fails                                                   | Item is a draft, a meeting request, or in a shared/delegate mailbox where the API is not supported. Use *Send manuelt i stedet*.                                                                                                             |
| New mail opens without the attachment                                    | Item not saved on the server (draft) or the client could not resolve the `itemId`. Report only received messages.                                                                                                                          |
| Button *Rapportér mail* missing                                          | Deployment not propagated yet (up to 24 h); reading pane turned off (Office.js needs it); user not in the assigned group; item is not a message (appointments never show the button).                                                        |
| Pane blank / `Office is not defined` in console                          | `https://appsforoffice.microsoft.com` blocked by proxy or CSP.                                                                                                                                                                             |
| Pane shows raw keys like `chipAuto`                                      | `STRINGS` missing that key in `config.js` – usually after an incomplete translation.                                                                                                                                                       |
| Manifest upload rejected in the admin center                             | Run `npm run validate` locally for the exact error. Common causes: placeholder host left in, icon URL not reachable, `Version` not increased on update, string length limits exceeded.                                                     |
| Old version keeps loading after an update                                | Outlook caches add-in files. Hard-reload the pane (right-click → Reload), or in classic Outlook clear `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.                                                                                          |

## Checking the host from a browser

Open these directly; each must return 200 with a valid certificate:

```text
https://<host>/manifest.xml
https://<host>/taskpane.html
https://<host>/config.js
https://<host>/lib/msal-browser.min.js
https://<host>/assets/icon-64.png
```

## Checking the report at Cisco

Reports appear in the [Talos Email Status Portal](https://talosintelligence.com/email_status_portal)
under *Submissions* for the sender's domain. If nothing shows up while the log says
`Report sent via Graph`, check the sender's Sent Items (with "keep a copy" on) or an address in
`CC_ADDRESSES` to confirm the mail left the mailbox, and then mail flow rules / DLP that might
block a `message/rfc822` attachment to an external domain.

## Escalating

Collect: the pane log, the Outlook client and version (from the `Start` line), the send mode,
the category, the time, and – for graph mode – the `AADSTS` or Graph error code. That is enough
to reproduce almost every issue.
