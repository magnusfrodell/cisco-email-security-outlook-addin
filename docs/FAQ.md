# FAQ

Questions that come up when an organisation evaluates this add-in against Cisco's own
*Secure Email Submission* add-in. Technical detail is in `docs/ARCHITECTURE.md`; the deployment
steps are in `docs/DEPLOYMENT.md`.

## Why does automatic send need an Entra app registration? Cisco's add-in "just works".

Because Cisco's add-in needs one too – it is Cisco's app, consented in your tenant. Office.js
has no API that lets a task pane send a message silently or move an item; everything that
happens after the click is done through Microsoft Graph with a delegated token, and a
delegated token requires an app registration that your tenant has consented to. Look under
*Enterprise applications* in Entra: Cisco's add-in appears there with *Mail.ReadWrite* and
*Mail.Send*. With this add-in the app is yours (or a shared multi-tenant one you consent to),
which is the same amount of consent work – a redirect URI, two delegated permissions, no
secrets.

## What does Cisco's add-in do under the hood?

From the JavaScript bundle served at `res.cisco.com/submission/` and Cisco's admin guide:

- Two mechanisms exist in the code. The older one calls Exchange Web Services through
  `Office.context.mailbox.makeEwsRequestAsync` – this is what the manifest permission
  `ReadWriteMailbox` unlocks, and it is authenticated with the legacy Exchange callback token.
  The newer one uses Microsoft Graph: `GET /me/messages/{id}/$value` to fetch the MIME,
  `POST /me/sendMail` to submit, `POST /me/messages/{id}/move` to move the message, with
  Cisco's own multi-tenant Entra app (client ID `ec007a11-e3ed-4943-b3da3c79c960c4ca`,
  delegated *User.Read*, *Mail.ReadWrite*, *Mail.Send*).
- There is no Cisco backend in the loop. The submission is an ordinary e-mail from the user's
  mailbox to `addin_spam@`, `addin_ham@` or `addin_mktg@access.ironport.com` with the original
  message as a `message/rfc822` attachment (`orig_msg_*.raw.eml`, plus an
  `X-MS-AddIn-Base64Encode: true` header). `res.cisco.com` only serves static files.
- Which path runs today: Graph. Microsoft turned legacy Exchange tokens off for all tenants by
  October 2025 with no re-enable option, so the EWS path cannot work in Exchange Online any
  more – and EWS itself is retired in Exchange Online from 1 October 2026. When Cisco's add-in
  submits and moves a message, that is Graph with Cisco's Entra app.

This add-in mirrors the Graph path and the message format, and uses Cisco's public submission
addresses (`spam@`, `phish@`, `virus@`, `ads@`, `ham@`, `not_ads@access.ironport.com`) instead
of the `addin_*` ones, which are reserved for Cisco's add-in.

## Can we keep Cisco's interaction model but avoid any Entra app?

Not with a task-pane add-in. One click → submit + move with no outgoing mail shown requires
Graph, and Graph requires an app. The alternatives change the model:

| Option                                            | Entra app | What changes                                                                                          |
| ------------------------------------------------- | --------- | ----------------------------------------------------------------------------------------------------- |
| This add-in, `SEND_MODE: "graph"` (default)       | yes       | Nothing – same one-click model as Cisco's add-in, in your language, with your settings.               |
| This add-in, `SEND_MODE: "compose"`               | no        | A pre-addressed message opens and the user presses Send; the message is not moved automatically.      |
| Outlook's integrated spam-reporting surface + a self-hosted relay | no | Native *Report* button instead of a task pane, and a server of yours that relays the EML to Cisco – trades the app registration for infrastructure. Not part of this project. |
| EWS via `ReadWriteMailbox`                        | no        | Not possible: legacy tokens are off since October 2025 and EWS is retired in Exchange Online from October 2026. |

## Can we self-host Cisco's task pane instead?

No. The code is Cisco's and is not licensed for redistribution, and its MSAL configuration is
bound to `res.cisco.com` as the redirect origin – a copy served from another host fails token
acquisition with `AADSTS50011`. A manifest of your own that points at `res.cisco.com` is
possible (it is what Cisco's manifest does) but changes nothing about the pane, its language or
its behaviour.

## Consent to Cisco's app, our own app, or a shared app – what is the difference?

| Model                                   | Who owns the app        | Admin work in your tenant                        | Redirect URI must match                  |
| --------------------------------------- | ----------------------- | ------------------------------------------------ | ---------------------------------------- |
| Cisco's add-in                          | Cisco                   | Admin consent once                               | `res.cisco.com` (Cisco's host)           |
| This add-in, own registration (model A) | You                     | Register app, two permissions, consent           | Your host                                |
| This add-in, shared registration (model B) | Whoever hosts the files | Admin consent once via the consent URL       | The shared host                          |

None of the three involves a client secret or an application permission; all use delegated
permissions bound to the signed-in user and brokered by Outlook (Nested App Authentication).

## Does Talos treat reports from this add-in like reports from Cisco's?

Yes. Reports arrive at Cisco's documented submission addresses in the same message shape,
sent from a mailbox in your domain, and appear in the Talos Email Status Portal for that domain
like any other submission. With `SEND_AS` they can all be attributed to one SOC mailbox.

## What leaves the organisation?

The reported message – headers, body and attachments – goes to Cisco Talos and to any
`SOC_ADDRESSES` you configure. Nothing goes to the add-in's web host except requests for the
static files. Details and permissions: `SECURITY.md`.

## Is this supported by Cisco?

It is Cisco sample code under the Cisco Sample Code License, provided as-is; Cisco TAC does not
support it. Cisco's own add-in remains the supported product. Reports submitted through this
add-in are handled by Talos exactly as any other submission to those addresses.
