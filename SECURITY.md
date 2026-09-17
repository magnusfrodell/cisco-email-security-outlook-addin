# Security

## What the add-in does with data

| Data                                   | Where it goes                                   | When                                                |
| -------------------------------------- | ----------------------------------------------- | --------------------------------------------------- |
| The full reported e-mail (headers + body + attachments) as `.eml` | Cisco Talos, via one of the `*@access.ironport.com` addresses, plus any `CC_ADDRESSES` | Only when the user clicks a category button |
| Subject and sender of the selected mail | Shown in the pane; written to the in-pane log and browser console | While the pane is open |
| Microsoft Graph access token            | Kept in MSAL's cache (`localStorage` of the add-in origin) | Automatic mode only |
| "Keep a copy in Sent Items" choice       | `Office.context.roamingSettings` (the user's mailbox) | When changed |

Nothing is sent to the add-in's own web host except requests for the static files. There is no
backend, no telemetry and no third-party script apart from Office.js
(`appsforoffice.microsoft.com`) and the self-hosted MSAL library.

## Permissions

| Scope / permission                  | Used for                                             | Required |
| ----------------------------------- | ---------------------------------------------------- | -------- |
| Manifest `ReadWriteItem`            | Read the selected item, open a new message form      | yes      |
| Graph delegated `Mail.Send`         | `POST /me/sendMail` on behalf of the user            | automatic mode |
| Graph delegated `Mail.ReadWrite`    | `POST /me/messages/{id}/move` to Junk/Inbox          | automatic mode with `MOVE_AFTER_REPORT: true` |

Tokens are obtained through Nested App Authentication (brokered by Outlook) for the
organisation's own Entra app registration; no client secret exists anywhere in the project.

## Things to be aware of

- A reported mail leaves the organisation and is processed by Cisco under Cisco's terms. This is
  the same data flow as Cisco's own Secure Email Submission add-in and as manual forwarding to
  the same addresses.
- In compose mode the report is an ordinary outgoing mail; it is subject to the organisation's
  DLP and mail-flow rules like any other message.
- Anyone who can modify the files on the add-in web host can change where reports are sent.
  Treat the host like any other production web property (access control, TLS, change tracking).

## Reporting a vulnerability

Open a private security advisory on the repository, or contact the maintainer directly. Please
do not open a public issue for security problems.
