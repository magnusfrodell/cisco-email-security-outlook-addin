# Security

## What the add-in does with data

| Data                                   | Where it goes                                   | When                                                |
| -------------------------------------- | ----------------------------------------------- | --------------------------------------------------- |
| The full reported e-mail (headers + body + attachments) as `.eml`, plus a details block (category, reporting user, original subject/sender/date, Message-ID) | Cisco Talos, via one of the `*@access.ironport.com` addresses, and any `SOC_ADDRESSES` (Cc or Bcc) the administrator configured | Only when the user clicks a category button |
| Subject and sender of the selected mail | Shown in the pane; written to the in-pane log and browser console | While the pane is open |
| A mail the user reports from a shared mailbox | Moved within that shared mailbox, and read from it if Outlook cannot supply it as a file – through Microsoft Graph, with the user's own access to that mailbox | Only when the user clicks a category button on such a mail |
| Microsoft Graph access token            | Kept in MSAL's cache (`localStorage` of the add-in origin) | Automatic mode only |
| "Keep a copy in Sent Items" and language choices | `Office.context.roamingSettings` (the user's mailbox) | When changed |

Nothing is sent to the add-in's own web host except requests for the static files. There is no
backend, no telemetry and no third-party script apart from Office.js
(`appsforoffice.microsoft.com`) and the self-hosted MSAL library.

## Permissions

| Scope / permission                  | Used for                                             | Required |
| ----------------------------------- | ---------------------------------------------------- | -------- |
| Manifest `ReadWriteMailbox`         | Read the selected item, open a new message form. The level is `ReadWriteMailbox` rather than `ReadWriteItem` because Microsoft requires it for add-ins that act on items in shared mailboxes; the add-in calls no EWS and no mailbox-wide Office.js API. | yes (`ReadWriteItem` suffices if `SupportsSharedFolders` is removed from the manifest) |
| Graph delegated `Mail.Send`         | `POST /me/sendMail` on behalf of the user            | automatic mode |
| Graph delegated `Mail.Send.Shared`  | `POST /me/sendMail` with `from` set to a shared mailbox the user has Send As rights on | automatic mode with `SEND_AS` (replaces `Mail.Send`) |
| Graph delegated `Mail.ReadWrite`    | `POST /me/messages/{id}/move` to Junk/Inbox; `GET /me/messages/{id}/$value` if Outlook cannot supply the mail as a file | automatic mode with `MOVE_AFTER_REPORT: true` |
| Graph delegated `Mail.ReadWrite.Shared` | `POST /users/{mailbox}/messages/{id}/move`, and `GET …/$value` if Outlook cannot supply the mail as a file – in mailboxes the user already has access to | reporting from shared mailboxes (`Mail.Read.Shared` suffices with `MOVE_IN_SHARED_MAILBOX: false`) |

Tokens are obtained through Nested App Authentication (brokered by Outlook) for the
organisation's own Entra app registration; no client secret exists anywhere in the project.

## Things to be aware of

- A reported mail leaves the organisation and is processed by Cisco under Cisco's terms. This is
  the same data flow as Cisco's own Secure Email Submission add-in and as manual forwarding to
  the same addresses.
- In compose mode the report is an ordinary outgoing mail; it is subject to the organisation's
  DLP and mail-flow rules like any other message.
- The SOC copy is configured by the administrator and is not shown to the user unless
  `SOC_SHOW_IN_STATUS` is on. Decide deliberately whether users should be told.
- A mail reported from a shared mailbox is moved for everyone who uses that mailbox. Set
  `MOVE_IN_SHARED_MAILBOX: false` if the team should keep handling such mails itself.
- `Mail.ReadWrite.Shared` is a delegated permission. It gives no user access to a mailbox they
  could not already open; it lets the add-in act, as that user, in the mailboxes the user can
  open. The token for it is requested silently and only when a mail in a shared mailbox is
  reported.
- Anyone who can modify the files on the add-in web host can change where reports are sent.
  Treat the host like any other production web property (access control, TLS, change tracking).

## Reporting a vulnerability

Open a private security advisory on the repository, or contact the maintainer directly. Please
do not open a public issue for security problems.
