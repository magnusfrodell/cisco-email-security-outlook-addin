# Deployment guide

Step-by-step, from a clone of this repository to a centrally deployed add-in. The Danish
customer-facing version of this guide is `docs/INSTALLATION.da.md`.

Time needed: about 45 minutes including the Entra app registration that automatic send (the
default) relies on. Central deployment can then take a few hours to reach all users.

## Prerequisites

- A static HTTPS web host with a valid certificate (Azure Static Web Apps, an Azure Storage
  static website, IIS, nginx, GitHub Pages for demos – anything that serves files).
- Microsoft 365 with Exchange Online mailboxes.
- Global Administrator or Exchange Administrator for central deployment; Application
  Administrator (or higher) for the Entra app registration.
- Users on supported Outlook clients (see `docs/ARCHITECTURE.md`, section 7).

## Step 1 – Host the files

1. Pick the hostname, e.g. `addin.contoso.com`. The contents of `src/` must be served from the
   root of that origin:
   `https://addin.contoso.com/taskpane.html`, `https://addin.contoso.com/assets/icon-64.png`, …
2. Upload everything under `src/` to the host.
3. Put the real host into the manifest:

   ```bash
   scripts/set-host.sh addin.contoso.com                 # edits src/manifest.xml in place
   scripts/set-host.sh addin.contoso.com manifest-prod.xml   # or write a separate copy
   ```

   PowerShell: `.\scripts\set-host.ps1 -Host addin.contoso.com`.
   The script replaces every `https://addin.example.com` URL and leaves the paths alone.
4. Optionally change `<ProviderName>` in the manifest to the organisation's name.
5. Verify in a browser that `https://addin.contoso.com/taskpane.html` and
   `https://addin.contoso.com/assets/icon-64.png` return 200 with a trusted certificate.

If the host sends a Content-Security-Policy, it must allow scripts from
`https://appsforoffice.microsoft.com` and connections to `https://login.microsoftonline.com`
and `https://graph.microsoft.com`. MSAL is self-hosted, so no other script origin is needed.

### Alternative for demos: GitHub Pages

The workflow `.github/workflows/pages.yml` publishes `src/` to
`https://<owner>.github.io/<repo>/` and rewrites the manifest for that host. Enable Pages in the
repository settings with *Source: GitHub Actions*, run the workflow, then sideload
`https://<owner>.github.io/<repo>/manifest.xml`. For automatic mode the Entra redirect URI is
`brk-multihub://<owner>.github.io`.

## Step 2 – Entra app registration

Automatic send is the default and needs an Entra app registration to obtain Graph tokens
through Outlook. Until `CLIENT_ID` is set, the pane shows a warning and opens a new message
instead. (If someone asks why an app is needed when Cisco's add-in "just works": Cisco's add-in
uses its own Entra app, consented in your tenant – see `docs/FAQ.md`.) Two models:

**A. One registration per customer tenant** (self-contained, recommended when the customer
hosts the files):

1. [entra.microsoft.com](https://entra.microsoft.com) → **Identity → Applications →
   App registrations → New registration**.
2. **Name:** `Cisco Email Security Reporter` (or anything recognisable).
   **Supported account types:** *Accounts in this organizational directory only*.
   **Redirect URI:** platform **Single-page application (SPA)**, value
   `brk-multihub://addin.contoso.com`.
   Origin only: no scheme, no path; include the port if it is not 443.
3. **Register**, then copy the **Application (client) ID** and **Directory (tenant) ID**.
4. **API permissions → Add a permission → Microsoft Graph → Delegated permissions:**
   - `Mail.Send` – send the report on the user's behalf
   - `Mail.ReadWrite` – move the reported message (omit if you set `MOVE_AFTER_REPORT: false`)
   - `Mail.ReadWrite.Shared` – read and move mails in shared mailboxes (only if users are to
     report from shared mailboxes; see "Shared mailboxes" under step 3)
5. **Grant admin consent for \<organisation\>** so users are never prompted.

**B. One shared multi-tenant registration** (the model Cisco's own add-in uses – one app owned
by whoever hosts the files, consented by each customer tenant):

1. Register the app as in A, but with **Supported account types:** *Accounts in any
   organizational directory (multitenant)*, and add one SPA redirect URI
   `brk-multihub://<host>` for every host the files are served from (e.g. the GitHub Pages
   host `brk-multihub://<owner>.github.io`).
2. Add the same delegated permissions. Do **not** add a client secret; none is needed.
3. Ship its client ID as the default in `config.js` with
   `AUTHORITY: "https://login.microsoftonline.com/organizations"`.
4. Each customer's Global Administrator grants consent once by opening

   ```text
   https://login.microsoftonline.com/<customer-tenant-id-or-domain>/adminconsent?client_id=<client-id>
   ```

   and accepting. The app then appears under *Enterprise applications* in their tenant with
   those permissions, and their users are never prompted.

Model B means one host and one app for everybody; model A means the customer controls both.
Mixing is fine: a customer can start on B and move to A later by changing `CLIENT_ID`,
`AUTHORITY` and the host.

Why the odd redirect URI: `brk-multihub://<origin>` tells the Microsoft identity platform that
the add-in may obtain tokens *through* Outlook (Nested App Authentication). Without it, token
requests fail with `AADSTS50011`. Compose-only deployments (`SEND_MODE: "compose"`) skip this
step entirely.

## Step 3 – Configure

Edit `src/config.js` on the host:

```javascript
LANGUAGE: "auto",                                             // or "da" / "en" / "sv" for everyone
SEND_MODE: "graph",                                           // default; "compose" = always a new message
COMPOSE_FALLBACK: true,                                       // new message when automatic send is unavailable
CLIENT_ID: "00000000-0000-0000-0000-000000000000",           // from step 2
AUTHORITY: "https://login.microsoftonline.com/<tenant-id>",  // model A; "…/organizations" for model B
MOVE_AFTER_REPORT: true,
SOC_ADDRESSES: [],                                            // e.g. ["soc@contoso.com"] – copy of every report
SOC_COPY_MODE: "cc",                                          // or "bcc" to hide the SOC from Cisco
SEND_AS: "",                                                  // e.g. "soc@contoso.com" – submit from the SOC mailbox
```

See `docs/CONFIGURATION.md` for everything else. Configuration changes never require a manifest
update.

### SOC visibility (optional)

Two administrator-only settings give the SOC insight into what users report; users cannot see or
change them.

**A copy of every report** – set `SOC_ADDRESSES` to the SOC mailbox (or a ticketing or SIEM
address). The SOC then receives the same mail Cisco receives, with the original as a `.raw.eml`
attachment and a `--- Report details ---` block (category, reporter, original subject, sender,
Message-ID) in the body. `SOC_COPY_MODE: "bcc"` keeps the SOC address off the mail Cisco sees;
`SOC_SHOW_IN_STATUS: true` tells users a copy went to the SOC.

**Submitting from the SOC mailbox** – set `SEND_AS` to a shared mailbox and every automatic
report is sent from that mailbox, so the Talos Email Status Portal shows all submissions under
one address the SOC owns. This needs three things:

1. A shared mailbox, e.g. `soc@contoso.com` (Exchange admin center → *Recipients → Mailboxes →
   Add a shared mailbox*).
2. **Send As** permission for the reporting users on that mailbox – in the Exchange admin center
   under the mailbox's *Delegation*, or for a group of users with PowerShell:

   ```powershell
   Add-RecipientPermission soc@contoso.com -AccessRights SendAs -Trustee "All Reporters" -Confirm:$false
   ```

   *Send on Behalf* works too, but recipients (Talos included) then see "user on behalf of
   soc@contoso.com". Full Access is not required.
3. The Entra app from step 2 needs the delegated permission **`Mail.Send.Shared`** instead of
   `Mail.Send`, with admin consent granted again.

Permission changes in Exchange can take up to an hour to be effective. The compose fallback
cannot change the sender; if it kicks in, the report leaves from the user's own mailbox and the
log says so.

### Shared mailboxes (optional)

Users can report a mail they read in a shared mailbox (support@, info@ …) or in a folder that
another user shares with them. The manifest already declares this
(`<SupportsSharedFolders>true</SupportsSharedFolders>`, with `ReadWriteMailbox` as the manifest
permission, which Microsoft requires for add-ins that act on shared items). What remains is
access:

1. **Graph permission.** Add the delegated permission **`Mail.ReadWrite.Shared`** to the Entra
   app from step 2 and grant admin consent again. With `MOVE_IN_SHARED_MAILBOX: false` the
   read-only `Mail.Read.Shared` is enough.
2. **Mailbox access.** The reporting user needs Full Access to the shared mailbox – the way
   members normally get a shared mailbox – or, for a shared folder, folder permissions that let
   them read and move the mail.
3. **Deployment.** Deploy the add-in to the *users* (step 5). It then also appears on mails in
   the shared mailboxes they open. Do not deploy it to the shared mailbox itself.
4. **Clients.** Outlook on the web, new Outlook on Windows, classic Outlook on Windows Version
   2304 (Build 16327.20248) or later, Outlook on Mac 16.47 or later. Not Outlook mobile.

What the user gets: the pane shows *Shared mailbox: support@contoso.com* under the selected
mail. The report is sent from the user's own mailbox – so Talos tracks it under the user, or
under the `SEND_AS` mailbox if that is set – and the mail is moved to Junk Email or the Inbox
*of the shared mailbox*, that is, for everyone who uses it. Set `MOVE_IN_SHARED_MAILBOX: false`
to report without moving.

Microsoft's limitations apply: a shared mailbox that is hidden from the address list, or one
the user signed in to with the mailbox's own password instead of opening it through their own
account, gives add-ins no item to work with – the pane then says that no mail is selected.

To leave shared mailboxes out altogether, remove the `SupportsSharedFolders` element from the
manifest and set `<Permissions>` back to `ReadWriteItem`.

Test it as in step 4, on a mail in a shared mailbox, and read the technical log: the line
`Item is in a shared mailbox or folder` shows the owner the pane found, and `Message moved`
names the mailbox. The automated test covers this path against a mocked Microsoft Graph only,
so confirm it once in your tenant before rolling it out.

## Step 4 – Test by sideloading

1. Open <https://aka.ms/olksideload> – Outlook on the web opens the *Add-ins* dialog.
2. **My add-ins → Add a custom add-in → Add from file** and pick your manifest.
3. Select a message, click **Report to Cisco** (group *Email security*; *Rapportér til Cisco* /
   *Rapportera till Cisco* in a Danish or Swedish Outlook) and report it as, say, *Marketing*.
4. The header chip must say *Automatic send*. If a warning appears under the header instead,
   automatic send is not active – expand **Technical log** at the bottom of the pane: it shows
   the language chosen and why, the send mode, the reasons, and every step of the send.
5. Switch the language in the pane's settings and check that the ribbon (which follows the
   Outlook display language) and the pane show what you expect.
6. During testing, set `SOC_ADDRESSES` to your own address to see the outgoing report exactly
   as Cisco and the SOC receive it.

Test in Outlook on the web first (fastest feedback), then in the desktop clients the users
actually have. Automatic mode can only be tested this way; the repository's smoke test covers
compose mode only.

## Step 5 – Central deployment

1. <https://admin.microsoft.com> → **Settings → Integrated apps → Upload custom apps**.
2. **App type:** Office Add-in → **Upload manifest file (.xml) from device** → select the
   manifest with the real host.
3. Assign users (everyone, or specific groups) → **Accept permissions** → **Finish deployment**.
4. Allow up to 24 hours for the button to appear for everyone; a few hours is typical.

From 1.4.0 the manifest asks for the *read/write mailbox* permission level (see "Shared
mailboxes" under step 3), so that is what the admin center shows when you accept the
permissions. It is the same level Cisco's own manifest uses.

Cisco's own add-in can be removed from *Integrated apps* once this one is verified. The two can
also coexist during a transition.

## Step 6 – Verify at Cisco

Sign in to the [Cisco Talos Email Status Portal](https://talosintelligence.com/email_status_portal)
with a Cisco account registered for the organisation's domain. User reports appear under
*Submissions* within minutes. Cisco does not send acknowledgement mails.

## Updating

- **Files only** (`config.js`, locale files, HTML, CSS, JS, icons): upload to the host. Outlook
  fetches the current version each time a pane opens.
- **Manifest** (new button text, new icons, new permissions): bump `<Version>` (e.g. `1.0.1.0`),
  then *Integrated apps → select the app → Update* and upload the new manifest.
- **Upgrading to 1.4.0** is a manifest update as well as a file update: shared-mailbox support
  and the `ReadWriteMailbox` permission level live in the manifest, so upload the new manifest
  and accept the permissions. Until then the new files run with the old manifest exactly as
  before, only without shared mailboxes.

## Removing

*Integrated apps → select the app → Remove app*. The files on the host can be deleted afterwards.
The Entra app registration can be deleted once no other add-in uses it.
