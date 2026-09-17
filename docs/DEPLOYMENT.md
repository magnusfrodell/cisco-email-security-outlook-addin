# Deployment guide

Step-by-step, from a clone of this repository to a centrally deployed add-in. The Danish
customer-facing version of this guide is `docs/INSTALLATION.da.md`.

Time needed: about 30 minutes for compose mode, plus 15 minutes for the Entra app registration
if you want automatic mode. Central deployment can then take a few hours to reach all users.

## Prerequisites

- A static HTTPS web host with a valid certificate (Azure Static Web Apps, an Azure Storage
  static website, IIS, nginx, GitHub Pages for demos – anything that serves files).
- Microsoft 365 with Exchange Online mailboxes.
- Global Administrator or Exchange Administrator for central deployment; Application
  Administrator (or higher) for the Entra app registration.
- Users on supported Outlook clients (see `docs/ARCHITECTURE.md`, section 7).

## Step 1 – Host the files

1. Pick the hostname, e.g. `addin.firma.dk`. The contents of `src/` must be served from the
   root of that origin:
   `https://addin.firma.dk/taskpane.html`, `https://addin.firma.dk/assets/icon-64.png`, …
2. Upload everything under `src/` to the host.
3. Put the real host into the manifest:

   ```bash
   scripts/set-host.sh addin.firma.dk                 # edits src/manifest.xml in place
   scripts/set-host.sh addin.firma.dk manifest-prod.xml   # or write a separate copy
   ```

   PowerShell: `.\scripts\set-host.ps1 -Host addin.firma.dk`.
   The script replaces every `https://addin.example.com` URL and leaves the paths alone.
4. Optionally change `<ProviderName>` in the manifest to the organisation's name.
5. Verify in a browser that `https://addin.firma.dk/taskpane.html` and
   `https://addin.firma.dk/assets/icon-64.png` return 200 with a trusted certificate.

If the host sends a Content-Security-Policy, it must allow scripts from
`https://appsforoffice.microsoft.com` and connections to `https://login.microsoftonline.com`
and `https://graph.microsoft.com`. MSAL is self-hosted, so no other script origin is needed.

### Alternative for demos: GitHub Pages

The workflow `.github/workflows/pages.yml` publishes `src/` to
`https://<owner>.github.io/<repo>/` and rewrites the manifest for that host. Enable Pages in the
repository settings with *Source: GitHub Actions*, run the workflow, then sideload
`https://<owner>.github.io/<repo>/manifest.xml`. For automatic mode the Entra redirect URI is
`brk-multihub://<owner>.github.io`.

## Step 2 – Entra app registration (automatic mode only)

Skip this step if compose mode is enough. Leave `CLIENT_ID` empty in that case.

1. [entra.microsoft.com](https://entra.microsoft.com) → **Identity → Applications →
   App registrations → New registration**.
2. **Name:** `Rapportér mail til Cisco` (or anything recognisable).
   **Supported account types:** *Accounts in this organizational directory only*.
   **Redirect URI:** platform **Single-page application (SPA)**, value
   `brk-multihub://addin.firma.dk`.
   Origin only: no scheme, no path; include the port if it is not 443.
3. **Register**, then copy the **Application (client) ID** and **Directory (tenant) ID**.
4. **API permissions → Add a permission → Microsoft Graph → Delegated permissions:**
   - `Mail.Send` – send the report on the user's behalf
   - `Mail.ReadWrite` – move the reported message (omit if you set `MOVE_AFTER_REPORT: false`)
5. **Grant admin consent for \<organisation\>** so users are never prompted.

Why the odd redirect URI: `brk-multihub://<origin>` tells the Microsoft identity platform that
the add-in may obtain tokens *through* Outlook (Nested App Authentication). Without it, token
requests fail with `AADSTS50011`.

## Step 3 – Configure

Edit `src/config.js` on the host:

```javascript
CLIENT_ID: "00000000-0000-0000-0000-000000000000",           // from step 2, or "" for compose mode
AUTHORITY: "https://login.microsoftonline.com/<tenant-id>",  // tenant ID from step 2
MOVE_AFTER_REPORT: true,
CC_ADDRESSES: [],                                             // e.g. ["soc@firma.dk"]
```

See `docs/CONFIGURATION.md` for everything else. Configuration changes never require a manifest
update.

## Step 4 – Test by sideloading

1. Open <https://aka.ms/olksideload> – Outlook on the web opens the *Add-ins* dialog.
2. **My add-ins → Add a custom add-in → Add from file** and pick your manifest.
3. Select a message, click **Rapportér mail** (group *Mailsikkerhed*) and report it as, say,
   *Marketing*.
4. Expand **Teknisk log** at the bottom of the pane. It shows the chosen send mode, the reason if
   automatic mode is not active, and every step of the send.
5. During testing, set `CC_ADDRESSES` to your own address to see the outgoing report.

Test in Outlook on the web first (fastest feedback), then in the desktop clients the users
actually have. Automatic mode can only be tested this way; the repository's smoke test covers
compose mode only.

## Step 5 – Central deployment

1. <https://admin.microsoft.com> → **Settings → Integrated apps → Upload custom apps**.
2. **App type:** Office Add-in → **Upload manifest file (.xml) from device** → select the
   manifest with the real host.
3. Assign users (everyone, or specific groups) → **Accept permissions** → **Finish deployment**.
4. Allow up to 24 hours for the button to appear for everyone; a few hours is typical.

Cisco's own add-in can be removed from *Integrated apps* once this one is verified. The two can
also coexist during a transition.

## Step 6 – Verify at Cisco

Sign in to the [Cisco Talos Email Status Portal](https://talosintelligence.com/email_status_portal)
with a Cisco account registered for the organisation's domain. User reports appear under
*Submissions* within minutes. Cisco does not send acknowledgement mails.

## Updating

- **Files only** (`config.js`, HTML, CSS, JS, icons): upload to the host. Outlook fetches the
  current version each time a pane opens.
- **Manifest** (new button text, new icons, new permissions): bump `<Version>` (e.g. `1.0.1.0`),
  then *Integrated apps → select the app → Update* and upload the new manifest.

## Removing

*Integrated apps → select the app → Remove app*. The files on the host can be deleted afterwards.
The Entra app registration can be deleted once no other add-in uses it.
