# Installationsvejledning (dansk)

*Cisco Email Security Reporter* er et Outlook-tilføjelsesprogram, der lader brugerne rapportere
spam, phishing, virus og marketing til Cisco Talos – og fortælle Cisco, når en legitim mail
fejlagtigt blev stoppet. Brugerfladen findes på dansk, engelsk og svensk og følger automatisk
brugerens sprog i Outlook; brugeren kan også selv vælge sprog i ruden. Rapporterne sendes til
Ciscos officielle indsendelsesadresser og kan følges i
[Cisco Talos Email Status Portal](https://talosintelligence.com/email_status_portal).

Det erstatter Ciscos eget *Cisco Secure Email Submission Add-In*, hvis brugerflade kun findes på
engelsk og ikke kan tilpasses. Selve indsendelsen sker på præcis samme måde som i Ciscos add-in: den rapporterede mail
sendes som en `message/rfc822`-vedhæftning (`.raw.eml`) til Cisco.

## Kategorier og adresser

| Knap i ruden                       | Sendes til                    | Mailen flyttes til |
| ---------------------------------- | ----------------------------- | ------------------ |
| Spam                               | spam@access.ironport.com      | Uønsket mail       |
| Phishing                           | phish@access.ironport.com     | Uønsket mail       |
| Virus eller skadelig vedhæftning   | virus@access.ironport.com     | Uønsket mail       |
| Marketing                          | ads@access.ironport.com       | Uønsket mail       |
| Legitim mail (ikke spam)           | ham@access.ironport.com       | Indbakke           |
| Ikke marketing                     | not_ads@access.ironport.com   | Indbakke           |

Adresserne er Ciscos dokumenterede adresser til direkte indsendelse (Cisco-dokument 214133).
Knapper kan skjules med `enabled: false` i `config.js`.

## Sådan virker det

Standarden er **automatisk afsendelse**: mailen hentes som EML, sendes via Microsoft Graph på
brugerens vegne og flyttes til Uønsket mail / Indbakke – ingen klik ud over kategorivalget. Det
kræver en Entra-appregistrering (trin 2) og en Outlook-klient med Mailbox 1.14 og Nested App
Auth.

| Måde            | Hvad sker der                                                                                                                                                    | Hvornår                                                                                                          |
| --------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| **Automatisk**  | Sendes via Graph, flyttes bagefter. Chippen øverst siger *Automatisk afsendelse*.                                                                                 | Standard, når `CLIENT_ID` er sat og klienten understøtter det.                                                    |
| **Ny mail**     | Der åbnes en ny mail til Ciscos adresse med den valgte besked vedhæftet (= "videresend som vedhæftning"). Brugeren trykker Send.                                  | Reserve: når automatisk afsendelse ikke er tilgængelig (der vises en advarsel i ruden), når mailen er over ca. 3 MB, eller via *Send manuelt i stedet* efter en fejl. Kan slås fra med `COMPOSE_FALLBACK: false` eller gøres til eneste måde med `SEND_MODE: "compose"`. |

Understøttede klienter: Outlook til Windows (klassisk og nyt Outlook, Microsoft 365), Outlook
til Mac (Microsoft 365) og Outlook på nettet. Outlook til iOS/Android er ikke understøttet.

## Filer, der skal publiceres

Alt i mappen `src/`:

| Fil                        | Formål                                                                    |
| -------------------------- | ------------------------------------------------------------------------- |
| `manifest.xml`             | Add-in-manifest (dansk). Uploades til Microsoft 365 Admin Center.         |
| `config.js`                | Al konfiguration: sprog, adresser, kategorier, Entra client-ID, Cc, emne. |
| `locales/da.js` m.fl.      | Én fil pr. sprog med alle tekster i ruden (dansk, engelsk, svensk).        |
| `taskpane.html/.css/.js`   | Selve ruden.                                                              |
| `lib/msal-browser.min.js`  | MSAL.js (selvhostet). Bruges kun til automatisk afsendelse.               |
| `assets/icon-*.png`        | Ikoner.                                                                   |
| `support.html`             | Supportside, som manifestet peger på.                                     |

## Trin 1 – Publicér filerne på et HTTPS-websted

1. Vælg et hostnavn, fx `addin.firma.dk`. Alle filer i `src/` skal ligge på roden af
   webstedet (`https://addin.firma.dk/taskpane.html`, `https://addin.firma.dk/assets/…`).
   Enhver statisk HTTPS-host duer: Azure Static Web Apps, statisk websted i Azure Storage, IIS,
   nginx osv. Der kræves ingen serverlogik.
2. Indsæt hostnavnet i manifestet:

   ```bash
   scripts/set-host.sh addin.firma.dk
   ```

   PowerShell: `.\scripts\set-host.ps1 -Host addin.firma.dk`.
   Scriptet erstatter alle `https://addin.example.com`-adresser i `src/manifest.xml`.
3. Ret evt. `<ProviderName>` i `manifest.xml` til jeres organisations navn.
4. Kontrollér i en browser, at `https://addin.firma.dk/taskpane.html` og
   `https://addin.firma.dk/assets/icon-64.png` svarer med 200 og et gyldigt certifikat.

Hvis webstedet har en Content-Security-Policy, skal den tillade
`https://appsforoffice.microsoft.com` (script), `https://login.microsoftonline.com` og
`https://graph.microsoft.com` (connect).

## Trin 2 – Entra-appregistrering

Automatisk afsendelse er standard og kræver en appregistrering, så tilføjelsesprogrammet kan få
Graph-tokens gennem Outlook. Indtil `CLIENT_ID` er sat, viser ruden en advarsel og åbner en ny
mail i stedet. Beskrivelsen nedenfor er én registrering i jeres egen tenant; den engelske
vejledning (`docs/DEPLOYMENT.md`, trin 2, model B) beskriver også en delt multi-tenant-app, hvor
I kun giver admin consent.

1. Gå til [entra.microsoft.com](https://entra.microsoft.com) → **Identity → Applications →
   App registrations → New registration**.
2. **Name:** `Rapportér mail til Cisco`.
   **Supported account types:** *Accounts in this organizational directory only*.
   **Redirect URI:** platform **Single-page application (SPA)**, værdi
   `brk-multihub://addin.firma.dk` (kun origin – ingen sti, ingen `https://`; medtag port, hvis
   der bruges en anden port end 443).
3. Klik **Register** og notér **Application (client) ID** og **Directory (tenant) ID**.
4. **API permissions → Add a permission → Microsoft Graph → Delegated permissions:**
   - `Mail.Send` – sende rapporten på brugerens vegne
   - `Mail.ReadWrite` – flytte mailen til Uønsket mail / Indbakke (kan udelades, hvis
     `MOVE_AFTER_REPORT` sættes til `false`)
5. Klik **Grant admin consent for \<organisation\>**, så brugerne ikke selv skal godkende.

`brk-multihub://…`-adressen fortæller Microsoft Identity Platform, at add-in'et må få tokens
gennem Outlook (Nested App Authentication). Uden den fejler tokenhentning med AADSTS50011.

## Trin 3 – Udfyld `config.js`

```javascript
LANGUAGE: "auto",            // følger Outlooks sprog – eller "da" for dansk til alle
SEND_MODE: "graph",          // standard; "compose" = altid ny mail
COMPOSE_FALLBACK: true,      // ny mail, hvis automatisk afsendelse ikke er tilgængelig
CLIENT_ID: "00000000-0000-0000-0000-000000000000",          // fra trin 2
AUTHORITY: "https://login.microsoftonline.com/<tenant-id>", // jeres tenant-ID
MOVE_AFTER_REPORT: true,
CC_ADDRESSES: [],            // fx ["soc@firma.dk"] for en intern kopi
```

`LANGUAGE: "auto"` giver dansk til brugere med dansk Outlook, engelsk/svensk til de øvrige.
`LANGUAGE: "da"` giver dansk til alle. Sprogvælgeren i ruden kan slås fra med
`SHOW_LANGUAGE_SELECTOR: false`.

Upload filen igen til webstedet. Ingen ændringer i manifestet er nødvendige, når konfigurationen
ændres.

## Trin 4 – Test ved sideloading

1. Åbn <https://aka.ms/olksideload> (Outlook på nettet åbner dialogen *Tilføjelsesprogrammer*).
2. **Mine tilføjelsesprogrammer → Tilføj et brugerdefineret tilføjelsesprogram → Tilføj fra fil**
   og vælg manifestet.
3. Vælg en mail, klik **Rapportér til Cisco** i båndet (under *Mailsikkerhed*) og rapportér den
   som fx *Marketing*.
4. Chippen øverst skal sige *Automatisk afsendelse*. Vises der i stedet en advarsel under
   overskriften, er automatisk afsendelse ikke aktiv – åbn **Teknisk log** nederst i ruden; den
   viser afsendelsesmåden, årsagen og hvert trin i afsendelsen.
5. Sæt evt. `CC_ADDRESSES` til jeres egen adresse under test for at se den udgående rapport.

## Trin 5 – Udrul centralt via Microsoft 365 Admin Center

1. Log ind på <https://admin.microsoft.com> → **Settings → Integrated apps → Upload custom apps**.
2. **App type:** Office Add-in → **Upload manifest file (.xml) from device** → vælg manifestet.
3. Tildel brugere (alle, eller udvalgte grupper) → **Accept permissions** → **Finish deployment**.
4. Udrulningen kan tage op til 24 timer, typisk et par timer.

Ciscos eget add-in kan fjernes fra *Integrated apps*, når det nye er verificeret. De to kan også
køre side om side i en overgangsperiode.

## Trin 6 – Verificér i Talos Email Status Portal

Log ind på <https://talosintelligence.com/email_status_portal> med en Cisco-konto, der er
registreret for jeres domæne. Rapporter fra brugerne vises som *Submissions* inden for få
minutter. Cisco sender ikke kvitteringsmails.

## Opdatering

Ret `<Version>` i `manifest.xml` (fx `1.0.1.0`), og upload det igen under *Integrated apps* →
vælg appen → **Update**. Ændringer i `config.js`, HTML, CSS og JS kræver kun ny upload til
webstedet – Outlook henter altid den seneste version derfra.

## Fejlfinding

| Symptom                                                                    | Årsag og løsning                                                                                                                                                                    |
| -------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Advarslen *Automatisk afsendelse er ikke tilgængelig …* vises               | Åbn *Teknisk log*. Linjen *Send mode* viser årsagen: `CLIENT_ID` mangler (trin 2–3), Outlook-klienten mangler Mailbox 1.14 eller NestedAppAuth 1.1 (opdater klienten), eller `lib/msal-browser.min.js` blev ikke indlæst (404). Rapporter virker stadig via ny mail, medmindre `COMPOSE_FALLBACK` er `false`. |
| Fejl med **AADSTS65001** / "mangler samtykke"                               | Admin consent er ikke givet. Trin 2, punkt 5.                                                                                                                                         |
| Fejl med **AADSTS50011** (redirect URI)                                     | `brk-multihub://<host>` matcher ikke det hostnavn, `taskpane.html` er publiceret på. Skal være identisk med origin (inkl. port).                                                     |
| Rapporten sendes, men mailen kunne ikke flyttes                            | `Mail.ReadWrite` mangler eller er ikke godkendt. Giv rettigheden, eller sæt `MOVE_AFTER_REPORT: false`.                                                                              |
| Store mails åbner altid en ny mail                                         | Forventet over 3 MB (`MAX_GRAPH_ATTACHMENT_BYTES`). Microsoft Graph afviser større vedhæftninger i `sendMail`.                                                                        |
| Knappen *Rapportér til Cisco* vises ikke                                   | Udrulningen er ikke slået igennem endnu (op til 24 t), læseruden er slået fra i Outlook, eller manifestet er ikke tildelt brugeren.                                                    |
| Ruden er på engelsk, selv om Outlook er dansk                              | Brugeren har valgt et sprog i ruden (vælg *Automatisk* igen), eller `LANGUAGE` i `config.js` er sat fast. *Teknisk log* viser linjen *Language* med årsagen.                            |
| Ruden er tom / fejl om `Office is not defined`                             | `https://appsforoffice.microsoft.com` er blokeret af proxy/CSP.                                                                                                                       |
| Ny mail åbnes uden vedhæftning                                             | Mailen er en kladde eller endnu ikke gemt på serveren. Rapportér kun modtagne mails.                                                                                                  |

Alle hændelser logges også til browserkonsollen med præfikset `[CiscoReporter]`.
Den engelske fejlfindingsvejledning (`docs/TROUBLESHOOTING.md`) er mere udførlig.

## Forskelle i forhold til Ciscos eget add-in

- Simulerede phishing-mails fra *Cisco Secure Awareness* håndteres ikke særskilt – de
  rapporteres som almindelig phishing.
- Der er seks kategorier (Ciscos add-in har tre: spam/phish/virus, legitim, marketing). Skjul dem,
  I ikke ønsker, med `enabled: false`.
- Tilføjelsesprogrammet er Cisco-eksempelkode under *Cisco Sample Code License 1.1* – ikke et
  Cisco-produkt, og det supporteres ikke af Cisco TAC. Rapporterne behandles dog på samme måde
  hos Talos, da de sendes til de officielle adresser i samme format.
