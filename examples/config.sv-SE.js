/*
 * EXEMPEL: svensk (sv-SE) version av src/config.js.
 * Kopiera till src/config.js och byt manifeststrängarna enligt docs/LOCALIZATION.md.
 * Inställningarna (CLIENT_ID, CC_ADDRESSES …) är identiska med den danska filen.
 */
window.RAPPORT_CONFIG = {

  /* ======================================================================
   * 1. Afsendelse
   * ==================================================================== */

  /* Application (client) ID fra jeres Entra-appregistrering (docs/DEPLOYMENT.md, trin 2).
   * Tom streng = automatisk afsendelse er slået fra. Add-in'et åbner i stedet en
   * ny mail med den rapporterede besked vedhæftet, og brugeren trykker selv Send.
   */
  CLIENT_ID: "",

  /* Single-tenant appregistrering: "https://login.microsoftonline.com/<tenant-id>"
   * Multi-tenant: "https://login.microsoftonline.com/organizations"
   */
  AUTHORITY: "https://login.microsoftonline.com/organizations",

  /* Flyt mailen når rapporten er sendt (kræver rettigheden Mail.ReadWrite):
   * spam / phishing / virus / marketing  -> Skräppost
   * legitim / inte marknadsföring         -> Inkorgen
   */
  MOVE_AFTER_REPORT: true,

  /* Standardværdi for "Behold en kopi i Sendt post". Brugeren kan selv ændre den
   * i ruden; valget gemmes i brugerens Outlook-indstillinger (roamingSettings).
   */
  SAVE_TO_SENT_DEFAULT: false,

  /* Ekstra modtagere, fx en intern SOC-postkasse. Sættes som Cc på alle rapporter.
   * Eksempel: ["soc@firma.dk"]
   */
  CC_ADDRESSES: [],

  /* Emne og brødtekst på rapport-mailen. {category} erstattes med kategorinavnet. */
  SUBJECT: "Cisco Secure Email submission – {category}",
  BODY_TEXT: "Rapport skickad från Outlook-tillägget. Det rapporterade mejlet är bifogat som .eml. Den här texten kan ignoreras.",

  /* Graph sendMail accepterer højst ca. 3 MB vedhæftning i selve kaldet.
   * Større mails falder automatisk tilbage til manuel afsendelse (ny mail).
   * Cisco accepterer højst 10 MB i alt pr. indsendelse.
   */
  MAX_GRAPH_ATTACHMENT_BYTES: 3 * 1024 * 1024,

  /* ======================================================================
   * 2. Kategorier
   *    address : Ciscos officielle modtageradresse – må ikke ændres.
   *    moveTo  : "junkemail" | "inbox" | null
   *    enabled : false skjuler knappen uden at fjerne konfigurationen.
   * ==================================================================== */
  CATEGORIES: [
    {
      id: "spam", group: "missed", enabled: true,
      label: "Skräppost",
      hint: "Oönskad massutskickad e-post som slapp igenom filtret",
      address: "spam@access.ironport.com", moveTo: "junkemail"
    },
    {
      id: "phish", group: "missed", enabled: true,
      label: "Nätfiske",
      hint: "Försök att lura av dig lösenord, betalningsuppgifter m.m.",
      address: "phish@access.ironport.com", moveTo: "junkemail"
    },
    {
      id: "virus", group: "missed", enabled: true,
      label: "Virus eller skadlig bilaga",
      hint: "Mejlet eller en bilaga innehåller skadlig kod",
      address: "virus@access.ironport.com", moveTo: "junkemail"
    },
    {
      id: "ads", group: "missed", enabled: true,
      label: "Marknadsföring",
      hint: "Nyhetsbrev eller reklam som inte markerades som marknadsföring",
      address: "ads@access.ironport.com", moveTo: "junkemail"
    },
    {
      id: "ham", group: "false_positive", enabled: true,
      label: "Legitimt mejl (inte skräppost)",
      hint: "Ett riktigt mejl som felaktigt markerades som skräppost",
      address: "ham@access.ironport.com", moveTo: "inbox"
    },
    {
      id: "not_ads", group: "false_positive", enabled: true,
      label: "Inte marknadsföring",
      hint: "Ett mejl som felaktigt markerades som marknadsföring eller massutskick",
      address: "not_ads@access.ironport.com", moveTo: "inbox"
    }
  ],

  /* Overskrifter over de to knapgrupper. */
  GROUPS: {
    missed: "Felaktigt levererat till inkorgen",
    false_positive: "Felaktigt stoppat eller markerat"
  },

  /* ======================================================================
   * 3. Tekster i ruden
   *    {navn} i en tekst erstattes af add-in'et med den aktuelle værdi.
   *    Sprogkode bruges til <html lang="…">.
   * ==================================================================== */
  LANG: "sv",
  STRINGS: {
    title: "Rapportera mejl till Cisco",

    chipStarting: "Startar …",
    chipAuto: "Skickas automatiskt",
    chipCompose: "Öppnar nytt mejl för sändning",

    selectedMail: "Valt mejl",
    noMailSelected: "Inget mejl valt",
    selectMailHint: "Välj ett mejl i listan för att rapportera det.",
    noSubject: "(utan ämne)",
    fromPrefix: "Från: ",

    missedHelp: "Mejlet ligger i inkorgen men borde ha stoppats.",
    falsePositiveHelp: "Mejlet markerades felaktigt som skräppost eller marknadsföring.",
    sentTo: "Skickas till {address}",

    settings: "Inställningar",
    keepCopy: "Behåll en kopia av rapporten i Skickat",
    logSummary: "Teknisk logg",
    footer: "Rapporter skickas till Cisco Talos. Följ status på ",
    footerLink: "Email Status Portal",

    chipOk: "Skickat",
    chipWarn: "Obs",
    chipErr: "Fel",
    chipWait: "Skickar",
    chipInfo: "Info",
    chipPressSend: "Tryck Skicka",

    sending: "Skickar rapport som {category} …",
    composeOpened: "Ett nytt mejl till {address} har öppnats med det rapporterade mejlet bifogat. Tryck Skicka i det nya mejlet för att slutföra rapporteringen.",
    reported: "Rapporterat som {category}. Skickat till {address}",
    movedTo: " och flyttat till {folder}",
    folderJunk: "Skräppost",
    folderInbox: "Inkorgen",
    selectFirst: "Välj först ett mejl i listan.",
    notSentAuto: "Rapporten kunde inte skickas automatiskt: ",
    notSent: "Rapporten skickades inte: ",
    retryCompose: "Skicka manuellt i stället",
    tooLarge: "mejlet är {mb} MB och för stort för automatisk sändning",
    manualChosen: "manuell sändning vald",
    reportedMailName: "Rapporterat mejl",

    errConsent: "Det saknas medgivande för att skicka mejl för din räkning. Be er administratör att godkänna tillägget.",
    errGraphDenied: "Outlook nekades av Microsoft Graph ({status}). Kontrollera att behörigheterna Mail.Send och Mail.ReadWrite är beviljade.",
    errGetFile: "Mejlet kunde inte hämtas som fil i den här Outlook-versionen.",
    errNotOutlook: "Det här tillägget måste öppnas från Outlook.",
    errStart: "Tillägget kunde inte starta: "
  }
};
