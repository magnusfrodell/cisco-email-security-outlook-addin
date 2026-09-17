/*
 * Konfiguration for "Rapportér mail til Cisco".
 *
 * Dette er den eneste fil, der skal redigeres ved installation eller oversættelse.
 * Rediger ikke taskpane.js. Se docs/CONFIGURATION.md og docs/LOCALIZATION.md.
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
   * spam / phishing / virus / marketing  -> Uønsket mail
   * legitim / ikke marketing             -> Indbakke
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
  BODY_TEXT: "Rapport sendt fra Outlook-tilføjelsesprogrammet. Den rapporterede mail er vedhæftet som .eml. Denne tekst kan ignoreres.",

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
      label: "Spam",
      hint: "Uønsket massemail, der slap igennem filteret",
      address: "spam@access.ironport.com", moveTo: "junkemail"
    },
    {
      id: "phish", group: "missed", enabled: true,
      label: "Phishing",
      hint: "Forsøg på at franarre adgangskoder, betalingsoplysninger m.m.",
      address: "phish@access.ironport.com", moveTo: "junkemail"
    },
    {
      id: "virus", group: "missed", enabled: true,
      label: "Virus eller skadelig vedhæftning",
      hint: "Mailen eller en vedhæftning indeholder skadelig kode",
      address: "virus@access.ironport.com", moveTo: "junkemail"
    },
    {
      id: "ads", group: "missed", enabled: true,
      label: "Marketing",
      hint: "Nyhedsbrev eller reklame, der ikke blev markeret som marketing",
      address: "ads@access.ironport.com", moveTo: "junkemail"
    },
    {
      id: "ham", group: "false_positive", enabled: true,
      label: "Legitim mail (ikke spam)",
      hint: "En rigtig mail, der fejlagtigt blev markeret som spam",
      address: "ham@access.ironport.com", moveTo: "inbox"
    },
    {
      id: "not_ads", group: "false_positive", enabled: true,
      label: "Ikke marketing",
      hint: "En mail, der fejlagtigt blev markeret som marketing eller bulk",
      address: "not_ads@access.ironport.com", moveTo: "inbox"
    }
  ],

  /* Overskrifter over de to knapgrupper. */
  GROUPS: {
    missed: "Fejlagtigt leveret til indbakken",
    false_positive: "Fejlagtigt stoppet eller markeret"
  },

  /* ======================================================================
   * 3. Tekster i ruden
   *    {navn} i en tekst erstattes af add-in'et med den aktuelle værdi.
   *    Sprogkode bruges til <html lang="…">.
   * ==================================================================== */
  LANG: "da",
  STRINGS: {
    title: "Rapportér mail til Cisco",

    chipStarting: "Starter …",
    chipAuto: "Automatisk afsendelse",
    chipCompose: "Åbner ny mail til afsendelse",

    selectedMail: "Valgt mail",
    noMailSelected: "Ingen mail valgt",
    selectMailHint: "Vælg en mail i listen for at rapportere den.",
    noSubject: "(uden emne)",
    fromPrefix: "Fra: ",

    missedHelp: "Mailen ligger i indbakken, men burde have været stoppet.",
    falsePositiveHelp: "Mailen blev fejlagtigt markeret som spam eller marketing.",
    sentTo: "Sendes til {address}",

    settings: "Indstillinger",
    keepCopy: "Behold en kopi af rapporten i Sendt post",
    logSummary: "Teknisk log",
    footer: "Rapporter sendes til Cisco Talos. Følg status på ",
    footerLink: "Email Status Portal",

    chipOk: "Sendt",
    chipWarn: "Bemærk",
    chipErr: "Fejl",
    chipWait: "Sender",
    chipInfo: "Info",
    chipPressSend: "Tryk Send",

    sending: "Sender rapport som {category} …",
    composeOpened: "En ny mail til {address} er åbnet med den rapporterede mail vedhæftet. Tryk Send i den nye mail for at fuldføre rapporteringen.",
    reported: "Rapporteret som {category}. Sendt til {address}",
    movedTo: " og flyttet til {folder}",
    folderJunk: "Uønsket mail",
    folderInbox: "Indbakke",
    selectFirst: "Vælg først en mail i listen.",
    notSentAuto: "Rapporten blev ikke sendt automatisk: ",
    notSent: "Rapporten blev ikke sendt: ",
    retryCompose: "Send manuelt i stedet",
    tooLarge: "mailen er {mb} MB og for stor til automatisk afsendelse",
    manualChosen: "manuel afsendelse valgt",
    reportedMailName: "Rapporteret mail",

    errConsent: "Der mangler samtykke til at sende mail på dine vegne. Bed jeres administrator om at godkende tilføjelsesprogrammet.",
    errGraphDenied: "Outlook fik afslag fra Microsoft Graph ({status}). Tjek at rettighederne Mail.Send og Mail.ReadWrite er givet.",
    errGetFile: "Mailen kunne ikke hentes som fil i denne Outlook-version.",
    errNotOutlook: "Dette tilføjelsesprogram skal åbnes fra Outlook.",
    errStart: "Tilføjelsesprogrammet kunne ikke starte: "
  }
};
