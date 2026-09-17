/**
 * @license
 * Copyright (c) 2026 Cisco and/or its affiliates.
 *
 * This software is licensed to you under the terms of the Cisco Sample
 * Code License, Version 1.1 (the "License"). You may obtain a copy of the
 * License at
 *
 *                https://developer.cisco.com/docs/licenses
 *
 * All use of the material herein must be in accordance with the terms of
 * the License. All rights not expressly granted by the License are
 * reserved. Unless required by applicable law or agreed to separately in
 * writing, software distributed under the License is distributed on an "AS
 * IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
 * or implied.
 */
/*
 * Svenska (sv) – språkfil för Cisco Email Security Reporter.
 * Alla texter som användaren ser i rutan. Se docs/LOCALIZATION.md.
 */
window.REPORTER_LOCALES = window.REPORTER_LOCALES || {};
window.REPORTER_LOCALES.sv = {
  name: "Svenska",

  /* Brödtext i rapportmejlet till Cisco */
  bodyText: "Rapport skickad från Outlook-tillägget. Det rapporterade mejlet är bifogat som .eml. Den här texten kan ignoreras.",

  /* Rubriker över de två knappgrupperna */
  groups: {
    missed: "Felaktigt levererat till inkorgen",
    false_positive: "Felaktigt stoppat eller markerat"
  },

  /* Knapptexter per kategori-id (id:n definieras i config.js) */
  categories: {
    spam:    { label: "Skräppost",                      hint: "Oönskad massutskickad e-post som slapp igenom filtret" },
    phish:   { label: "Nätfiske",                       hint: "Försök att lura av dig lösenord, betalningsuppgifter m.m." },
    virus:   { label: "Virus eller skadlig bilaga",     hint: "Mejlet eller en bilaga innehåller skadlig kod" },
    ads:     { label: "Marknadsföring",                 hint: "Nyhetsbrev eller reklam som inte markerades som marknadsföring" },
    ham:     { label: "Legitimt mejl (inte skräppost)", hint: "Ett riktigt mejl som felaktigt markerades som skräppost" },
    not_ads: { label: "Inte marknadsföring",            hint: "Ett mejl som felaktigt markerades som marknadsföring eller massutskick" }
  },

  /* Övriga texter. {namn} ersätts med aktuellt värde. */
  strings: {
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
    language: "Språk",
    languageAuto: "Automatiskt (följer Outlook)",
    logSummary: "Teknisk logg",
    footer: "Rapporter skickas till Cisco Talos. Följ status på ",
    footerLink: "Email Status Portal",

    chipOk: "Skickat",
    chipWarn: "Obs",
    chipErr: "Fel",
    chipWait: "Skickar",
    chipInfo: "Info",
    chipPressSend: "Tryck Skicka",
    chipUnavailable: "Inte tillgängligt",
    autoUnavailable: "Automatisk sändning är inte tillgänglig: {reason}. Rapporter öppnas som ett nytt mejl i stället.",
    autoUnavailableNoFallback: "Automatisk sändning är inte tillgänglig: {reason}. Kontakta er IT-support.",
    composeUnavailable: "Den här Outlook-versionen stöder inte tillägget.",
    reasonNotConfigured: "tillägget är inte färdigkonfigurerat (CLIENT_ID saknas)",
    reasonClientUnsupported: "den här Outlook-versionen stöder det inte",
    reasonMsalMissing: "en nödvändig komponent kunde inte läsas in",

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
