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
 * Dansk (da) – sprogfil for Cisco Email Security Reporter.
 * Alle tekster, som brugeren ser i ruden. Se docs/LOCALIZATION.md.
 */
window.REPORTER_LOCALES = window.REPORTER_LOCALES || {};
window.REPORTER_LOCALES.da = {
  name: "Dansk",

  /* Brødtekst i rapport-mailen til Cisco */
  bodyText: "Rapport sendt fra Outlook-tilføjelsesprogrammet. Den rapporterede mail er vedhæftet som .eml. Denne tekst kan ignoreres.",

  /* Overskrifter over de to knapgrupper */
  groups: {
    missed: "Fejlagtigt leveret til indbakken",
    false_positive: "Fejlagtigt stoppet eller markeret"
  },

  /* Knaptekster pr. kategori-id (id'erne er defineret i config.js) */
  categories: {
    spam:    { label: "Spam",                              hint: "Uønsket massemail, der slap igennem filteret" },
    phish:   { label: "Phishing",                          hint: "Forsøg på at franarre adgangskoder, betalingsoplysninger m.m." },
    virus:   { label: "Virus eller skadelig vedhæftning",  hint: "Mailen eller en vedhæftning indeholder skadelig kode" },
    ads:     { label: "Marketing",                         hint: "Nyhedsbrev eller reklame, der ikke blev markeret som marketing" },
    ham:     { label: "Legitim mail (ikke spam)",          hint: "En rigtig mail, der fejlagtigt blev markeret som spam" },
    not_ads: { label: "Ikke marketing",                    hint: "En mail, der fejlagtigt blev markeret som marketing eller bulk" }
  },

  /* Øvrige tekster. {navn} erstattes med den aktuelle værdi. */
  strings: {
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
    language: "Sprog",
    languageAuto: "Automatisk (følger Outlook)",
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
