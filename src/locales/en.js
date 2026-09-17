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
 * English (en) – language file for Cisco Email Security Reporter.
 * Every text the user sees in the pane. See docs/LOCALIZATION.md.
 */
window.REPORTER_LOCALES = window.REPORTER_LOCALES || {};
window.REPORTER_LOCALES.en = {
  name: "English",

  /* Body of the report mail sent to Cisco */
  bodyText: "Report sent from the Outlook add-in. The reported mail is attached as .eml. This text can be ignored.",

  /* Headings of the two button groups */
  groups: {
    missed: "Wrongly delivered to the inbox",
    false_positive: "Wrongly stopped or marked"
  },

  /* Button texts per category id (ids are defined in config.js) */
  categories: {
    spam:    { label: "Spam",                          hint: "Unwanted bulk mail that got through the filter" },
    phish:   { label: "Phishing",                      hint: "Attempt to steal passwords, payment details or similar" },
    virus:   { label: "Virus or malicious attachment", hint: "The mail or an attachment contains malicious code" },
    ads:     { label: "Marketing",                     hint: "Newsletter or advertising that was not marked as marketing" },
    ham:     { label: "Legitimate mail (not spam)",    hint: "A genuine mail that was wrongly marked as spam" },
    not_ads: { label: "Not marketing",                 hint: "A mail that was wrongly marked as marketing or bulk" }
  },

  /* Everything else. {name} is replaced with the current value. */
  strings: {
    title: "Report mail to Cisco",

    chipStarting: "Starting …",
    chipAuto: "Automatic send",
    chipCompose: "Opens a new message to send",

    selectedMail: "Selected mail",
    noMailSelected: "No mail selected",
    selectMailHint: "Select a mail in the list to report it.",
    noSubject: "(no subject)",
    fromPrefix: "From: ",

    missedHelp: "The mail is in the inbox but should have been stopped.",
    falsePositiveHelp: "The mail was wrongly marked as spam or marketing.",
    sentTo: "Sent to {address}",

    settings: "Settings",
    keepCopy: "Keep a copy of the report in Sent Items",
    language: "Language",
    languageAuto: "Automatic (follows Outlook)",
    logSummary: "Technical log",
    footer: "Reports go to Cisco Talos. Track their status in the ",
    footerLink: "Email Status Portal",

    chipOk: "Sent",
    chipWarn: "Note",
    chipErr: "Error",
    chipWait: "Sending",
    chipInfo: "Info",
    chipPressSend: "Press Send",

    sending: "Sending report as {category} …",
    composeOpened: "A new message to {address} has been opened with the reported mail attached. Press Send in the new message to complete the report.",
    reported: "Reported as {category}. Sent to {address}",
    movedTo: " and moved to {folder}",
    folderJunk: "Junk Email",
    folderInbox: "Inbox",
    selectFirst: "Select a mail in the list first.",
    notSentAuto: "The report was not sent automatically: ",
    notSent: "The report was not sent: ",
    retryCompose: "Send manually instead",
    tooLarge: "the mail is {mb} MB, too large for automatic send",
    manualChosen: "manual send chosen",
    reportedMailName: "Reported mail",

    errConsent: "Consent to send mail on your behalf is missing. Ask your administrator to approve the add-in.",
    errGraphDenied: "Microsoft Graph refused the request ({status}). Check that the Mail.Send and Mail.ReadWrite permissions are granted.",
    errGetFile: "The mail could not be retrieved as a file in this version of Outlook.",
    errNotOutlook: "This add-in must be opened from Outlook.",
    errStart: "The add-in could not start: "
  }
};
