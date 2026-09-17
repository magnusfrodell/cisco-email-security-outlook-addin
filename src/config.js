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
 * Configuration for Cisco Email Security Reporter (Outlook add-in).
 *
 * This is the only file to edit at installation time. Texts live in src/locales/<lang>.js.
 * Do not edit taskpane.js. See docs/CONFIGURATION.md and docs/LOCALIZATION.md.
 */
window.REPORTER_CONFIG = {

  /* ======================================================================
   * 1. Language
   * ==================================================================== */

  /* "auto" follows the user's Outlook display language (da-DK -> da, sv-SE -> sv, …).
   * Or force one language for everybody: "da", "en" or "sv" – any file in src/locales/.
   */
  LANGUAGE: "auto",

  /* Used when "auto" cannot match the Outlook language to a locale file. */
  DEFAULT_LANGUAGE: "en",

  /* Show a language selector in the pane. A user's choice is stored in their
   * mailbox (roamingSettings) and overrides LANGUAGE for that user only.
   */
  SHOW_LANGUAGE_SELECTOR: true,

  /* ======================================================================
   * 2. Sending
   * ==================================================================== */

  /* "graph"   (default): reports are sent automatically through Microsoft Graph on the
   *                      user's behalf – one click, no new message. Needs CLIENT_ID.
   * "compose": every report opens a new message with the reported mail attached and
   *            the user presses Send. No app registration needed.
   */
  SEND_MODE: "graph",

  /* If automatic send is unavailable (CLIENT_ID missing, unsupported Outlook client) or
   * a send fails, fall back to a new message. false = show an error instead and never
   * open a new message.
   */
  COMPOSE_FALLBACK: true,

  /* Application (client) ID of the Entra app registration that automatic send uses
   * (docs/DEPLOYMENT.md, step 2). Required for SEND_MODE "graph". Leave empty only for
   * SEND_MODE "compose" – with "graph" an empty value shows a warning in the pane.
   */
  CLIENT_ID: "",

  /* Single-tenant registration: "https://login.microsoftonline.com/<tenant-id>"
   * Multi-tenant: "https://login.microsoftonline.com/organizations"
   */
  AUTHORITY: "https://login.microsoftonline.com/organizations",

  /* Move the message after a successful report (needs the Mail.ReadWrite permission):
   * spam / phishing / virus / marketing  -> Junk Email
   * legitimate / not marketing           -> Inbox
   */
  MOVE_AFTER_REPORT: true,

  /* Default for "Keep a copy in Sent Items". Users can change it in the pane;
   * the choice is stored in their mailbox (roamingSettings).
   */
  SAVE_TO_SENT_DEFAULT: false,

  /* Subject of the report mail. Placeholders: {category} (label in the active language),
   * {subject} (original subject, shortened), {sender} (original sender address),
   * {reporter} (the reporting user's address).
   */
  SUBJECT: "Cisco Secure Email submission – {category}: {subject}",

  /* Append a machine-readable block to the report body (category, Cisco address, reporter,
   * original subject/sender/date, Message-ID, client). Meant for the SOC copy; Cisco
   * ignores the body and reads the attached .eml.
   */
  REPORT_DETAILS: true,

  /* ======================================================================
   * 3. SOC visibility (set by the administrator, not visible to users)
   * ==================================================================== */

  /* Every report is also delivered to these addresses, e.g. a SOC mailbox or a
   * SIEM ingestion address. Applies to both send modes. Example: ["soc@example.com"]
   */
  SOC_ADDRESSES: [],

  /* "cc": the SOC address is visible to Cisco on the report. "bcc": hidden. */
  SOC_COPY_MODE: "cc",

  /* Tell the user in the status text that a copy went to the SOC. */
  SOC_SHOW_IN_STATUS: false,

  /* Send reports FROM a shared mailbox (e.g. the SOC mailbox) instead of the user's own,
   * so that every submission is tracked under that address in the Talos Email Status
   * Portal. Automatic send only. Requires: the shared mailbox exists, each user has
   * "Send As" permission on it (Send on Behalf also works but marks the mail "on behalf of"),
   * and the Entra app has the delegated permission Mail.Send.Shared (instead of Mail.Send)
   * with admin consent. Empty = send as the user.
   * In the compose fallback the report is always sent from the user's own mailbox.
   */
  SEND_AS: "",

  /* Graph sendMail accepts at most ~3 MB of attachment inline. Larger mails fall back
   * to a new message automatically. Cisco accepts at most 10 MB per submission.
   */
  MAX_GRAPH_ATTACHMENT_BYTES: 3 * 1024 * 1024,

  /* ======================================================================
   * 4. Categories
   *    id      : key used in the locale files for label and hint
   *    group   : "missed" | "false_positive" (which panel the button is in)
   *    address : Cisco's official submission address – do not change
   *    moveTo  : "junkemail" | "inbox" | null
   *    enabled : false hides the button without removing the entry
   * ==================================================================== */
  CATEGORIES: [
    { id: "spam",    group: "missed",         enabled: true, address: "spam@access.ironport.com",    moveTo: "junkemail" },
    { id: "phish",   group: "missed",         enabled: true, address: "phish@access.ironport.com",   moveTo: "junkemail" },
    { id: "virus",   group: "missed",         enabled: true, address: "virus@access.ironport.com",   moveTo: "junkemail" },
    { id: "ads",     group: "missed",         enabled: true, address: "ads@access.ironport.com",     moveTo: "junkemail" },
    { id: "ham",     group: "false_positive", enabled: true, address: "ham@access.ironport.com",     moveTo: "inbox" },
    { id: "not_ads", group: "false_positive", enabled: true, address: "not_ads@access.ironport.com", moveTo: "inbox" }
  ]
};
