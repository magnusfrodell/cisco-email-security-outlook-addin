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

  /* Application (client) ID of your Entra app registration (docs/DEPLOYMENT.md, step 2).
   * Empty string = automatic send is off. The add-in then opens a new message with the
   * reported mail attached, and the user presses Send.
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

  /* Extra recipients, e.g. an internal SOC mailbox. Added as Cc on every report.
   * Example: ["soc@example.com"]
   */
  CC_ADDRESSES: [],

  /* Subject of the report mail. {category} is replaced with the category label
   * in the active language.
   */
  SUBJECT: "Cisco Secure Email submission – {category}",

  /* Graph sendMail accepts at most ~3 MB of attachment inline. Larger mails fall back
   * to a new message automatically. Cisco accepts at most 10 MB per submission.
   */
  MAX_GRAPH_ATTACHMENT_BYTES: 3 * 1024 * 1024,

  /* ======================================================================
   * 3. Categories
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
