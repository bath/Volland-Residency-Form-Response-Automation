/**
 * CONFIG - customize these
 */
const CONFIG = {
  TEMPLATE_DOC_ID: "INSERT_DOC_ID_HERE",
  DROPBOX_TOKEN: "INSERT_VALUE_HERE",
  DROPBOX_BASE_FOLDER: `/INSERT_FOLDER_PATHING_HERE/${new Date().getFullYear()}`,

  // The title that appears at the top of the Doc/PDF
  DOC_TITLE: "Submitted Residency Form",

  // These must match your Google Form/Sheet headers exactly
  KEY_FIELDS: {
    NAME: "Name",
    EMAIL: "Email Address",
    LOR_CONTACT: "Letter of Recommendation Contact (Name, Email Address and Phone Number)"
  },

  // Headers you do NOT want repeated in the Answers section
  EXCLUDE_FROM_ANSWERS: ["Timestamp"],

  // If true, blanks will still show as "Question: (blank)"
  INCLUDE_BLANK_ANSWERS: true,
  BLANK_ANSWER_TEXT: "The applicant did not provide an answer to this question.",
};

/**
 * Runs automatically on each new form submit
 * Set up trigger: Triggers -> Add Trigger -> onFormSubmit -> From spreadsheet -> On form submit
 */
function onFormSubmit(e) {
  const values = e.namedValues; // {Header: [value]}

  const timestamp =
    (values["Timestamp"] && values["Timestamp"][0]) ||
    (values["Submitted at"] && values["Submitted at"][0]) ||
    new Date().toISOString();

  const submitterName = safeText((values[CONFIG.KEY_FIELDS.NAME] || ["Unknown"])[0]).trim() || "Unknown";
  const email = safeText((values[CONFIG.KEY_FIELDS.EMAIL] || [""])[0]).trim();
  const lorContact = safeText((values[CONFIG.KEY_FIELDS.LOR_CONTACT] || [""])[0]);

  const docId = createDocFromTemplate_({
    docTitle: CONFIG.DOC_TITLE,
    name: submitterName,
    email,
    timestamp,
    lorContact,
    namedValues: values,
  });


  // Export to PDF
  const fileName = `${CONFIG.DOC_TITLE} - ${sanitizePathSegment(submitterName)} - ${formatDateForFilename(new Date(timestamp))}.pdf`;
  const pdfBlob = DriveApp.getFileById(docId).getAs(MimeType.PDF).setName(fileName);

  // Upload to Dropbox folder named after submitter
  const folderPath = `${CONFIG.DROPBOX_BASE_FOLDER}/${sanitizePathSegment(submitterName)}`;
  ensureDropboxFolder_(folderPath);
  uploadToDropbox_(folderPath, fileName, pdfBlob);

  // Cleanup temp doc
  DriveApp.getFileById(docId).setTrashed(true);
}

function buildAnswersBlock_(namedValues) {
  const exclude = new Set(CONFIG.EXCLUDE_FROM_ANSWERS);

  // Also exclude key fields to avoid duplication
  Object.values(CONFIG.KEY_FIELDS).forEach(h => exclude.add(h));

  const lines = [];

  Object.keys(namedValues).forEach(header => {
    if (exclude.has(header)) return;

    const raw = namedValues[header] || [""];
    // Form can return multiple values for checkboxes / multiple choice etc.
    const answer = raw.filter(Boolean).join(", ").trim();

    if (!CONFIG.INCLUDE_BLANK_ANSWERS && !answer) return;

    lines.push(`${header}: ${answer || CONFIG.BLANK_ANSWER_TEXT}`);
  });

  // If you want a blank line between entries, use "\n\n"
  return lines.join("\n");
}

function createDocFromTemplate_({ docTitle, name, email, timestamp, lorContact, namedValues }) {
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
  const copy = templateFile.makeCopy(`TEMP - ${name} - ${new Date().toISOString()}`);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  body.replaceText("{{DOC_TITLE}}", safeText(docTitle));
  body.replaceText("{{NAME}}", safeText(name));
  body.replaceText("{{EMAIL}}", safeText(email));
  body.replaceText("{{TIMESTAMP}}", safeText(timestamp));
  body.replaceText("{{LOR_CONTACT}}", safeText(lorContact));

  // Insert formatted Q/A instead of raw text
  insertFormattedAnswers_(doc, namedValues);

  doc.saveAndClose();
  return copy.getId();
}

function insertFormattedAnswers_(doc, namedValues) {
  const body = doc.getBody();

  const found = body.findText("{{ANSWERS}}");
  if (!found) throw new Error("{{ANSWERS}} placeholder not found in template");

  const textEl = found.getElement().asText();
  const placeholderParagraph = textEl.getParent().asParagraph();
  textEl.replaceText("\\{\\{ANSWERS\\}\\}", "");

  let insertIndex = body.getChildIndex(placeholderParagraph) + 1;

  const exclude = new Set(CONFIG.EXCLUDE_FROM_ANSWERS || []);
  Object.values(CONFIG.KEY_FIELDS).forEach(h => exclude.add(h));

  const ANSWER_INDENT_POINTS = 18;
  const QUESTION_SPACING_BEFORE = 10;
  const QUESTION_SPACING_AFTER = 2;
  const ANSWER_SPACING_AFTER = 10;

  const blankText = CONFIG.BLANK_ANSWER_TEXT || "User did not answer this question.";

  Object.keys(namedValues).forEach(header => {
    if (exclude.has(header)) return;

    const raw = namedValues[header] || [""];
    const answer = raw
      .map(v => safeText(v).trim())
      .filter(v => v.length > 0)
      .join(", ");

    // If you *don’t* want blanks at all, keep this behavior
    if (!CONFIG.INCLUDE_BLANK_ANSWERS && !answer) return;

    const q = body.insertParagraph(insertIndex++, header);
    q.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    q.setSpacingBefore(QUESTION_SPACING_BEFORE);
    q.setSpacingAfter(QUESTION_SPACING_AFTER);

    // ✅ This is the key change
    const displayAnswer = answer || blankText;

    const a = body.insertParagraph(insertIndex++, displayAnswer);
    a.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    a.setIndentStart(ANSWER_INDENT_POINTS);
    a.setSpacingBefore(0);
    a.setSpacingAfter(ANSWER_SPACING_AFTER);

    linkifyUrlsInParagraph_(a);
  });
}


/**
 * Turns any http(s) URLs in a paragraph into clickable links.
 * Safe to call even if there are no URLs.
 */
function linkifyUrlsInParagraph_(paragraph) {
  const text = paragraph.editAsText().getText();
  const urlRegex = /(https?:\/\/[^\s]+)/g;

  let match;
  while ((match = urlRegex.exec(text)) !== null) {
    const url = match[0];
    const start = match.index;
    const end = start + url.length - 1;

    const t = paragraph.editAsText();
    t.setLinkUrl(start, end, url);
  }
}


/**
 * Dropbox upload helpers
 */
function uploadToDropbox_(folderPath, fileName, blob) {
  const dropboxPath = `${folderPath}/${fileName}`;

  const res = UrlFetchApp.fetch("https://content.dropboxapi.com/2/files/upload", {
    method: "post",
    contentType: "application/octet-stream",
    headers: {
      Authorization: `Bearer ${CONFIG.DROPBOX_TOKEN}`,
      "Dropbox-API-Arg": JSON.stringify({
        path: dropboxPath,
        mode: "add",
        autorename: true,
        mute: false,
      }),
    },
    payload: blob.getBytes(),
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() >= 300) {
    throw new Error(`Dropbox upload failed: ${res.getResponseCode()} ${res.getContentText()}`);
  }
}

function ensureDropboxFolder_(folderPath) {
  const res = UrlFetchApp.fetch("https://api.dropboxapi.com/2/files/create_folder_v2", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: `Bearer ${CONFIG.DROPBOX_TOKEN}` },
    payload: JSON.stringify({ path: folderPath, autorename: false }),
    muteHttpExceptions: true,
  });

  // 409 = already exists
  if (res.getResponseCode() === 409) return;

  if (res.getResponseCode() >= 300) {
    throw new Error(`Dropbox create folder failed: ${res.getResponseCode()} ${res.getContentText()}`);
  }
}

/**
 * Utilities
 */
function safeText(v) {
  return v === null || v === undefined ? "" : String(v);
}

function sanitizePathSegment(s) {
  return safeText(s).replace(/[\/\\:*?"<>|#%\u0000-\u001F]/g, "-").trim() || "Unknown";
}

function formatDateForFilename(d) {
  const pad = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
}

function reprocessSelectedRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const range = sheet.getActiveRange();

  if (!range || range.getRow() === 1) {
    throw new Error("Please select a data row (not the header).");
  }

  const row = range.getRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const namedValues = {};
  headers.forEach((h, i) => namedValues[h] = [values[i]]);

  // Fake the onFormSubmit payload
  onFormSubmit({ namedValues });
}
