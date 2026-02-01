# Google Form → PDF in Dropbox Automation Guide

## What this does
Whenever someone submits your Google Form, the system automatically:

1. Takes their answers
2. Creates a nicely formatted Google Doc (with a title + key fields like Name/Email + a clean “Answers” section)
3. Exports that Doc as a PDF
4. Uploads the PDF into Dropbox, organized into folders by the submitter’s name (and optionally by year)

You can also regenerate a PDF for an old submission without resubmitting the form.

---

## What you need before you start

- A Google Form that stores responses in a Google Sheet
- Access to Dropbox
- The Apps Script (already created) inside the Google Sheet
- A Google Doc “template” file (the script copies it each time)

---

# Part 1 — The Google template document

## Open the Google Doc template

In Google Drive, open the template document your automation uses.

It includes placeholders like:

- `{{DOC_TITLE}}`
- `{{NAME}}`
- `{{EMAIL}}`
- `{{RECOMMENDATION}}` (optional)
- `{{TIMESTAMP}}`
- `{{ANSWERS}}`

### Editing rules

✅ You *can* change fonts, spacing, headings, add lines, add a logo, etc.  
❌ Don’t delete the placeholders unless you also update the script.

### Suggested layout

- Big title at top
- Small “info block” underneath (Name, Email, Letter of Recommendation Contact, Timestamp)
- Divider line
- “Answers:” section

---

# Part 2 — Dropbox setup

## Confirm where files will go in Dropbox

Dropbox apps typically upload to:

```
Dropbox → Apps → <Your App Name>
```

If you want files to land in a normal existing Dropbox folder (not under Apps), your Dropbox app must be set to **Full Dropbox access**, and you must generate a new access token.

### Important behavior

- With **App Folder** access, everything stays under `Apps/<AppName>`
- With **Full Dropbox**, paths start from your Dropbox root

---

## Dropbox token (important)

The script uses a **Dropbox access token**. This is basically a password.

✅ Store it in **Apps Script → Script Properties**  
❌ Don’t email it  
❌ Don’t paste it into chat  
❌ Don’t commit it to git

---

# Part 3 — Google Sheets script setup

## Open the response spreadsheet

1. Open the Google Form
2. Go to **Responses**
3. Click the green Sheets icon

---

## Open Apps Script

Inside the spreadsheet:

```
Extensions → Apps Script
```

This opens the automation code.

---

# Part 4 — Configuration (the only technical-ish part)

## Update the CONFIG section

In the script you’ll see a `CONFIG` object. This controls:

- Template Doc ID
- Dropbox base folder
- Which columns are promoted (Name, Email, Recommendation)
- Whether blank answers appear

### Header matching matters

Values in `KEY_FIELDS` must match **exactly** the column headers in row 1 of your response sheet.

Example:

If the sheet header is:

```
Email Address
```

Then CONFIG must contain:

```js
EMAIL: "Email Address"
```

---

## Dropbox base folder

Example:

```js
DROPBOX_BASE_FOLDER: "/Photos/Sample Album"
```

The script may automatically append:

- Current year
- Submitter name

Result:

```
/Photos/Sample Album/2026/Jane Doe/
```

Folders are created automatically.

---

# Part 5 — First-time authorization

Google blocks new scripts until approved.

In Apps Script:

1. Select any function (for example `reprocessSelectedRow`)
2. Click ▶ Run
3. Click **Review permissions**
4. Choose your Google account
5. Click **Advanced**
6. Click **Go to project (unsafe)** (this is normal)
7. Click **Allow**

This grants access to:

- Google Drive
- Google Docs
- PDF export
- Dropbox uploads

You only do this once.

---

# Part 6 — Turning on automation (Trigger)

## Confirm trigger exists

Apps Script → left sidebar → **Triggers**

You should see:

- Function: `onFormSubmit`
- Source: From spreadsheet
- Event: On form submit

If missing:

1. Click **Add Trigger**
2. Choose:
   - Function: `onFormSubmit`
   - Source: From spreadsheet
   - Event: On form submit
3. Save

Now every new form submission runs automatically.

---

# Part 7 — Testing without filling the form again

## Reprocess an existing row

You don’t need to resubmit the form.

1. Open the response sheet
2. Click any cell in the row you want (not the header)
3. Open Apps Script
4. Select `reprocessSelectedRow`
5. Click ▶ Run

This:

- Generates the Doc
- Exports PDF
- Uploads to Dropbox

Creates a new file each time (never overwrites).

---

# Part 8 — Output format

The generated PDF:

- Questions appear as headings
- Answers are indented
- Clean spacing between each Q/A
- Links (Drive, URLs) are clickable
- Key fields appear at top (Name, Email, Recommendation, Timestamp)

It reads like a real application packet, not a spreadsheet dump.

---

# Safety: Nothing is deleted

Dropbox:

- Files are uploaded using “add” mode
- Existing files are never overwritten
- No delete API calls exist in the script

Google:

- Temporary Docs may be trashed after PDF creation
- Remove that line if you want to keep them

---

# Troubleshooting

## Nothing appears in Dropbox

Check:

1. Trigger exists
2. Script authorized
3. Dropbox token correct
4. Folder path starts with `/`

Then open Apps Script → **Executions** to see errors.

---

## Files go into Apps folder

That means your Dropbox app is still “App Folder” access.

Switch to **Full Dropbox**, regenerate token, update script.

---

## Fields missing in PDF

Almost always header mismatch:

Sheet header must match CONFIG exactly.

---

## Formatting looks strange

Edit the template Doc:

- Adjust Heading 3 style
- Add divider
- Change fonts
- Resize title

No code changes needed.

---

# Daily-use cheat sheet

### To confirm it’s working
Submit a form OR run `reprocessSelectedRow`.

### To regenerate a PDF
1. Click a row
2. Run `reprocessSelectedRow`
3. Check Dropbox

That’s it.

---

If needed later, the system can be extended to:

- Add month folders
- Attach uploaded files
- Page breaks between sections
- Two-column headers
- Submission IDs

But none of that is required for normal operation.
