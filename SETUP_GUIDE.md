# Student PDF Splitter — Setup Guide

## What This Does

1. You upload a single scanned PDF (all students' papers) to Google Drive
2. Open the sidebar tool in your Google Sheet
3. Assign page ranges to each student (takes ~30 sec per student)
4. Click **Split PDFs & Link to Sheet** — the tool:
   - Splits the PDF into individual files named after each student
   - Saves them in the same Drive folder as the original
   - Writes a clickable hyperlink into your chosen column in the sheet

---

## Files You Need

| File | Purpose |
|------|---------|
| `Code.gs` | Main script — menu, sheet reading, linking |
| `PDFSplitter.gs` | PDF splitting engine |
| `Sidebar.html` | The sidebar UI |

---

## Step-by-Step Setup

### 1. Open Your Google Sheet

Open the Google Sheet where your student marks are stored.

### 2. Open Apps Script

- Click **Extensions → Apps Script**
- This opens the Apps Script editor

### 3. Add the Files

You need **3 files** in the Apps Script project:

**File 1 — Code.gs** (rename the default `Code.gs` that already exists)
- Paste the contents of `Code.gs` into it

**File 2 — PDFSplitter.gs** (click the `+` next to Files → Script)
- Name it `PDFSplitter`
- Paste the contents of `PDFSplitter.gs`

**File 3 — Sidebar.html** (click the `+` next to Files → HTML)
- Name it `Sidebar`
- Paste the contents of `Sidebar.html`

### 4. Configure Your Column Settings

At the top of `Code.gs`, edit the CONFIG block:

```javascript
const CONFIG = {
  STUDENT_NAME_COLUMN: 'B',   // ← Column letter with student names
  LINK_COLUMN:         'H',   // ← Column where links will be written
  DATA_START_ROW:      2,     // ← First row of data (2 if row 1 is a header)
  SHEET_NAME:          '',    // ← Sheet tab name, or leave blank for active sheet
};
```

### 5. Save and Authorize

- Press **Ctrl+S** (or Cmd+S) to save
- Click **Run → onOpen** to trigger authorization
- Google will ask you to grant permissions — click **Allow**
  - It needs: access to Drive (read/write files) and Sheets (write links)

### 6. Reload Your Sheet

Close the Apps Script tab and reload your Google Sheet.
You should see a new menu: **📄 PDF Splitter**

---

## Using the Tool

### Every Time You Have Papers to Process:

**1. Scan all papers into one PDF**
Scan them in student name order (matching your sheet) to make page assignment easier.

**2. Upload the PDF to Google Drive**
Put it in whatever folder you want the split files to live in.

**3. Get the File ID**
Open the PDF in Drive → look at the URL:
```
https://drive.google.com/file/d/[THIS-IS-THE-FILE-ID]/view
```
Copy that long string.

**4. Open the tool**
In your Sheet: **📄 PDF Splitter → Open PDF Splitter Tool**

**5. Load the PDF**
Paste the file ID and click **Load**. The tool will confirm the file name.

**6. Add bookmarks**
For each student:
- Select their name from the dropdown (pulled from your sheet) or type it
- Enter their **Start Page** and **End Page**
- Click **+ Add Bookmark**

**7. Split & Link**
Click **▶ Split PDFs & Link to Sheet**

The tool will:
- Create `StudentName.pdf` for each student in the same Drive folder
- Add a clickable link in column H (or whichever you configured) on their row

---

## Tips

- **Page ranges can overlap** if a student's work spans the same pages as another (unusual but handled)
- **Student names must match exactly** what's in column B of your sheet (case-insensitive)
- **If a file already exists** with the same name, it will be replaced automatically
- **If a student name isn't found** in the sheet, the script logs a warning but still creates the PDF file

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "Cannot access file" | Make sure the PDF is in your Drive (not just shared with you) |
| "File is not a PDF" | Only PDF files are supported |
| "Could not split PDF" | Re-save the PDF: open it → File → Print → Save as PDF → use this new file |
| Links not appearing | Check that `STUDENT_NAME_COLUMN` and `LINK_COLUMN` are set correctly |
| Menu not showing | Re-run `onOpen` from the Apps Script editor |

---

## Advanced: Faster PDF Splitting (Optional)

The built-in PDF splitter handles most scanner-produced PDFs. If you run into issues with complex PDFs, you can deploy a helper Web App:

1. In Apps Script, create a new file `WebApp.gs`
2. Add a `doPost(e)` function that uses the `pdf-lib` library
3. Deploy it as a Web App (Execute as: Me, Access: Anyone)
4. Paste the Web App URL into `PDFSplitter.gs` at the `WEB_APP_URL` line

This is optional — the built-in parser works for the vast majority of scanned PDFs.
