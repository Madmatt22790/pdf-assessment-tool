// ============================================================
// CODE.GS — Student PDF Splitter + Assessment Generator + PTI Generator
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Assessment Tools')
    .addItem('Generate Assessment PDFs', 'showGeneratorSidebar')
    .addItem('Split & Link Scanned PDFs', 'showSplitterSidebar')
    .addSeparator()
    .addItem('Generate PTI Documents', 'showPtiSidebar')
    .addToUi();
}

// ── GitHub source ────────────────────────────────────────────
// HTML files are fetched from GitHub so all classes stay in sync automatically.
// Update GITHUB_RAW to point to your own repository.
var GITHUB_RAW = 'https://raw.githubusercontent.com/Madmatt22790/pdf-assessment-tool/main/';

function fetchSidebarHtml_(filename) {
  try {
    return UrlFetchApp.fetch(GITHUB_RAW + filename).getContentText();
  } catch (e) {
    throw new Error('Could not load ' + filename + ' from GitHub: ' + e.message);
  }
}

function showSplitterSidebar() {
  var html = HtmlService.createHtmlOutput(fetchSidebarHtml_('Sidebar.html'))
    .setTitle('Student PDF Splitter')
    .setWidth(440);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showGeneratorSidebar() {
  var html = HtmlService.createHtmlOutput(fetchSidebarHtml_('AssessmentGenerator.html'))
    .setTitle('Assessment Generator')
    .setWidth(440);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showPtiSidebar() {
  var html = HtmlService.createHtmlOutput(fetchSidebarHtml_('PTIGenerator.html'))
    .setTitle('PTI Document Generator')
    .setWidth(440);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── Sheet helpers ────────────────────────────────────────────

function getSheetNames() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) { throw new Error('No active spreadsheet found.'); }
    return ss.getSheets().map(function(s) { return s.getName(); });
  } catch (e) {
    throw new Error('getSheetNames failed: ' + e.message);
  }
}

function getActiveSheetName() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  } catch (e) {
    return null;
  }
}

function getSheetNamesWithActive() {
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var names  = ss.getSheets().map(function(s) { return s.getName(); });
    var active = ss.getActiveSheet().getName();
    return { names: names, active: active };
  } catch (e) {
    throw new Error('getSheetNamesWithActive failed: ' + e.message);
  }
}

/**
 * Returns all students as [{name, id}] using sidebar-provided config.
 * Stops reading at the first completely blank row — handles sheets where
 * multiple classes are stacked with a blank row separator between them.
 * config: {sheetName, nameCol, idCol, dataStartRow}
 */
function getStudentsWithConfig(config) {
  var sheet    = getSheet(config.sheetName);
  var lastRow  = sheet.getLastRow();
  if (lastRow < config.dataStartRow) { return []; }

  var nameColIdx = colLetterToIndex(config.nameCol);
  var idColIdx   = colLetterToIndex(config.idCol);
  var lastCol    = sheet.getLastColumn();
  var count      = lastRow - config.dataStartRow + 1;

  // Fetch the full row range so we can detect completely blank rows
  var allVals  = sheet.getRange(config.dataStartRow, 1, count, lastCol).getValues();

  var students = [];
  for (var i = 0; i < allVals.length; i++) {
    var row = allVals[i];

    // Check if every cell in this row is blank — if so, stop
    var rowIsBlank = row.every(function(cell) {
      return String(cell).trim() === '';
    });
    if (rowIsBlank) break;

    var name = String(row[nameColIdx - 1]).trim();
    var id   = String(row[idColIdx   - 1]).trim();
    if (name !== '' && id !== '') {
      students.push({ name: name, id: id });
    }
  }
  return students;
}

// ── PDF access ───────────────────────────────────────────────

function validatePdfFile(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var mime = file.getMimeType();
    if (mime !== 'application/pdf') {
      throw new Error('File is not a PDF (detected: ' + mime + ')');
    }
    return { name: file.getName(), id: file.getId() };
  } catch (e) {
    throw new Error('Cannot access file: ' + e.message);
  }
}

function getPdfAsBase64(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var str  = blob.getDataAsString('ISO-8859-1');
    return Utilities.base64Encode(str);
  } catch (e) {
    throw new Error('Could not fetch PDF: ' + (e && e.message ? e.message : e));
  }
}

/**
 * Returns the file size in bytes without downloading the file content.
 * Used by the chunked download path in the sidebar.
 */
function getPdfFileSize(fileId) {
  try {
    return DriveApp.getFileById(fileId).getSize();
  } catch (e) {
    throw new Error('Could not get file size: ' + e.message);
  }
}

/**
 * Downloads a byte range of a Drive file using the Drive API with OAuth,
 * and returns it as a base64 string.
 * This lets the sidebar fetch large PDFs in small pieces without exhausting
 * Apps Script memory.
 */
function getPdfChunk(fileId, offset, length) {
  try {
    var token = ScriptApp.getOAuthToken();
    var url   = 'https://www.googleapis.com/drive/v3/files/' + fileId + '?alt=media';
    var resp  = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token,
        'Range': 'bytes=' + offset + '-' + (offset + length - 1)
      },
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code !== 206 && code !== 200) {
      throw new Error('HTTP ' + code);
    }
    return Utilities.base64Encode(resp.getContent());
  } catch (e) {
    throw new Error('Could not fetch chunk at offset ' + offset + ': ' + e.message);
  }
}

/**
 * Returns the number of pages in a PDF.
 * Scans the full file for all /Count entries and returns the largest,
 * which is always the root Pages node total.
 */
function getPdfPageCount(fileId) {
  try {
    // Read the blob as a single-byte string (ISO-8859-1) so each byte maps
    // directly to a character. This is much faster and avoids manual byte
    // concatenation loops that can crash the Apps Script runtime for large files.
    var blob = DriveApp.getFileById(fileId).getBlob();
    var str = blob.getDataAsString('ISO-8859-1');

    var re = /\/Count\s+(\d+)/g;
    var maxCount = 0;
    var m;
    while ((m = re.exec(str)) !== null) {
      var n = parseInt(m[1], 10);
      if (n > maxCount) maxCount = n;
    }
    if (maxCount > 0) return maxCount;
    throw new Error('Could not determine page count from this PDF.');
  } catch (e) {
    // Re-throw with extra context for the client UI
    throw new Error('getPdfPageCount failed: ' + (e && e.message ? e.message : e));
  }
}

// ── Utilities ────────────────────────────────────────────────

function getSheet(sheetName) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  if (!sheet) { throw new Error('Sheet "' + sheetName + '" not found.'); }
  return sheet;
}

function colLetterToIndex(letter) {
  letter = letter.toUpperCase();
  var index = 0;
  for (var i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index;
}

function sanitizeFilename(name) {
  return name.replace(/[\/\\:*?"<>|]/g, '_');
}
