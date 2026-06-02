// ============================================================
// PDFSPLITTER.GS
// Server-side functions for the PDF Splitter sidebar.
// PDF splitting itself happens client-side via pdf-lib (CDN).
// This file handles Drive upload and sheet link writing.
// ============================================================

/**
 * Receives a split student PDF (already processed by pdf-lib in the sidebar),
 * uploads it to Drive, finds the student's row in the sheet, and writes
 * a HYPERLINK formula - all in one call to minimise round-trips.
 *
 * @param {string} base64Pdf   - Base64-encoded PDF bytes from pdf-lib
 * @param {string} studentName - Student's name (used as filename + row lookup)
 * @param {string} pdfFileId   - Original scanned PDF file ID (to find folder)
 * @param {Object} config      - {sheetName, nameCol, idCol, linkCol, dataStartRow}
 * @returns {Object}            - {success, studentName, error?}
 */
function uploadAndLink(base64Pdf, studentName, pdfFileId, config) {
  try {
    // ── Upload to Drive ──────────────────────────────────────
    var originalFile = DriveApp.getFileById(pdfFileId);
    var parentFolder = originalFile.getParents().next();
    var fileName     = sanitizeFilename(studentName) + '.pdf';

    var bytes = Utilities.base64Decode(base64Pdf);
    var blob  = Utilities.newBlob(bytes, 'application/pdf', fileName);

    // Replace any existing file with the same name
    var existing = parentFolder.getFilesByName(fileName);
    while (existing.hasNext()) { existing.next().setTrashed(true); }

    var newFile = parentFolder.createFile(blob);
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = newFile.getUrl();

    // ── Write link to sheet ──────────────────────────────────
    var sheet    = getSheet(config.sheetName);
    var lastRow  = sheet.getLastRow();
    var nameCol  = colLetterToIndex(config.nameCol);
    var linkCol  = colLetterToIndex(config.linkCol);
    var found    = false;

    for (var row = config.dataStartRow; row <= lastRow; row++) {
      var cellVal = String(sheet.getRange(row, nameCol).getValue()).trim();
      if (cellVal.toLowerCase() === studentName.trim().toLowerCase()) {
        sheet.getRange(row, linkCol)
             .setFormula('=HYPERLINK("' + fileUrl + '","Link")');
        found = true;
        break;
      }
    }

    if (!found) {
      Logger.log('No matching row for student: ' + studentName);
    }

    return { success: true, studentName: studentName };

  } catch (e) {
    return { success: false, studentName: studentName, error: e.message };
  }
}

/**
 * Writes raw scores into the configured score column.
 * scores: [{name, score}]  — only students with a score entered are passed.
 * config: {sheetName, nameCol, scoreCol, dataStartRow}
 * Returns {written: n}
 */
function writeScoresToSheet(scores, config) {
  if (!config.scoreCol) throw new Error('No score column configured.');

  var sheet     = getSheet(config.sheetName);
  var lastRow   = sheet.getLastRow();
  var nameColIdx  = colLetterToIndex(config.nameCol);
  var scoreColIdx = colLetterToIndex(config.scoreCol);
  var written   = 0;

  for (var s = 0; s < scores.length; s++) {
    var targetName = scores[s].name.trim().toLowerCase();
    var scoreVal   = Number(scores[s].score);

    for (var row = config.dataStartRow; row <= lastRow; row++) {
      var cellName = String(sheet.getRange(row, nameColIdx).getValue()).trim().toLowerCase();
      if (cellName === targetName) {
        sheet.getRange(row, scoreColIdx).setValue(scoreVal);
        written++;
        break;
      }
    }
  }

  return { written: written };
}

/**
 * Writes Y or N into the outcomes sheet for each student/outcome pair.
 *
 * results: [{
 *   studentName,          // full name from main sheet
 *   sheetName,            // outcomes sheet tab name
 *   outcomes: [{text, achieved}]  // achieved = true → Y, false → N
 * }]
 *
 * Outcomes sheet layout:
 *   Row 1: student first names starting at column C (col index 3)
 *   Col B (index 2): outcome text, starting from row 2
 *
 * Student matching: fuzzy match on first name (col C onward in row 1)
 * Outcome matching: fuzzy match on outcome text in col B
 */
function writeOutcomesToSheet(results) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var written = 0;
  var errors  = [];

  // Cache sheets to avoid repeated lookups
  var sheetCache = {};

  for (var r = 0; r < results.length; r++) {
    var result    = results[r];
    var sheetName = result.sheetName;

    if (!sheetCache[sheetName]) {
      var s = ss.getSheetByName(sheetName);
      if (!s) { errors.push('Sheet not found: ' + sheetName); continue; }
      sheetCache[sheetName] = s;
    }
    var sheet   = sheetCache[sheetName];
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();

    // Find student column — row 1, starting at col C (3)
    var studentCol = findStudentColumn(sheet, result.studentName, lastCol);
    if (!studentCol) {
      errors.push('Student not found in outcomes sheet: ' + result.studentName);
      continue;
    }

    // Build outcome text list from col B (2), rows 2 onward
    var outcomeTexts = [];
    for (var row = 2; row <= lastRow; row++) {
      outcomeTexts.push({
        row:  row,
        text: String(sheet.getRange(row, 2).getValue()).trim()
      });
    }

    // Write Y or N for each outcome
    for (var o = 0; o < result.outcomes.length; o++) {
      var outcomeResult = result.outcomes[o];
      var bestRow = fuzzyMatchOutcome(outcomeResult.text, outcomeTexts);
      if (bestRow) {
        sheet.getRange(bestRow, studentCol).setValue(outcomeResult.achieved ? 'Y' : 'N');
        written++;
      } else {
        errors.push('No match for outcome: "' + outcomeResult.text + '"');
      }
    }
  }

  return { written: written, errors: errors };
}

/**
 * Finds the column of a student in row 1 of the outcomes sheet.
 * Matches on first name only (case-insensitive).
 */
function findStudentColumn(sheet, fullName, lastCol) {
  // Try matching first name against row 1
  var firstName = fullName.trim().split(/\s+/)[0].toLowerCase();
  for (var col = 3; col <= lastCol; col++) {
    var cell = String(sheet.getRange(1, col).getValue()).trim().toLowerCase();
    if (cell === firstName || cell === fullName.toLowerCase()) return col;
  }
  return null;
}

/**
 * Fuzzy-matches outcomeText against a list of {row, text} objects.
 * Returns the row number of the best match above a similarity threshold.
 * Uses word overlap scoring — robust to minor wording differences.
 */
function fuzzyMatchOutcome(queryText, outcomeTexts) {
  var queryWords = tokenise(queryText);
  var bestScore  = 0;
  var bestRow    = null;
  var THRESHOLD  = 0.35;  // minimum overlap to count as a match

  for (var i = 0; i < outcomeTexts.length; i++) {
    var candidate = outcomeTexts[i];
    if (!candidate.text) continue;
    var score = wordOverlapScore(queryWords, tokenise(candidate.text));
    if (score > bestScore) { bestScore = score; bestRow = candidate.row; }
  }

  return bestScore >= THRESHOLD ? bestRow : null;
}

function tokenise(str) {
  // Lowercase, remove punctuation, split on whitespace, filter stop words
  var stops = { 'and':1,'the':1,'a':1,'an':1,'of':1,'to':1,'in':1,'for':1,
                'by':1,'or':1,'that':1,'is':1,'are':1,'using':1,'with':1 };
  return str.toLowerCase().replace(/[^a-z0-9\s]/g, ' ').split(/\s+/)
    .filter(function(w) { return w.length > 2 && !stops[w]; });
}

function wordOverlapScore(wordsA, wordsB) {
  if (!wordsA.length || !wordsB.length) return 0;
  var setB = {};
  wordsB.forEach(function(w) { setB[w] = 1; });
  var matches = 0;
  wordsA.forEach(function(w) { if (setB[w]) matches++; });
  // Jaccard-like: intersection / union
  var union = wordsA.length + wordsB.length - matches;
  return union > 0 ? matches / union : 0;
}
