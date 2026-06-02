// ============================================================
//  PTI Document Generator — Server-side functions
//  Bound to the same Google Sheet as the other Assessment Tools
// ============================================================

/**
 * Validates a Google Doc template and returns its name.
 * Accepts a full Drive URL or bare file ID.
 */
function validatePtiTemplate(rawId) {
  var id = extractFileId(rawId);
  try {
    var file = DriveApp.getFileById(id);
    // Confirm it is a Google Doc
    if (file.getMimeType() !== 'application/vnd.google-apps.document') {
      throw new Error('File is not a Google Doc. Please link to a Google Doc template.');
    }
    return { id: id, name: file.getName() };
  } catch (e) {
    throw new Error('Could not access file: ' + e.message);
  }
}

/**
 * Reads student names (and optional data columns) from the sheet.
 * config: { sheetName, nameCol, dataStartRow }
 * Returns [{name, rowIndex}]
 */
function getPtiStudents(config) {
  var sheet    = getSheet(config.sheetName);
  var lastRow  = sheet.getLastRow();
  var nameIdx  = colLetterToIndex(config.nameCol);
  var students = [];

  for (var r = config.dataStartRow; r <= lastRow; r++) {
    var name = String(sheet.getRange(r, nameIdx).getValue()).trim();
    if (name) students.push({ name: name, rowIndex: r });
  }
  return students;
}

/**
 * Diagnostic function — call this from Apps Script editor to check
 * what names are found in each sheet and whether they match.
 * Returns a string report you can read in the execution log.
 */
function diagnosePtiData(config) {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var nameSheet    = ss.getSheetByName(config.sheetName);
  var spellSheet   = ss.getSheetByName(config.spellingSheet);
  var orfSheetObj  = ss.getSheetByName(config.orfSheet);

  var nameColIdx      = colLetterToIndex(config.nameCol);
  var spellNameColIdx = colLetterToIndex(config.spellingNameCol || config.nameCol);
  var orfNameColIdx   = colLetterToIndex(config.orfNameCol      || config.nameCol);
  var lines           = [];

  lines.push('=== NAME SHEET: ' + config.sheetName + ' (col ' + config.nameCol + ', from row ' + config.dataStartRow + ') ===');
  var lastRow = nameSheet.getLastRow();
  for (var r = config.dataStartRow; r <= lastRow; r++) {
    var v = nameSheet.getRange(r, nameColIdx).getValue();
    lines.push('  row ' + r + ': [' + v + ']');
  }

  lines.push('');
  lines.push('=== SPELLING SHEET: ' + config.spellingSheet + ' (col ' + (config.spellingNameCol || config.nameCol) + ', from row ' + config.spellingStartRow + ') ===');
  var spellMap = buildNameRowMap(spellSheet, spellNameColIdx, config.spellingStartRow);
  Object.keys(spellMap).forEach(function(k) { lines.push('  [' + k + '] → row ' + spellMap[k]); });

  lines.push('');
  lines.push('=== ORF SHEET: ' + config.orfSheet + ' (col ' + (config.orfNameCol || config.nameCol) + ', from row ' + config.orfStartRow + ') ===');
  var orfMap = buildNameRowMap(orfSheetObj, orfNameColIdx, config.orfStartRow);
  Object.keys(orfMap).forEach(function(k) { lines.push('  [' + k + '] → row ' + orfMap[k]); });

  lines.push('');
  lines.push('=== MATCH RESULTS ===');
  for (var r2 = config.dataStartRow; r2 <= lastRow; r2++) {
    var name = String(nameSheet.getRange(r2, nameColIdx).getValue()).trim();
    if (!name) continue;
    var norm = name.toLowerCase();
    lines.push('  ' + name + ' → spelling: ' + (spellMap[norm] ? 'row ' + spellMap[norm] : 'NOT FOUND')
                             + ' | orf: '     + (orfMap[norm]   ? 'row ' + orfMap[norm]   : 'NOT FOUND'));
  }

  var report = lines.join('\n');
  Logger.log(report);
  return report;
}

/**
 * Main generation function.
 *
 * config: {
 *   templateId,
 *   sheetName,        // student names sheet
 *   nameCol,
 *   dataStartRow,
 *   spellingSheet,    // sheet tab for SpellingScore, ChronAge, SpellingAge, Difference
 *   orfSheet,         // sheet tab for ORFCorrect, ORFError, MAZECorrect, MAZEError
 *   suffix,
 *   dateTime,
 *   fieldMap: { SpellingScore, ChronAge, SpellingAge, Difference,
 *               ORFCorrect, ORFError, MAZECorrect, MAZEError }
 * }
 */
function generatePtiDocuments(config) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var nameSheet    = ss.getSheetByName(config.sheetName);
  var spellSheet   = ss.getSheetByName(config.spellingSheet);
  var orfSheetObj  = ss.getSheetByName(config.orfSheet);

  if (!nameSheet)   throw new Error('Sheet not found: ' + config.sheetName);
  if (!spellSheet)  throw new Error('Sheet not found: ' + config.spellingSheet);
  if (!orfSheetObj) throw new Error('Sheet not found: ' + config.orfSheet);

  var nameColIdx      = colLetterToIndex(config.nameCol);
  var spellNameColIdx = colLetterToIndex(config.spellingNameCol || config.nameCol);
  var orfNameColIdx   = colLetterToIndex(config.orfNameCol      || config.nameCol);
  var lastRow         = nameSheet.getLastRow();

  var spellMap = buildNameRowMap(spellSheet,  spellNameColIdx, config.spellingStartRow);
  var orfMap   = buildNameRowMap(orfSheetObj, orfNameColIdx,   config.orfStartRow);

  // Resolve column indices (null = not mapped)
  var spellingKeys = ['SpellingScore','ChronAge','SpellingAge','Difference'];
  var orfKeys      = ['ORFCorrect','ORFError','MAZECorrect','MAZEError'];

  // Maps fieldMap key → placeholder string in the doc
  var placeholderMap = {
    'SpellingScore': '{{SS}}',
    'ChronAge':      '{{CA}}',
    'SpellingAge':   '{{SA}}',
    'Difference':    '{{DIF}}',
    'ORFCorrect':    '{{ORFC}}',
    'ORFError':      '{{ORFE}}',
    'MAZECorrect':   '{{MAZEC}}',
    'MAZEError':     '{{MAZEE}}',
  };

  function colIdx(key) {
    var col = config.fieldMap[key];
    return (col && col.trim()) ? colLetterToIndex(col.trim().toUpperCase()) : null;
  }

  var templateFile = DriveApp.getFileById(config.templateId);
  var parentFolder = templateFile.getParents().next();
  var results      = [];

  for (var r = config.dataStartRow; r <= lastRow; r++) {
    var rawName = String(nameSheet.getRange(r, nameColIdx).getValue()).trim();
    if (!rawName) continue;

    var fileName = formatStudentName(rawName) + ' - ' + config.suffix;
    var normName = rawName.toLowerCase().trim();

    try {
      var newFile = templateFile.makeCopy(fileName, parentFolder);
      var doc     = DocumentApp.openById(newFile.getId());
      var body    = doc.getBody();

      var rep = {};
      rep['{{Name}}']     = rawName;
      rep['{{DateTime}}'] = config.dateTime || '';

      // Spelling fields
      var spellRow = spellMap[normName] || null;
      spellingKeys.forEach(function(key) {
        var idx = colIdx(key);
        rep[placeholderMap[key]] = (spellRow && idx !== null)
          ? spellSheet.getRange(spellRow, idx).getDisplayValue()
          : '';
      });

      // ORF/MAZE fields
      var orfRow = orfMap[normName] || null;
      orfKeys.forEach(function(key) {
        var idx = colIdx(key);
        rep[placeholderMap[key]] = (orfRow && idx !== null)
          ? orfSheetObj.getRange(orfRow, idx).getDisplayValue()
          : '';
      });

      // Apply all replacements
      Object.keys(rep).forEach(function(placeholder) {
        body.replaceText(escapeRegex(placeholder), rep[placeholder]);
      });

      doc.saveAndClose();
      results.push({ success: true, name: rawName, fileName: fileName, url: newFile.getUrl() });

    } catch (e) {
      results.push({ success: false, name: rawName, error: e.message });
    }
  }

  return results;
}

/**
 * Builds a map of { normalisedStudentName: rowNumber } for a given sheet,
 * reading names from the specified column starting at dataStartRow.
 */
function buildNameRowMap(sheet, nameColIdx, dataStartRow) {
  var map     = {};
  var lastRow = sheet.getLastRow();
  for (var r = dataStartRow; r <= lastRow; r++) {
    var name = String(sheet.getRange(r, nameColIdx).getValue()).trim().toLowerCase();
    if (name) map[name] = r;
  }
  return map;
}

// ── Utilities ────────────────────────────────────────────────────────────────

/**
 * Converts a student name to "Last First" order for filenames.
 * Handles "First Last", "First Middle Last", and "Last, First" formats.
 */
function formatStudentName(name) {
  // If already "Last, First" format
  if (name.indexOf(',') !== -1) {
    var parts = name.split(',');
    return parts[0].trim() + ' ' + parts[1].trim();
  }
  // Otherwise assume "First [Middle] Last" → swap to "Last First [Middle]"
  var parts = name.trim().split(/\s+/);
  if (parts.length === 1) return parts[0];
  var last  = parts[parts.length - 1];
  var first = parts.slice(0, parts.length - 1).join(' ');
  return last + ' ' + first;
}

/**
 * Extracts a Drive file ID from a URL or returns the string as-is if it looks like an ID.
 */
function extractFileId(raw) {
  if (!raw) throw new Error('No file ID or URL provided.');
  raw = raw.trim();
  var match = raw.match(/[-\w]{25,}/);
  if (!match) throw new Error('Could not find a valid file ID in: ' + raw);
  return match[0];
}

/**
 * Escapes special regex characters in a placeholder string.
 */
function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
