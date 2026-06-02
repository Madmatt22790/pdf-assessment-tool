// ============================================================
// ASSESSMENTGENERATOR.GS
// Server-side functions for the Assessment Generator sidebar.
// ============================================================

/**
 * Fetches a PDF from Drive and returns it as base64.
 * Same pattern as the splitter — avoids CORS in the sidebar.
 */
function getAssessmentPdfAsBase64(fileId) {
  try {
    var file  = DriveApp.getFileById(fileId);
    var bytes = file.getBlob().getBytes();
    return Utilities.base64Encode(bytes);
  } catch (e) {
    throw new Error('Could not fetch PDF: ' + (e && e.message ? e.message : e));
  }
}

/**
 * Validates a Drive file is an accessible PDF.
 */
function validateAssessmentFile(fileId) {
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

/**
 * Receives the fully assembled PDF (all students combined) as base64,
 * saves it to the same Drive folder as the blank assessment,
 * and returns the download URL.
 *
 * @param {string} base64Pdf   - The complete generated PDF
 * @param {string} sourceFileId - Original blank assessment file ID (used to find folder)
 * @param {string} fileName    - Name for the output file
 * @returns {Object} {url, fileName}
 */
function saveGeneratedAssessment(base64Pdf, sourceFileId, fileName) {
  var sourceFile   = DriveApp.getFileById(sourceFileId);
  var parentFolder = sourceFile.getParents().next();

  var bytes   = Utilities.base64Decode(base64Pdf);
  var blob    = Utilities.newBlob(bytes, 'application/pdf', fileName);

  // Replace existing file with same name if present
  var existing = parentFolder.getFilesByName(fileName);
  while (existing.hasNext()) { existing.next().setTrashed(true); }

  var newFile = parentFolder.createFile(blob);
  newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { url: newFile.getUrl(), fileName: fileName };
}
