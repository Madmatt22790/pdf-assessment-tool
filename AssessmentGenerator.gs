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
 * Only used as a fallback for small files — large files use the chunked path below.
 */
function saveGeneratedAssessment(base64Pdf, sourceFileId, fileName) {
  var sourceFile   = DriveApp.getFileById(sourceFileId);
  var parentFolder = sourceFile.getParents().next();

  var bytes   = Utilities.base64Decode(base64Pdf);
  var blob    = Utilities.newBlob(bytes, 'application/pdf', fileName);

  var existing = parentFolder.getFilesByName(fileName);
  while (existing.hasNext()) { existing.next().setTrashed(true); }

  var newFile = parentFolder.createFile(blob);
  newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { url: newFile.getUrl(), fileName: fileName };
}

/**
 * Returns the OAuth token and parent folder ID needed for a direct
 * browser-to-Drive upload. The file bytes never pass through Apps Script.
 */
function getDriveUploadInfo(sourceFileId) {
  try {
    var folderId = DriveApp.getFileById(sourceFileId).getParents().next().getId();
    return { token: ScriptApp.getOAuthToken(), folderId: folderId };
  } catch (e) {
    throw new Error('Could not get upload info: ' + e.message);
  }
}

/**
 * Sets sharing on the uploaded file and returns its URL.
 * Called after the browser has finished uploading directly to Drive.
 */
function finalizeAssessmentFile(fileId, fileName) {
  try {
    var file = DriveApp.getFileById(fileId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { url: file.getUrl(), fileName: fileName };
  } catch (e) {
    throw new Error('Could not finalise file: ' + e.message);
  }
}

/**
 * Starts a Drive resumable upload session and returns the session URL.
 * The client then sends chunks via uploadAssessmentChunk().
 */
function startAssessmentUpload(sourceFileId, fileName, fileSize) {
  try {
    var folderId = DriveApp.getFileById(sourceFileId).getParents().next().getId();
    var token    = ScriptApp.getOAuthToken();
    var resp = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable',
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json',
          'X-Upload-Content-Type': 'application/pdf',
          'X-Upload-Content-Length': String(fileSize)
        },
        payload: JSON.stringify({ name: fileName, parents: [folderId] }),
        muteHttpExceptions: true
      }
    );
    var headers = resp.getHeaders();
    var loc = headers['Location'] || headers['location'];
    if (!loc) throw new Error('No upload URL returned (HTTP ' + resp.getResponseCode() + ')');
    return loc;
  } catch (e) {
    throw new Error('Could not start upload: ' + e.message);
  }
}

/**
 * Sends one chunk to an in-progress Drive resumable upload session.
 * Returns {done: false} while more chunks are needed, or
 * {done: true, url, fileName} when the upload is complete.
 */
function uploadAssessmentChunk(uploadUrl, base64Chunk, offset, totalSize, fileName) {
  try {
    var token = ScriptApp.getOAuthToken();
    var bytes = Utilities.base64Decode(base64Chunk);
    var end   = offset + bytes.length - 1;
    var resp  = UrlFetchApp.fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Range': 'bytes ' + offset + '-' + end + '/' + totalSize,
        'Content-Type': 'application/pdf'
      },
      payload: bytes,
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code === 200 || code === 201) {
      var fileId = JSON.parse(resp.getContentText()).id;
      var file   = DriveApp.getFileById(fileId);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return { done: true, url: file.getUrl(), fileName: fileName };
    }
    if (code === 308) return { done: false };
    throw new Error('HTTP ' + code + ' at offset ' + offset + ': ' + resp.getContentText());
  } catch (e) {
    throw new Error('Chunk upload failed at offset ' + offset + ': ' + e.message);
  }
}
