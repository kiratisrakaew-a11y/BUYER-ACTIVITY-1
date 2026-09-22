/**
 * DriveService.js — one Drive folder per Case, and the files attached to it.
 *
 * Files are shared to the company domain only (SPEC §11); the DB spreadsheet
 * itself is never shared, buyers only ever reach data through the web app.
 */
var DriveService = (function () {

  var ROOT_SETTING = 'DRIVE_ROOT_FOLDER_ID';
  var ROOT_FOLDER_NAME = 'Buyer Procurement Activity Tracker';
  var DESCRIPTION_LIMIT = 50;             // SPEC §11 — folder name is "{Case_ID} - {first 50 chars}"
  var MAX_UPLOAD_BYTES = 10 * 1024 * 1024;

  /**
   * The folder all Case folders live under. If an administrator has not set
   * DRIVE_ROOT_FOLDER_ID, one is created once and written back to Config_Settings
   * so the system works out of the box instead of failing on the first upload.
   */
  function getRootFolder() {
    var id = Config.get(ROOT_SETTING, '');
    if (id) {
      try {
        return DriveApp.getFolderById(id);
      } catch (e) {
        throw Err.internal('เปิดโฟลเดอร์หลักใน Drive ไม่ได้ (' + ROOT_SETTING + ' = ' + id + ')');
      }
    }
    var folder = DriveApp.createFolder(ROOT_FOLDER_NAME);
    folder.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
    Config.setSetting(ROOT_SETTING, folder.getId(), 'สร้างอัตโนมัติเมื่อใช้งานครั้งแรก');
    return folder;
  }

  function folderName(caseId, description) {
    var short = Utils.truncate(String(description || '').replace(/\s+/g, ' ').trim(), DESCRIPTION_LIMIT);
    return short ? caseId + ' - ' + short : caseId;
  }

  /** Creates the Case folder. Returns null instead of throwing: a Drive outage
   *  must not stop a buyer from opening a Case. */
  function createCaseFolder(caseId, description) {
    try {
      var folder = getRootFolder().createFolder(folderName(caseId, description));
      folder.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
      return folder.getId();
    } catch (e) {
      console.error('createCaseFolder(' + caseId + ') failed: ' + ((e && e.message) || e));
      return null;
    }
  }

  /** The Case folder, created on demand if it does not exist yet. */
  function ensureCaseFolder(caseRecord) {
    if (caseRecord.Drive_Folder_ID) {
      try {
        return DriveApp.getFolderById(caseRecord.Drive_Folder_ID);
      } catch (e) {
        console.error('Case folder ' + caseRecord.Drive_Folder_ID + ' is unreachable, creating a new one');
      }
    }
    var folder = getRootFolder().createFolder(folderName(caseRecord.Case_ID, caseRecord.Description));
    folder.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
    Repository.update('Cases', caseRecord.Case_ID, { Drive_Folder_ID: folder.getId() }, null,
      { actor: ChangeLog.SYSTEM_USER, reason: 'สร้างโฟลเดอร์เก็บไฟล์ของงาน' });
    return folder;
  }

  function folderUrl(folderId) {
    if (!folderId) return '';
    try {
      return DriveApp.getFolderById(folderId).getUrl();
    } catch (e) {
      return '';
    }
  }

  /** Stores a base64 upload in the Case folder and returns its URL. */
  function uploadToCase(caseRecord, fileName, mimeType, base64) {
    if (Utils.isBlank(fileName)) throw Err.validation('ไม่พบชื่อไฟล์');
    if (Utils.isBlank(base64)) throw Err.validation('ไม่พบข้อมูลไฟล์');

    var bytes;
    try {
      bytes = Utilities.base64Decode(base64);
    } catch (e) {
      throw Err.validation('ไฟล์ที่อัปโหลดไม่ถูกต้อง');
    }
    if (bytes.length > MAX_UPLOAD_BYTES) {
      throw Err.validation('ไฟล์ใหญ่เกิน ' + Math.round(MAX_UPLOAD_BYTES / 1024 / 1024) + ' MB');
    }

    var safeName = Utils.truncate(String(fileName).replace(/[\\/:*?"<>|]/g, '_'), 120);
    var blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', safeName);
    var file = ensureCaseFolder(caseRecord).createFile(blob);
    file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
    return { fileId: file.getId(), url: file.getUrl(), name: safeName };
  }

  return {
    ROOT_SETTING: ROOT_SETTING,
    MAX_UPLOAD_BYTES: MAX_UPLOAD_BYTES,
    getRootFolder: getRootFolder,
    folderName: folderName,
    createCaseFolder: createCaseFolder,
    ensureCaseFolder: ensureCaseFolder,
    folderUrl: folderUrl,
    uploadToCase: uploadToCase
  };
})();
