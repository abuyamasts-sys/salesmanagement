function getApproverBackupDataFromDashboard(userId) {
  requireCurrentUserRole_(['Approver'], userId);
  return toClientValue_(getBackupDashboardData_());
}

function createDatabaseBackupFromDashboard(userId) {
  var currentUser = requireCurrentUserRole_(['Approver'], userId);
  return toClientValue_(createSpreadsheetBackup_(currentUser));
}

function getBackupDashboardData_() {
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.BACKUP_LOG, APP_CONFIG.HEADERS.BACKUP_LOG);

  var history = getSheetData_(APP_CONFIG.SHEETS.BACKUP_LOG)
    .slice()
    .sort(function(left, right) {
      return String(right.timestamp_backup || '').localeCompare(String(left.timestamp_backup || ''));
    });
  var latest = history.length ? history[0] : null;

  return {
    summary: {
      total_backup: history.length,
      backup_terakhir: latest ? (latest.timestamp_backup || '') : '',
      file_terakhir: latest ? (latest.nama_file || '') : '',
      status_terakhir: latest ? (latest.status_backup || '') : '',
      folder_name: latest ? (latest.folder_name || '') : (APP_CONFIG.BACKUP.FOLDER_NAME || ''),
      folder_id: latest ? (latest.folder_id || '') : '',
      folder_url: latest && latest.folder_id ? ('https://drive.google.com/drive/folders/' + latest.folder_id) : '',
      jumlah_sheet_operasional: getOperationalSheetNames_().length
    },
    history: history.slice(0, 20)
  };
}

function createSpreadsheetBackup_(currentUser) {
  ensureSheetHeadersContain_(APP_CONFIG.SHEETS.BACKUP_LOG, APP_CONFIG.HEADERS.BACKUP_LOG);

  var now = new Date();
  var parts = getNowParts_();
  var timestampLabel = Utilities.formatDate(now, APP_CONFIG.TIMEZONE, 'yyyy-MM-dd_HH-mm-ss');
  var fileName = (APP_CONFIG.BACKUP.FILE_PREFIX || 'BACKUP_MVP_SALES_ORDER') + '_' + timestampLabel;
  var sourceFile = DriveApp.getFileById(APP_CONFIG.SPREADSHEET_ID);
  var folder = ensureBackupFolder_();
  var backupFile;
  var backupSpreadsheet;
  var operationalSheets = getOperationalSheetNames_();
  var actualBackupSheets = [];

  try {
    backupFile = sourceFile.makeCopy(fileName, folder);
    backupSpreadsheet = SpreadsheetApp.openById(backupFile.getId());
    trimBackupSpreadsheetSheets_(backupSpreadsheet, operationalSheets);
    actualBackupSheets = backupSpreadsheet.getSheets().map(function(sheet) {
      return String(sheet.getName() || '').trim();
    });

    var logRow = {
      backup_id: generateDocNumber_('BKP'),
      tanggal_backup: parts.tanggal,
      jam_backup: parts.jam,
      timestamp_backup: nowIso_(),
      nama_file: backupFile.getName(),
      file_id: backupFile.getId(),
      file_url: backupFile.getUrl(),
      folder_id: folder.getId(),
      folder_name: folder.getName(),
      jumlah_sheet: actualBackupSheets.length,
      daftar_sheet: actualBackupSheets.join(', '),
      dibackup_oleh: String(currentUser && currentUser.user_id || '').trim(),
      status_backup: 'Berhasil',
      catatan: 'Backup spreadsheet operasional berhasil dibuat.'
    };

    appendRowByHeaders_(APP_CONFIG.SHEETS.BACKUP_LOG, logRow);

    return {
      success: true,
      message: 'Backup database berhasil dibuat.',
      file_name: backupFile.getName(),
      file_id: backupFile.getId(),
      file_url: backupFile.getUrl(),
      folder_id: folder.getId(),
      folder_name: folder.getName(),
      folder_url: 'https://drive.google.com/drive/folders/' + folder.getId(),
      jumlah_sheet: actualBackupSheets.length,
      daftar_sheet: actualBackupSheets
    };
  } catch (error) {
    appendRowByHeaders_(APP_CONFIG.SHEETS.BACKUP_LOG, {
      backup_id: generateDocNumber_('BKP'),
      tanggal_backup: parts.tanggal,
      jam_backup: parts.jam,
      timestamp_backup: nowIso_(),
      nama_file: fileName,
      file_id: backupFile ? backupFile.getId() : '',
      file_url: backupFile ? backupFile.getUrl() : '',
      folder_id: folder ? folder.getId() : '',
      folder_name: folder ? folder.getName() : '',
      jumlah_sheet: actualBackupSheets.length || operationalSheets.length,
      daftar_sheet: operationalSheets.join(', '),
      dibackup_oleh: String(currentUser && currentUser.user_id || '').trim(),
      status_backup: 'Gagal',
      catatan: String(error && error.message ? error.message : error)
    });

    throw new Error('Backup database gagal: ' + (error && error.message ? error.message : error));
  }
}

function ensureBackupFolder_() {
  var folderName = APP_CONFIG.BACKUP.FOLDER_NAME || 'BACKUP_MVP_SALES_ORDER';
  var folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(folderName);
}

function getOperationalSheetNames_() {
  return Object.keys(APP_CONFIG.SHEETS).map(function(key) {
    return APP_CONFIG.SHEETS[key];
  }).filter(function(sheetName, index, list) {
    return !!sheetName && list.indexOf(sheetName) === index;
  });
}

function trimBackupSpreadsheetSheets_(spreadsheet, allowedSheetNames) {
  var allowedMap = {};

  (allowedSheetNames || []).forEach(function(sheetName) {
    allowedMap[String(sheetName || '').trim()] = true;
  });

  spreadsheet.getSheets().forEach(function(sheet) {
    var name = String(sheet.getName() || '').trim();

    if (!allowedMap[name]) {
      spreadsheet.deleteSheet(sheet);
    }
  });
}
