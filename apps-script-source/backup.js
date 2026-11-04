/**
 * Spreadsheet backup utilities (Project Snowbee A2.3)
 * Run backupSpreadsheetToDrive() manually or via daily trigger.
 */

var BACKUP_FOLDER_NAME = 'POS_TS3_Backups';

function getOrCreateBackupFolder_() {
  var folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

/**
 * Copy the POS spreadsheet into Drive folder POS_TS3_Backups.
 * @returns {{success: boolean, fileName: string, fileId: string, folderName: string}}
 */
function backupSpreadsheetToDrive() {
  var ss = getGlobalSpreadsheet();
  var tz = Session.getScriptTimeZone();
  var stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HHmm');
  var fileName = ss.getName() + '_backup_' + stamp;
  var folder = getOrCreateBackupFolder_();
  var copy = DriveApp.getFileById(ss.getId()).makeCopy(fileName, folder);
  Logger.log('Backup created: ' + fileName + ' (' + copy.getId() + ')');
  return {
    success: true,
    fileName: fileName,
    fileId: copy.getId(),
    folderName: BACKUP_FOLDER_NAME
  };
}

/**
 * Install a daily backup trigger (default 03:00 script timezone).
 */
function installDailyBackupTrigger(hour) {
  removeDailyBackupTrigger();
  var h = (hour === undefined || hour === null) ? 3 : Number(hour);
  ScriptApp.newTrigger('backupSpreadsheetToDrive')
    .timeBased()
    .everyDays(1)
    .atHour(h)
    .create();
  Logger.log('Daily backup trigger installed at hour ' + h);
  return { success: true, hour: h };
}

function removeDailyBackupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'backupSpreadsheetToDrive') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  return { success: true, removed: removed };
}

function getBackupTriggerInfo() {
  var triggers = ScriptApp.getProjectTriggers();
  var backupTriggers = triggers.filter(function(t) {
    return t.getHandlerFunction() === 'backupSpreadsheetToDrive';
  });
  return {
    installed: backupTriggers.length > 0,
    count: backupTriggers.length
  };
}
