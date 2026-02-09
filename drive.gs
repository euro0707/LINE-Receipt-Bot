function saveReceiptImageToDrive_(imageBytes, mimeType, dateObj, messageId) {
  const targetDate = dateObj || new Date();
  const rootFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  const monthFolder = getOrCreateMonthFolder_(rootFolder, targetDate);

  const extension = mimeTypeToExtension_(mimeType);
  const timestamp = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const safeMessageId = safeText_(messageId, Utilities.getUuid().slice(0, 8));
  const fileName = timestamp + '_' + safeMessageId + '.' + extension;

  const file = executeWithRetry_(function() {
    const blob = Utilities.newBlob(imageBytes, mimeType || 'image/jpeg', fileName);
    return monthFolder.createFile(blob);
  }, SETTINGS.DRIVE_MAX_RETRY, 1000, 'Drive save');

  return {
    id: file.getId(),
    name: file.getName(),
    url: file.getUrl()
  };
}

function getOrCreateMonthFolder_(rootFolder, targetDate) {
  const folderName = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy年MM月');
  const folders = rootFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return rootFolder.createFolder(folderName);
}

function mimeTypeToExtension_(mimeType) {
  if (!mimeType) {
    return 'jpg';
  }
  if (mimeType.indexOf('png') >= 0) {
    return 'png';
  }
  if (mimeType.indexOf('gif') >= 0) {
    return 'gif';
  }
  if (mimeType.indexOf('webp') >= 0) {
    return 'webp';
  }
  return 'jpg';
}
