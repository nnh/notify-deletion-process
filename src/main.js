/**
 * Send a chat.
 * @param {string} The text of the post.
 * @return none.
 */
function notifyWorkInformation(strPayload) {
  // Webhook URL
  const postUrl =
    PropertiesService.getScriptProperties().getProperty('targetUrl');
  const payload = {
    text: strPayload,
  };
  const options = {
    method: 'POST',
    headers: { 'Content-Type': 'application/json; charset=UTF-8' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
  const result = UrlFetchApp.fetch(postUrl, options);
}
/**
 * Get information about the spreadsheet.
 * @param none.
 * @return none.
 */
function getAssetManagementInfo() {
  let arrayInfo = {};
  arrayInfo.ssUrl =
    PropertiesService.getScriptProperties().getProperty('assetManagement');
  arrayInfo.sheetName = 'SoftwareUsers';
  arrayInfo.targetItemindex = 0;
  arrayInfo.targetItemName = null;
  const sheetValues = getSoftwareUsers(arrayInfo);
  const planToDelete = getDelTarget(sheetValues, 5, 6);
  if (planToDelete.length > 1) {
    notifyWorkInformation(arrayInfo.ssUrl);
  }
}
function getboxCollaboratorInfo() {
  let arrayInfo = {};
  arrayInfo.ssUrl =
    PropertiesService.getScriptProperties().getProperty('boxCollaborator');
  arrayInfo.sheetName = 'フォームの回答 2';
  arrayInfo.targetItemindex = 5;
  arrayInfo.targetItemName = '登録';
  const sheetValues = getSoftwareUsers(arrayInfo);
  const planToDelete = getDelTarget(sheetValues, 13, 15);
  if (planToDelete.length > 1) {
    notifyWorkInformation(arrayInfo.ssUrl);
  }
}
/**
 * Get information about the spreadsheet.
 * @param {Object} associative array.
 * @return {Array.<string>} Spreadsheet contents.
 */
function getSoftwareUsers(arrayInfo) {
  const ss = SpreadsheetApp.openByUrl(arrayInfo.ssUrl);
  const sheet = ss.getSheetByName(arrayInfo.sheetName);
  const targetRange = sheet.getDataRange();
  const targetValues =
    arrayInfo.targetItemName != null
      ? targetRange
          .getValues()
          .filter(
            (x, idx) =>
              x[parseInt(arrayInfo.targetItemindex)] ==
                arrayInfo.targetItemName || idx == 0
          )
      : targetRange.getValues();
  return targetValues;
}
/**
 * Extract the records whose scheduled deletion date is earlier than 7 days after the processing date.
 * @param {Array.<string>} Spreadsheet contents.
 * @param {number} Column of deletion plan date.
 * @param {number} Column of deletion date.
 * @return {Array.<string>} Target values.
 */
function getDelTarget(inputValues, planToDeleteCol, deleteCol) {
  let checkDate = new Date();
  checkDate.setDate(checkDate.getDate() + 1);
  const planToDeleteValues = inputValues.filter(
    (x, idx) => x[planToDeleteCol] != '' || idx == 0
  );
  let notDeleted = planToDeleteValues.filter(
    (x, idx) => !x[deleteCol] || idx == 0
  );
  notDeleted = notDeleted.filter(
    (x, idx) => x[planToDeleteCol] < checkDate || idx == 0
  );
  return notDeleted;
}
/**
 * Set the properties.
 * @param none.
 * @return none.
 */
function registerScriptProperty() {
  PropertiesService.getScriptProperties().deleteAllProperties;
  // set Webhook URL
  PropertiesService.getScriptProperties().setProperty(
    'targetUrl',
    'https://chat.googleapis.com/...'
  );
  // Set the URL of the spreadsheet that contains the Prime Drive information.
  PropertiesService.getScriptProperties().setProperty(
    'assetManagement',
    'https://docs.google.com/spreadsheets/...'
  );
  // Set the URL of the spreadsheet that contains the Box information.
  PropertiesService.getScriptProperties().setProperty(
    'boxCollaborator',
    'https://docs.google.com/spreadsheets/...'
  );
}
