/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

var DIALOG_TITLE = 'Example Dialog';
var SIDEBAR_TITLE = 'Example Sidebar';

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
//      .addItem('Enter source folder ID', 'showFolderIdDialog')
      .addItem('Copy folder', 'showFolderCopyDialog')
      .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle(SIDEBAR_TITLE)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Opens a dialog. The dialog structure is described in the Dialog.html
 * project file.
 */
function showFolderIdDialog() {
  var ui = HtmlService.createTemplateFromFile('FolderIdDialog')
      .evaluate()
      .setWidth(400)
      .setHeight(190)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(ui, 'FOLDER ID');
}

/**
 * Opens a dialog. The dialog structure is described in the Dialog.html
 * project file.
 */
function showFolderCopyDialog() {
  var ui = HtmlService.createTemplateFromFile('FolderCopyDialog')
      .evaluate()
      .setWidth(400)
      .setHeight(350)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(ui, 'FOLDER COPY');
}

function getProperty(key) {
  var properties = PropertiesService.getDocumentProperties();
//  properties.deleteAllProperties();
//  properties.setProperty('targetFolderId', '1MPJ3--dp8W9_BB5zi_HDqS3_STkGGcJg');
  var value = properties.getProperty(key);
  value = value ? value : '';
  return value;
}

function onEdit(e) {
  var ui = SpreadsheetApp.getUi();
//  ui.alert(typeof e.value)
  if (e.range.getRow() > 1 && e.range.getColumn() == 1 && e.value !== undefined)
    collapseFolderContent(e.range, e.value == 'FALSE');
}

function collapseFolderContent (range, isTrue) {
  var sheet = range.getSheet();
  var row = range.getRow();
  var data = sheet.getDataRange().getValues();
  var level;
  for (var level = 1; level < data[0].length; level++) {
    if (data[row - 1][level] != '') break;
  }
  var numRows = 0;
  for (var i = row; i < data.length; i++) {
    if (data[i][level] != '') {
      numRows = i - row;
      break;
    }
  }
  sheet = SpreadsheetApp.getActiveSheet();
  if (isTrue) {
    sheet.hideRows(row + 1, numRows);
  } else {
    sheet.showRows(row + 1, numRows);
  }
}

function copyFolder(targetId, driveObject) {
  var targetFolder = targetId == '' ? DriveApp.getRootFolder() : DriveApp.getFolderById(targetId);
  Logger.log('Target folder, %s, opened', targetFolder.getName());
  var properties = PropertiesService.getDocumentProperties();
  properties.setProperty('targetFolderId', targetId);
//  var sourceId = properties.getProperty('sourceFolderId');
//  if (sourceId == null) throw 'Source folder is not set!';
//  var sourceFolder = DriveApp.getFolderById(sourceId);
//  Logger.log('Source folder, %s, opened', sourceFolder.getName());
  if (driveObject == undefined) {
    driveObject = getJsonProperty('driveObject', properties);
    if (!driveObject) throw 'DriveObject was not created!<br>Please select a source folder.';
    Logger.log('DriveObject created');
  }
//  var folder = targetFolder.createFolder(sourceFolder.getName());
  copyDriveObject(targetFolder, driveObject);
}

function copyDriveObject(targetFolder, driveObject) {
  if (driveObject.type == 'FOLDER') {
    var folder = targetFolder.createFolder(driveObject.object.getName());
    for (var i in driveObject.content) {
      copyDriveObject(folder, driveObject.content[i]);
    }
  } else if (driveObject.type == 'FILE') {
    var file = driveObject.object;//DriveApp.getFileById(driveObject.id);
    file.makeCopy(file.getName(), targetFolder);
  }
}

function saveSourceFolderId(sourceId, driveOnly, targetId) {
  if (sourceId == '') throw 'Invalid folder ID';
  var folder = DriveApp.getFolderById(sourceId);
  var properties = PropertiesService.getDocumentProperties();
  properties.setProperty('sourceFolderId', sourceId);
  properties.setProperty('driveOnly', driveOnly);
  Logger.log('Folder ID, %s, saved.', sourceId);
  var driveObject = getFolderContent(folder, driveOnly);
  if (targetId) {
    copyFolder(targetId, driveObject);
  } else {
    var list = listContent(driveObject);
    Logger.log('List set on sheet');
//    setJsonProperty(driveObject, 'driveObject', properties);
//      Logger.log('DriveObject saved.');
  }
}

function setJsonProperty(object, key, properties) {
  properties = properties == undefined ? PropertiesService.getDocumentProperties() : properties;
  var text = JSON.stringify(object);
  var props = {};
  var i;
  var max = 9000;
  for (i = 0; i * max < text.length; i++) {
    var start = i * max;
    var end = (i + 1) * max < text.length ? (i + 1) * max : text.length;
    var subtext = text.substring(start, end);
//    Logger.log(subtext)
    props[key + '-' + i] = subtext;
  }
  Logger.log('JSON object is set in %s part(s).', i)
  props[key + '-num'] = i;
  properties.setProperties(props);
}

function getJsonProperty(key, properties) {
  properties = properties == undefined ? PropertiesService.getDocumentProperties() : properties;
//    var properties = PropertiesService.getDocumentProperties();
  var text = '';
  var num = parseInt(properties.getProperty(key + '-num'));
  if (!num) return null;
  for (var i = 0; i < num; i++) {
    text += properties.getProperty(key + '-' + i);
//    Logger.log(text)
  }
  Logger.log('JSON object is retrieved from %s part(s).', num)
  var object = JSON.parse(text);
  return object;
}

function listContent(driveObject) {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear().clearNotes().getRange('A:A').clearDataValidations();
  var list = [];
  
  list.push([driveObject.object.getName()]);
//  delete driveObject.name;//***
  var level = addToList(list, driveObject.content, 1);
  for (var i in list) {
//    Logger.log('before' + list[i].length)
    while (list[i].length <= level) {
      list[i].push(null);
    }
//    Logger.log('after' + list[i].length)
  }
//  Logger.log(list);
  sheet.setColumnWidths(1, list[0].length, 30);
  sheet.getRange(2, 1, list.length - 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false));
  sheet.getRange(1, 1, list.length, list[0].length)
      .setValues(list);
  list = [];
  list.toString();
  
  return list;
}

function addToList(list, content, level) {
  var highest = level;
  var start = list.length;
  var count;
  for (count in content) {
    var row = [true];
    for (var j = 1; j < level; j++) {
      row.push(null);
    }
    row.push(content[count].name);
//    delete content[count].name;//***
    list.push(row);
    if (content[count].content.length > 0) {
      var newLevel = addToList(list, content[count].content, level + 1);
      if (newLevel > highest) highest = newLevel;
    }
  }
  return highest;
}

function test(obj) {
//  saveSourceFolderId('1tnKmY0qEw2sJJT2FXgMULmpfjSkAGYMO');
  
//  var sheet = SpreadsheetApp.getActiveSheet();
//  sheet.clear().clearNotes().clearConditionalFormatRules();
  
//  var properties = PropertiesService.getDocumentProperties();
//  properties.deleteAllProperties();
////  Logger.log(properties.getProperty('ole'))
//  var s = '';
//  for (var i = 0; i < 80; i++) {
//    s += 'a';
//  }
//  var object = {text:s};
//  setJsonProperty(object, 'key', properties);
//  var object2 = getJsonProperty('key', properties);
//  Logger.log(object2.text.length);
//  var list = [];
//  for (var i in list) {
//    Logger.log('worked');
//  }
  
//  var file = DriveApp.getFileById('1ogzQrI_QoFVbzA0IJQ9ehdpKnh3jN2iv1ItpGiiEb_M');
//  var folder = DriveApp.getFolderById('1UL3UWASMU1TUF-Glnp6Ha9zk5ljXekGt');
//  file.makeCopy(file.getName(), folder);

//  var driveOnly = true;
//  var file = DriveApp.getFileById('1AeoJaaYhky_qnaIZUHzk9o1uySu5fVVLcfQq62HrtAQ');
//  Logger.log(file.getMimeType())
//  if (!driveOnly || file.getMimeType().indexOf('google') >= 0) {
//    Logger.log('FILE object added')
//  }
  Logger.log(getProperty('driveOnly'))
}

function linkFormula(url, link_label) {
  return '=HYPERLINK("' + url + '","' + link_label + '")';
}

function getFolderContent(folder, driveOnly) {
  var driveObject = getDriveObject('FOLDER', folder);
  var folders = folder.getFolders();
  getSubFolders(folders, driveObject.content);
  var files = folder.getFiles();
  getFiles(files, driveObject.content, driveOnly);
  return driveObject;
}

function getSubFolders(folders, content) {
  while (folders.hasNext()) {
    var folder = folders.next();
    var driveObject = getDriveObject('FOLDER', folder);
    content.push(driveObject);
//    Logger.log('FOLDER object added')
    getSubFolders(folder.getFolders(), driveObject.content);
    getFiles(folder.getFiles(), driveObject.content);
  }
}

function getFiles(files, content, driveOnly) {
  while (files.hasNext()) {
    var file = files.next();
    if (!driveOnly || file.getMimeType().indexOf('google') >= 0) {
      var driveObject = getDriveObject('FILE', file);
      content.push(driveObject);
//      Logger.log('FILE object added')
    }
  }
}

function getDriveObject(type, object) {
  var driveObject = {
    type: type,
    object: object,
//    name: object.getName(),
//    id: object.getId(),
    content: []
  };
  Logger.log('%s drive object created', type);
  return driveObject;
}

/**
 * Returns the value in the active cell.
 *
 * @return {String} The value of the active cell.
 */
function getActiveValue() {
  // Retrieve and return the information requested by the sidebar.
  var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  return cell.getValue();
}

/**
 * Replaces the active cell value with the given value.
 *
 * @param {Number} value A reference number to replace with.
 */
function setActiveValue(value) {
  // Use data collected from sidebar to manipulate the sheet.
  var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  cell.setValue(value);
}

/**
 * Executes the specified action (create a new sheet, copy the active sheet, or
 * clear the current sheet).
 *
 * @param {String} action An identifier for the action to take.
 */
function modifySheets(action) {
  // Use data collected from dialog to manipulate the spreadsheet.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();
  if (action == "create") {
    ss.insertSheet();
  } else if (action == "copy") {
    currentSheet.copyTo(ss);
  } else if (action == "clear") {
    currentSheet.clear();
  }
}
