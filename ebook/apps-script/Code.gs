const SPREADSHEET_ID = '1TeTgln1MkV67tmidOhYnuBteyzKP77OPknU71mDRliU';
const ESSAYS_SHEET = 'Essays';
const UPLOAD_FOLDER_NAME = 'eKarangan BM Tahun 4 Uploads';
const API_VERSION = 1;

function setup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(ESSAYS_SHEET);
  if (!sheet) sheet = ss.insertSheet(ESSAYS_SHEET);
  const headers = ['id','title','theme','standard','easy','imageFileId','imageUrl','createdAt','updatedAt','active'];
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (headers.some((h, i) => current[i] !== h)) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TEACHER_PIN')) props.setProperty('TEACHER_PIN', '1234');
  if (!props.getProperty('UPLOAD_FOLDER_ID')) {
    const folder = getOrCreateUploadFolder_();
    props.setProperty('UPLOAD_FOLDER_ID', folder.getId());
  }
  return {
    ok: true,
    spreadsheetId: SPREADSHEET_ID,
    uploadFolderId: props.getProperty('UPLOAD_FOLDER_ID'),
    apiVersion: API_VERSION
  };
}

function setTeacherPin(newPin) {
  const pin = String(newPin || '').trim();
  if (!/^\d{4,8}$/.test(pin)) throw new Error('PIN mesti 4 hingga 8 digit.');
  PropertiesService.getScriptProperties().setProperty('TEACHER_PIN', pin);
  return true;
}

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'list').toLowerCase();
    if (action === 'ping') return json_({ok:true, apiVersion:API_VERSION, service:'eKarangan'});
    if (action === 'list') return json_({ok:true, apiVersion:API_VERSION, essays:listEssays_()});
    return json_({ok:false, error:'Unknown action'});
  } catch (err) {
    return json_({ok:false, error:String(err && err.message || err)});
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(body.action || '').toLowerCase();
    if (action === 'auth') {
      checkPin_(body.pin);
      return json_({ok:true});
    }
    if (action === 'save') {
      checkPin_(body.pin);
      return json_({ok:true, essay:saveEssay_(body.essay || {}, body.imageData || '')});
    }
    if (action === 'delete') {
      checkPin_(body.pin);
      deleteEssay_(String(body.id || ''));
      return json_({ok:true, id:String(body.id || '')});
    }
    return json_({ok:false, error:'Unknown action'});
  } catch (err) {
    return json_({ok:false, error:String(err && err.message || err)});
  }
}

function listEssays_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ESSAYS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
  return values.filter(r => String(r[0] || '').trim()).map(rowToEssay_);
}

function saveEssay_(essay, imageData) {
  const id = String(essay.id || '').trim();
  const title = String(essay.title || '').trim();
  if (!id) throw new Error('ID karangan tiada.');
  if (!title) throw new Error('Tajuk karangan diperlukan.');
  if (!String(essay.standard || '').trim() && !String(essay.easy || '').trim()) {
    throw new Error('Sekurang-kurangnya satu versi karangan diperlukan.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ESSAYS_SHEET);
    const found = findRowById_(sheet, id);
    const old = found ? sheet.getRange(found, 1, 1, 10).getValues()[0] : ['', '', '', '', '', '', '', '', '', true];
    let imageFileId = String(old[5] || '');
    let imageUrl = String(essay.imageUrl || old[6] || '');

    if (essay.imageRemoved === true) {
      trashFileSafe_(imageFileId);
      imageFileId = '';
      imageUrl = '';
    } else if (imageData && /^data:image\//i.test(imageData)) {
      const uploaded = saveImage_(imageData, title);
      trashFileSafe_(imageFileId);
      imageFileId = uploaded.fileId;
      imageUrl = uploaded.url;
    }

    const now = new Date().toISOString();
    const createdAt = String(old[7] || now);
    const row = [
      id,
      title,
      String(essay.theme || '').trim(),
      String(essay.standard || '').trim(),
      String(essay.easy || '').trim(),
      imageFileId,
      imageUrl,
      createdAt,
      now,
      true
    ];
    const rowNum = found || Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(rowNum, 1, 1, 10).setValues([row]);
    return rowToEssay_(row);
  } finally {
    lock.releaseLock();
  }
}

function deleteEssay_(id) {
  if (!id) throw new Error('ID karangan tiada.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ESSAYS_SHEET);
    const found = findRowById_(sheet, id);
    const now = new Date().toISOString();
    if (found) {
      const row = sheet.getRange(found, 1, 1, 10).getValues()[0];
      row[8] = now;
      row[9] = false;
      sheet.getRange(found, 1, 1, 10).setValues([row]);
    } else {
      sheet.appendRow([id,'','','','','','',now,now,false]);
    }
  } finally {
    lock.releaseLock();
  }
}

function findRowById_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '') === id) return i + 2;
  }
  return 0;
}

function rowToEssay_(r) {
  return {
    id: String(r[0] || ''),
    title: String(r[1] || ''),
    theme: String(r[2] || ''),
    standard: String(r[3] || ''),
    easy: String(r[4] || ''),
    imageFileId: String(r[5] || ''),
    imageUrl: String(r[6] || ''),
    createdAt: r[7] ? String(r[7]) : '',
    updatedAt: r[8] ? String(r[8]) : '',
    active: !(r[9] === false || String(r[9]).toLowerCase() === 'false')
  };
}

function saveImage_(dataUrl, title) {
  const match = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!match) throw new Error('Format imej tidak sah.');
  const mime = match[1];
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 3 * 1024 * 1024) throw new Error('Imej melebihi 3 MB.');
  const ext = mime.indexOf('png') >= 0 ? 'png' : mime.indexOf('webp') >= 0 ? 'webp' : 'jpg';
  const safe = String(title || 'karangan').replace(/[^a-zA-Z0-9 _-]+/g, '').trim().replace(/\s+/g, '_').slice(0, 60) || 'karangan';
  const blob = Utilities.newBlob(bytes, mime, safe + '_' + Date.now() + '.' + ext);
  const folder = getUploadFolder_();
  const file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
  return {
    fileId: file.getId(),
    url: 'https://drive.google.com/uc?export=view&id=' + encodeURIComponent(file.getId())
  };
}

function getUploadFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('UPLOAD_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (err) {}
  }
  const it = DriveApp.getFoldersByName(UPLOAD_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(UPLOAD_FOLDER_NAME);
  props.setProperty('UPLOAD_FOLDER_ID', folder.getId());
  return folder;
}

function getOrCreateUploadFolder_() {
  return getUploadFolder_();
}

function trashFileSafe_(fileId) {
  if (!fileId) return;
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) {}
}

function checkPin_(pin) {
  const expected = PropertiesService.getScriptProperties().getProperty('TEACHER_PIN') || '1234';
  if (String(pin || '') !== expected) throw new Error('PIN guru salah.');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
