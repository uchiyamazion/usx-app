/**
 * ===== スプレッドシートDB（records / photos） =====
 *
 * server.py の PostgreSQL/SQLite版DBをGoogleスプレッドシートで再現する。
 *
 * テーブル構造：
 *   records シート: id | name | created_at | data(JSON文字列)
 *   photos  シート: id | record_id | filename | caption | created_at | drive_file_id
 *
 * 画像本体はスプレッドシートのセルに直接入れず、Google Driveの専用フォルダに
 * ファイルとして保存し、photosシートにはDriveのファイルIDだけを持たせる。
 * （Sheetsのセルはサイズ上限があり、Base64画像を直接入れるのは非効率かつ危険）
 *
 * 事前準備：
 *   1. 新しい空のGoogleスプレッドシートを1つ作成し、DB専用にする
 *      （名前は何でも良い。例：「USX_DB」）
 *   2. そのスプレッドシートのIDをコピー
 *   3. 写真保存用のGoogle Driveフォルダを1つ作成（PhotoExportTest.gsの
 *      OUTPUT_FOLDER_IDとは別の、永続的な保存先として新規に用意することを推奨）
 *   4. setupDatabaseProperties 内のIDを書き換えて1回実行
 *   5. initDb を実行してシート（ヘッダー行）を作成
 *   6. testDatabaseCrud を実行して動作確認
 */

const RECORDS_SHEET_NAME = 'records';
const PHOTOS_SHEET_NAME = 'photos';

function getDbSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
  if (!id) throw new Error('スクリプトプロパティ DB_SPREADSHEET_ID が未設定です。');
  return SpreadsheetApp.openById(id);
}

function getPhotosFolder_() {
  const id = PropertiesService.getScriptProperties().getProperty('PHOTOS_FOLDER_ID');
  if (!id) throw new Error('スクリプトプロパティ PHOTOS_FOLDER_ID が未設定です。');
  return DriveApp.getFolderById(id);
}

/**
 * records / photos シートが無ければヘッダー付きで作成する。
 * 既に存在する場合は何もしない（実行しても壊れない）。
 */
function initDb() {
  const ss = getDbSpreadsheet_();

  let recSheet = ss.getSheetByName(RECORDS_SHEET_NAME);
  if (!recSheet) {
    recSheet = ss.insertSheet(RECORDS_SHEET_NAME);
    recSheet.appendRow(['id', 'name', 'created_at', 'data']);
    recSheet.setFrozenRows(1);
  }

  let photoSheet = ss.getSheetByName(PHOTOS_SHEET_NAME);
  if (!photoSheet) {
    photoSheet = ss.insertSheet(PHOTOS_SHEET_NAME);
    photoSheet.appendRow(['id', 'record_id', 'filename', 'caption', 'created_at', 'drive_file_id']);
    photoSheet.setFrozenRows(1);
  }

  // デフォルトの「シート1」が残っていて空なら削除（任意）
  const defaultSheet = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 2) {
    const isEmpty = defaultSheet.getDataRange().getNumRows() <= 1 && defaultSheet.getDataRange().getNumColumns() <= 1;
    if (isEmpty) ss.deleteSheet(defaultSheet);
  }

  Logger.log('DB初期化完了: ' + ss.getUrl());
}

/** シート全体を見出し付きオブジェクト配列として取得する共通ヘルパー */
function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .filter(function (row) { return row[0] !== '' && row[0] !== null; })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    });
}

// ===== records =====

/**
 * 記録を1件保存する。idは未指定なら新規発行（新規作成）、
 * 指定されていれば該当行を上書き（更新）する。
 * @param {Object} record { id?, name, data(オブジェクト) }
 * @return {string} 保存したレコードのid
 */
function saveRecord(record) {
  const ss = getDbSpreadsheet_();
  const sheet = ss.getSheetByName(RECORDS_SHEET_NAME);
  const id = record.id || Utilities.getUuid();
  const createdAt = record.created_at || new Date().toISOString();
  const dataJson = JSON.stringify(record.data || {});
  const name = record.name || (record.data && record.data['現場名']) || '';

  const values = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let r = 1; r < values.length; r++) {
    if (values[r][0] === id) { targetRow = r + 1; break; }
  }

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, 4).setValues([[id, name, values[targetRow - 1][2], dataJson]]);
  } else {
    sheet.appendRow([id, name, createdAt, dataJson]);
  }
  return id;
}

/** 記録一覧を新しい順に返す（id, name, created_atのみ。dataは含めない） */
function getRecords() {
  const ss = getDbSpreadsheet_();
  const sheet = ss.getSheetByName(RECORDS_SHEET_NAME);
  const list = sheetToObjects_(sheet).map(function (r) {
    return { id: r.id, name: r.name, created_at: r.created_at };
  });
  list.sort(function (a, b) { return new Date(b.created_at) - new Date(a.created_at); });
  return list;
}

/** 1件取得（dataはJSON.parse済みのオブジェクトで返す） */
function getRecord(id) {
  const ss = getDbSpreadsheet_();
  const sheet = ss.getSheetByName(RECORDS_SHEET_NAME);
  const found = sheetToObjects_(sheet).filter(function (r) { return r.id === id; })[0];
  if (!found) return null;
  return {
    id: found.id,
    name: found.name,
    created_at: found.created_at,
    data: JSON.parse(found.data || '{}')
  };
}

/** 記録を削除する。関連する写真も一緒に削除する。 */
function deleteRecord(id) {
  const ss = getDbSpreadsheet_();
  const sheet = ss.getSheetByName(RECORDS_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  for (let r = values.length - 1; r >= 1; r--) {
    if (values[r][0] === id) { sheet.deleteRow(r + 1); break; }
  }
  getPhotos(id).forEach(function (p) { deletePhoto(p.id); });
}

// ===== photos =====

/**
 * 写真を1枚保存する。画像本体はDriveに保存し、メタデータをphotosシートに追加。
 * @param {string} recordId
 * @param {Blob} blob 画像のBlob
 * @param {string} filename
 * @param {string} caption
 * @return {string} 保存した写真のid
 */
function savePhoto(recordId, blob, filename, caption) {
  const ss = getDbSpreadsheet_();
  const sheet = ss.getSheetByName(PHOTOS_SHEET_NAME);
  const folder = getPhotosFolder_();

  const id = Utilities.getUuid();
  const createdAt = new Date().toISOString();
  const driveFile = folder.createFile(blob.setName(filename || (id + '.jpg')));

  sheet.appendRow([id, recordId, filename || driveFile.getName(), caption || '', createdAt, driveFile.getId()]);
  return id;
}

/** 指定した記録に紐づく写真一覧を作成日時順に返す（画像本体は含まない） */
function getPhotos(recordId) {
  const ss = getDbSpreadsheet_();
  const sheet = ss.getSheetByName(PHOTOS_SHEET_NAME);
  const list = sheetToObjects_(sheet).filter(function (p) { return p.record_id === recordId; });
  list.sort(function (a, b) { return new Date(a.created_at) - new Date(b.created_at); });
  return list;
}

/** 写真本体（Blob）をDriveから取得する */
function getPhotoFile(photoId) {
  const ss = getDbSpreadsheet_();
  const sheet = ss.getSheetByName(PHOTOS_SHEET_NAME);
  const found = sheetToObjects_(sheet).filter(function (p) { return p.id === photoId; })[0];
  if (!found) return null;
  return DriveApp.getFileById(found.drive_file_id).getBlob();
}

/** 写真を削除する（Drive上のファイルとphotosシートの行の両方） */
function deletePhoto(photoId) {
  const ss = getDbSpreadsheet_();
  const sheet = ss.getSheetByName(PHOTOS_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  for (let r = values.length - 1; r >= 1; r--) {
    if (values[r][0] === photoId) {
      const driveFileId = values[r][5];
      try { DriveApp.getFileById(driveFileId).setTrashed(true); } catch (e) { /* 既に削除済みなら無視 */ }
      sheet.deleteRow(r + 1);
      break;
    }
  }
}

/**
 * スクリプトプロパティを設定するヘルパー。
 */
function setupDatabaseProperties() {
  PropertiesService.getScriptProperties().setProperties({
    DB_SPREADSHEET_ID: 'ここに新規作成したDB用スプレッドシートのIDを貼る',
    PHOTOS_FOLDER_ID: 'ここに写真保存用DriveフォルダのIDを貼る'
  });
}
