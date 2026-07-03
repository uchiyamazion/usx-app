/**
 * ===== Webアプリ エントリポイント (doGet / doPost) =====
 *
 * フロントエンドはGitHub Pages（usx-app/index.html）に統一済み。
 * このGASはAPI専用（doPostのみ）として動作する。
 * doGetはヘルスチェック用に簡易メッセージを返すのみで、HTML描画は行わない。
 *
 * デプロイ方法：
 *   Apps Scriptエディタ右上「デプロイ」→「新しいデプロイ」→
 *   種類「ウェブアプリ」→ 実行ユーザー「自分」、アクセスできるユーザーは
 *   運用方針に応じて選択（社内のみなら「Google Workspace内の全員」等）
 */

function doGet(e) {
  return ContentService.createTextOutput('USX API is running. Frontend: https://uchiyamazion.github.io/usx-app/')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ error: 'リクエストの解析に失敗しました: ' + err.message });
  }

  var action = body.action;
  var payload = body.payload || {};

  try {
    switch (action) {
      case 'getRecords':
        return jsonOutput_(getRecords());

      case 'getRecord':
        return jsonOutput_(getRecord(payload.id));

      case 'saveRecord': {
        var id = saveRecord({ id: payload.id, name: payload.name, data: payload.data });
        return jsonOutput_({ id: id });
      }

      case 'deleteRecord':
        deleteRecord(payload.id);
        return jsonOutput_({ ok: true });

      case 'getPhotos':
        return jsonOutput_(getPhotos(payload.recordId));

      case 'uploadPhoto': {
        var blob = dataUrlToBlob_(payload.image);
        var photoId = savePhoto(payload.recordId, blob, 'photo_' + Date.now() + '.jpg', payload.caption);
        return jsonOutput_({ id: photoId });
      }

      case 'deletePhoto':
        deletePhoto(payload.id);
        return jsonOutput_({ ok: true });

      case 'getPhotoFile': {
        var fileBlob = getPhotoFile(payload.id);
        if (!fileBlob) return jsonOutput_({ error: '写真が見つかりません' });
        return jsonOutput_({
          base64: Utilities.base64Encode(fileBlob.getBytes()),
          mimeType: fileBlob.getContentType()
        });
      }

      case 'generateExcel':
        return jsonOutput_(generateOperationExcel_(payload.data, null));

      case 'generateExcelWithPhotos':
        return jsonOutput_(generateOperationExcel_(payload.data, payload.recordId));

      default:
        return jsonOutput_({ error: '不明なactionです: ' + action });
    }
  } catch (err) {
    return jsonOutput_({ error: err.message });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** data:image/jpeg;base64,xxxx 形式の文字列をGASのBlobに変換する */
function dataUrlToBlob_(dataUrl) {
  var m = String(dataUrl).match(/^data:([^;]+);base64,(.*)$/);
  var mime = m ? m[1] : 'image/jpeg';
  var b64 = m ? m[2] : dataUrl;
  return Utilities.newBlob(Utilities.base64Decode(b64), mime);
}

/**
 * 運転記録表（本体＋必要なら写真台帳シートを結合）をxlsxとして生成し、
 * base64文字列で返す。Flask版 make_excel() のGAS版相当。
 * recordId が渡された場合のみ、DBに紐づく写真を取得してシートを追加する
 * （server.pyのgenerate_with_photosと同じ「同一ファイルに結合」方式）。
 */
function generateOperationExcel_(data, recordId) {
  var props = PropertiesService.getScriptProperties();
  var opTemplateId = props.getProperty('OPERATION_TEMPLATE_SHEET_ID');
  var photoTemplateId = props.getProperty('PHOTO_TEMPLATE_SHEET_ID');
  var outputFolderId = props.getProperty('OUTPUT_FOLDER_ID') || props.getProperty('PHOTOS_FOLDER_ID');

  if (!opTemplateId) throw new Error('OPERATION_TEMPLATE_SHEET_ID が未設定です（スクリプトプロパティを確認してください）');

  var folder = outputFolderId ? DriveApp.getFolderById(outputFolderId) : DriveApp.getRootFolder();
  var ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var siteName = String(data && data['現場名'] || '').replace(/[\/\s]/g, '_').slice(0, 20);
  var baseName = '運転記録表_' + (siteName || '現場') + '_' + ts;

  var copyFile = DriveApp.getFileById(opTemplateId).makeCopy(baseName, folder);
  var ss = SpreadsheetApp.openById(copyFile.getId());
  var sheets = ss.getSheets();

  [[0, 4], [4, 8], [8, 12]].forEach(function (range, idx) {
    if (idx < sheets.length) writeOperationRecordSheet(sheets[idx], data || {}, range[0], range[1]);
  });

  // 写真台帳シートを同じファイルに追加結合する（recordIdがあり、写真が存在する場合のみ）
  if (recordId && photoTemplateId) {
    var photos = getPhotos(recordId);
    if (photos.length > 0) {
      var sheetsNeeded = Math.ceil(photos.length / 3);
      var photoTemplateSS = SpreadsheetApp.openById(photoTemplateId);
      var photoSourceSheets = photoTemplateSS.getSheets();
      var addedSheets = [];

      for (var s = 0; s < sheetsNeeded && s < photoSourceSheets.length; s++) {
        var newSheet = photoSourceSheets[s].copyTo(ss);
        newSheet.setName('写真台帳' + (s + 1));
        addedSheets.push(newSheet);
      }

      var sheetIdx = 0;
      for (var i = 0; i < photos.length; i++) {
        var slotIdx = i % 3;
        if (slotIdx === 0 && i > 0) sheetIdx++;
        if (sheetIdx >= addedSheets.length) break;

        var sheet = addedSheets[sheetIdx];
        var range2 = PHOTO_SLOT_RANGES[slotIdx];
        var info = PHOTO_INFO_ROWS[slotIdx];
        var fileBlob = getPhotoFile(photos[i].id);
        if (fileBlob) {
          insertPhotoFitToRange(sheet, fileBlob, range2[0], range2[1], PHOTO_COL_START, PHOTO_COL_END);
        }
        sheet.getRange(info.date, CAPTION_COL).setValue(String(photos[i].created_at || '').slice(0, 10));
        sheet.getRange(info.content, CAPTION_COL).setValue(photos[i].caption || '');
      }
    }
  }

  SpreadsheetApp.flush();

  var xlsxBlob = exportSpreadsheetAsXlsx(copyFile.getId());
  return {
    base64: Utilities.base64Encode(xlsxBlob.getBytes()),
    filename: baseName + '.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  };
}
