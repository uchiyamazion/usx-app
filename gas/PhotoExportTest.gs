/**
 * ===== 写真台帳エクスポート検証スクリプト =====
 *
 * 目的：GAS移行で最も再現が難しい「写真をセル範囲にぴったり合わせて貼り付け、
 *       xlsxとして書き出す」処理が可能かどうかを検証する。
 *
 * 現行Flask版(server.py make_excel)の仕様：
 *   - photo_template.xlsx は16シート×1シート3枚 = 最大48枚
 *   - 1シート内の3枚の貼付位置（B列〜F列 × 指定行範囲）
 *       1枚目: 行3〜23
 *       2枚目: 行25〜45
 *       3枚目: 行47〜67
 *   - 日付/キャプションはH列の指定行に書き込み
 *       1枚目: 日付=行4, キャプション=行7
 *       2枚目: 日付=行26, キャプション=行29
 *       3枚目: 日付=行48, キャプション=行51
 *   - openpyxlのTwoCellAnchorで「セル範囲にフィット」させていたため、
 *     559×391pxという数値はその結果であり固定値ではない（フォールバック時のみ使用）。
 *
 * 事前準備：
 *   1. photo_template.xlsx をGoogleドライブにアップロード
 *   2. 右クリック →「アプリで開く」→「Google スプレッドシート」で変換コピーを作成
 *      （元のxlsxとは別に、変換された新しいGoogleスプレッドシートが作られる）
 *   3. 変換後のスプレッドシートのIDをコピー（URLの /d/ と /edit の間の文字列）
 *   4. 出力先にしたいGoogle DriveフォルダのIDも用意
 *   5. Apps Scriptエディタのプロジェクトの設定 →「スクリプト プロパティ」で以下を登録
 *        PHOTO_TEMPLATE_SHEET_ID = (変換後スプレッドシートのID)
 *        OUTPUT_FOLDER_ID        = (出力先フォルダのID)
 *   6. testPhotoExport を実行（初回は権限許可が必要）
 *   7. 実行ログ(表示 → ログ)に出力されるxlsxファイルのURLを開き、
 *      写真がセル範囲にぴったり収まっているか目視確認する
 */

const PHOTO_SLOT_RANGES = [[3, 23], [25, 45], [47, 67]]; // [開始行, 終了行] 1スロットにつき3枚分
const PHOTO_INFO_ROWS = [
  { date: 4, content: 7 },
  { date: 26, content: 29 },
  { date: 48, content: 51 }
];
const PHOTO_COL_START = 2; // B列 (1-indexed)
const PHOTO_COL_END = 6;   // F列 (1-indexed)
const CAPTION_COL = 8;     // H列 (1-indexed)

/**
 * 検証用メイン関数。テスト用のサンプル画像3枚を1シート目に貼り付け、
 * xlsxとしてエクスポートしてDriveに保存する。
 */
function testPhotoExport() {
  const props = PropertiesService.getScriptProperties();
  const templateId = props.getProperty('PHOTO_TEMPLATE_SHEET_ID');
  const outputFolderId = props.getProperty('OUTPUT_FOLDER_ID');

  if (!templateId) {
    throw new Error('スクリプトプロパティ PHOTO_TEMPLATE_SHEET_ID が未設定です。');
  }

  const folder = outputFolderId ? DriveApp.getFolderById(outputFolderId) : DriveApp.getRootFolder();
  const ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');

  // テンプレートをコピー（元データを汚さない）
  const copyFile = DriveApp.getFileById(templateId).makeCopy('写真台帳_テスト_' + ts, folder);
  const ss = SpreadsheetApp.openById(copyFile.getId());
  const sheets = ss.getSheets();

  // サンプル画像3枚（外部のダミー画像APIから取得。実運用では実際の写真Blobに置き換える）
  const testPhotos = [
    { url: 'https://picsum.photos/seed/usx1/640/480', caption: 'テスト写真1：屋外熱源機', date: '2026-06-24' },
    { url: 'https://picsum.photos/seed/usx2/640/480', caption: 'テスト写真2：配管接続部', date: '2026-06-24' },
    { url: 'https://picsum.photos/seed/usx3/640/480', caption: 'テスト写真3：銘板',     date: '2026-06-24' }
  ];

  let sheetIdx = 0;
  for (let i = 0; i < testPhotos.length; i++) {
    const slotIdx = i % 3;
    if (slotIdx === 0 && i > 0) sheetIdx++;
    if (sheetIdx >= sheets.length) break;

    const sheet = sheets[sheetIdx];
    const range = PHOTO_SLOT_RANGES[slotIdx];
    const info = PHOTO_INFO_ROWS[slotIdx];

    const blob = UrlFetchApp.fetch(testPhotos[i].url).getBlob();
    insertPhotoFitToRange(sheet, blob, range[0], range[1], PHOTO_COL_START, PHOTO_COL_END);

    sheet.getRange(info.date, CAPTION_COL).setValue(testPhotos[i].date);
    sheet.getRange(info.content, CAPTION_COL).setValue(testPhotos[i].caption);
  }

  SpreadsheetApp.flush();

  const xlsxBlob = exportSpreadsheetAsXlsx(copyFile.getId());
  const xlsxFile = folder.createFile(xlsxBlob).setName('写真台帳_テスト_' + ts + '.xlsx');

  Logger.log('検証用スプレッドシート: ' + ss.getUrl());
  Logger.log('xlsx出力ファイル: ' + xlsxFile.getUrl());
  return xlsxFile.getUrl();
}

/**
 * 指定したセル範囲（行・列）にぴったり収まるように画像を貼り付ける。
 * TwoCellAnchor(editAs='twoCell')の「セル範囲にフィットして伸縮する」挙動を、
 * 列幅・行高の合計値を画像の幅・高さに設定することで再現する。
 */
function insertPhotoFitToRange(sheet, blob, rowStart, rowEnd, colStart, colEnd) {
  const img = sheet.insertImage(blob, colStart, rowStart);

  let width = 0;
  for (let c = colStart; c <= colEnd; c++) width += sheet.getColumnWidth(c);

  let height = 0;
  for (let r = rowStart; r <= rowEnd; r++) height += sheet.getRowHeight(r);

  img.setWidth(width);
  img.setHeight(height);
  return img;
}

/**
 * Google スプレッドシートをxlsx形式のBlobとしてエクスポートする。
 * Drive REST APIのexportエンドポイントを、現在のユーザーのOAuthトークンで叩く。
 */
function exportSpreadsheetAsXlsx(fileId) {
  const url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token }
  });
  return response.getBlob();
}

/**
 * スクリプトプロパティを設定するためのヘルパー。
 * 一度だけ実行値を書き換えて実行すればOK（実行後は消してよい）。
 */
function setupScriptProperties() {
  PropertiesService.getScriptProperties().setProperties({
    PHOTO_TEMPLATE_SHEET_ID: 'ここに変換後スプレッドシートのIDを貼る',
    OUTPUT_FOLDER_ID: 'ここに出力先フォルダのIDを貼る'
  });
}
