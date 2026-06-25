/**
 * ===== DB（Database.gs）の動作確認テスト =====
 *
 * 記録の保存・一覧・取得・更新・削除、写真の保存・一覧・取得・削除を
 * 一通り実行し、実行ログで結果を確認できるようにする。
 *
 * 事前準備：
 *   1. Database.gs の説明に従い、DB_SPREADSHEET_ID / PHOTOS_FOLDER_ID を設定
 *   2. initDb を実行（records / photos シートを作成）
 *   3. この testDatabaseCrud を実行
 */
function testDatabaseCrud() {
  Logger.log('--- 1. 記録を新規保存 ---');
  const newId = saveRecord({
    name: 'CRUDテスト現場',
    data: { '現場名': 'CRUDテスト現場', '住所': 'テスト県テスト市1-1-1', '作業者': '内山太郎' }
  });
  Logger.log('保存した記録ID: ' + newId);

  Logger.log('--- 2. 記録一覧を取得 ---');
  const list1 = getRecords();
  Logger.log('一覧件数: ' + list1.length);
  Logger.log(JSON.stringify(list1.slice(0, 3)));

  Logger.log('--- 3. 記録を1件取得 ---');
  const got = getRecord(newId);
  Logger.log(JSON.stringify(got));

  Logger.log('--- 4. 記録を更新（同じidで再保存） ---');
  saveRecord({
    id: newId,
    name: 'CRUDテスト現場（更新後）',
    data: { '現場名': 'CRUDテスト現場（更新後）', '備考': '更新テスト済み' }
  });
  const updated = getRecord(newId);
  Logger.log('更新後のname: ' + updated.name + ' / data.備考: ' + updated.data['備考']);

  Logger.log('--- 5. 写真を保存（サンプル画像） ---');
  const blob = UrlFetchApp.fetch('https://picsum.photos/seed/dbcrud/400/300').getBlob();
  const photoId = savePhoto(newId, blob, 'test.jpg', 'CRUDテスト写真');
  Logger.log('保存した写真ID: ' + photoId);

  Logger.log('--- 6. 写真一覧を取得 ---');
  const photos = getPhotos(newId);
  Logger.log('この記録の写真件数: ' + photos.length);
  Logger.log(JSON.stringify(photos));

  Logger.log('--- 7. 写真本体を取得 ---');
  const photoBlob = getPhotoFile(photoId);
  Logger.log('取得した写真サイズ(byte): ' + (photoBlob ? photoBlob.getBytes().length : 'なし'));

  Logger.log('--- 8. 写真を削除 ---');
  deletePhoto(photoId);
  Logger.log('削除後の写真件数: ' + getPhotos(newId).length);

  Logger.log('--- 9. 記録を削除 ---');
  deleteRecord(newId);
  const afterDelete = getRecord(newId);
  Logger.log('削除後にgetRecordした結果: ' + (afterDelete === null ? '正常にnull(削除成功)' : 'エラー:まだ残っている'));

  Logger.log('=== 全テスト完了 ===');
}
