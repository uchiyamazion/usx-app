/**
 * ===== 運転記録表本体 セル書き込み 検証スクリプト =====
 *
 * server.py の write_sheet() / make_excel() のロジックをGAS用に移植したもの。
 * 写真台帳と違い、こちらは単純なセル書き込み（マージセル対応含む）なので
 * 技術的難易度は低いが、フィールド数が多いので転記漏れがないか検証する。
 *
 * 事前準備：
 *   1. template.xlsx をGoogleドライブにアップロード
 *   2. 右クリック →「アプリで開く」→「Google スプレッドシート」で変換
 *      （元のtemplate.xlsxは3シート構成: 1-4号機/5-8号機/9-12号機。
 *       変換後もシートはそのまま3枚保持される想定）
 *   3. 変換後スプレッドシートのIDをコピー
 *   4. setupOperationRecordProperties 内のIDを書き換えて1回実行
 *   5. testOperationRecordExport を実行
 *   6. 出力されたxlsxを開き、各セルが正しく入っているか確認
 */

// 熱源機1台ごとの列マッピング（No.1=G/H, No.2=K/L, No.3=O/P, No.4=S/T ... No.12=AY/AZ）
const UNIT_COLS = [
  { info: 'G', no_val: 'H', A: 'G', B: 'H', C: 'I', D: 'J' },
  { info: 'K', no_val: 'L', A: 'K', B: 'L', C: 'M', D: 'N' },
  { info: 'O', no_val: 'P', A: 'O', B: 'P', C: 'Q', D: 'R' },
  { info: 'S', no_val: 'T', A: 'S', B: 'T', C: 'U', D: 'V' },
  { info: 'W', no_val: 'X', A: 'W', B: 'X', C: 'Y', D: 'Z' },
  { info: 'AA', no_val: 'AB', A: 'AA', B: 'AB', C: 'AC', D: 'AD' },
  { info: 'AE', no_val: 'AF', A: 'AE', B: 'AF', C: 'AG', D: 'AH' },
  { info: 'AI', no_val: 'AJ', A: 'AI', B: 'AJ', C: 'AK', D: 'AL' },
  { info: 'AM', no_val: 'AN', A: 'AM', B: 'AN', C: 'AO', D: 'AP' },
  { info: 'AQ', no_val: 'AR', A: 'AQ', B: 'AR', C: 'AS', D: 'AT' },
  { info: 'AU', no_val: 'AV', A: 'AU', B: 'AV', C: 'AW', D: 'AX' },
  { info: 'AY', no_val: 'AZ', A: 'AY', B: 'AZ', C: 'BA', D: 'BB' }
];

// [フィールド名, 行番号, 回路A~D個別か否か]
const FIELD_ROWS = [
  ['運転時間', 16, true], ['運転回数', 17, true],
  ['冷温水入口圧力', 18, false], ['冷温水出口圧力', 19, false], ['換算流量', 20, false],
  ['圧縮機電流_R', 21, true], ['圧縮機電流_S', 22, true], ['圧縮機電流_T', 23, true],
  ['圧縮機電圧_RS', 24, false], ['圧縮機電圧_ST', 25, false], ['圧縮機電圧_TR', 26, false],
  ['冷温水入口温度', 27, false], ['冷温水中間温度', 28, false], ['冷温水出口温度', 29, false],
  ['外気温度', 30, false],
  ['冷媒高圧圧力', 31, true], ['冷媒低圧圧力', 32, true],
  ['冷媒吐出ガス温度', 33, true], ['冷媒吸入ガス温度', 34, true],
  ['冷媒コイルガス温度1', 35, true], ['冷媒コイルガス温度2', 36, true],
  ['ファン回転数', 37, true], ['圧縮機運転周波数', 38, true],
  ['膨脹弁1開度', 39, true], ['膨脹弁2開度', 40, true],
  ['高圧圧力異常', 42, true], ['絶縁抵抗', 44, true],
  ['出荷時充填量', 46, true], ['初期充填量', 47, true], ['総充填量', 48, true]
];

const CHECK_CELLS = {
  '圧縮機関係': { '異常音': 'F50', '異常振動': 'H50', '圧力不良': 'K50', 'オイル量': 'N50', '異常過熱': 'Q50' },
  '凝縮器関係': { '汚れ': 'F51', '流量不足': 'H51', 'ストレ詰り': 'K51', '風量不足': 'N51' },
  '蒸発器関係': { '汚れ': 'F52', '流量不足': 'H52', 'ストレ詰り': 'K52', '風量不足': 'N52' },
  '送風機関係': { '異常音': 'F53', '異常振動': 'H53', 'ショートサイクル': 'K53' },
  '冷媒配管系統部品関係': { 'ストレーナ': 'F54', '四方弁': 'H54', '電磁弁': 'K54', '膨張弁': 'N54' },
  '電装品関係': { 'センサー': 'F55', 'リレー': 'H55', 'トランス': 'K55', '制御基板': 'N55', '通信不具合': 'Q55' },
  '補機関係': { 'ポンプ': 'F56', '圧力計': 'H56', 'アキュムレータ': 'K56' },
  'その他': { 'その他': 'F57' }
};

const PRECHECK_MAP = {
  '電装品結線状態確認': ['F59', 'H59'], '通信線アドレス設定': ['F60', 'H60'],
  '熱源機外観点検': ['F61', 'H61'], '内蔵ポンプ動作確認': ['F62', 'H62'],
  '水配管フラッシング': ['Q59', 'T59'], '12時間前通電': ['Q60', 'T60'],
  '冷媒漏洩点検': ['Q61', 'T61'], 'インターロック動作確認': ['Q62', 'T62']
};

/**
 * マージセルを考慮してセルに値をセットする。
 * Googleスプレッドシートではマージ範囲内のどのセルを指定してもsetValueは
 * 左上セルに対して適用されるため、openpyxlのsafe_set相当の特別処理は
 * 基本的には不要だが、念のためA1記法のセルを取得して直接setValueする。
 */
function safeSet(sheet, a1, value) {
  if (value === undefined || value === null || value === '') return;
  sheet.getRange(a1).setValue(value);
}

/**
 * 1シート分の書き込み。server.py の write_sheet() に対応。
 * unitStart/unitEnd は0-indexed（例: 1-4号機シートなら 0, 4）
 */
function writeOperationRecordSheet(sheet, data, unitStart, unitEnd) {
  safeSet(sheet, 'D2', data['現場名']);
  safeSet(sheet, 'D3', data['住所']);
  safeSet(sheet, 'T3', data['作業日時']);
  safeSet(sheet, 'T4', data['作業者']);
  safeSet(sheet, 'E5', data['系統名']);
  safeSet(sheet, 'F6', data['型式']);
  safeSet(sheet, 'E7', data['運転状態']);
  if (data['設定温度'] !== undefined && data['設定温度'] !== null) {
    safeSet(sheet, 'F7', String(data['設定温度']) + '℃');
  }
  safeSet(sheet, 'T7', data['連結台数']);
  safeSet(sheet, 'E8', data['設置年月日']);
  safeSet(sheet, 'I8', data['冷媒種類']);
  safeSet(sheet, 'T70', data['作成者']);
  safeSet(sheet, 'B71', data['備考']);

  const workTypeCells = { '試運転': 'D4', '定期点検': 'F4', '簡易点検': 'H4', '修理': 'J4', '整備': 'L4', '故障判定': 'N4' };
  Object.keys(workTypeCells).forEach(function (k) {
    safeSet(sheet, workTypeCells[k], (data['作業区分'] === k ? '■' : '□') + k);
  });

  const freonCells = { '定期点検': 'Q8', '簡易点検': 'T8' };
  Object.keys(freonCells).forEach(function (k) {
    safeSet(sheet, freonCells[k], (data['フロン区分'] === k ? '■' : '□') + k);
  });

  const units = data['熱源機'] || [];
  for (let li = 0, ui = unitStart; ui < unitEnd; li++, ui++) {
    if (ui >= units.length) break;
    const u = units[ui];
    const cols = UNIT_COLS[li];
    if (u['No']) safeSet(sheet, cols.no_val + '11', u['No']);
    if (u['UC製造番号']) safeSet(sheet, cols.info + '12', u['UC製造番号']);
    if (u['冷却加熱']) safeSet(sheet, cols.info + '13', u['冷却加熱']);
    if (u['製造年']) safeSet(sheet, cols.info + '14', u['製造年']);
    if (u['記録時間']) safeSet(sheet, cols.info + '49', u['記録時間']);

    FIELD_ROWS.forEach(function (fr) {
      const field = fr[0], row = fr[1], isCircuit = fr[2];
      if (isCircuit) {
        ['A', 'B', 'C', 'D'].forEach(function (c) {
          const val = u[field + '_' + c];
          if (val !== undefined && val !== null && val !== '') {
            safeSet(sheet, cols[c] + row, val);
          }
        });
      } else {
        const val = (u[field + '_A'] !== undefined ? u[field + '_A'] : u[field]);
        if (val !== undefined && val !== null && val !== '') {
          safeSet(sheet, cols['A'] + row, val);
        }
      }
    });
  }

  const checks = data['チェック結果'] || {};
  Object.keys(CHECK_CELLS).forEach(function (group) {
    const items = CHECK_CELLS[group];
    const sel = checks[group] || [];
    Object.keys(items).forEach(function (item) {
      safeSet(sheet, items[item], (sel.indexOf(item) !== -1 ? '■' : '□') + item);
    });
  });

  const pre = data['試運転事前点検'] || {};
  Object.keys(PRECHECK_MAP).forEach(function (item) {
    const cells = PRECHECK_MAP[item];
    const r = pre[item] || '';
    safeSet(sheet, cells[0], r === 'OK' ? '■ＯＫ' : '□ＯＫ');
    safeSet(sheet, cells[1], r === 'NG' ? '■NG' : '□NG');
  });
}

/**
 * 検証用メイン関数。サンプルデータを3シート（1-4/5-8/9-12号機）に書き込み、
 * xlsxとしてエクスポートする。
 */
function testOperationRecordExport() {
  const props = PropertiesService.getScriptProperties();
  const templateId = props.getProperty('OPERATION_TEMPLATE_SHEET_ID');
  const outputFolderId = props.getProperty('OUTPUT_FOLDER_ID');

  if (!templateId) {
    throw new Error('スクリプトプロパティ OPERATION_TEMPLATE_SHEET_ID が未設定です。');
  }

  const folder = outputFolderId ? DriveApp.getFolderById(outputFolderId) : DriveApp.getRootFolder();
  const ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');

  const copyFile = DriveApp.getFileById(templateId).makeCopy('運転記録表_テスト_' + ts, folder);
  const ss = SpreadsheetApp.openById(copyFile.getId());
  const sheets = ss.getSheets();

  const sampleData = {
    '現場名': 'テスト現場株式会社',
    '住所': '愛知県名古屋市中区テスト1-2-3',
    '作業日時': '2026-06-25 10:00',
    '作業者': '内山太郎',
    '系統名': 'A系統',
    '型式': 'RUA-XXXP1',
    '運転状態': '正常',
    '設定温度': 7,
    '連結台数': 4,
    '設置年月日': '2024-04-01',
    '冷媒種類': 'R410A',
    '作成者': '内山太郎',
    '備考': 'GAS移行検証テストデータ',
    '作業区分': '試運転',
    'フロン区分': '定期点検',
    '熱源機': [
      {
        'No': 1, 'UC製造番号': 'UC-0001', '冷却加熱': '冷却', '製造年': '2024',
        '記録時間': '10:00',
        '運転時間_A': 1234, '運転回数_A': 56,
        '冷温水入口圧力': 0.25, '冷温水出口圧力': 0.2, '換算流量': 120,
        '圧縮機電流_R_A': 10.5, '圧縮機電流_S_A': 10.6, '圧縮機電流_T_A': 10.4,
        '圧縮機電圧_RS': 200, '圧縮機電圧_ST': 201, '圧縮機電圧_TR': 199,
        '冷温水入口温度': 12, '冷温水中間温度': 9, '冷温水出口温度': 7,
        '外気温度': 28,
        '冷媒高圧圧力_A': 2.8, '冷媒低圧圧力_A': 0.6,
        '冷媒吐出ガス温度_A': 75, '冷媒吸入ガス温度_A': 15,
        '冷媒コイルガス温度1_A': 30, '冷媒コイルガス温度2_A': 31,
        'ファン回転数_A': 900, '圧縮機運転周波数_A': 60,
        '膨脹弁1開度_A': 50, '膨脹弁2開度_A': 48,
        '高圧圧力異常_A': '正常', '絶縁抵抗_A': '2.5MΩ',
        '出荷時充填量_A': '10kg', '初期充填量_A': '10kg', '総充填量_A': '10kg'
      },
      {
        'No': 2, 'UC製造番号': 'UC-0002', '冷却加熱': '冷却', '製造年': '2024',
        '記録時間': '10:05',
        '運転時間_A': 1200, '運転回数_A': 54
      }
    ],
    'チェック結果': {
      '圧縮機関係': ['異常音', 'オイル量'],
      '送風機関係': ['異常音']
    },
    '試運転事前点検': {
      '電装品結線状態確認': 'OK',
      '熱源機外観点検': 'OK',
      '水配管フラッシング': 'NG'
    }
  };

  [[0, 4], [4, 8], [8, 12]].forEach(function (range, idx) {
    if (idx < sheets.length) {
      writeOperationRecordSheet(sheets[idx], sampleData, range[0], range[1]);
    }
  });

  SpreadsheetApp.flush();

  const xlsxBlob = exportSpreadsheetAsXlsx(copyFile.getId());
  const xlsxFile = folder.createFile(xlsxBlob).setName('運転記録表_テスト_' + ts + '.xlsx');

  Logger.log('検証用スプレッドシート: ' + ss.getUrl());
  Logger.log('xlsx出力ファイル: ' + xlsxFile.getUrl());
  return xlsxFile.getUrl();
}

/**
 * スクリプトプロパティを設定するヘルパー（PhotoExportTest.gsのOUTPUT_FOLDER_IDと共用）。
 */
function setupOperationRecordProperties() {
  PropertiesService.getScriptProperties().setProperty(
    'OPERATION_TEMPLATE_SHEET_ID',
    'ここに変換後スプレッドシートのIDを貼る'
  );
}
