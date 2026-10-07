/**
 * ===== 表紙・報告書（新書式）の追加 =====
 *
 * 運転記録表のExcelに「表紙」「保守点検作業報告書」の2シートを先頭に追加する。
 * 書式は日本キヤリア様式（2026年7月〜の新書式）。
 *
 * 事前準備：
 *   1. リポジトリの report_template.xlsx をGoogleドライブにアップロード
 *   2. 右クリック →「アプリで開く」→「Google スプレッドシート」で変換
 *   3. 変換後スプレッドシートのIDをコピー
 *   4. スクリプトプロパティ REPORT_TEMPLATE_SHEET_ID にそのIDを設定
 *      （プロジェクトの設定 →「スクリプト プロパティ」→ 追加）
 *
 * テンプレートのセル配置（保守点検作業報告書シート）：
 *   B7        宛先（表紙 A2 にも同じ値）
 *   C17       現場名       J17  作業日
 *   C18       住所         J19  作業者
 *   D21       件名（◇の見出し行）
 *   C23:D31   作業内容（1行1項目。C=番号、D=項目）… 9行。超えたら行を挿入
 *   C34       作業結果の1行目
 *   C36:D51   作業結果の2行目以降（※の行はC=※、D=本文。続き行はD）… 16行。超えたら行を挿入
 *   52行目〜  備考（添付書類）
 */

const REPORT_COVER_SHEET = '表紙';
const REPORT_BODY_SHEET = '保守点検作業報告書';

const REPORT_ITEM_FIRST_ROW = 23;
const REPORT_ITEM_ROWS = 9;          // 23〜31
const REPORT_RESULT_FIRST_ROW = 36;
const REPORT_RESULT_ROWS = 16;       // 36〜51
const REPORT_WRAP_WIDTH = 72;        // 半角=1, 全角=2 として1行あたりの最大幅（D〜M列に収まる目安）
const REPORT_LAST_COL = 13;          // M列

/** 出力ブック ss の先頭に、表紙と報告書を追加して値を書き込む */
function addReportSheets_(ss, data, templateId) {
  const tss = SpreadsheetApp.openById(templateId);
  const srcCover = tss.getSheetByName(REPORT_COVER_SHEET);
  const srcBody = tss.getSheetByName(REPORT_BODY_SHEET);
  if (!srcCover || !srcBody) {
    throw new Error('報告書テンプレートに「' + REPORT_COVER_SHEET + '」「' + REPORT_BODY_SHEET + '」シートが見つかりません');
  }

  const cover = srcCover.copyTo(ss).setName(REPORT_COVER_SHEET);
  const body = srcBody.copyTo(ss).setName(REPORT_BODY_SHEET);

  // 印刷範囲はコピーされないため、M列より右（テンプレートの余白）を削除して、1ページ幅に収める
  const extraCols = body.getMaxColumns() - REPORT_LAST_COL;
  if (extraCols > 0) body.deleteColumns(REPORT_LAST_COL + 1, extraCols);

  const layout = buildReportLayout_(data);
  writeReportCover_(cover, layout);
  writeReportBody_(body, layout);

  // 先頭（表紙→報告書の順）へ移動
  cover.activate();
  ss.moveActiveSheet(1);
  body.activate();
  ss.moveActiveSheet(2);
  cover.activate();
}

/** 表紙の書き込み */
function writeReportCover_(sheet, layout) {
  sheet.getRange('A2').setValue(layout.addressee);
}

/** 報告書の書き込み。下の項目（作業結果）から先に書き、行を挿入しても上の位置がずれないようにする */
function writeReportBody_(sheet, layout) {
  sheet.getRange('B7').setValue(layout.addressee);
  sheet.getRange('C17').setValue(layout.siteName);
  sheet.getRange('C18').setValue(layout.address);
  // 日付として解釈されて「Wednesday, October 7, 2026」のような形式になるのを防ぐため、書式を文字列にしてから書き込む
  sheet.getRange('J17').setNumberFormat('@').setValue(layout.workDate);
  sheet.getRange('J19').setValue(layout.workers);
  sheet.getRange('D21').setValue(layout.subject);

  // --- 作業結果（先に処理） ---
  const resultEntries = layout.result;
  if (resultEntries.length > 0) {
    sheet.getRange('C34').setValue(resultEntries[0].text);
  }
  const resultRest = resultEntries.slice(1);
  const resultEndRow = REPORT_RESULT_FIRST_ROW + REPORT_RESULT_ROWS; // 52 = 備考の行
  if (resultRest.length > REPORT_RESULT_ROWS) {
    insertReportRows_(sheet, resultEndRow, resultRest.length - REPORT_RESULT_ROWS);
  }
  resultRest.forEach(function (e, i) {
    const r = REPORT_RESULT_FIRST_ROW + i;
    if (e.marker) sheet.getRange(r, 3).setValue(e.marker);
    sheet.getRange(r, 4).setValue(e.text);
  });

  // --- 作業内容 ---
  const items = layout.items;
  const itemEndRow = REPORT_ITEM_FIRST_ROW + REPORT_ITEM_ROWS; // 32 = 作業結果の行
  if (items.length > REPORT_ITEM_ROWS) {
    insertReportRows_(sheet, itemEndRow, items.length - REPORT_ITEM_ROWS);
  }
  items.forEach(function (it, i) {
    const r = REPORT_ITEM_FIRST_ROW + i;
    sheet.getRange(r, 3).setValue(it.no);
    sheet.getRange(r, 4).setValue(it.text);
  });
}

/** beforeRow の手前に n 行挿入し、直前の行の書式をコピーする */
function insertReportRows_(sheet, beforeRow, n) {
  sheet.insertRowsBefore(beforeRow, n);
  const src = sheet.getRange(beforeRow - 1, 1, 1, REPORT_LAST_COL);
  src.copyTo(sheet.getRange(beforeRow, 1, n, REPORT_LAST_COL), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

// ---------------------------------------------------------------------------
// 以下は純粋関数（Sheet APIを使わない）。レイアウトの計算だけを行う
// ---------------------------------------------------------------------------

/** 入力データから、報告書に書く内容を組み立てる */
function buildReportLayout_(data) {
  data = data || {};
  const lineSplit = function (v) {
    return String(v || '').split(/\r?\n/).map(function (l) { return l.replace(/\s+$/, ''); }).filter(function (l) { return l.trim() !== ''; });
  };

  const items = [];
  lineSplit(data['報告_作業内容']).forEach(function (raw) {
    const text = raw.trim().replace(/^(\d+[．.]|[・･])\s*/, '');
    reportWrap_(text, REPORT_WRAP_WIDTH).forEach(function (piece, k) {
      if (k === 0) items.push({ no: (items.filter(function (x) { return x.no; }).length + 1) + '．', text: piece });
      else items.push({ no: '', text: piece });
    });
  });

  const result = [];
  lineSplit(data['報告_作業結果']).forEach(function (raw, idx) {
    let text = raw.trim();
    let marker = '';
    if (idx > 0 && (text.charAt(0) === '※' || text.charAt(0) === '■')) {
      marker = text.charAt(0);
      text = text.slice(1).trim();
    }
    reportWrap_(text, REPORT_WRAP_WIDTH).forEach(function (piece, k) {
      result.push({ marker: (k === 0 ? marker : ''), text: piece });
    });
  });

  return {
    addressee: reportNormalizeSpaces_(data['報告_宛先']),
    siteName: reportSiteName_(data['現場名']),
    address: reportNormalizeSpaces_(data['住所']),
    workDate: reportJapaneseDate_(data['作業日時']),
    workers: reportNormalizeSpaces_(data['作業者']),
    subject: reportNormalizeSpaces_(data['報告_件名']),
    items: items,
    result: result
  };
}

function reportNormalizeSpaces_(v) {
  return String(v || '').replace(/\u3000/g, ' ').replace(/\s+/g, ' ').trim();
}

/** 現場名の末尾に敬称が無ければ「 様」を付ける */
function reportSiteName_(v) {
  const name = reportNormalizeSpaces_(v);
  if (!name) return '';
  return /(様|殿|御中)$/.test(name) ? name : name + ' 様';
}

/** yyyy-mm-dd / yyyy/mm/dd を「2026年7月7日」に。既に日本語表記や範囲ならそのまま */
function reportJapaneseDate_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m) return m[1] + '年' + Number(m[2]) + '月' + Number(m[3]) + '日';
  return s;
}

/** 表示幅（半角=1, 全角=2）で折り返す */
function reportWrap_(text, maxWidth) {
  const out = [];
  let cur = '';
  let w = 0;
  const chars = Array.from(String(text));
  const NO_LINE_START = '。、，．）」』】';
  const isAscii = function (c) { return /[\x21-\x7E]/.test(c); };  // 半角の英数字・記号（空白は除く）
  chars.forEach(function (ch) {
    const cw = ch.charCodeAt(0) > 255 && !(ch.charCodeAt(0) >= 0xFF61 && ch.charCodeAt(0) <= 0xFF9F) ? 2 : 1;
    if (w + cw > maxWidth && cur !== '') {
      if (NO_LINE_START.indexOf(ch) >= 0) {
        // 句読点・閉じ括弧は行頭に置かず、前の行にぶら下げる
        cur += ch;
        out.push(cur);
        cur = '';
        w = 0;
        return;
      }
      // 半角英数字の途中で切らない（「2,500hr」「ATR-1-2」など）。直前の連続部分を次の行へ送る
      let carry = '';
      if (isAscii(ch)) {
        const m = cur.match(/[\x21-\x7E]+$/);
        if (m && m[0].length < cur.length) { carry = m[0]; cur = cur.slice(0, cur.length - carry.length); }
      }
      out.push(cur);
      cur = carry;
      w = Array.from(carry).length;
    }
    cur += ch;
    w += cw;
  });
  if (cur !== '' || out.length === 0) out.push(cur);
  return out;
}
