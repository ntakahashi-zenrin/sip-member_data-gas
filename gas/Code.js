/**
 * Google Drive フォルダ内の会員データ（CSVから変換されたスプレッドシート）を読み込み、
 * このスプレッドシートの各シート（canceler / first_free / new_member / total_member）に
 * 集約して出力する。
 *
 * 実行方法：スプレッドシートのメニュー「データ更新」→「Driveから取り込み」から手動実行。
 * 実行のたびに対象シートの内容を全て消してから、Drive内の全該当ファイルを日付昇順で書き出す。
 */

var DRIVE_FOLDER_ID = '1-A0EVklg__0mg-xQEANjDBbFfgkEqDPF';

var IMPORT_CONFIGS = [
  {
    sheetName: 'canceler',
    filePattern: /^canceler_(\d{6})$/,
    headers: ['日付', 'Google会員数', 'Apple会員数', 'クレカ会員数', '法人会員数'],
  },
  {
    sheetName: 'first_free',
    filePattern: /^first_free_(\d{6})$/,
    headers: ['日付', 'Google会員数', 'Apple会員数', 'クレカ会員数'],
  },
  {
    sheetName: 'new_member',
    filePattern: /^new_member_(\d{6})$/,
    headers: ['日付', '無料会員数', 'Google会員数', 'Apple会員数', 'クレカ会員数', '法人会員数'],
  },
  {
    sheetName: 'total_member',
    filePattern: /^total_member_(\d{6})$/,
    headers: ['日付', '無料会員数', 'Google会員数', 'Apple会員数', 'クレカ会員数', '法人会員数'],
  },
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('データ更新')
    .addItem('Driveから取り込み', 'importAllMemberData')
    .addToUi();
}

function importAllMemberData() {
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  IMPORT_CONFIGS.forEach(function (config) {
    var rows = collectRowsForConfig(folder, config);
    writeRowsToSheet(ss, config, rows);
  });

  SpreadsheetApp.getUi().alert('データの取り込みが完了しました。');
}

function collectRowsForConfig(folder, config) {
  var matchedFiles = [];
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    var file = files.next();
    var match = config.filePattern.exec(file.getName());
    if (match) {
      matchedFiles.push({ yyyymm: match[1], file: file });
    }
  }

  matchedFiles.sort(function (a, b) {
    return a.yyyymm.localeCompare(b.yyyymm);
  });

  var rows = [];
  matchedFiles.forEach(function (entry) {
    var sourceSheet = SpreadsheetApp.openById(entry.file.getId()).getSheets()[0];
    var values = sourceSheet.getDataRange().getValues();
    values.forEach(function (row) {
      if (row[0] === '' || row[0] === null) {
        return;
      }
      var normalizedRow = row.slice(0, config.headers.length);
      normalizedRow[0] = String(normalizedRow[0]);
      rows.push(normalizedRow);
    });
  });

  return rows;
}

function writeRowsToSheet(ss, config, rows) {
  var sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(config.sheetName);
  }
  sheet.clearContents();

  var totalRows = rows.length + 1;
  sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
  sheet.getRange(1, 1, totalRows, 1).setNumberFormat('@');
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, config.headers.length).setValues(rows);
  }
}
