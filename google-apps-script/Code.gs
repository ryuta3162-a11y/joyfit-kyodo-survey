/**
 * JOYFIT24 経堂 アンケート → スプレッドシート
 * ・「100％／50％／0％」は列見出し1つずつ、そのセル内にサービス名をカンマ区切りでまとめる
 * ・送信のたびに 1 行目がヘッダーでなければ、行を挿入してヘッダーを自動配置（既存データは1行下にずれる）
 * ・手動で 1 行目だけ直したいときは setupResponseHeaders() を実行（1行目を上書き）
 */

const SPREADSHEET_ID = '1h_hG-ac0DwPLLPuxsMNf4umsgQjJkv7bdTE69eATqIQ';
const SHEET_NAME = 'responses';
const SHEET_NAME_VOL2 = 'responses_vol2';
const VOL2_DRIVE_FOLDER_NAME = 'JOYFIT24経堂_アンケートVOL2_参考画像';

/** アンケートのサービス順（index.html の services と同じ順） */
const SERVICE_ORDER = [
  'FWエリア',
  'マシンエリア',
  '有酸素エリア',
  'ストレッチエリア',
  'サウナ',
  'ミストサウナ',
  'スタジオ',
  'マシンピラティス(リフォーマー)',
  'エステ',
  'タンニング'
];

/**
 * 利用度オブジェクトから区分別のサービス名リストを返す
 */
function splitUsageByLevel_(dep) {
  const always = [];
  const sometimes = [];
  const never = [];
  SERVICE_ORDER.forEach(function (name) {
    const n = Number(dep[name]);
    if (n === 100) always.push(name);
    else if (n === 50) sometimes.push(name);
    else never.push(name);
  });
  return { always: always, sometimes: sometimes, never: never };
}

/** 1セルに複数サービス（英カンマ＋スペース区切り） */
function joinServicesForCell_(arr) {
  return arr.length ? arr.join(', ') : '（なし）';
}

/** 1行目の見出し配列（列順＝doPost の appendRow と一致） */
function getResponseHeaders_() {
  return [].concat(
    ['受信日時', '送信日時（ISO）', '必ず利用（100％）', 'たまに利用（50％）', '利用しない（0％）'],
    ['BEST5マシン・1位', 'BEST5マシン・2位', 'BEST5マシン・3位', 'BEST5マシン・4位', 'BEST5マシン・5位'],
    [
      'レッスン曜日',
      'スタジオ・お気に入りレッスン',
      'スタジオ・お気に入りインストラクター',
      'ピラティス・お気に入りレッスン',
      'ピラティス・お気に入りインストラクター',
      'ご希望のマシン',
      'ご希望の設備・環境',
      'その他ご意見'
    ]
  );
}

/**
 * 1行目が正しいヘッダーでなければ、上に1行挿入してヘッダーを書く
 */
function ensureResponseHeaderRow_(sheet) {
  const headers = getResponseHeaders_();
  const colCount = headers.length;

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, colCount).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }

  const a1 = String(sheet.getRange(1, 1).getDisplayValue() || '').trim();
  if (a1 === '受信日時') {
    sheet.setFrozenRows(1);
    return;
  }

  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, colCount).setValues([headers]);
  sheet.setFrozenRows(1);
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const data = JSON.parse(raw);

    if (data.version === 'vol2') {
      return doPostVol2_(data);
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet not found: ' + SHEET_NAME);

    ensureResponseHeaderRow_(sheet);

    const dep = data.dependency || {};
    const parts = splitUsageByLevel_(dep);

    const top5 = data.top5 || [];
    const top5cols = [0, 1, 2, 3, 4].map(function (i) {
      return top5[i] ? String(top5[i]) : '';
    });

    const row = [].concat(
      [
        new Date(),
        data.submittedAt || '',
        joinServicesForCell_(parts.always),
        joinServicesForCell_(parts.sometimes),
        joinServicesForCell_(parts.never)
      ],
      top5cols,
      [
        (data.lessonWeekdays || []).join(', '),
        (data.studioFavoriteLessons || []).join('\n'),
        (data.studioFavoriteInstructors || []).join(', '),
        (data.pilatesFavoriteLessons || []).join('\n'),
        (data.pilatesFavoriteInstructors || []).join(', '),
        data.wantedMachines || '',
        data.wantedOptions || '',
        data.freeComment || ''
      ]
    );

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 1行目を見出しで上書きする（手動メンテ用）。
 * 注意: 1行目にあった内容は消えます。データは2行目以降に置いてください。
 */
function setupResponseHeaders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + SHEET_NAME);

  const headers = getResponseHeaders_();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

/** VOL.2 見出し（設置してほしいマシン聞き取り） */
function getVol2ResponseHeaders_() {
  return [
    '受信日時',
    '送信日時（ISO）',
    '設置してほしいマシン',
    '参考URL',
    '参考画像リンク',
    '補足・理由'
  ];
}

function getOrCreateSheet_(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  var a1 = String(sheet.getRange(1, 1).getDisplayValue() || '').trim();
  if (a1 !== headers[0]) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.setFrozenRows(1);
  return sheet;
}

function getOrCreateVol2ImageFolder_() {
  var folders = DriveApp.getFoldersByName(VOL2_DRIVE_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(VOL2_DRIVE_FOLDER_NAME);
}

function saveVol2Images_(images, submittedAt) {
  if (!images || !images.length) return '';
  var folder = getOrCreateVol2ImageFolder_();
  var stamp = Utilities.formatDate(new Date(submittedAt || Date.now()), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  var urls = [];
  images.forEach(function (img, i) {
    if (!img || !img.dataBase64) return;
    var mime = img.mimeType || 'image/jpeg';
    var ext = mime.indexOf('png') >= 0 ? 'png' : (mime.indexOf('webp') >= 0 ? 'webp' : 'jpg');
    var name = String(img.name || ('image_' + (i + 1) + '.' + ext)).replace(/[\\/:*?"<>|]/g, '_');
    var blob = Utilities.newBlob(Utilities.base64Decode(img.dataBase64), mime, stamp + '_' + name);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urls.push(file.getUrl());
  });
  return urls.join('\n');
}

function doPostVol2_(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var headers = getVol2ResponseHeaders_();
  var sheet = getOrCreateSheet_(ss, SHEET_NAME_VOL2, headers);

  var imageLinks = saveVol2Images_(data.images || [], data.submittedAt);
  var urls = (data.referenceUrls || []).map(function (u) {
    return String(u || '').trim();
  }).filter(Boolean);

  sheet.appendRow([
    new Date(),
    data.submittedAt || '',
    String(data.machineRequest || '').trim(),
    urls.join('\n'),
    imageLinks,
    String(data.note || '').trim()
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** VOL.2 シートの1行目を見出しで上書き（手動メンテ用） */
function setupVol2ResponseHeaders() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var headers = getVol2ResponseHeaders_();
  var sheet = getOrCreateSheet_(ss, SHEET_NAME_VOL2, headers);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

/**
 * 初回だけエディタから実行してください（参考画像を Drive に保存する権限の許可）。
 * 実行 → 「権限を確認」→ 許可 → ログに OK と出れば完了。
 */
function authorizeVol2Drive() {
  var folder = getOrCreateVol2ImageFolder_();
  Logger.log('OK: Drive folder = ' + folder.getName() + ' (' + folder.getId() + ')');
}

/**
 * GET でダッシュボード HTML を返す。
 * デプロイ: 「ウェブアプリ」で新バージョン作成。アクセス権は共有範囲に合わせて設定。
 * （アンケート送信用の doPost と同じプロジェクト・別 URL になる場合は、同一デプロイに doGet / doPost 両方を含める）
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('JOYFIT24 経堂｜アンケート結果')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ダッシュボード集計用の列名（doPost のヘッダーと一致） */
var DASH_KEYS = {
  u100: '必ず利用（100％）',
  u50: 'たまに利用（50％）',
  u0: '利用しない（0％）',
  b1: 'BEST5マシン・1位',
  b2: 'BEST5マシン・2位',
  b3: 'BEST5マシン・3位',
  b4: 'BEST5マシン・4位',
  b5: 'BEST5マシン・5位',
  sl: 'スタジオ・お気に入りレッスン',
  si: 'スタジオ・お気に入りインストラクター',
  pl: 'ピラティス・お気に入りレッスン',
  pi: 'ピラティス・お気に入りインストラクター',
  wm: 'ご希望のマシン',
  wo: 'ご希望の設備・環境',
  wmLegacy: '追加して欲しいマシン',
  woLegacy: '追加して欲しい設備'
};

/** 列見出しリネーム前のシートでも集計できるよう、新見出し優先でセル値を取得 */
function cellOrLegacy_(row, primary, legacy) {
  var p = row[primary];
  if (p !== undefined && p !== null && String(p).trim() !== '') return String(p);
  if (!legacy) return '';
  var l = row[legacy];
  return l !== undefined && l !== null ? String(l) : '';
}

function getSheetRowsAsObjects_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    return { ok: false, error: 'シートが見つかりません: ' + SHEET_NAME };
  }
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { ok: true, rows: [] };
  }
  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0].map(function (h) {
    return String(h || '').trim();
  });
  const rows = [];
  for (var r = 1; r < values.length; r++) {
    const rowArr = values[r];
    if (rowArr.every(function (c) {
      return String(c || '').trim() === '';
    })) {
      continue;
    }
    const obj = {};
    for (var c = 0; c < headers.length; c++) {
      const key = headers[c] || '列' + (c + 1);
      obj[key] = rowArr[c] !== undefined && rowArr[c] !== null ? String(rowArr[c]) : '';
    }
    rows.push(obj);
  }
  return { ok: true, rows: rows };
}

function splitCommaServices_(cell) {
  const s = String(cell || '').trim();
  if (!s || s === '（なし）') return [];
  return s.split(/,\s*/).map(function (x) {
    return x.trim();
  }).filter(Boolean);
}

function bumpMap_(map, keys) {
  keys.forEach(function (k) {
    if (!k) return;
    map[k] = (map[k] || 0) + 1;
  });
}

function mapToSortedPairs_(map, maxLen) {
  const list = Object.keys(map).map(function (name) {
    return { name: name, count: map[name] };
  });
  list.sort(function (a, b) {
    return b.count - a.count || a.name.localeCompare(b.name, 'ja');
  });
  return maxLen ? list.slice(0, maxLen) : list;
}

function countLinesInColumn_(map, cell) {
  String(cell || '').split(/\n/).forEach(function (line) {
    const t = line.trim();
    if (t.length >= 2) {
      map[t] = (map[t] || 0) + 1;
    }
  });
}

function splitCommaOrJa_(cell) {
  const s = String(cell || '').trim();
  if (!s) return [];
  return s.split(/[,、]\s*/).map(function (x) {
    return x.trim();
  }).filter(function (x) {
    return x.length >= 1;
  });
}

/**
 * ダッシュボード用: 全回答を集計したサマリー（件数・ランキング）。
 */
function getDashboardData() {
  const got = getSheetRowsAsObjects_();
  if (!got.ok) {
    return { ok: false, error: got.error || '読み込みエラー' };
  }
  const rows = got.rows || [];
  const K = DASH_KEYS;

  const u100 = {};
  const u50 = {};
  const u0 = {};
  const best5 = {};
  const studioLessons = {};
  const pilatesLessons = {};
  const studioInst = {};
  const pilatesInst = {};
  const wantedMach = {};
  const wantedOpt = {};

  rows.forEach(function (row) {
    bumpMap_(u100, splitCommaServices_(row[K.u100]));
    bumpMap_(u50, splitCommaServices_(row[K.u50]));
    bumpMap_(u0, splitCommaServices_(row[K.u0]));

    [K.b1, K.b2, K.b3, K.b4, K.b5].forEach(function (bk) {
      const m = String(row[bk] || '').trim();
      if (m) {
        best5[m] = (best5[m] || 0) + 1;
      }
    });

    countLinesInColumn_(studioLessons, row[K.sl]);
    countLinesInColumn_(pilatesLessons, row[K.pl]);
    bumpMap_(studioInst, splitCommaOrJa_(row[K.si]));
    bumpMap_(pilatesInst, splitCommaOrJa_(row[K.pi]));

    countLinesInColumn_(wantedMach, cellOrLegacy_(row, K.wm, K.wmLegacy));
    countLinesInColumn_(wantedOpt, cellOrLegacy_(row, K.wo, K.woLegacy));
  });

  const maxRank = 20;
  return {
    ok: true,
    sheetUrl: spreadsheetUrl_(),
    responseCount: rows.length,
    usage100: mapToSortedPairs_(u100, maxRank),
    usage50: mapToSortedPairs_(u50, maxRank),
    usage0: mapToSortedPairs_(u0, maxRank),
    best5Machines: mapToSortedPairs_(best5, maxRank),
    studioLessons: mapToSortedPairs_(studioLessons, maxRank),
    pilatesLessons: mapToSortedPairs_(pilatesLessons, maxRank),
    studioInstructors: mapToSortedPairs_(studioInst, maxRank),
    pilatesInstructors: mapToSortedPairs_(pilatesInst, maxRank),
    wantedMachines: mapToSortedPairs_(wantedMach, maxRank),
    wantedOptions: mapToSortedPairs_(wantedOpt, maxRank)
  };
}

function spreadsheetUrl_() {
  return 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/edit';
}
