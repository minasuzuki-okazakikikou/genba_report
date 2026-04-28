// ========================================
// 現場日報 LINE貼り付け → 自動集計スクリプト
// ========================================

// ========================================
// 初期設定（最初に1回だけ実行）
// ========================================
function 初期設定() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 「日報入力」シート
  let inputSheet = ss.getSheetByName('日報入力');
  if (!inputSheet) {
    inputSheet = ss.insertSheet('日報入力');
  } else {
    inputSheet.clearContents();
    inputSheet.clearFormats();
  }
  inputSheet.getRange('A1').setValue('LINEのテキストをB1に貼り付けてください');
  inputSheet.getRange('A1').setFontWeight('bold');
  inputSheet.getRange('B1').setBackground('#fff9c4');
  inputSheet.getRange('A3').setValue('貼り付けたら メニュー「現場日報」→「日報を登録」を実行');
  inputSheet.getRange('A3').setFontColor('#e65100');
  inputSheet.setColumnWidth(1, 300);
  inputSheet.setColumnWidth(2, 500);

  // 「集計」シート
  let dataSheet = ss.getSheetByName('集計');
  if (!dataSheet) {
    dataSheet = ss.insertSheet('集計');
  } else {
    dataSheet.clearContents();
    dataSheet.clearFormats();
  }
  // ★ Excelファイルの実際の列順に合わせたヘッダー
  const headers = [
    '登録日時', '工事日', '現場名', '工事内容',
    '作業開始', '作業終了', '実働時間', '休憩',
    'クレーン会社', '車両', 'ガードマン',
    '自社人数', '協力会社人数', '自社スタッフ',  // ★ 修正：協力会社人数→自社スタッフの順
    '協力会社詳細',
    '合計人数', '備考'
  ];
  const headerRange = dataSheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1a56c4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  dataSheet.setFrozenRows(1);

  // 「現場集計」シート
  let summarySheet = ss.getSheetByName('現場集計');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('現場集計');
  } else {
    summarySheet.clearContents();
    summarySheet.clearFormats();
  }
  const summaryHeaders = [
    '現場名', '作業日数',
    '自社 延べ人数', '協力会社 延べ人数', '合計 延べ人数',
    '実働時間 合計', 'クレーン使用回数', '車両使用回数'
  ];
  const sHeaderRange = summarySheet.getRange(1, 1, 1, summaryHeaders.length);
  sHeaderRange.setValues([summaryHeaders]);
  sHeaderRange.setBackground('#0f6e56');
  sHeaderRange.setFontColor('#ffffff');
  sHeaderRange.setFontWeight('bold');
  summarySheet.setFrozenRows(1);
  summarySheet.getRange('A3').setValue('※ メニュー「現場日報」→「現場集計を更新」で最新データに更新できます');
  summarySheet.getRange('A3').setFontColor('#888888').setFontStyle('italic');

  SpreadsheetApp.getUi().alert('初期設定完了！\n「日報入力」シートのB1にLINEテキストを貼り付け、\nメニュー「現場日報」→「日報を登録」を実行してください。');
}

// ========================================
// 日報登録
// ========================================
function 日報を登録() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('日報入力');
  const dataSheet = ss.getSheetByName('集計');

  if (!inputSheet || !dataSheet) {
    SpreadsheetApp.getUi().alert('シートが見つかりません。先に「初期設定」を実行してください。');
    return;
  }

  const rawText = inputSheet.getRange('B1').getValue().toString().trim();
  if (!rawText) {
    SpreadsheetApp.getUi().alert('B1にLINEのテキストを貼り付けてください。');
    return;
  }
  if (!rawText.includes('【現場日報】')) {
    SpreadsheetApp.getUi().alert('現場日報のテキストではないようです。\n【現場日報】から始まるテキストを貼り付けてください。');
    return;
  }

  const data = parseNippo(rawText);
  const newRow = dataSheet.getLastRow() + 1;

  // ★ 修正：Excelの列順（自社人数→協力会社人数→自社スタッフ→協力会社詳細）に合わせて書き込み
  dataSheet.getRange(newRow, 1, 1, 17).setValues([[
    data.登録日時, data.工事日, data.現場名, data.工事内容,
    data.作業開始, data.作業終了, data.実働時間, data.休憩,
    data.クレーン会社, data.車両, data.ガードマン,
    data.自社人数, data.協力会社人数, data.自社スタッフ,  // ★ 修正
    data.協力会社詳細,
    data.合計人数, data.備考
  ]]);

  if (newRow % 2 === 0) {
    dataSheet.getRange(newRow, 1, 1, 17).setBackground('#f8f9ff');
  }

  inputSheet.getRange('B1').setValue('');

  // 登録後に現場集計も自動更新
  現場集計を更新();

  ss.setActiveSheet(dataSheet);
  SpreadsheetApp.getUi().alert('登録完了！\n現場：' + (data.現場名 || '（未入力）') + '\n工事日：' + (data.工事日 || '（未入力）') + '\n合計人数：' + data.合計人数 + '名');
}

// ========================================
// 現場集計を更新
// ========================================
function 現場集計を更新() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('集計');
  const summarySheet = ss.getSheetByName('現場集計');

  if (!dataSheet || !summarySheet) {
    SpreadsheetApp.getUi().alert('シートが見つかりません。先に「初期設定」を実行してください。');
    return;
  }

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('集計データがまだありません。');
    return;
  }

  // 集計シートの全データ取得（ヘッダー除く）
  const allData = dataSheet.getRange(2, 1, lastRow - 1, 17).getValues();

  // 現場ごとに集計
  const genbaMap = {};

  allData.forEach(function(row) {
    const genba       = row[2];   // C列：現場名
    const jidoStr     = row[6];   // G列：実働時間
    const crane       = row[8];   // I列：クレーン会社
    const sharyo      = row[9];   // J列：車両
    const jishaNin    = parseInt(row[11]) || 0;  // L列：自社人数
    const kyoryokuNin = parseInt(row[12]) || 0;  // ★ 修正：M列：協力会社人数（旧row[13]→row[12]）
    const totalNin    = parseInt(row[15]) || 0;  // P列：合計人数

    if (!genba) return;

    if (!genbaMap[genba]) {
      genbaMap[genba] = {
        作業日数: 0,
        自社延べ: 0,
        協力延べ: 0,
        合計延べ: 0,
        実働分合計: 0,
        クレーン回数: 0,
        車両回数: 0
      };
    }

    const g = genbaMap[genba];
    g.作業日数++;
    g.自社延べ    += jishaNin;
    g.協力延べ    += kyoryokuNin;
    g.合計延べ    += totalNin;
    g.実働分合計  += parseJidoToMin(jidoStr);
    if (crane && crane !== '') g.クレーン回数++;
    if (sharyo && sharyo !== '') g.車両回数++;
  });

  // 現場集計シートに書き出し
  summarySheet.clearContents();
  summarySheet.clearFormats();

  // ヘッダー
  const summaryHeaders = [
    '現場名', '作業日数',
    '自社 延べ人数', '協力会社 延べ人数', '合計 延べ人数',
    '実働時間 合計', 'クレーン使用回数', '車両使用回数'
  ];
  const sHeaderRange = summarySheet.getRange(1, 1, 1, summaryHeaders.length);
  sHeaderRange.setValues([summaryHeaders]);
  sHeaderRange.setBackground('#0f6e56');
  sHeaderRange.setFontColor('#ffffff');
  sHeaderRange.setFontWeight('bold');
  summarySheet.setFrozenRows(1);

  // データ行
  const genbaNames = Object.keys(genbaMap).sort();
  const rows = genbaNames.map(function(name) {
    const g = genbaMap[name];
    return [
      name,
      g.作業日数,
      g.自社延べ,
      g.協力延べ,
      g.合計延べ,
      minToJido(g.実働分合計),
      g.クレーン回数,
      g.車両回数
    ];
  });

  if (rows.length > 0) {
    summarySheet.getRange(2, 1, rows.length, 8).setValues(rows);

    // 偶数行に薄い色
    for (var i = 0; i < rows.length; i++) {
      if ((i + 2) % 2 === 0) {
        summarySheet.getRange(i + 2, 1, 1, 8).setBackground('#f0faf5');
      }
    }

    // 列幅
    summarySheet.setColumnWidth(1, 160);
    summarySheet.setColumnWidth(2, 80);
    summarySheet.setColumnWidth(3, 110);
    summarySheet.setColumnWidth(4, 120);
    summarySheet.setColumnWidth(5, 100);
    summarySheet.setColumnWidth(6, 110);
    summarySheet.setColumnWidth(7, 110);
    summarySheet.setColumnWidth(8, 90);

    // 更新日時
    const lastRowNum = rows.length + 3;
    summarySheet.getRange(lastRowNum, 1).setValue('最終更新：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
    summarySheet.getRange(lastRowNum, 1).setFontColor('#aaaaaa').setFontStyle('italic');
  }

  ss.setActiveSheet(summarySheet);
}

// ========================================
// ユーティリティ：実働時間 → 分に変換
// ========================================
function parseJidoToMin(str) {
  if (!str) return 0;
  str = str.toString();
  var h = 0, m = 0;
  var hMatch = str.match(/(\d+)時間/);
  var mMatch = str.match(/(\d+)分/);
  if (hMatch) h = parseInt(hMatch[1]);
  if (mMatch) m = parseInt(mMatch[1]);
  return h * 60 + m;
}

// ========================================
// ユーティリティ：分 → 時間表示に変換
// ========================================
function minToJido(min) {
  if (!min || min === 0) return '0時間';
  var h = Math.floor(min / 60);
  var m = min % 60;
  return m > 0 ? h + '時間' + m + '分' : h + '時間';
}

// ========================================
// テキスト解析関数
// ========================================
function parseNippo(text) {
  const lines = text.split('\n').map(function(l) { return l.trim(); }).filter(function(l) { return l !== ''; });

  const result = {
    登録日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
    工事日: '', 現場名: '', 工事内容: '',
    作業開始: '', 作業終了: '', 実働時間: '', 休憩: '',
    クレーン会社: '', 車両: '', ガードマン: '',
    自社人数: 0, 自社スタッフ: '',
    協力会社人数: 0, 協力会社詳細: '',
    合計人数: 0, 備考: ''
  };

  for (var i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.indexOf('工事日') === 0) {
      result.工事日 = line.replace(/^工事日[：:]/, '').trim();
    } else if (line.indexOf('現場名') === 0) {
      result.現場名 = line.replace(/^現場名[：:]/, '').trim();
    } else if (line.indexOf('工事内容') === 0) {
      result.工事内容 = line.replace(/^工事内容[：:]/, '').trim();
    } else if (line.indexOf('作業時間') === 0) {
      const timeStr = line.replace(/^作業時間[：:]/, '').trim();
      const timeMatch = timeStr.match(/(\d{1,2}:\d{2})\s*[〜~]\s*(\d{1,2}:\d{2})/);
      if (timeMatch) {
        result.作業開始 = timeMatch[1];
        result.作業終了 = timeMatch[2];
      }
      const jidoMatch = timeStr.match(/実働\s*([0-9時間分]+)/);
      if (jidoMatch) result.実働時間 = jidoMatch[1];
    } else if (line.indexOf('休憩') === 0) {
      result.休憩 = line.replace(/^休憩[：:]/, '').trim();
    } else if (line.indexOf('クレーン会社') === 0) {
      result.クレーン会社 = line.replace(/^クレーン会社[：:]/, '').trim();
    } else if (line.indexOf('車両') === 0) {
      result.車両 = line.replace(/^車両[：:]/, '').trim();
    } else if (line.indexOf('ガードマン') === 0) {
      result.ガードマン = line.replace(/^ガードマン[：:]/, '').trim();
    } else if (line.indexOf('出面') === 0) {
      const jishaMatch = line.match(/自社\s*(\d+)名/);
      const kyoryokuMatch = line.match(/協力会社\s*(\d+)名/);
      const totalMatch = line.match(/合計\s*(\d+)名/);
      if (jishaMatch) result.自社人数 = parseInt(jishaMatch[1]);
      if (kyoryokuMatch) result.協力会社人数 = parseInt(kyoryokuMatch[1]);
      if (totalMatch) result.合計人数 = parseInt(totalMatch[1]);
    } else if (line.indexOf('自社：') === 0 || line.indexOf('自社:') === 0) {
      result.自社スタッフ = line.replace(/^自社[：:]/, '').trim();
    } else if (line.indexOf('協力会社：') === 0 || line.indexOf('協力会社:') === 0) {
      result.協力会社詳細 = line.replace(/^協力会社[：:]/, '').trim();
    } else if (line.indexOf('備考') === 0) {
      result.備考 = line.replace(/^備考[：:]/, '').trim();
    }
  }
  return result;
}

// ========================================
// 協力会社一覧を更新
// ========================================
function 協力会社一覧を更新() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('集計');

  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('「集計」シートが見つかりません。');
    return;
  }

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('集計データがまだありません。');
    return;
  }

  // 「協力会社一覧」シートを取得or作成
  let listSheet = ss.getSheetByName('協力会社一覧');
  if (!listSheet) {
    listSheet = ss.insertSheet('協力会社一覧');
  } else {
    listSheet.clearContents();
    listSheet.clearFormats();
  }

  // ヘッダー
  const headers = ['工事日', '現場名', '協力会社人数', '協力会社詳細'];
  const headerRange = listSheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#7b4fa6');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  listSheet.setFrozenRows(1);

  // 集計シートから必要列を抽出
  // B列:工事日(1), C列:現場名(2), M列:協力会社人数(12), O列:協力会社詳細(14)
  const allData = dataSheet.getRange(2, 1, lastRow - 1, 17).getValues();
  const rows = [];

  allData.forEach(function(row) {
    const kojibi    = row[1];   // B列：工事日
    const genba     = row[2];   // C列：現場名
    const kyoNin    = row[12];  // M列：協力会社人数
    const kyoDetail = row[14];  // O列：協力会社詳細

    if (!genba) return;
    rows.push([kojibi, genba, kyoNin, kyoDetail]);
  });

  if (rows.length > 0) {
    listSheet.getRange(2, 1, rows.length, 4).setValues(rows);

    // 交互に薄い色
    for (var i = 0; i < rows.length; i++) {
      if ((i + 2) % 2 === 0) {
        listSheet.getRange(i + 2, 1, 1, 4).setBackground('#f5f0fa');
      }
    }

    // 列幅
    listSheet.setColumnWidth(1, 120);
    listSheet.setColumnWidth(2, 160);
    listSheet.setColumnWidth(3, 100);
    listSheet.setColumnWidth(4, 300);
  }

  // 更新日時
  const lastRowNum = rows.length + 3;
  listSheet.getRange(lastRowNum, 1).setValue('最終更新：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
  listSheet.getRange(lastRowNum, 1).setFontColor('#aaaaaa').setFontStyle('italic');

  // 続けて協力会社集計も更新
  協力会社集計を更新();

  ss.setActiveSheet(listSheet);
  SpreadsheetApp.getUi().alert('協力会社一覧・集計を更新しました！');
}

// ========================================
// 協力会社集計を更新（会社別×現場別）
// ========================================
function 協力会社集計を更新() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('集計');

  if (!dataSheet) return;

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;

  const allData = dataSheet.getRange(2, 1, lastRow - 1, 17).getValues();

  // 会社名リストと現場名リストを収集
  // 協力会社詳細は「・」や「、」区切りで複数会社が入る場合も想定
  const companySet = {};  // { 会社名: { 現場名: 人数合計 } }
  const genbaSet   = {};  // 現場名の出現順管理

  allData.forEach(function(row) {
    const genba      = row[2];   // C列：現場名
    const kyoNin     = parseInt(row[12]) || 0;  // M列：協力会社人数
    const kyoDetail  = row[14] ? row[14].toString().trim() : '';  // O列：協力会社詳細

    if (!genba || !kyoDetail) return;

    genbaSet[genba] = true;

    // 会社名を区切り文字で分割（・ / 、 / , / 　/ スペース）
    const companies = kyoDetail.split(/[・、,，\s　]+/).map(function(s) { return s.trim(); }).filter(function(s) { return s !== ''; });

    // 同じ現場・同じ日の協力会社人数を会社数で按分（端数は最初の会社に加算）
    const perCompany = companies.length > 0 ? Math.floor(kyoNin / companies.length) : 0;
    const remainder  = companies.length > 0 ? kyoNin % companies.length : 0;

    companies.forEach(function(company, idx) {
      if (!companySet[company]) companySet[company] = {};
      if (!companySet[company][genba]) companySet[company][genba] = 0;
      companySet[company][genba] += perCompany + (idx === 0 ? remainder : 0);
    });
  });

  // 「協力会社集計」シートを取得or作成
  let kyoSheet = ss.getSheetByName('協力会社集計');
  if (!kyoSheet) {
    kyoSheet = ss.insertSheet('協力会社集計');
  } else {
    kyoSheet.clearContents();
    kyoSheet.clearFormats();
  }

  const genbaNames   = Object.keys(genbaSet).sort();
  const companyNames = Object.keys(companySet).sort();

  // ★ 縦軸＝現場名、横軸＝会社名 に変更
  // ヘッダー行：「現場名」＋会社名一覧＋「合計」
  const headers = ['現場名'].concat(companyNames).concat(['合計']);
  const headerRange = kyoSheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#c0392b');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  kyoSheet.setFrozenRows(1);
  kyoSheet.setFrozenColumns(1);

  // データ行：現場別×会社別
  const rows = genbaNames.map(function(genba) {
    const counts = companyNames.map(function(c) { return (companySet[c][genba] || 0); });
    const total  = counts.reduce(function(a, b) { return a + b; }, 0);
    return [genba].concat(counts).concat([total]);
  });

  // 合計行
  const totalRow = ['合計'].concat(companyNames.map(function(c) {
    return genbaNames.reduce(function(sum, g) { return sum + (companySet[c][g] || 0); }, 0);
  })).concat([
    companyNames.reduce(function(sum, c) {
      return sum + genbaNames.reduce(function(s2, g) { return s2 + (companySet[c][g] || 0); }, 0);
    }, 0)
  ]);

  const allRows = rows.concat([totalRow]);

  if (allRows.length > 0) {
    kyoSheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);

    // 合計行に色
    const totalRowIdx = allRows.length + 1;
    kyoSheet.getRange(totalRowIdx, 1, 1, headers.length).setBackground('#fde8e8').setFontWeight('bold');

    // 合計列（最終列）に色
    kyoSheet.getRange(2, headers.length, rows.length, 1).setBackground('#fff3cd').setFontWeight('bold');

    // 交互色（データ行のみ）
    for (var i = 0; i < rows.length; i++) {
      if ((i + 2) % 2 === 0) {
        kyoSheet.getRange(i + 2, 1, 1, headers.length).setBackground('#fdf5f5');
      }
    }

    // 列幅
    kyoSheet.setColumnWidth(1, 160);  // 現場名列は少し広め
    for (var j = 2; j <= headers.length; j++) {
      kyoSheet.setColumnWidth(j, 90);
    }
  }

  // 更新日時
  const lastRowNum = allRows.length + 3;
  kyoSheet.getRange(lastRowNum, 1).setValue('最終更新：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
  kyoSheet.getRange(lastRowNum, 1).setFontColor('#aaaaaa').setFontStyle('italic');
}

// ========================================
// メニュー追加
// ========================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('現場日報')
    .addItem('日報を登録', '日報を登録')
    .addItem('現場集計を更新', '現場集計を更新')
    .addItem('協力会社一覧・集計を更新', '協力会社一覧を更新')
    .addSeparator()
    .addItem('初期設定（初回のみ）', '初期設定')
    .addToUi();
}
