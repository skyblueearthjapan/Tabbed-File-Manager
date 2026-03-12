/**
 * Tabbed File Manager Dashboard
 * スプレッドシートのtabsシートからタブ情報を読み取り、
 * iframe切り替え式ダッシュボードを表示する
 */

// スプレッドシートIDをスクリプトプロパティから取得、または直接指定
function getSpreadsheetId_() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('SPREADSHEET_ID') || '1Rz_LpVMogNFXZpbOrj3Uv1EPByrBRehNrgDnoR9EjEo';
}

/**
 * Webアプリのエントリーポイント
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  template.tabsData = JSON.stringify(getTabsData_());
  return template.evaluate()
    .setTitle('Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * tabsシートからタブ情報を取得
 */
function getTabsData_() {
  const ssId = getSpreadsheetId_();
  if (!ssId) return [];

  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName('tabs');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // ヘッダーのみ

  const tabs = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1]) continue; // labelが空ならスキップ

    tabs.push({
      order:     Number(row[0]) || i,
      label:     String(row[1]),
      url:       String(row[2]),
      active:    row[3] === true || String(row[3]).toUpperCase() === 'TRUE',
      color:     String(row[4] || ''),
      icon:      String(row[5] || ''),
      colorCode: String(row[6] || '#607D8B'),
      iconCode:  String(row[7] || 'tab')
    });
  }

  // order順にソート
  tabs.sort((a, b) => a.order - b.order);
  return tabs;
}

/**
 * 初期セットアップ: スプレッドシートにシートと参考データを作成
 * メニューまたは手動で1回だけ実行
 */
function setupSpreadsheet() {
  const ssId = '1Rz_LpVMogNFXZpbOrj3Uv1EPByrBRehNrgDnoR9EjEo';
  const ss = SpreadsheetApp.openById(ssId);

  // スクリプトプロパティに保存
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ssId);

  setupTabsSheet_(ss);
  setupColorReference_(ss);
  setupIconReference_(ss);

  SpreadsheetApp.getUi().alert(
    'セットアップ完了\n\nスプレッドシートID: ' + ssId +
    '\n\ntabsシートにタブ情報を入力してください。'
  );
}

/**
 * tabsシートを作成
 */
function setupTabsSheet_(ss) {
  let sheet = ss.getSheetByName('tabs');
  if (!sheet) {
    sheet = ss.insertSheet('tabs');
  } else {
    sheet.clear();
  }

  // ヘッダー
  const headers = ['order', 'label', 'url', 'active', 'color', 'icon', 'color_code（自動）', 'icon_code（自動）'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#37474F')
    .setFontColor('#FFFFFF');

  // サンプルデータ
  const samples = [
    [1, 'サンプルA', 'https://example.com', true,  '青',   'ホーム', '', ''],
    [2, 'サンプルB', 'https://example.com', false, '緑',   'グラフ', '', ''],
    [3, 'サンプルC', 'https://example.com', false, 'オレンジ', 'メール', '', '']
  ];
  sheet.getRange(2, 1, samples.length, samples[0].length).setValues(samples);

  // VLOOKUP数式をG列・H列に設定（100行分）
  for (let r = 2; r <= 100; r++) {
    sheet.getRange(r, 7).setFormula(
      '=IFERROR(VLOOKUP(E' + r + ',色参考!A:B,2,FALSE),"")'
    );
    sheet.getRange(r, 8).setFormula(
      '=IFERROR(VLOOKUP(F' + r + ',アイコン参考!A:B,2,FALSE),"")'
    );
  }

  // 列幅調整
  sheet.setColumnWidth(1, 60);   // order
  sheet.setColumnWidth(2, 140);  // label
  sheet.setColumnWidth(3, 350);  // url
  sheet.setColumnWidth(4, 70);   // active
  sheet.setColumnWidth(5, 90);   // color
  sheet.setColumnWidth(6, 90);   // icon
  sheet.setColumnWidth(7, 120);  // color_code
  sheet.setColumnWidth(8, 120);  // icon_code

  // G,H列の背景をグレーにして自動列であることを示す
  sheet.getRange('G:H').setBackground('#F5F5F5').setFontColor('#888888');
}

/**
 * 色参考シートを作成（30色以上）
 */
function setupColorReference_(ss) {
  let sheet = ss.getSheetByName('色参考');
  if (!sheet) {
    sheet = ss.insertSheet('色参考');
  } else {
    sheet.clear();
  }

  const headers = ['色名', 'カラーコード', 'プレビュー'];
  sheet.getRange(1, 1, 1, 3).setValues([headers]);
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#37474F')
    .setFontColor('#FFFFFF');

  const colors = [
    // 赤系
    ['赤',         '#E53935'],
    ['ダークレッド', '#B71C1C'],
    ['ライトレッド', '#EF9A9A'],
    ['ローズ',      '#E91E63'],
    ['ピンク',      '#F48FB1'],
    ['ライトピンク', '#F8BBD0'],
    // オレンジ・黄系
    ['オレンジ',    '#FB8C00'],
    ['ダークオレンジ','#E65100'],
    ['ライトオレンジ','#FFCC80'],
    ['アンバー',    '#FFB300'],
    ['黄',         '#FDD835'],
    ['ライトイエロー','#FFF9C4'],
    // 緑系
    ['緑',         '#43A047'],
    ['ダークグリーン','#1B5E20'],
    ['ライトグリーン','#A5D6A7'],
    ['ティール',    '#00897B'],
    ['ミント',      '#80CBC4'],
    ['ライム',      '#C0CA33'],
    // 青系
    ['青',         '#1E88E5'],
    ['ダークブルー', '#0D47A1'],
    ['ライトブルー', '#90CAF9'],
    ['スカイブルー', '#03A9F4'],
    ['ネイビー',    '#1A237E'],
    ['シアン',      '#00BCD4'],
    // 紫系
    ['紫',         '#8E24AA'],
    ['ダークパープル','#4A148C'],
    ['ライトパープル','#CE93D8'],
    ['インディゴ',   '#3F51B5'],
    ['ラベンダー',   '#B39DDB'],
    // 茶・グレー系
    ['ブラウン',    '#6D4C41'],
    ['グレー',      '#757575'],
    ['ダークグレー', '#424242'],
    ['ライトグレー', '#BDBDBD'],
    ['ブルーグレー', '#607D8B'],
    // 特殊
    ['ゴールド',    '#FFD600'],
    ['シルバー',    '#B0BEC5'],
    ['コーラル',    '#FF7043'],
    ['サーモン',    '#FF8A65'],
    ['マルーン',    '#880E4F'],
    ['オリーブ',    '#827717'],
    ['カーキ',      '#F0E68C'],
    ['ターコイズ',   '#26C6DA'],
    ['ワイン',      '#AD1457'],
    ['チョコレート', '#4E342E'],
    ['エメラルド',   '#00C853'],
    ['サファイア',   '#2962FF'],
    ['ルビー',      '#D50000'],
    ['アクア',      '#00E5FF']
  ];

  sheet.getRange(2, 1, colors.length, 2).setValues(colors);

  // C列にプレビュー色を設定（背景色で表現）
  for (let i = 0; i < colors.length; i++) {
    const cell = sheet.getRange(i + 2, 3);
    cell.setValue('████████');
    cell.setFontColor(colors[i][1]);
    cell.setBackground('#FFFFFF');
  }

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
}

/**
 * アイコン参考シートを作成（50種以上、Material Icons）
 */
function setupIconReference_(ss) {
  let sheet = ss.getSheetByName('アイコン参考');
  if (!sheet) {
    sheet = ss.insertSheet('アイコン参考');
  } else {
    sheet.clear();
  }

  const headers = ['アイコン名', 'Material Iconコード', 'カテゴリ'];
  sheet.getRange(1, 1, 1, 3).setValues([headers]);
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground('#37474F')
    .setFontColor('#FFFFFF');

  const icons = [
    // ナビゲーション・一般
    ['ホーム',       'home',              'ナビゲーション'],
    ['ダッシュボード', 'dashboard',        'ナビゲーション'],
    ['メニュー',     'menu',              'ナビゲーション'],
    ['設定',        'settings',           'ナビゲーション'],
    ['検索',        'search',             'ナビゲーション'],
    ['お気に入り',   'favorite',           'ナビゲーション'],
    ['スター',      'star',               'ナビゲーション'],
    ['ブックマーク',  'bookmark',          'ナビゲーション'],
    // 編集・作成
    ['ペン',        'edit',               '編集'],
    ['作成',        'create',             '編集'],
    ['追加',        'add_circle',         '編集'],
    ['削除',        'delete',             '編集'],
    ['コピー',      'content_copy',       '編集'],
    ['保存',        'save',               '編集'],
    // ファイル・ドキュメント
    ['ファイル',     'description',        'ファイル'],
    ['フォルダ',     'folder',             'ファイル'],
    ['添付',        'attach_file',        'ファイル'],
    ['ダウンロード',  'download',          'ファイル'],
    ['アップロード',  'upload',            'ファイル'],
    ['クラウド',     'cloud',              'ファイル'],
    ['ドキュメント',  'article',           'ファイル'],
    // コミュニケーション
    ['メール',       'mail',              'コミュニケーション'],
    ['チャット',     'chat',              'コミュニケーション'],
    ['通知',        'notifications',      'コミュニケーション'],
    ['電話',        'phone',             'コミュニケーション'],
    ['ビデオ',      'videocam',           'コミュニケーション'],
    ['共有',        'share',             'コミュニケーション'],
    ['フォーラム',   'forum',             'コミュニケーション'],
    // ビジネス・データ
    ['グラフ',      'bar_chart',          'ビジネス'],
    ['円グラフ',     'pie_chart',         'ビジネス'],
    ['折れ線グラフ',  'show_chart',        'ビジネス'],
    ['トレンド',     'trending_up',       'ビジネス'],
    ['カレンダー',   'calendar_today',     'ビジネス'],
    ['タスク',      'task_alt',           'ビジネス'],
    ['チェック',     'check_circle',      'ビジネス'],
    ['レポート',     'assessment',        'ビジネス'],
    ['お金',        'payments',          'ビジネス'],
    ['ショップ',     'storefront',        'ビジネス'],
    ['在庫',        'inventory_2',       'ビジネス'],
    // 人・組織
    ['ユーザー',     'person',            '人'],
    ['グループ',     'group',             '人'],
    ['チーム',       'groups',            '人'],
    ['管理者',       'admin_panel_settings','人'],
    ['アカウント',   'account_circle',     '人'],
    // テクノロジー
    ['コード',       'code',              'テクノロジー'],
    ['ターミナル',    'terminal',          'テクノロジー'],
    ['データベース',  'storage',           'テクノロジー'],
    ['API',         'api',               'テクノロジー'],
    ['セキュリティ',  'security',          'テクノロジー'],
    ['バグ',         'bug_report',        'テクノロジー'],
    ['速度',         'speed',             'テクノロジー'],
    // 場所・施設
    ['地図',         'map',               '場所'],
    ['ビル',         'business',          '場所'],
    ['会議室',       'meeting_room',       '場所'],
    ['学校',         'school',             '場所'],
    ['病院',         'local_hospital',     '場所'],
    // その他
    ['電球',         'lightbulb',          'その他'],
    ['ヘルプ',       'help',               'その他'],
    ['情報',         'info',               'その他'],
    ['警告',         'warning',            'その他'],
    ['時計',         'schedule',           'その他'],
    ['リンク',       'link',               'その他'],
    ['写真',         'photo_camera',       'その他'],
    ['音楽',         'music_note',         'その他'],
    ['印刷',         'print',              'その他'],
    ['QR',           'qr_code',           'その他'],
    ['ロケット',      'rocket_launch',     'その他'],
    ['ツール',        'build',             'その他'],
    ['パレット',      'palette',           'その他'],
    ['リスト',        'list',              'その他'],
    ['テーブル',      'table_chart',       'その他'],
    ['Wi-Fi',        'wifi',              'その他'],
    ['地球',         'public',             'その他'],
    ['鍵',           'vpn_key',            'その他']
  ];

  sheet.getRange(2, 1, icons.length, 3).setValues(icons);

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 140);
}

/**
 * スプレッドシートを開いた時のカスタムメニュー
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗂️ タブ管理')
    .addItem('初期セットアップ', 'setupSpreadsheet')
    .addToUi();
}
