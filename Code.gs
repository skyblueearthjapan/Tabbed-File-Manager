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
  const title = getDashboardTitle_();
  const template = HtmlService.createTemplateFromFile('index');
  template.tabsData = JSON.stringify(getTabsData_());
  return template.evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 設定シートからダッシュボードタイトルを取得
 */
function getDashboardTitle_() {
  const ssId = getSpreadsheetId_();
  if (!ssId) return 'Dashboard';

  try {
    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName('設定');
    if (!sheet) return 'Dashboard';

    // B1セルにタイトルを格納（A1はラベル「タイトル」）
    const value = sheet.getRange('B1').getValue();
    return value ? String(value) : 'Dashboard';
  } catch (e) {
    return 'Dashboard';
  }
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

  setupSettingsSheet_(ss);
  setupTabsSheet_(ss);
  setupColorReference_(ss);
  setupIconReference_(ss);

  SpreadsheetApp.getUi().alert(
    'セットアップ完了\n\nスプレッドシートID: ' + ssId +
    '\n\ntabsシートにタブ情報を入力してください。'
  );
}

/**
 * 設定シートを作成（タイトルなど）
 */
function setupSettingsSheet_(ss) {
  let sheet = ss.getSheetByName('設定');
  if (!sheet) {
    sheet = ss.insertSheet('設定');
    sheet.getRange('A1').setValue('タイトル')
      .setFontWeight('bold')
      .setBackground('#37474F')
      .setFontColor('#FFFFFF');
    sheet.getRange('B1').setValue('Dashboard');
    sheet.getRange('C1').setValue('← B1セルを編集するとダッシュボードのタイトルが変わります')
      .setFontColor('#888888');
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 240);
    sheet.setColumnWidth(3, 400);
  }
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

  // C列にプレビュー色を設定（セル背景色で直感的に表現）
  for (let i = 0; i < colors.length; i++) {
    const cell = sheet.getRange(i + 2, 3);
    cell.setValue('');
    cell.setBackground(colors[i][1]);
  }

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
}

/**
 * アイコン参考シートを作成（250種以上、Material Icons）
 */
function setupIconReference_(ss) {
  let sheet = ss.getSheetByName('アイコン参考');
  if (!sheet) {
    sheet = ss.insertSheet('アイコン参考');
  } else {
    sheet.clear();
  }

  const headers = ['アイコン名', 'Material Iconコード', 'カテゴリ', 'プレビュー'];
  sheet.getRange(1, 1, 1, 4).setValues([headers]);
  sheet.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#37474F')
    .setFontColor('#FFFFFF');

  const icons = [
    // ナビゲーション・一般
    ['ホーム',          'home',                 'ナビゲーション'],
    ['ダッシュボード',    'dashboard',            'ナビゲーション'],
    ['メニュー',         'menu',                 'ナビゲーション'],
    ['設定',            'settings',             'ナビゲーション'],
    ['検索',            'search',               'ナビゲーション'],
    ['お気に入り',       'favorite',             'ナビゲーション'],
    ['スター',          'star',                 'ナビゲーション'],
    ['ブックマーク',      'bookmark',             'ナビゲーション'],
    ['戻る',            'arrow_back',           'ナビゲーション'],
    ['進む',            'arrow_forward',        'ナビゲーション'],
    ['上へ',            'arrow_upward',         'ナビゲーション'],
    ['下へ',            'arrow_downward',       'ナビゲーション'],
    ['更新',            'refresh',              'ナビゲーション'],
    ['閉じる',          'close',                 'ナビゲーション'],
    ['もっと見る',       'more_horiz',           'ナビゲーション'],
    ['展開',            'expand_more',          'ナビゲーション'],
    ['全画面',          'fullscreen',           'ナビゲーション'],
    ['ピン',            'push_pin',             'ナビゲーション'],
    // 編集・作成
    ['ペン',            'edit',                 '編集'],
    ['作成',            'create',               '編集'],
    ['追加',            'add_circle',           '編集'],
    ['削除',            'delete',               '編集'],
    ['コピー',          'content_copy',         '編集'],
    ['切り取り',         'content_cut',          '編集'],
    ['貼り付け',         'content_paste',        '編集'],
    ['保存',            'save',                 '編集'],
    ['元に戻す',         'undo',                 '編集'],
    ['やり直し',         'redo',                 '編集'],
    ['署名',            'draw',                 '編集'],
    ['フォーマット',      'format_paint',         '編集'],
    ['選択',            'check_box',            '編集'],
    ['フィルタ',         'filter_alt',           '編集'],
    ['ソート',          'sort',                 '編集'],
    // ファイル・ドキュメント
    ['ファイル',         'description',          'ファイル'],
    ['フォルダ',         'folder',               'ファイル'],
    ['フォルダ追加',      'create_new_folder',    'ファイル'],
    ['フォルダ共有',      'folder_shared',        'ファイル'],
    ['添付',            'attach_file',          'ファイル'],
    ['ダウンロード',      'download',             'ファイル'],
    ['アップロード',      'upload',               'ファイル'],
    ['クラウド',         'cloud',                'ファイル'],
    ['クラウド保存',      'cloud_upload',         'ファイル'],
    ['クラウドDL',       'cloud_download',       'ファイル'],
    ['ドキュメント',      'article',              'ファイル'],
    ['PDF',            'picture_as_pdf',       'ファイル'],
    ['スプレッドシート',   'table_view',           'ファイル'],
    ['アーカイブ',       'archive',              'ファイル'],
    ['履歴',            'history',              'ファイル'],
    ['ゴミ箱',          'delete_forever',       'ファイル'],
    ['バックアップ',      'backup',               'ファイル'],
    ['同期',            'sync',                 'ファイル'],
    // コミュニケーション
    ['メール',          'mail',                 'コミュニケーション'],
    ['受信箱',          'inbox',                'コミュニケーション'],
    ['送信',            'send',                 'コミュニケーション'],
    ['下書き',          'drafts',               'コミュニケーション'],
    ['チャット',         'chat',                 'コミュニケーション'],
    ['吹き出し',         'chat_bubble',          'コミュニケーション'],
    ['コメント',         'comment',              'コミュニケーション'],
    ['通知',            'notifications',        'コミュニケーション'],
    ['通知オフ',         'notifications_off',    'コミュニケーション'],
    ['電話',            'phone',                'コミュニケーション'],
    ['通話中',          'phone_in_talk',        'コミュニケーション'],
    ['ビデオ',          'videocam',             'コミュニケーション'],
    ['ビデオ通話',       'video_call',           'コミュニケーション'],
    ['マイク',          'mic',                  'コミュニケーション'],
    ['共有',            'share',                'コミュニケーション'],
    ['フォーラム',       'forum',                'コミュニケーション'],
    ['翻訳',            'translate',            'コミュニケーション'],
    ['連絡先',          'contacts',             'コミュニケーション'],
    // ビジネス・データ
    ['グラフ',          'bar_chart',            'ビジネス'],
    ['円グラフ',         'pie_chart',            'ビジネス'],
    ['折れ線グラフ',      'show_chart',           'ビジネス'],
    ['ドーナツグラフ',    'donut_large',          'ビジネス'],
    ['散布図',          'scatter_plot',         'ビジネス'],
    ['バブルチャート',    'bubble_chart',         'ビジネス'],
    ['トレンド',         'trending_up',          'ビジネス'],
    ['トレンド下降',      'trending_down',        'ビジネス'],
    ['分析',            'analytics',            'ビジネス'],
    ['インサイト',       'insights',             'ビジネス'],
    ['カレンダー',       'calendar_today',       'ビジネス'],
    ['予定',            'event',                'ビジネス'],
    ['タスク',          'task_alt',             'ビジネス'],
    ['チェック',         'check_circle',         'ビジネス'],
    ['チェックリスト',    'checklist',            'ビジネス'],
    ['レポート',         'assessment',           'ビジネス'],
    ['要約',            'summarize',            'ビジネス'],
    ['お金',            'payments',             'ビジネス'],
    ['現金',            'attach_money',         'ビジネス'],
    ['通貨円',          'currency_yen',         'ビジネス'],
    ['クレジットカード',  'credit_card',          'ビジネス'],
    ['レシート',         'receipt',              'ビジネス'],
    ['電卓',            'calculate',            'ビジネス'],
    ['ショップ',         'storefront',           'ビジネス'],
    ['カート',          'shopping_cart',        'ビジネス'],
    ['バッグ',          'shopping_bag',         'ビジネス'],
    ['在庫',            'inventory_2',          'ビジネス'],
    ['配送',            'local_shipping',       'ビジネス'],
    ['梱包',            'inventory',            'ビジネス'],
    ['契約書',          'gavel',                'ビジネス'],
    ['工場',            'factory',              'ビジネス'],
    ['倉庫',            'warehouse',            'ビジネス'],
    ['バーコード',       'barcode_reader',       'ビジネス'],
    // 人・組織
    ['ユーザー',         'person',               '人'],
    ['ユーザー追加',      'person_add',           '人'],
    ['グループ',         'group',                '人'],
    ['チーム',          'groups',               '人'],
    ['管理者',          'admin_panel_settings', '人'],
    ['アカウント',       'account_circle',       '人'],
    ['顔',              'face',                 '人'],
    ['指紋',            'fingerprint',          '人'],
    ['バッジ',          'badge',                '人'],
    ['ハンドシェイク',    'handshake',            '人'],
    ['サポート',         'support_agent',        '人'],
    ['エンジニア',       'engineering',          '人'],
    // テクノロジー
    ['コード',          'code',                 'テクノロジー'],
    ['開発',            'developer_mode',       'テクノロジー'],
    ['ターミナル',       'terminal',             'テクノロジー'],
    ['データベース',      'storage',              'テクノロジー'],
    ['API',            'api',                  'テクノロジー'],
    ['セキュリティ',      'security',             'テクノロジー'],
    ['シールド',         'shield',               'テクノロジー'],
    ['ロック',          'lock',                 'テクノロジー'],
    ['ロック解除',       'lock_open',            'テクノロジー'],
    ['バグ',            'bug_report',           'テクノロジー'],
    ['速度',            'speed',                'テクノロジー'],
    ['メモリ',          'memory',               'テクノロジー'],
    ['ルーター',         'router',               'テクノロジー'],
    ['デバイス',         'devices',              'テクノロジー'],
    ['PC',             'computer',             'テクノロジー'],
    ['ノートPC',        'laptop',               'テクノロジー'],
    ['タブレット',       'tablet',               'テクノロジー'],
    ['スマホ',          'smartphone',           'テクノロジー'],
    ['マウス',          'mouse',                'テクノロジー'],
    ['キーボード',       'keyboard',             'テクノロジー'],
    ['USB',            'usb',                  'テクノロジー'],
    ['バッテリー',       'battery_full',         'テクノロジー'],
    ['電源',            'power',                'テクノロジー'],
    ['スマートホーム',    'smart_toy',            'テクノロジー'],
    // 場所・施設
    ['地図',            'map',                  '場所'],
    ['マップピン',       'place',                '場所'],
    ['位置情報',         'my_location',          '場所'],
    ['ルート',          'directions',           '場所'],
    ['ビル',            'business',             '場所'],
    ['会議室',          'meeting_room',         '場所'],
    ['学校',            'school',               '場所'],
    ['病院',            'local_hospital',       '場所'],
    ['自宅',            'house',                '場所'],
    ['アパート',         'apartment',            '場所'],
    ['店舗',            'store',                '場所'],
    ['レストラン',       'restaurant',           '場所'],
    ['カフェ',          'local_cafe',           '場所'],
    ['ガソリンスタンド',  'local_gas_station',    '場所'],
    ['駐車場',          'local_parking',        '場所'],
    ['空港',            'flight',               '場所'],
    ['ホテル',          'hotel',                '場所'],
    ['公園',            'park',                 '場所'],
    ['図書館',          'local_library',        '場所'],
    ['銀行',            'account_balance',      '場所'],
    ['郵便局',          'local_post_office',    '場所'],
    // 交通
    ['車',              'directions_car',       '交通'],
    ['タクシー',         'local_taxi',           '交通'],
    ['バス',            'directions_bus',       '交通'],
    ['電車',            'train',                '交通'],
    ['新幹線',          'directions_railway',   '交通'],
    ['自転車',          'directions_bike',      '交通'],
    ['バイク',          'two_wheeler',          '交通'],
    ['徒歩',            'directions_walk',      '交通'],
    ['船',              'directions_boat',      '交通'],
    ['信号',            'traffic',              '交通'],
    // 時間・スケジュール
    ['時計',            'schedule',             '時間'],
    ['アラーム',         'alarm',                '時間'],
    ['タイマー',         'timer',                '時間'],
    ['ストップウォッチ',  'timelapse',            '時間'],
    ['砂時計',          'hourglass_empty',      '時間'],
    ['日時',            'date_range',           '時間'],
    ['更新時間',         'update',               '時間'],
    // 天気・自然
    ['晴れ',            'wb_sunny',             '天気'],
    ['曇り',            'cloud_queue',          '天気'],
    ['雨',              'umbrella',             '天気'],
    ['雪',              'ac_unit',              '天気'],
    ['雷',              'flash_on',             '天気'],
    ['水',              'water_drop',           '天気'],
    ['風',              'air',                  '天気'],
    ['月',              'nights_stay',          '天気'],
    ['森',              'forest',               '天気'],
    ['花',              'local_florist',        '天気'],
    ['エコ',            'eco',                  '天気'],
    ['地球',            'public',               '天気'],
    // エンタメ・ホビー
    ['音楽',            'music_note',           'エンタメ'],
    ['ライブラリ',       'library_music',        'エンタメ'],
    ['再生',            'play_arrow',           'エンタメ'],
    ['一時停止',         'pause',                'エンタメ'],
    ['停止',            'stop',                 'エンタメ'],
    ['早送り',          'fast_forward',         'エンタメ'],
    ['映画',            'movie',                'エンタメ'],
    ['TV',             'tv',                   'エンタメ'],
    ['カメラ',          'photo_camera',         'エンタメ'],
    ['写真',            'photo',                'エンタメ'],
    ['ギャラリー',       'collections',          'エンタメ'],
    ['ゲーム',          'sports_esports',       'エンタメ'],
    ['スポーツ',         'sports_soccer',        'エンタメ'],
    ['本',              'menu_book',            'エンタメ'],
    ['パレット',         'palette',              'エンタメ'],
    ['ブラシ',          'brush',                'エンタメ'],
    // 食べ物・生活
    ['食事',            'restaurant_menu',      '生活'],
    ['ファストフード',    'fastfood',             '生活'],
    ['ピザ',            'local_pizza',          '生活'],
    ['ドリンク',         'local_drink',          '生活'],
    ['バー',            'local_bar',            '生活'],
    ['アイス',          'icecream',             '生活'],
    ['ケーキ',          'cake',                 '生活'],
    ['洗濯',            'local_laundry_service', '生活'],
    ['ベッド',          'bed',                  '生活'],
    ['風呂',            'bathtub',              '生活'],
    ['医療',            'medical_services',     '生活'],
    ['薬',              'medication',           '生活'],
    ['フィットネス',      'fitness_center',       '生活'],
    ['スパ',            'spa',                  '生活'],
    ['ペット',          'pets',                 '生活'],
    // 安全・警告
    ['警告',            'warning',              '安全'],
    ['エラー',          'error',                '安全'],
    ['禁止',            'block',                '安全'],
    ['消火器',          'fire_extinguisher',    '安全'],
    ['ヘルメット',       'hardware',             '安全'],
    ['救急',            'emergency',            '安全'],
    ['報告',            'report',               '安全'],
    // 工具・作業
    ['ツール',          'build',                '工具'],
    ['ハンマー',         'handyman',             '工具'],
    ['工事',            'construction',         '工具'],
    ['ネジ',            'screwdriver',          '工具'],
    ['センサー',         'sensors',              '工具'],
    ['計測',            'straighten',           '工具'],
    ['水道',            'plumbing',             '工具'],
    ['電気',            'electrical_services',  '工具'],
    // その他
    ['電球',            'lightbulb',            'その他'],
    ['ヘルプ',          'help',                 'その他'],
    ['情報',            'info',                 'その他'],
    ['リンク',          'link',                 'その他'],
    ['印刷',            'print',                'その他'],
    ['QR',             'qr_code',              'その他'],
    ['ロケット',         'rocket_launch',        'その他'],
    ['リスト',          'list',                 'その他'],
    ['テーブル',         'table_chart',          'その他'],
    ['Wi-Fi',          'wifi',                 'その他'],
    ['ブルートゥース',   'bluetooth',            'その他'],
    ['鍵',              'vpn_key',              'その他'],
    ['プレゼント',       'card_giftcard',        'その他'],
    ['ラベル',          'label',                'その他'],
    ['タグ',            'sell',                 'その他'],
    ['言語',            'g_translate',          'その他'],
    ['無限',            'all_inclusive',        'その他'],
    ['ハート',          'favorite_border',      'その他'],
    ['サムズアップ',      'thumb_up',             'その他'],
    ['スマイル',         'mood',                 'その他'],
    ['AI',              'auto_awesome',         'その他'],
    ['魔法',            'auto_fix_high',        'その他']
  ];

  sheet.getRange(2, 1, icons.length, 3).setValues(icons);

  // D列にIMAGE関数でアイコンプレビューを表示
  for (let i = 0; i < icons.length; i++) {
    const iconCode = icons[i][1];
    const row = i + 2;
    sheet.getRange(row, 4).setFormula(
      '=IMAGE("https://fonts.gstatic.com/s/i/short-term/release/materialsymbolsoutlined/' + iconCode + '/default/24px.svg")'
    );
  }

  // 行の高さを調整してアイコンが見やすいように
  for (let i = 2; i <= icons.length + 1; i++) {
    sheet.setRowHeight(i, 30);
  }

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 80);
}

/**
 * スプレッドシートを開いた時のカスタムメニュー
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗂️ タブ管理')
    .addItem('初期セットアップ', 'setupSpreadsheet')
    .addToUi();
}
