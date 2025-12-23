/**
 * スプレッドシート初期設定スクリプト
 *
 * 【使い方】
 * 1. スプレッドシートを開く
 * 2. 拡張機能 → Apps Script をクリック
 * 3. このコードを全て貼り付け
 * 4. 「setupAllSheets」を選択して実行ボタン（▶）をクリック
 * 5. 権限を許可する
 */

function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 既存のシートを取得（最初のシートを「問い合わせ一覧」として使う）
  let sheet1 = ss.getSheets()[0];
  sheet1.setName('問い合わせ一覧');

  // シート2〜4を作成
  let sheet2 = ss.getSheetByName('工数マスタ') || ss.insertSheet('工数マスタ');
  let sheet3 = ss.getSheetByName('メールテンプレート') || ss.insertSheet('メールテンプレート');
  let sheet4 = ss.getSheetByName('設定') || ss.insertSheet('設定');

  // 各シートを設定
  setup問い合わせ一覧(sheet1);
  setup工数マスタ(sheet2);
  setupメールテンプレート(sheet3);
  setup設定(sheet4);

  // 完了メッセージ
  SpreadsheetApp.getUi().alert('✅ 全てのシートの初期設定が完了しました！');
}

function setup問い合わせ一覧(sheet) {
  // ヘッダー設定
  const headers = [
    'タイムスタンプ', '会社名', '担当者名', 'メールアドレス', '電話番号',
    'カテゴリ', '利用者', '同時利用人数', '想定データ量', '処理タイミング',
    '既存データ移行', 'デザイン要件', 'NGフィルター回答', '自由記述',
    '判定結果', '工数(分)', '判定理由', 'ステータス'
  ];

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 150);  // タイムスタンプ
  sheet.setColumnWidth(2, 150);  // 会社名
  sheet.setColumnWidth(3, 100);  // 担当者名
  sheet.setColumnWidth(4, 200);  // メールアドレス
  sheet.setColumnWidth(14, 300); // 自由記述
  sheet.setColumnWidth(17, 300); // 判定理由

  // 行を固定
  sheet.setFrozenRows(1);
}

function setup工数マスタ(sheet) {
  sheet.clear();

  // データ構造
  const data = [
    ['【工数マスタ】', '', ''],
    ['', '', ''],
    ['■ データ構造', '工数（分）', ''],
    ['単一テーブル（マスタ1個）', 15, ''],
    ['2〜3テーブル', 30, ''],
    ['4〜5テーブル＋リレーション', 60, ''],
    ['複雑な正規化・多対多', 120, ''],
    ['', '', ''],
    ['■ 画面', '工数（分）', ''],
    ['一覧画面（テーブル表示）', 20, ''],
    ['入力フォーム（基本）', 20, ''],
    ['入力フォーム（バリデーション込み）', 40, ''],
    ['詳細・編集画面', 30, ''],
    ['ダッシュボード（集計表示）', 45, ''],
    ['ダッシュボード（グラフ付き）', 60, ''],
    ['検索・フィルタ・ソート', 30, ''],
    ['モーダル・ポップアップ', 15, ''],
    ['', '', ''],
    ['■ 機能', '工数（分）', ''],
    ['CRUD基本セット', 30, ''],
    ['メール自動送信', 20, ''],
    ['Slack通知', 20, ''],
    ['LINE通知', 30, ''],
    ['PDF生成', 45, ''],
    ['スプレッドシート出力', 20, ''],
    ['定期実行トリガー設定', 10, ''],
    ['外部API連携（ドキュメント明確）', 45, ''],
    ['外部API連携（認証複雑）', 90, ''],
    ['承認ワークフロー（1段階）', 40, ''],
    ['承認ワークフロー（多段階・分岐あり）', 90, ''],
    ['権限管理（2〜3ロール）', 30, ''],
    ['権限管理（複雑な権限マトリクス）', 75, ''],
    ['ファイルアップロード', 30, ''],
    ['カレンダー連携', 40, ''],
    ['既存データ移行・インポート', 60, ''],
    ['', '', ''],
    ['■ 難易度係数', '係数', ''],
    ['要件が明文化されている', 1.0, ''],
    ['口頭ベース、でも具体的', 1.3, ''],
    ['「いい感じにして」系', 2.0, ''],
    ['既存業務の置き換え（暗黙知多い）', 1.8, ''],
    ['', '', ''],
    ['■ バッファ', '工数', ''],
    ['環境構築・初期設定', 30, '分'],
    ['テスト・デバッグ', 0.2, '×見積もり'],
    ['軽微な修正対応', 30, '分']
  ];

  sheet.getRange(1, 1, data.length, 3).setValues(data);

  // 書式設定
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A3').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A9').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A19').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A37').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A43').setFontWeight('bold').setBackground('#e8f0fe');

  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 100);
}

function setupメールテンプレート(sheet) {
  sheet.clear();

  const data = [
    ['パターン', '件名', '本文'],
    [
      'OK（無料対応可能）',
      '【無料開発確定】システム開発のお申込みありがとうございます',
      '{{会社名}} {{担当者名}}様\n\nお問い合わせいただきありがとうございます。\n合同会社リバイラル システム開発事業部です。\n\nお問い合わせ内容を確認した結果、無料での開発が確定いたしました。\n\n【納期について】\nお申込みから一週間後の納品となります。\n\n下記URLより、一週間以降のご都合の良い日時を選択し、\n納品日のご予約をお願いいたします。\n\n【納品日予約URL】\nhttps://timerex.net/s/cz1917903_47c5/b09b13d3\n\nご予約完了後、開発を開始いたします。\nご不明な点がございましたら、お気軽にお問い合わせください。\n\n--\n合同会社リバイラル\nシステム開発事業部'
    ],
    [
      'BORDERLINE（要ヒアリング）',
      '【ご相談】システム開発のお申込みについて',
      '{{会社名}} {{担当者名}}様\n\nお問い合わせいただきありがとうございます。\n合同会社リバイラル システム開発事業部です。\n\nお問い合わせ内容を確認した結果、機能によっては無料対応が可能です。\n\n一度オンラインで詳細をお伺いし、無料対応範囲をご提案させてください。\n以下のURLからご都合の良い日時をお選びください。\n\n【日程調整URL】\n{{日程調整URL}}\n\nご不明な点がございましたら、お気軽にお問い合わせください。\n\n--\n合同会社リバイラル\nシステム開発事業部'
    ],
    [
      'NG（工数オーバー）',
      '【有料プランのご案内】システム開発のお申込みについて',
      '{{会社名}} {{担当者名}}様\n\nお問い合わせいただきありがとうございます。\n合同会社リバイラル システム開発事業部です。\n\nお問い合わせ内容を確認した結果、無料対応の範囲を超えるご要望でした。\n\n有料プランでのご対応となりますが、ご興味があれば詳細をご案内いたします。\n以下のURLからご都合の良い日時をお選びください。\n\n【日程調整URL】\n{{日程調整URL}}\n\nご不明な点がございましたら、お気軽にお問い合わせください。\n\n--\n合同会社リバイラル\nシステム開発事業部'
    ],
    [
      'NG（技術制約）',
      '【ご案内】システム開発のお申込みについて',
      '{{会社名}} {{担当者名}}様\n\nお問い合わせいただきありがとうございます。\n合同会社リバイラル システム開発事業部です。\n\nお問い合わせ内容を確認した結果、当サービスの対応範囲外となります。\n\n【対応不可の理由】\n{{NG理由}}\n\n当サービスはGoogle Apps Script（GAS）を使用したシステム開発に特化しております。\nご要望に沿えず申し訳ございません。\n\n--\n合同会社リバイラル\nシステム開発事業部'
    ]
  ];

  sheet.getRange(1, 1, data.length, 3).setValues(data);

  // 書式設定
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 350);
  sheet.setColumnWidth(3, 600);

  // 本文列を折り返し表示
  sheet.getRange('C:C').setWrap(true);

  sheet.setFrozenRows(1);
}

function setup設定(sheet) {
  sheet.clear();

  const data = [
    ['設定項目', '値', '説明'],
    ['CLAUDE_API_KEY', '', 'Claude APIキーを入力してください'],
    ['日程調整URL', 'https://calendly.com/your-link', 'Calendly等の日程調整リンク'],
    ['Slack Webhook URL', '', 'Slack通知用のWebhook URL（Phase 3で使用）'],
    ['通知先メールアドレス', '', '社内通知用メールアドレス'],
    ['無料対応しきい値（分）', '360', 'この時間以下なら無料対応OK'],
    ['要ヒアリングしきい値（分）', '480', 'この時間以下なら要ヒアリング'],
    ['テストモード', 'TRUE', 'TRUEの場合、実際にメール送信しない']
  ];

  sheet.getRange(1, 1, data.length, 3).setValues(data);

  // 書式設定
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 350);
  sheet.setColumnWidth(3, 300);

  // 入力欄をハイライト
  sheet.getRange('B2:B8').setBackground('#fff3cd');

  sheet.setFrozenRows(1);
}
