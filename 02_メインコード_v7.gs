/**
 * GAS案件 自動見積もりシステム - メインコード（v7 顧客情報追加版）
 *
 * 変更点:
 * - 事業形態、年商、決算済み、従業員数の4項目を追加
 * - スプレッドシートの列構成を更新
 */

// ===========================================
// スプレッドシートIDを直接指定
// ===========================================
const SPREADSHEET_ID = '1A0QgC425A1yQDoZyg8ieB7sNPsyksH7b1uXrJeAJZL0';

// ===========================================
// スプレッドシートを取得
// ===========================================
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ===========================================
// 設定値の取得
// ===========================================
function getConfig(key) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('設定');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return null;
}

// ===========================================
// Webアプリとしてフォームを表示
// ===========================================
function doGet(e) {
  const page = e ? e.parameter.page : null;

  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('管理画面')
      .setTitle('案件管理画面')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'test') {
    return HtmlService.createHtmlOutputFromFile('テスト管理画面')
      .setTitle('テスト')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile('フォーム')
    .setTitle('無料システム開発 お問い合わせフォーム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===========================================
// フォーム送信処理
// ===========================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = processFormSubmission(data);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ===========================================
// フォーム送信のメイン処理
// ===========================================
function processFormSubmission(data) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('問い合わせ一覧');

  // Step 1: NGフィルターチェック
  const ngResult = checkNGFilter(data.ngFilters);

  let judgment = '';
  let totalMinutes = 0;
  let reason = '';

  if (ngResult.isNG) {
    // NGフィルターに該当
    judgment = 'NG（技術制約）';
    reason = ngResult.reason;
  } else {
    // Step 2: 規模感チェック（GAS技術制約）
    const scaleResult = checkScaleConstraints(data);

    if (scaleResult.isNG) {
      judgment = 'NG（技術制約）';
      reason = scaleResult.reason;
    } else {
      // Step 3: Claude APIで工数判定
      const claudeResult = judgeWithClaudeAPI(data);

      if (claudeResult) {
        totalMinutes = claudeResult.totalMinutes;
        reason = claudeResult.reason;

        // 規模感による工数加算
        if (scaleResult.additionalMinutes > 0) {
          totalMinutes += scaleResult.additionalMinutes;
          reason += ' ' + scaleResult.warningMessage;
        }

        // 最終判定
        const okThreshold = parseInt(getConfig('無料対応しきい値（分）')) || 360;
        const borderlineThreshold = parseInt(getConfig('要ヒアリングしきい値（分）')) || 480;

        if (totalMinutes <= okThreshold) {
          judgment = 'OK';
        } else if (totalMinutes <= borderlineThreshold) {
          judgment = 'BORDERLINE';
        } else {
          judgment = 'NG（工数オーバー）';
        }
      } else {
        // Claude APIが使えない場合はシンプル判定
        const simpleResult = simpleJudgment(data);
        judgment = simpleResult.judgment;
        totalMinutes = simpleResult.totalMinutes;
        reason = simpleResult.reason;

        // 規模感による工数加算
        if (scaleResult.additionalMinutes > 0) {
          totalMinutes += scaleResult.additionalMinutes;
          reason += ' ' + scaleResult.warningMessage;
        }
      }
    }
  }

  // スプレッドシートに記録（新しい列構成）
  const timestamp = new Date();
  const rowData = [
    timestamp,                                      // A: タイムスタンプ
    data.companyName,                               // B: 会社名
    data.contactName,                               // C: 担当者名
    data.email,                                     // D: メールアドレス
    data.phone || '',                               // E: 電話番号
    data.businessType || '',                        // F: 事業形態（新規）
    data.annualRevenue || '',                       // G: 年商（新規）
    data.taxFiled || '',                            // H: 決算済みか（新規）
    data.employeeCount || '',                       // I: 従業員数（新規）
    data.categories ? data.categories.join(', ') : '',  // J: カテゴリ
    data.userType,                                  // K: 利用者
    data.concurrentUsers,                           // L: 同時利用人数
    data.dataVolume,                                // M: 想定データ量
    data.processTiming,                             // N: 処理タイミング
    data.dataMigration,                             // O: 既存データ移行
    data.designRequirement,                         // P: デザイン要件
    data.ngFilters ? data.ngFilters.join(', ') : 'なし',  // Q: NGフィルター
    data.description,                               // R: 詳細
    judgment,                                       // S: 判定結果
    totalMinutes,                                   // T: 工数（分）
    reason,                                         // U: 理由
    '未対応'                                        // V: ステータス
  ];

  sheet.appendRow(rowData);
  sendAutoReplyEmail(data, judgment, reason);

  if (judgment === 'OK' || judgment === 'BORDERLINE') {
    sendChatworkNotification(data, judgment, totalMinutes, reason);
  }

  return {
    success: true,
    judgment: judgment,
    totalMinutes: totalMinutes,
    reason: reason
  };
}

// ===========================================
// NGフィルターチェック
// ===========================================
function checkNGFilter(ngFilters) {
  if (!ngFilters || ngFilters.length === 0) {
    return { isNG: false };
  }

  const ngReasons = {
    'native_app': 'スマホアプリ（ネイティブ）の開発は、Google Apps Scriptでは対応できません。',
    'public_system': '一般公開システムは、アクセス数やセキュリティの観点からGASの対応範囲外となります。',
    'existing_system': '既存システム（kintone、Salesforce等）の改修は、各プラットフォーム固有の専門知識が必要なため対応外です。',
    'realtime_sync': 'リアルタイム同期（チャット、在庫同期等）は、GASの実行制限により実現が困難です。',
    'external_users': 'Google Workspace外のユーザー利用は、認証・権限管理の観点から対応外となります。'
  };

  for (const filter of ngFilters) {
    if (ngReasons[filter]) {
      return {
        isNG: true,
        reason: ngReasons[filter]
      };
    }
  }

  return { isNG: false };
}

// ===========================================
// 規模感チェック（GAS技術制約）
// ===========================================
function checkScaleConstraints(data) {
  let isNG = false;
  let reason = '';
  let additionalMinutes = 0;
  let warningMessage = '';
  let warnings = [];

  // 1. データ量チェック
  if (data.dataVolume === '1万行以上') {
    // GASの制限に近いため警告
    additionalMinutes += 60;
    warnings.push('データ量が多いため、パフォーマンス対策が必要です（+60分）');
  }

  // 2. 同時利用人数チェック
  if (data.concurrentUsers === '数十人以上') {
    additionalMinutes += 45;
    warnings.push('同時利用人数が多いため、排他制御の考慮が必要です（+45分）');
  }

  // 3. 処理タイミングチェック
  if (data.processTiming === 'リアルタイム必須') {
    // NGフィルターで既にチェックしているが、念のため
    additionalMinutes += 90;
    warnings.push('リアルタイム処理はGASの制限があるため、工数が増加します（+90分）');
  }

  // 4. 既存データ移行チェック
  if (data.dataMigration === 'あり（大量）') {
    additionalMinutes += 60;
    warnings.push('大量データ移行のため、インポート処理が必要です（+60分）');
  }

  // 5. デザイン要件チェック
  if (data.designRequirement === 'こだわりたい') {
    additionalMinutes += 60;
    warnings.push('デザインにこだわるため、UI調整の工数が増加します（+60分）');
  }

  // 6. 社外利用チェック
  if (data.userType === '社外（取引先等）も利用') {
    additionalMinutes += 45;
    warnings.push('社外利用のため、セキュリティ・権限管理が必要です（+45分）');
  }

  // 複合条件でのNG判定
  // データ量1万行以上 + 同時利用数十人以上 + リアルタイム = NG
  if (data.dataVolume === '1万行以上' &&
      data.concurrentUsers === '数十人以上' &&
      data.processTiming === 'リアルタイム必須') {
    isNG = true;
    reason = 'データ量・同時利用人数・リアルタイム処理の組み合わせは、GASの技術制約を超えるため対応困難です。';
  }

  // 警告メッセージを結合
  if (warnings.length > 0) {
    warningMessage = '【規模感による追加工数】' + warnings.join(' / ');
  }

  return {
    isNG: isNG,
    reason: reason,
    additionalMinutes: additionalMinutes,
    warningMessage: warningMessage
  };
}

// ===========================================
// シンプル判定
// ===========================================
function simpleJudgment(data) {
  let baseMinutes = 120;

  const categoryCount = data.categories ? data.categories.length : 1;
  baseMinutes += categoryCount * 60;

  // 規模感での加算は checkScaleConstraints で行うため、ここでは基本のみ
  const totalMinutes = Math.round(baseMinutes * 1.2 + 60);

  const okThreshold = parseInt(getConfig('無料対応しきい値（分）')) || 360;
  const borderlineThreshold = parseInt(getConfig('要ヒアリングしきい値（分）')) || 480;

  let judgment = '';
  let reason = '';

  if (totalMinutes <= okThreshold) {
    judgment = 'OK';
    reason = `推定工数 ${totalMinutes}分（約${Math.round(totalMinutes/60)}時間）で、無料対応の範囲内です。`;
  } else if (totalMinutes <= borderlineThreshold) {
    judgment = 'BORDERLINE';
    reason = `推定工数 ${totalMinutes}分（約${Math.round(totalMinutes/60)}時間）です。機能を絞れば無料対応可能な可能性があります。`;
  } else {
    judgment = 'NG（工数オーバー）';
    reason = `推定工数 ${totalMinutes}分（約${Math.round(totalMinutes/60)}時間）で、無料対応の範囲を超えています。`;
  }

  return {
    judgment: judgment,
    totalMinutes: totalMinutes,
    reason: reason
  };
}

// ===========================================
// 自動返信メール送信
// ===========================================
function sendAutoReplyEmail(data, judgment, ngReason) {
  const ss = getSpreadsheet();
  const templateSheet = ss.getSheetByName('メールテンプレート');
  const templates = templateSheet.getDataRange().getValues();

  const isTestMode = getConfig('テストモード') === 'TRUE';

  let templateRow = null;
  for (let i = 1; i < templates.length; i++) {
    const pattern = templates[i][0];
    if (judgment === 'OK' && pattern.includes('OK')) {
      templateRow = templates[i];
      break;
    } else if (judgment === 'BORDERLINE' && pattern.includes('BORDERLINE')) {
      templateRow = templates[i];
      break;
    } else if (judgment === 'NG（工数オーバー）' && pattern.includes('工数オーバー')) {
      templateRow = templates[i];
      break;
    } else if (judgment === 'NG（技術制約）' && pattern.includes('技術制約')) {
      templateRow = templates[i];
      break;
    }
  }

  if (!templateRow) {
    console.log('テンプレートが見つかりませんでした: ' + judgment);
    return;
  }

  const schedulingUrl = getConfig('日程調整URL') || 'https://calendly.com/your-link';

  let subject = templateRow[1];
  let body = templateRow[2];

  body = body.replace(/\{\{会社名\}\}/g, data.companyName);
  body = body.replace(/\{\{担当者名\}\}/g, data.contactName);
  body = body.replace(/\{\{日程調整URL\}\}/g, schedulingUrl);
  body = body.replace(/\{\{NG理由\}\}/g, ngReason || '');

  if (isTestMode) {
    console.log('=== テストモード：メール送信スキップ ===');
    console.log('宛先: ' + data.email);
    console.log('件名: ' + subject);
    console.log('本文: ' + body);
  } else {
    MailApp.sendEmail({
      to: data.email,
      subject: subject,
      body: body
    });
  }
}

// ===========================================
// 管理画面用：問い合わせ一覧を取得
// ===========================================
function getInquiries() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('問い合わせ一覧');
    const data = sheet.getDataRange().getValues();

    console.log('データ行数: ' + data.length);

    if (data.length <= 1) {
      return [];
    }

    const headers = data[0];
    console.log('ヘッダー: ' + headers.join(', '));

    const inquiries = [];

    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const inquiry = {};

      for (let j = 0; j < headers.length; j++) {
        if (headers[j] === 'タイムスタンプ' && row[j] instanceof Date) {
          inquiry[headers[j]] = row[j].toISOString();
        } else {
          inquiry[headers[j]] = row[j];
        }
      }
      inquiry['rowIndex'] = i + 1;

      inquiries.push(inquiry);
    }

    console.log('取得件数: ' + inquiries.length);
    return inquiries;
  } catch (error) {
    console.error('getInquiries エラー: ' + error);
    return [];
  }
}

// ===========================================
// 管理画面用：ステータスを更新（新しい列位置に対応）
// ===========================================
function updateStatus(rowIndex, newStatus) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('問い合わせ一覧');

  // ステータスは22列目（V列）に変更
  sheet.getRange(rowIndex, 22).setValue(newStatus);

  return { success: true };
}

// ===========================================
// テスト用関数
// ===========================================
function testFormSubmission() {
  const testData = {
    companyName: 'テスト株式会社',
    contactName: 'テスト太郎',
    email: 'test@example.com',
    phone: '03-1234-5678',
    businessType: '法人',
    annualRevenue: '1,000万〜1,500万円',
    taxFiled: 'はい',
    employeeCount: '5',
    categories: ['データ集計・転記の自動化', 'フォーム・申請システム'],
    userType: '社内のみ',
    concurrentUsers: '5人以下',
    dataVolume: '1000行以下',
    processTiming: '手動実行',
    dataMigration: 'なし',
    designRequirement: '最低限でOK',
    ngFilters: [],
    description: '売上データを毎日集計して、レポートを自動作成したい。データはスプレッドシートに保存されており、毎週月曜日にPDFレポートを作成してメールで送信したい。また、特定の条件（売上が目標を下回った場合など）にはアラート通知も欲しい。さらに、データのフィルタリング機能も追加して、期間別・商品別に集計できるようにしたい。'
  };

  const result = processFormSubmission(testData);
  console.log(result);
}

// 規模感チェックのテスト
function testScaleCheck() {
  // 大規模なテストデータ
  const testData = {
    userType: '社外（取引先等）も利用',
    concurrentUsers: '数十人以上',
    dataVolume: '1万行以上',
    processTiming: '定期バッチ',
    dataMigration: 'あり（大量）',
    designRequirement: 'こだわりたい'
  };

  const result = checkScaleConstraints(testData);
  console.log('規模感チェック結果:', result);
}
