/**
 * Chatwork通知モジュール v2
 */

// ===========================================
// Chatwork通知を送信
// ===========================================
function sendChatworkNotification(data, judgment, totalMinutes, reason) {
  const apiToken = getConfig('Chatwork APIトークン');
  const roomId = getConfig('Chatwork ルームID');

  if (!apiToken || !roomId) {
    console.log('Chatwork設定が不完全です。通知をスキップします。');
    return;
  }

  // 判定に応じた絵文字
  var emoji = '';
  switch (judgment) {
    case 'OK':
      emoji = '✅ 無料対応OK';
      break;
    case 'BORDERLINE':
      emoji = '🤔 要ヒアリング';
      break;
    case 'NG（工数オーバー）':
      emoji = '📋 工数オーバー';
      break;
    case 'NG（技術制約）':
      emoji = '⚠️ 技術制約';
      break;
    default:
      emoji = '📩 新規問い合わせ';
  }

  // Chatworkメッセージを構築
  var message = '[info][title]' + emoji + '[/title]';
  message += '新規お問い合わせがありました\n\n';
  message += '会社名: ' + data.companyName + '\n';
  message += '担当者: ' + data.contactName + '\n';
  message += 'メール: ' + data.email + '\n\n';
  message += '【判定結果】' + judgment + '\n';
  message += '【推定工数】' + totalMinutes + '分（約' + Math.round(totalMinutes/60) + '時間）\n\n';
  message += '【カテゴリ】\n' + (data.categories ? data.categories.join(', ') : '未選択') + '\n\n';
  message += '【詳細】\n' + truncateText(data.description, 300) + '\n\n';
  message += '【判定理由】\n' + reason + '\n\n';
  message += 'スプレッドシート: ' + SpreadsheetApp.openById(SPREADSHEET_ID).getUrl();
  message += '[/info]';

  // Chatworkに送信
  try {
    var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/messages';

    var options = {
      method: 'post',
      headers: {
        'X-ChatWorkToken': apiToken,
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: 'body=' + encodeURIComponent(message),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      console.error('Chatwork通知エラー:', response.getContentText());
    } else {
      console.log('Chatwork通知を送信しました');
    }
  } catch (error) {
    console.error('Chatwork通知エラー:', error);
  }
}

// ===========================================
// テキストを指定文字数で切り詰め
// ===========================================
function truncateText(text, maxLength) {
  if (!text) return '';
  if (text.length <= maxLength) return text;
  return text.substring(0, maxLength) + '...';
}

// ===========================================
// テスト用関数
// ===========================================
function testChatworkNotification() {
  var testData = {
    companyName: 'テスト株式会社',
    contactName: 'テスト太郎',
    email: 'test@example.com',
    categories: ['データ集計・転記の自動化'],
    description: 'これはテスト通知です。売上データを毎日集計して、レポートを自動作成したい。'
  };

  sendChatworkNotification(testData, 'OK', 240, '推定工数240分で、無料対応の範囲内です。');
}
