/**
 * 設定シートにChatwork項目を追加するスクリプト
 *
 * 【使い方】
 * 1. Apps Scriptで「addChatworkSettings」を選択
 * 2. 実行ボタン（▶）をクリック
 */

function addChatworkSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('設定');

  // 現在の最終行を取得
  const lastRow = sheet.getLastRow();

  // Chatwork設定を追加
  const newSettings = [
    ['Chatwork APIトークン', '', 'ChatworkのAPIトークンを入力'],
    ['Chatwork ルームID', '', '通知先のルームID（数字のみ）']
  ];

  sheet.getRange(lastRow + 1, 1, newSettings.length, 3).setValues(newSettings);

  // 入力欄をハイライト
  sheet.getRange(lastRow + 1, 2, newSettings.length, 1).setBackground('#fff3cd');

  SpreadsheetApp.getUi().alert('✅ Chatwork設定項目を追加しました！');
}
