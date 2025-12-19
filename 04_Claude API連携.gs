/**
 * Claude API連携モジュール
 *
 * 【このファイルの追加方法】
 * 1. Apps Script画面の左側「ファイル」の横にある「+」をクリック
 * 2. 「スクリプト」を選択
 * 3. ファイル名を「ClaudeAPI」に変更
 * 4. このコードを全て貼り付け
 */

// ===========================================
// Claude APIで工数判定
// ===========================================
function judgeWithClaudeAPI(data) {
  const apiKey = getConfig('CLAUDE_API_KEY');

  if (!apiKey) {
    console.log('Claude APIキーが設定されていません。シンプル判定を使用します。');
    return null;
  }

  const prompt = buildPrompt(data);

  try {
    const response = callClaudeAPI(apiKey, prompt);
    return parseClaudeResponse(response);
  } catch (error) {
    console.error('Claude API呼び出しエラー:', error);
    return null;
  }
}

// ===========================================
// プロンプト生成
// ===========================================
function buildPrompt(data) {
  const prompt = `あなたはGAS（Google Apps Script）開発の工数見積もりエキスパートです。
以下の案件情報をもとに、Claude Codeでバイブコーディングした場合の工数（分）を算出してください。

【案件情報】
- 会社名: ${data.companyName}
- カテゴリ: ${data.categories ? data.categories.join(', ') : '未選択'}
- 利用者: ${data.userType}
- 同時利用人数: ${data.concurrentUsers}
- 想定データ量: ${data.dataVolume}
- 処理タイミング: ${data.processTiming}
- 既存データ移行: ${data.dataMigration}
- デザイン要件: ${data.designRequirement}
- 詳細説明: ${data.description}

【工数テーブル】
■ データ構造
- 単一テーブル（マスタ1個）: 15分
- 2〜3テーブル: 30分
- 4〜5テーブル＋リレーション: 60分
- 複雑な正規化・多対多: 120分

■ 画面
- 一覧画面（テーブル表示）: 20分
- 入力フォーム（基本）: 20分
- 入力フォーム（バリデーション込み）: 40分
- 詳細・編集画面: 30分
- ダッシュボード（集計表示）: 45分
- ダッシュボード（グラフ付き）: 60分
- 検索・フィルタ・ソート: 30分
- モーダル・ポップアップ: 15分

■ 機能
- CRUD基本セット: 30分
- メール自動送信: 20分
- Slack通知: 20分
- LINE通知: 30分
- PDF生成: 45分
- スプレッドシート出力: 20分
- 定期実行トリガー設定: 10分
- 外部API連携（ドキュメント明確）: 45分
- 外部API連携（認証複雑）: 90分
- 承認ワークフロー（1段階）: 40分
- 承認ワークフロー（多段階・分岐あり）: 90分
- 権限管理（2〜3ロール）: 30分
- 権限管理（複雑な権限マトリクス）: 75分
- ファイルアップロード: 30分
- カレンダー連携: 40分
- 既存データ移行・インポート: 60分

■ 難易度係数
- 要件が明文化されている: ×1.0
- 口頭ベース、でも具体的: ×1.3
- 「いい感じにして」系: ×2.0
- 既存業務の置き換え（暗黙知多い）: ×1.8

■ バッファ
- 環境構築・初期設定: 30分
- テスト・デバッグ: 見積もり×0.2
- 軽微な修正対応: 30分

【判定基準】
- OK（360分以下）: 無料対応可能
- BORDERLINE（361〜480分）: 要ヒアリング・機能削減提案
- NG（481分以上）: 有料プラン案内

【出力形式】
必ず以下のJSON形式のみで出力してください。説明文は不要です。
{
  "totalMinutes": 数値,
  "breakdown": {
    "dataStructure": {"item": "項目名", "minutes": 数値},
    "screens": [{"item": "項目名", "minutes": 数値}],
    "functions": [{"item": "項目名", "minutes": 数値}],
    "difficultyFactor": 数値,
    "buffer": 数値
  },
  "feasibility": "OK" または "BORDERLINE" または "NG",
  "reason": "判定理由を1〜2文で"
}`;

  return prompt;
}

// ===========================================
// Claude API呼び出し
// ===========================================
function callClaudeAPI(apiKey, prompt) {
  const url = 'https://api.anthropic.com/v1/messages';

  const payload = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 1024,
    messages: [
      {
        role: 'user',
        content: prompt
      }
    ]
  };

  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error(`API Error (${responseCode}): ${responseText}`);
  }

  const responseJson = JSON.parse(responseText);
  return responseJson.content[0].text;
}

// ===========================================
// Claude レスポンス解析
// ===========================================
function parseClaudeResponse(responseText) {
  try {
    // JSONを抽出（余計なテキストがある場合に対応）
    const jsonMatch = responseText.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      throw new Error('JSON not found in response');
    }

    const result = JSON.parse(jsonMatch[0]);

    return {
      totalMinutes: result.totalMinutes,
      breakdown: result.breakdown,
      feasibility: result.feasibility,
      reason: result.reason
    };
  } catch (error) {
    console.error('レスポンス解析エラー:', error, responseText);
    return null;
  }
}

// ===========================================
// テスト用関数
// ===========================================
function testClaudeAPI() {
  const testData = {
    companyName: 'テスト株式会社',
    categories: ['データ集計・転記の自動化', 'ダッシュボード・レポート'],
    userType: '社内のみ',
    concurrentUsers: '10人程度',
    dataVolume: '1000行以下',
    processTiming: '定期バッチ',
    dataMigration: 'なし',
    designRequirement: 'ある程度きれいに',
    description: '毎月の売上データをスプレッドシートから集計して、部門別・商品別のレポートを自動生成したい。グラフも表示して、月初に自動でメール送信してほしい。'
  };

  const result = judgeWithClaudeAPI(testData);
  console.log('Claude API判定結果:', result);
}
