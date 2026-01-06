/**
 * 楽天レビュー収集 - Google Apps Script
 * Chrome拡張機能から送信されたレビューデータをスプレッドシートに保存する
 *
 * 使い方:
 * 1. Googleスプレッドシートを新規作成
 * 2. 拡張機能 → Apps Script を開く
 * 3. このコードを貼り付けて保存
 * 4. デプロイ → 新しいデプロイ → ウェブアプリ を選択
 * 5. アクセスできるユーザーを「全員」に設定してデプロイ
 * 6. 表示されたURLをChrome拡張機能の設定画面に貼り付け
 */

/**
 * POSTリクエストを処理
 * Chrome拡張機能からのレビューデータを受け取り、スプレッドシートに保存
 */
function doPost(e) {
  try {
    // リクエストボディをパース
    const data = JSON.parse(e.postData.contents);

    // スプレッドシートのURLを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetUrl = ss.getUrl();

    // テストリクエストの場合
    if (data.test) {
      return createResponse({
        success: true,
        message: '接続テスト成功',
        spreadsheetUrl: spreadsheetUrl
      });
    }

    // レビューデータがない場合
    if (!data.reviews || data.reviews.length === 0) {
      return createResponse({
        success: false,
        error: 'レビューデータがありません',
        spreadsheetUrl: spreadsheetUrl
      });
    }

    // スプレッドシートに保存
    const savedCount = saveReviews(data.reviews);

    return createResponse({
      success: true,
      message: `${savedCount}件のレビューを保存しました`,
      savedCount: savedCount,
      spreadsheetUrl: spreadsheetUrl
    });

  } catch (error) {
    console.error('エラー:', error);
    return createResponse({
      success: false,
      error: error.message
    });
  }
}

/**
 * GETリクエストを処理（テスト用）
 */
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return createResponse({
    success: true,
    message: '楽天レビュー収集 GAS API は正常に動作しています',
    timestamp: new Date().toISOString(),
    spreadsheetUrl: ss.getUrl()
  });
}

/**
 * レビューをスプレッドシートに保存
 */
function saveReviews(reviews) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('レビュー');

  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet('レビュー');
    // ヘッダーを追加
    addHeader(sheet);
  }

  // ヘッダーがなければ追加
  if (sheet.getLastRow() === 0) {
    addHeader(sheet);
  }

  // レビューデータを行に変換して追加
  const rows = reviews.map(review => [
    review.collectedAt || new Date().toISOString(),
    review.productName || '',
    review.productUrl || '',
    review.rating || '',
    review.title || '',
    review.body || '',
    review.author || '',
    review.reviewDate || '',
    review.purchaseInfo || '',
    review.helpfulCount || 0,
    review.pageUrl || ''
  ]);

  // データを追加
  if (rows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  return rows.length;
}

/**
 * ヘッダー行を追加
 */
function addHeader(sheet) {
  const headers = [
    '収集日時',
    '商品名',
    '商品URL',
    '評価',
    'タイトル',
    '本文',
    '投稿者',
    'レビュー日',
    '購入情報',
    '参考になった数',
    'ページURL'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダーのスタイルを設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅を調整
  sheet.setColumnWidth(1, 150);  // 収集日時
  sheet.setColumnWidth(2, 300);  // 商品名
  sheet.setColumnWidth(3, 200);  // 商品URL
  sheet.setColumnWidth(4, 50);   // 評価
  sheet.setColumnWidth(5, 200);  // タイトル
  sheet.setColumnWidth(6, 400);  // 本文
  sheet.setColumnWidth(7, 100);  // 投稿者
  sheet.setColumnWidth(8, 100);  // レビュー日
  sheet.setColumnWidth(9, 150);  // 購入情報
  sheet.setColumnWidth(10, 100); // 参考になった数
  sheet.setColumnWidth(11, 200); // ページURL

  // ヘッダー行を固定
  sheet.setFrozenRows(1);
}

/**
 * JSONレスポンスを作成
 */
function createResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * テスト用関数 - 手動でデータを追加してテスト
 */
function testAddReview() {
  const testData = {
    reviews: [
      {
        collectedAt: new Date().toISOString(),
        productName: 'テスト商品',
        productUrl: 'https://example.com/product',
        rating: 5,
        title: 'とても良い商品です',
        body: 'この商品を購入して大変満足しています。品質も良く、配送も早かったです。',
        author: 'テストユーザー',
        reviewDate: '2024-01-01',
        purchaseInfo: 'サイズ: M, カラー: ブラック',
        helpfulCount: 10,
        pageUrl: 'https://review.rakuten.co.jp/test'
      }
    ]
  };

  const result = saveReviews(testData.reviews);
  Logger.log('保存件数: ' + result);
}

/**
 * シートをリセット（テスト用）
 * 注意: すべてのデータが削除されます
 */
function resetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('レビュー');

  if (sheet) {
    sheet.clear();
    addHeader(sheet);
    Logger.log('シートをリセットしました');
  } else {
    Logger.log('シートが見つかりません');
  }
}

/**
 * 重複レビューを削除（メンテナンス用）
 * 本文と投稿者が同じレビューを重複とみなす
 */
function removeDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('レビュー');

  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log('データがありません');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  // 重複をチェック（本文 + 投稿者 をキーとする）
  const seen = new Set();
  const uniqueRows = [];

  rows.forEach(row => {
    const body = row[5] || ''; // 本文
    const author = row[6] || ''; // 投稿者
    const key = body.substring(0, 100) + '|' + author;

    if (!seen.has(key)) {
      seen.add(key);
      uniqueRows.push(row);
    }
  });

  const removedCount = rows.length - uniqueRows.length;

  if (removedCount > 0) {
    // シートをクリアして、ヘッダーとユニークなデータを再挿入
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    if (uniqueRows.length > 0) {
      sheet.getRange(2, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
    }

    // ヘッダーのスタイルを再適用
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);

    Logger.log(removedCount + '件の重複を削除しました');
  } else {
    Logger.log('重複はありませんでした');
  }
}
