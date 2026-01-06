/**
 * 楽天レビュー収集 - Google Apps Script
 * Chrome拡張機能から送信されたレビューデータをスプレッドシートに保存する
 */

// スプレッドシートID
const SPREADSHEET_ID = '1o-VqcgiGf_1vItKItOIhz756bcV3KvqbBWuOH4KFnaQ';

/**
 * スプレッドシートを取得
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * POSTリクエストを処理
 * Chrome拡張機能からのレビューデータを受け取り、スプレッドシートに保存
 */
function doPost(e) {
  try {
    // リクエストボディをパース
    const data = JSON.parse(e.postData.contents);

    // スプレッドシートのURLを取得
    const ss = getSpreadsheet();
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

    // 商品ごとにシートを分けるかどうか（デフォルトはtrue）
    const separateSheets = data.separateSheets !== false;

    // スプレッドシートに保存
    const savedCount = saveReviews(data.reviews, separateSheets);

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
  const ss = getSpreadsheet();
  return createResponse({
    success: true,
    message: '楽天レビュー収集 GAS API は正常に動作しています',
    timestamp: new Date().toISOString(),
    spreadsheetUrl: ss.getUrl()
  });
}

/**
 * レビューをスプレッドシートに保存
 * @param {Array} reviews - レビューデータの配列
 * @param {boolean} separateSheets - 商品ごとにシートを分けるかどうか
 */
function saveReviews(reviews, separateSheets = true) {
  const ss = getSpreadsheet();

  if (separateSheets) {
    // 商品ごとにシートを分けて保存
    return saveReviewsByProduct(ss, reviews);
  } else {
    // 1つのシートにすべて保存
    return saveReviewsToSingleSheet(ss, reviews);
  }
}

/**
 * 商品ごとに別々のシートに保存
 */
function saveReviewsByProduct(ss, reviews) {
  let totalSaved = 0;

  // 商品管理番号ごとにレビューをグループ化
  const reviewsByProduct = {};
  reviews.forEach(review => {
    // レビューデータから商品管理番号を取得、なければURLから抽出
    const productId = review.productId || extractProductId(review.productUrl) || '不明な商品';
    if (!reviewsByProduct[productId]) {
      reviewsByProduct[productId] = [];
    }
    reviewsByProduct[productId].push(review);
  });

  // 各商品のシートに保存
  for (const productId in reviewsByProduct) {
    const productReviews = reviewsByProduct[productId];

    // シート名は商品管理番号をそのまま使用（31文字以内、特殊文字を除去）
    let sheetName = sanitizeSheetName(productId);

    // シートを取得または作成
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      addHeader(sheet);
    }

    // ヘッダーがなければ追加
    if (sheet.getLastRow() === 0) {
      addHeader(sheet);
    }

    // レビューデータを行に変換して追加
    const rows = productReviews.map(review => [
      review.reviewDate || '',
      review.productId || extractProductId(review.productUrl) || '',
      review.productName || '',
      review.productUrl || '',
      review.rating || '',
      review.title || '',
      review.body || '',
      review.author || '',
      review.age || '',
      review.gender || '',
      review.orderDate || '',
      review.variation || '',
      review.usage || '',
      review.recipient || '',
      review.purchaseCount || '',
      review.purchaseInfo || '',
      review.helpfulCount || 0,
      review.shopName || '',
      review.pageUrl || '',
      review.collectedAt || new Date().toISOString()
    ]);

    // データを追加
    if (rows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
      totalSaved += rows.length;
    }
  }

  return totalSaved;
}

/**
 * 1つのシートにすべて保存
 */
function saveReviewsToSingleSheet(ss, reviews) {
  let sheet = ss.getSheetByName('レビュー');

  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet('レビュー');
    addHeader(sheet);
  }

  // ヘッダーがなければ追加
  if (sheet.getLastRow() === 0) {
    addHeader(sheet);
  }

  // レビューデータを行に変換して追加
  const rows = reviews.map(review => [
    review.reviewDate || '',
    review.productId || extractProductId(review.productUrl) || '',
    review.productName || '',
    review.productUrl || '',
    review.rating || '',
    review.title || '',
    review.body || '',
    review.author || '',
    review.age || '',
    review.gender || '',
    review.orderDate || '',
    review.variation || '',
    review.usage || '',
    review.recipient || '',
    review.purchaseCount || '',
    review.purchaseInfo || '',
    review.helpfulCount || 0,
    review.shopName || '',
    review.pageUrl || '',
    review.collectedAt || new Date().toISOString()
  ]);

  // データを追加
  if (rows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  return rows.length;
}

/**
 * 商品URLから商品管理番号を抽出
 * 例: https://item.rakuten.co.jp/sakuradome/hug/ → hug
 */
function extractProductId(productUrl) {
  if (!productUrl) {
    return null;
  }

  try {
    // item.rakuten.co.jp/ショップ名/商品管理番号/ の形式から抽出
    const match = productUrl.match(/item\.rakuten\.co\.jp\/[^\/]+\/([^\/\?]+)/);
    if (match && match[1]) {
      return match[1];
    }

    // review.rakuten.co.jp/item/1/ショップID/商品ID/ の形式から抽出
    const reviewMatch = productUrl.match(/review\.rakuten\.co\.jp\/item\/\d+\/[^\/]+\/([^\/\?]+)/);
    if (reviewMatch && reviewMatch[1]) {
      return reviewMatch[1];
    }

    return null;
  } catch (e) {
    return null;
  }
}

/**
 * シート名をサニタイズ（特殊文字除去、31文字以内）
 */
function sanitizeSheetName(name) {
  // 使用できない文字を除去: * ? : \ / [ ]
  let sanitized = name.replace(/[*?:\\/\[\]]/g, '');

  // 31文字以内に切り詰め
  if (sanitized.length > 31) {
    sanitized = sanitized.substring(0, 31);
  }

  // 空文字になった場合
  if (!sanitized.trim()) {
    sanitized = '不明な商品';
  }

  return sanitized;
}

/**
 * ヘッダー行を追加
 */
function addHeader(sheet) {
  const headers = [
    'レビュー日',
    '商品管理番号',
    '商品名',
    '商品URL',
    '評価',
    'タイトル',
    '本文',
    '投稿者',
    '年代',
    '性別',
    '注文日',
    'バリエーション',
    '用途',
    '贈り先',
    '購入回数',
    '購入情報',
    '参考になった数',
    'ショップ名',
    'ページURL',
    '収集日時'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダーのスタイルを設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅を調整
  sheet.setColumnWidth(1, 100);  // レビュー日
  sheet.setColumnWidth(2, 120);  // 商品管理番号
  sheet.setColumnWidth(3, 300);  // 商品名
  sheet.setColumnWidth(4, 200);  // 商品URL
  sheet.setColumnWidth(5, 50);   // 評価
  sheet.setColumnWidth(6, 200);  // タイトル
  sheet.setColumnWidth(7, 400);  // 本文
  sheet.setColumnWidth(8, 100);  // 投稿者
  sheet.setColumnWidth(9, 60);   // 年代
  sheet.setColumnWidth(10, 60);  // 性別
  sheet.setColumnWidth(11, 100); // 注文日
  sheet.setColumnWidth(12, 150); // バリエーション
  sheet.setColumnWidth(13, 120); // 用途
  sheet.setColumnWidth(14, 80);  // 贈り先
  sheet.setColumnWidth(15, 80);  // 購入回数
  sheet.setColumnWidth(16, 150); // 購入情報
  sheet.setColumnWidth(17, 100); // 参考になった数
  sheet.setColumnWidth(18, 150); // ショップ名
  sheet.setColumnWidth(19, 200); // ページURL
  sheet.setColumnWidth(20, 150); // 収集日時

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

  const result = saveReviews(testData.reviews, true);
  Logger.log('保存件数: ' + result);
}

/**
 * シートをリセット（テスト用）
 * 注意: すべてのデータが削除されます
 */
function resetSheet() {
  const ss = getSpreadsheet();
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
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();

  let totalRemoved = 0;

  sheets.forEach(sheet => {
    if (sheet.getLastRow() <= 1) {
      return; // ヘッダーのみ or データなし
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

      totalRemoved += removedCount;
      Logger.log(sheet.getName() + ': ' + removedCount + '件の重複を削除');
    }
  });

  Logger.log('合計: ' + totalRemoved + '件の重複を削除しました');
}
