/**
 * 楽天レビュー収集 - Google Apps Script
 * Chrome拡張機能から送信されたレビューデータをスプレッドシートに保存する
 */

// スプレッドシートID（ここに自分のスプレッドシートIDを設定してください）
// スプレッドシートURLの https://docs.google.com/spreadsheets/d/XXXXX/edit の XXXXX 部分
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

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

  // デバッグ: 商品IDの一覧をログ出力
  const productIds = Object.keys(reviewsByProduct);
  Logger.log('=== saveReviewsByProduct デバッグ ===');
  Logger.log('受信レビュー数: ' + reviews.length);
  Logger.log('商品ID数: ' + productIds.length);
  Logger.log('商品ID一覧: ' + productIds.join(', '));

  // 各商品のシートに保存
  for (const productId in reviewsByProduct) {
    const productReviews = reviewsByProduct[productId];

    // シート名は商品管理番号をそのまま使用（31文字以内、特殊文字を除去）
    let sheetName = sanitizeSheetName(productId);

    // シートを取得または作成
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      // 「レビュー」シートが存在し、空（ヘッダーのみ）なら商品管理番号にリネームして使用
      const defaultSheet = ss.getSheetByName('レビュー');
      if (defaultSheet && defaultSheet.getLastRow() <= 1) {
        defaultSheet.setName(sheetName);
        sheet = defaultSheet;
        // ヘッダーを赤色で再設定（既存ヘッダーを上書き）
        addHeader(sheet);
      } else {
        sheet = ss.insertSheet(sheetName);
        addHeader(sheet);
      }
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
      review.helpfulCount || 0,
      review.shopReply || '',
      review.shopName || '',
      review.pageUrl || '',
      review.collectedAt || new Date().toISOString()
    ]);

    // データを追加
    if (rows.length > 0) {
      const lastRow = sheet.getLastRow();
      const dataRange = sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length);
      dataRange.setValues(rows);
      dataRange.setVerticalAlignment('middle');
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
    review.helpfulCount || 0,
    review.shopReply || '',
    review.shopName || '',
    review.pageUrl || '',
    review.collectedAt || new Date().toISOString()
  ]);

  // データを追加
  if (rows.length > 0) {
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length);
    dataRange.setValues(rows);
    dataRange.setVerticalAlignment('middle');
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
    '参考になった数',
    'ショップからの返信',
    'ショップ名',
    'レビュー掲載URL',
    '収集日時'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダーのスタイルを設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#BF0000');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setVerticalAlignment('middle');
  headerRange.setHorizontalAlignment('center');

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
  sheet.setColumnWidth(16, 100); // 参考になった数
  sheet.setColumnWidth(17, 300); // ショップからの返信
  sheet.setColumnWidth(18, 150); // ショップ名
  sheet.setColumnWidth(19, 250); // レビュー掲載URL
  sheet.setColumnWidth(20, 150); // 収集日時

  // 不要な列を削除（21列目以降）
  const maxColumns = sheet.getMaxColumns();
  if (maxColumns > 20) {
    sheet.deleteColumns(21, maxColumns - 20);
  }

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
 * スプレッドシートを初期化（すべてのシートを削除して空にする）
 * Apps Scriptエディタから手動で実行してください
 * 注意: すべてのレビューデータが削除されます！
 */
function initializeSpreadsheet() {
  const ui = SpreadsheetApp.getUi();

  // 確認ダイアログを表示
  const response = ui.alert(
    '⚠️ スプレッドシートの初期化',
    'すべてのシートとデータが削除されます。\nこの操作は取り消せません。\n\n本当に初期化しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('初期化をキャンセルしました');
    return;
  }

  const ss = getSpreadsheet();
  const sheets = ss.getSheets();

  // 「レビュー」シートを取得または作成
  let reviewSheet = ss.getSheetByName('レビュー');
  if (reviewSheet) {
    // 既存の「レビュー」シートをクリアして再利用
    reviewSheet.clear();
  } else {
    // なければ新規作成
    reviewSheet = ss.insertSheet('レビュー');
  }
  addHeader(reviewSheet);

  // 「レビュー」以外のすべてのシートを削除
  let deletedCount = 0;
  sheets.forEach(sheet => {
    if (sheet.getName() !== 'レビュー') {
      ss.deleteSheet(sheet);
      deletedCount++;
    }
  });

  ui.alert(
    '✅ 初期化完了',
    `${deletedCount}個のシートを削除しました。\nスプレッドシートは初期状態に戻りました。`,
    ui.ButtonSet.OK
  );

  Logger.log('スプレッドシートを初期化しました。削除したシート数: ' + deletedCount);
}

/**
 * 特定のシートを削除
 * @param {string} sheetName - 削除するシート名
 */
function deleteSheet(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    // 最後の1シートは削除できないため確認
    if (ss.getSheets().length <= 1) {
      Logger.log('最後のシートは削除できません');
      return false;
    }

    ss.deleteSheet(sheet);
    Logger.log('シート「' + sheetName + '」を削除しました');
    return true;
  } else {
    Logger.log('シート「' + sheetName + '」が見つかりません');
    return false;
  }
}

/**
 * 空のシートを一括削除（メンテナンス用）
 * ヘッダーのみのシートを削除
 */
function deleteEmptySheets() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();

  let deletedCount = 0;

  sheets.forEach(sheet => {
    // ヘッダー行のみ（1行以下）のシートを削除
    if (sheet.getLastRow() <= 1 && ss.getSheets().length > 1) {
      const name = sheet.getName();
      ss.deleteSheet(sheet);
      Logger.log('空のシート「' + name + '」を削除しました');
      deletedCount++;
    }
  });

  Logger.log('合計 ' + deletedCount + ' 個の空シートを削除しました');
}

/**
 * メニューを追加（スプレッドシートを開いたときに実行）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ レビュー管理')
    .addItem('📊 スプレッドシートを初期化', 'initializeSpreadsheet')
    .addItem('🔄 重複レビューを削除', 'removeDuplicates')
    .addToUi();
}

/**
 * 全シートのヘッダーを赤色に修正（メンテナンス用）
 */
function fixAllHeaders() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  let fixedCount = 0;

  sheets.forEach(sheet => {
    if (sheet.getLastRow() === 0) return;

    const lastCol = sheet.getLastColumn();
    if (lastCol === 0) return;

    const headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setBackground('#BF0000');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setVerticalAlignment('middle');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    // データがあれば上下中央揃え
    if (sheet.getLastRow() > 1) {
      const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol);
      dataRange.setVerticalAlignment('middle');
    }

    fixedCount++;
    Logger.log('シート「' + sheet.getName() + '」のヘッダーを修正しました');
  });

  const ui = SpreadsheetApp.getUi();
  ui.alert('✅ 完了', fixedCount + '個のシートのヘッダーを赤色に修正しました。', ui.ButtonSet.OK);
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
      headerRange.setBackground('#BF0000');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setVerticalAlignment('middle');
      headerRange.setHorizontalAlignment('center');
      sheet.setFrozenRows(1);

      // データの上下中央揃え
      if (uniqueRows.length > 0) {
        const dataRange = sheet.getRange(2, 1, uniqueRows.length, uniqueRows[0].length);
        dataRange.setVerticalAlignment('middle');
      }

      totalRemoved += removedCount;
      Logger.log(sheet.getName() + ': ' + removedCount + '件の重複を削除');
    }
  });

  Logger.log('合計: ' + totalRemoved + '件の重複を削除しました');
}

/**
 * デバッグ用：複数商品のレビュー保存テスト
 * GASエディタで実行してログを確認
 */
function debugMultipleProducts() {
  const testReviews = [
    {
      productId: 'product-A',
      productName: 'テスト商品A',
      productUrl: 'https://item.rakuten.co.jp/shop/product-A/',
      rating: 5,
      title: '商品Aのレビュー1',
      body: '商品Aのレビュー内容1',
      author: 'ユーザー1',
      reviewDate: '2024-01-01'
    },
    {
      productId: 'product-B',
      productName: 'テスト商品B',
      productUrl: 'https://item.rakuten.co.jp/shop/product-B/',
      rating: 4,
      title: '商品Bのレビュー1',
      body: '商品Bのレビュー内容1',
      author: 'ユーザー2',
      reviewDate: '2024-01-02'
    },
    {
      productId: 'product-A',
      productName: 'テスト商品A',
      productUrl: 'https://item.rakuten.co.jp/shop/product-A/',
      rating: 4,
      title: '商品Aのレビュー2',
      body: '商品Aのレビュー内容2',
      author: 'ユーザー3',
      reviewDate: '2024-01-03'
    },
    {
      productId: 'product-C',
      productName: 'テスト商品C',
      productUrl: 'https://item.rakuten.co.jp/shop/product-C/',
      rating: 3,
      title: '商品Cのレビュー1',
      body: '商品Cのレビュー内容1',
      author: 'ユーザー4',
      reviewDate: '2024-01-04'
    }
  ];

  const ss = getSpreadsheet();
  const savedCount = saveReviewsByProduct(ss, testReviews);
  Logger.log('保存件数: ' + savedCount);
  Logger.log('シート一覧: ' + ss.getSheets().map(s => s.getName()).join(', '));
}

/**
 * デバッグ用：初期化テスト（UIなし）
 * GASエディタで実行してログを確認
 */
function debugInitialize() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();

  Logger.log('=== デバッグ開始 ===');
  Logger.log('シート数: ' + sheets.length);

  sheets.forEach(sheet => {
    Logger.log('シート名: ' + sheet.getName() + ', 行数: ' + sheet.getLastRow());
  });

  // 「レビュー」シートを確認
  const reviewSheet = ss.getSheetByName('レビュー');
  Logger.log('「レビュー」シート存在: ' + (reviewSheet !== null));

  if (reviewSheet) {
    Logger.log('「レビュー」シートをクリアします');
    reviewSheet.clear();
    addHeader(reviewSheet);
    Logger.log('ヘッダー追加完了');
  } else {
    Logger.log('「レビュー」シートを新規作成します');
    const newSheet = ss.insertSheet('レビュー');
    addHeader(newSheet);
    Logger.log('新規作成完了');
  }

  // 他のシートを削除
  let deletedCount = 0;
  const currentSheets = ss.getSheets();
  currentSheets.forEach(sheet => {
    const name = sheet.getName();
    if (name !== 'レビュー') {
      Logger.log('削除: ' + name);
      ss.deleteSheet(sheet);
      deletedCount++;
    }
  });

  Logger.log('削除したシート数: ' + deletedCount);
  Logger.log('=== デバッグ完了 ===');
}
