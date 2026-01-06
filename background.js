/**
 * バックグラウンドサービスワーカー
 * レビューデータの保存、GASへの送信、CSVダウンロードを処理
 */

// メッセージリスナー
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  switch (message.action) {
    case 'saveReviews':
      handleSaveReviews(message.reviews)
        .then(() => sendResponse({ success: true }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true; // 非同期レスポンスのため

    case 'downloadCSV':
      handleDownloadCSV()
        .then(() => sendResponse({ success: true }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'collectionComplete':
      handleCollectionComplete();
      break;

    case 'updateProgress':
    case 'log':
      // ポップアップにそのまま転送（ポップアップが開いている場合）
      forwardToPopup(message);
      break;
  }
});

/**
 * レビューを保存
 */
async function handleSaveReviews(reviews) {
  if (!reviews || reviews.length === 0) {
    return;
  }

  // ローカルストレージに保存
  await saveToLocalStorage(reviews);

  // GAS URLが設定されている場合はスプレッドシートにも送信
  const { gasUrl, separateSheets } = await chrome.storage.sync.get(['gasUrl', 'separateSheets']);

  if (gasUrl) {
    try {
      await sendToGas(gasUrl, reviews, separateSheets !== false);
      log('スプレッドシートに保存しました', 'success');
    } catch (error) {
      log(`スプレッドシートへの保存に失敗: ${error.message}`, 'error');
      // エラーでもローカルには保存済みなので続行
    }
  }
}

/**
 * ローカルストレージに保存
 */
async function saveToLocalStorage(reviews) {
  return new Promise((resolve, reject) => {
    chrome.storage.local.get(['collectionState'], (result) => {
      const state = result.collectionState || {
        isRunning: true,
        reviewCount: 0,
        pageCount: 0,
        reviews: [],
        logs: []
      };

      // 既存のレビューに追加
      state.reviews = (state.reviews || []).concat(reviews);

      chrome.storage.local.set({ collectionState: state }, () => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          resolve();
        }
      });
    });
  });
}

/**
 * GASにデータを送信
 */
async function sendToGas(gasUrl, reviews, separateSheets = true) {
  const response = await fetch(gasUrl, {
    method: 'POST',
    mode: 'no-cors', // CORSを回避
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      reviews: reviews,
      separateSheets: separateSheets,
      timestamp: new Date().toISOString()
    })
  });

  // no-corsモードではレスポンスを読めないので、成功とみなす
  return true;
}

/**
 * CSVダウンロード
 */
async function handleDownloadCSV() {
  return new Promise((resolve, reject) => {
    chrome.storage.local.get(['collectionState'], async (result) => {
      const state = result.collectionState;

      if (!state || !state.reviews || state.reviews.length === 0) {
        reject(new Error('ダウンロードするデータがありません'));
        return;
      }

      try {
        const csv = convertToCSV(state.reviews);
        const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' }); // BOM付きUTF-8
        const url = URL.createObjectURL(blob);

        const filename = `rakuten_reviews_${formatDate(new Date())}.csv`;

        await chrome.downloads.download({
          url: url,
          filename: filename,
          saveAs: true
        });

        resolve();
      } catch (error) {
        reject(error);
      }
    });
  });
}

/**
 * レビューデータをCSV形式に変換
 */
function convertToCSV(reviews) {
  // ヘッダー
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

  // データ行
  const rows = reviews.map(review => [
    review.collectedAt || '',
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

  // CSVエスケープ処理
  const escapeCSV = (value) => {
    if (value === null || value === undefined) {
      return '';
    }
    const str = String(value);
    // ダブルクォート、カンマ、改行を含む場合はクォートで囲む
    if (str.includes('"') || str.includes(',') || str.includes('\n') || str.includes('\r')) {
      return '"' + str.replace(/"/g, '""') + '"';
    }
    return str;
  };

  // CSV文字列を生成
  const csvContent = [
    headers.map(escapeCSV).join(','),
    ...rows.map(row => row.map(escapeCSV).join(','))
  ].join('\r\n');

  return csvContent;
}

/**
 * 日付をフォーマット
 */
function formatDate(date) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}_${pad(date.getHours())}${pad(date.getMinutes())}`;
}

/**
 * 収集完了時の処理
 */
async function handleCollectionComplete() {
  // 状態を更新
  chrome.storage.local.get(['collectionState'], (result) => {
    const state = result.collectionState || {};
    state.isRunning = false;

    chrome.storage.local.set({ collectionState: state }, () => {
      // ポップアップに完了を通知
      forwardToPopup({
        action: 'collectionComplete',
        state: state
      });
    });
  });
}

/**
 * ポップアップにメッセージを転送
 */
function forwardToPopup(message) {
  // ポップアップが開いているかは不明なので、エラーは無視
  chrome.runtime.sendMessage(message).catch(() => {
    // ポップアップが閉じている場合はエラーになるが無視
  });
}

/**
 * ログをポップアップに送信
 */
function log(text, type = '') {
  console.log(`[楽天レビュー収集] ${text}`);
  forwardToPopup({
    action: 'log',
    text: text,
    type: type
  });
}

// 拡張機能インストール時の初期化
chrome.runtime.onInstalled.addListener(() => {
  console.log('楽天レビュー収集拡張機能がインストールされました');

  // 初期状態をセット
  chrome.storage.local.set({
    collectionState: {
      isRunning: false,
      reviewCount: 0,
      pageCount: 0,
      reviews: [],
      logs: []
    }
  });
});
