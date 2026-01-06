/**
 * バックグラウンドサービスワーカー
 * レビューデータの保存、GASへの送信、CSVダウンロード、キュー管理を処理
 */

// メッセージリスナー
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  switch (message.action) {
    case 'saveReviews':
      handleSaveReviews(message.reviews)
        .then(() => sendResponse({ success: true }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'downloadCSV':
      handleDownloadCSV()
        .then(() => sendResponse({ success: true }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'collectionComplete':
      handleCollectionComplete();
      sendResponse({ success: true });
      break;

    case 'startQueueCollection':
      startQueueCollection()
        .then(() => sendResponse({ success: true }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'fetchRanking':
      fetchRankingProducts(message.url, message.count)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'updateProgress':
    case 'log':
      forwardToAll(message);
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
        totalPages: 0,
        reviews: [],
        logs: []
      };

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
    mode: 'no-cors',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      reviews: reviews,
      separateSheets: separateSheets,
      timestamp: new Date().toISOString()
    })
  });

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
        const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
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
  const headers = [
    'レビュー日', '商品管理番号', '商品名', '商品URL', '評価', 'タイトル', '本文',
    '投稿者', '年代', '性別', '注文日', 'バリエーション', '用途', '贈り先',
    '購入回数', '購入情報', '参考になった数', 'ショップ名', 'ページURL', '収集日時'
  ];

  const rows = reviews.map(review => [
    review.reviewDate || '', review.productId || '', review.productName || '',
    review.productUrl || '', review.rating || '', review.title || '', review.body || '',
    review.author || '', review.age || '', review.gender || '', review.orderDate || '',
    review.variation || '', review.usage || '', review.recipient || '',
    review.purchaseCount || '', review.purchaseInfo || '', review.helpfulCount || 0,
    review.shopName || '', review.pageUrl || '', review.collectedAt || ''
  ]);

  const escapeCSV = (value) => {
    if (value === null || value === undefined) return '';
    const str = String(value);
    if (str.includes('"') || str.includes(',') || str.includes('\n') || str.includes('\r')) {
      return '"' + str.replace(/"/g, '""') + '"';
    }
    return str;
  };

  return [
    headers.map(escapeCSV).join(','),
    ...rows.map(row => row.map(escapeCSV).join(','))
  ].join('\r\n');
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
  const result = await chrome.storage.local.get(['collectionState', 'queue']);
  const state = result.collectionState || {};
  const queue = result.queue || [];

  state.isRunning = false;
  await chrome.storage.local.set({ collectionState: state });

  // ポップアップに完了を通知
  forwardToAll({
    action: 'collectionComplete',
    state: state
  });

  // キューに次の商品がある場合、次を処理
  if (queue.length > 0) {
    log('次の商品の収集を開始します...');
    setTimeout(() => {
      processNextInQueue();
    }, 3000);
  }
}

/**
 * キュー一括収集を開始
 */
async function startQueueCollection() {
  const result = await chrome.storage.local.get(['queue']);
  const queue = result.queue || [];

  if (queue.length === 0) {
    throw new Error('キューが空です');
  }

  log(`${queue.length}件のキュー収集を開始します`, 'success');
  await processNextInQueue();
}

/**
 * キューの次の商品を処理
 */
async function processNextInQueue() {
  const result = await chrome.storage.local.get(['queue']);
  const queue = result.queue || [];

  if (queue.length === 0) {
    log('キューの処理が完了しました', 'success');
    return;
  }

  const nextItem = queue[0];
  queue.shift();
  await chrome.storage.local.set({ queue });

  forwardToAll({ action: 'queueUpdated' });

  log(`収集中: ${nextItem.title || nextItem.url}`);

  // 新しいタブで商品ページを開いて収集開始
  const tab = await chrome.tabs.create({ url: nextItem.url, active: true });

  // 収集状態をセット
  await chrome.storage.local.set({
    collectionState: {
      isRunning: true,
      reviewCount: 0,
      pageCount: 0,
      totalPages: 0,
      reviews: [],
      logs: []
    }
  });

  // ページ読み込み完了後に収集開始を指示
  chrome.tabs.onUpdated.addListener(function listener(tabId, info) {
    if (tabId === tab.id && info.status === 'complete') {
      chrome.tabs.onUpdated.removeListener(listener);
      setTimeout(() => {
        chrome.tabs.sendMessage(tab.id, { action: 'startCollection' }).catch(() => {});
      }, 2000);
    }
  });
}

/**
 * ランキングから商品を取得してキューに追加
 */
async function fetchRankingProducts(url, count) {
  try {
    // ランキングページをfetchで取得
    const response = await fetch(url);
    const html = await response.text();

    // HTMLをパースして商品URLを抽出
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');

    const products = [];

    // ランキングの商品リンクを探す
    const productLinks = doc.querySelectorAll('a[href*="item.rakuten.co.jp"]');

    const seenUrls = new Set();

    for (const link of productLinks) {
      if (products.length >= count) break;

      let href = link.href;

      // 相対URLの場合は絶対URLに変換
      if (href.startsWith('/')) {
        href = 'https://ranking.rakuten.co.jp' + href;
      }

      // item.rakuten.co.jpを含むURLのみ
      if (!href.includes('item.rakuten.co.jp')) continue;

      // クエリパラメータを除去してURLを正規化
      const url = new URL(href);
      const cleanUrl = `${url.origin}${url.pathname}`;

      // 重複チェック
      if (seenUrls.has(cleanUrl)) continue;
      seenUrls.add(cleanUrl);

      // 商品名を取得
      let title = link.textContent.trim();
      if (!title || title.length < 3) {
        const img = link.querySelector('img');
        title = img ? img.alt : '商品';
      }

      products.push({
        url: cleanUrl,
        title: title.substring(0, 100),
        addedAt: new Date().toISOString()
      });
    }

    if (products.length === 0) {
      return { success: false, error: '商品が見つかりませんでした' };
    }

    // キューに追加
    const result = await chrome.storage.local.get(['queue']);
    const queue = result.queue || [];

    let addedCount = 0;
    for (const product of products) {
      const exists = queue.some(item => item.url === product.url);
      if (!exists) {
        queue.push(product);
        addedCount++;
      }
    }

    await chrome.storage.local.set({ queue });

    forwardToAll({ action: 'queueUpdated' });

    return { success: true, addedCount };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * すべてのページにメッセージを転送
 */
function forwardToAll(message) {
  chrome.runtime.sendMessage(message).catch(() => {});
}

/**
 * ログを送信
 */
function log(text, type = '') {
  console.log(`[楽天レビュー収集] ${text}`);

  // ログを保存
  chrome.storage.local.get(['logs'], (result) => {
    const logs = result.logs || [];
    const time = new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    logs.push({ time, text, type });

    if (logs.length > 100) {
      logs.splice(0, logs.length - 100);
    }

    chrome.storage.local.set({ logs });
  });

  forwardToAll({
    action: 'log',
    text: text,
    type: type
  });
}

// 拡張機能インストール時の初期化
chrome.runtime.onInstalled.addListener(() => {
  console.log('楽天レビュー収集拡張機能がインストールされました');

  chrome.storage.local.set({
    collectionState: {
      isRunning: false,
      reviewCount: 0,
      pageCount: 0,
      totalPages: 0,
      reviews: [],
      logs: []
    },
    queue: [],
    logs: []
  });
});
