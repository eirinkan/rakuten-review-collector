/**
 * バックグラウンドサービスワーカー
 * レビューデータの保存、GASへの送信、CSVダウンロード、キュー管理を処理
 */

// アクティブな収集タブを追跡
let activeCollectionTabs = new Set();

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
      handleCollectionComplete(sender.tab?.id);
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
async function handleCollectionComplete(tabId) {
  // アクティブタブから削除
  if (tabId) {
    activeCollectionTabs.delete(tabId);
    // タブごとの状態をクリーンアップ
    const stateKey = `collectionState_${tabId}`;
    await chrome.storage.local.remove(stateKey);
  }

  // 状態を更新
  const result = await chrome.storage.local.get(['collectionState', 'queue']);
  const state = result.collectionState || {};
  const queue = result.queue || [];

  state.isRunning = activeCollectionTabs.size > 0;
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
  } else if (activeCollectionTabs.size === 0) {
    log('すべての収集が完了しました', 'success');
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

  // 設定から同時収集数を取得
  const settings = await chrome.storage.sync.get(['maxConcurrent']);
  const maxConcurrent = settings.maxConcurrent || 1;

  log(`${queue.length}件のキュー収集を開始します（同時${maxConcurrent}件）`, 'success');

  // 同時収集数分だけ開始
  for (let i = 0; i < maxConcurrent && i < queue.length; i++) {
    await processNextInQueue();
  }
}

/**
 * キューの次の商品を処理
 */
async function processNextInQueue() {
  // 設定から同時収集数を取得
  const settings = await chrome.storage.sync.get(['maxConcurrent']);
  const maxConcurrent = settings.maxConcurrent || 1;

  // 現在のアクティブタブ数をチェック
  if (activeCollectionTabs.size >= maxConcurrent) {
    log(`同時収集数上限（${maxConcurrent}）に達しています`, 'info');
    return;
  }

  const result = await chrome.storage.local.get(['queue']);
  const queue = result.queue || [];

  if (queue.length === 0) {
    if (activeCollectionTabs.size === 0) {
      log('キューの処理が完了しました', 'success');
    }
    return;
  }

  const nextItem = queue[0];
  queue.shift();
  await chrome.storage.local.set({ queue });

  forwardToAll({ action: 'queueUpdated' });

  log(`収集中: ${nextItem.title || nextItem.url}`);

  // 新しいタブで商品ページを開いて収集開始（バックグラウンドで開く）
  const tab = await chrome.tabs.create({ url: nextItem.url, active: false });

  // アクティブタブとして追跡
  activeCollectionTabs.add(tab.id);

  // 収集状態をセット（タブIDごとに管理）
  const stateKey = `collectionState_${tab.id}`;
  await chrome.storage.local.set({
    [stateKey]: {
      isRunning: true,
      reviewCount: 0,
      pageCount: 0,
      totalPages: 0,
      reviews: [],
      tabId: tab.id
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
 * Service WorkerではDOMParserが使えないため、正規表現でパース
 */
async function fetchRankingProducts(url, count) {
  try {
    // ランキングページをfetchで取得
    const response = await fetch(url);
    const html = await response.text();

    const products = [];
    const seenUrls = new Set();

    // 正規表現でitem.rakuten.co.jpへのリンクを抽出
    // パターン: href="https://item.rakuten.co.jp/ショップ名/商品ID/"
    const linkPattern = /href="(https?:\/\/item\.rakuten\.co\.jp\/[^"]+)"/g;
    let match;

    while ((match = linkPattern.exec(html)) !== null && products.length < count) {
      let href = match[1];

      // クエリパラメータを除去してURLを正規化
      try {
        const urlObj = new URL(href);
        const cleanUrl = `${urlObj.origin}${urlObj.pathname}`;

        // 重複チェック
        if (seenUrls.has(cleanUrl)) continue;
        seenUrls.add(cleanUrl);

        // URLからショップ名と商品IDを取得してタイトルにする
        const pathMatch = cleanUrl.match(/item\.rakuten\.co\.jp\/([^\/]+)\/([^\/]+)/);
        let title = '商品';
        if (pathMatch) {
          title = `${pathMatch[1]} - ${pathMatch[2]}`;
        }

        products.push({
          url: cleanUrl,
          title: title.substring(0, 100),
          addedAt: new Date().toISOString()
        });
      } catch (e) {
        // URL解析エラーは無視
        continue;
      }
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

// タブが閉じられた時のクリーンアップ
chrome.tabs.onRemoved.addListener((tabId) => {
  if (activeCollectionTabs.has(tabId)) {
    activeCollectionTabs.delete(tabId);
    log(`タブ ${tabId} が閉じられました。収集を中断しました`, 'error');

    // タブごとの状態をクリーンアップ
    const stateKey = `collectionState_${tabId}`;
    chrome.storage.local.remove(stateKey);

    // キューに次がある場合は処理を続行
    setTimeout(() => {
      processNextInQueue();
    }, 1000);
  }
});
