/**
 * バックグラウンドサービスワーカー
 * レビューデータの保存、CSVダウンロード、キュー管理を処理
 * 楽天市場・Amazon両対応
 */

// アクティブな収集タブを追跡
let activeCollectionTabs = new Set();
// 収集用ウィンドウのID
let collectionWindowId = null;

// ===== activeCollectionTabs の永続化 =====
/**
 * activeCollectionTabsをchrome.storage.localに保存
 */
async function persistActiveCollectionTabs() {
  const tabIds = Array.from(activeCollectionTabs);
  await chrome.storage.local.set({ activeCollectionTabIds: tabIds });
  console.log('[persistActiveCollectionTabs] 保存:', tabIds);
}

/**
 * chrome.storage.localからactiveCollectionTabsを復元
 */
async function restoreActiveCollectionTabs() {
  const result = await chrome.storage.local.get(['activeCollectionTabIds']);
  const tabIds = result.activeCollectionTabIds || [];

  // 存在するタブのみを復元（閉じられたタブは除外）
  for (const tabId of tabIds) {
    try {
      await chrome.tabs.get(tabId);
      activeCollectionTabs.add(tabId);
    } catch (e) {
      // タブが存在しない場合は無視
      console.log('[restoreActiveCollectionTabs] タブが存在しない:', tabId);
    }
  }

  // 存在するタブのみを再保存
  await persistActiveCollectionTabs();
  console.log('[restoreActiveCollectionTabs] 復元完了:', Array.from(activeCollectionTabs));
}

// Service Worker起動時に復元
restoreActiveCollectionTabs();
// タブごとのスプレッドシートURL（定期収集用）
const tabSpreadsheetUrls = new Map();
// タブごとの最後のresumeCollection送信時刻（重複送信防止）
const lastResumeSentTime = new Map();
// 重複送信防止の閾値（ミリ秒）
const RESUME_DEBOUNCE_MS = 5000;
// 最後に使用したタブID（キュー収集でタブを再利用するため）
let lastUsedTabId = null;

// キュー処理用のアラーム名
const QUEUE_NEXT_ALARM_NAME = 'processNextInQueue';

// ===== Amazonボット対策: レート制限 =====
const AMAZON_DAILY_LIMIT = 100;  // 1日100ページまで（警戒圏）

/**
 * Amazonのレート制限をチェック・カウント
 * @returns {Promise<{allowed: boolean, remaining: number, count: number}>}
 */
async function checkAmazonRateLimit() {
  const data = await chrome.storage.local.get(['amazonRateLimit']);
  const today = new Date().toDateString();
  let rateLimit = data.amazonRateLimit || { count: 0, date: today };

  // 日付が変わったらリセット
  if (rateLimit.date !== today) {
    rateLimit = { count: 0, date: today };
  }

  // 制限チェック（カウントは増やさない）
  if (rateLimit.count >= AMAZON_DAILY_LIMIT) {
    return { allowed: false, remaining: 0, count: rateLimit.count };
  }

  return {
    allowed: true,
    remaining: AMAZON_DAILY_LIMIT - rateLimit.count,
    count: rateLimit.count
  };
}

/**
 * Amazonのレート制限カウントを増加
 */
async function incrementAmazonRateLimit() {
  const data = await chrome.storage.local.get(['amazonRateLimit']);
  const today = new Date().toDateString();
  let rateLimit = data.amazonRateLimit || { count: 0, date: today };

  // 日付が変わったらリセット
  if (rateLimit.date !== today) {
    rateLimit = { count: 0, date: today };
  }

  rateLimit.count++;
  await chrome.storage.local.set({ amazonRateLimit: rateLimit });

  return {
    count: rateLimit.count,
    remaining: AMAZON_DAILY_LIMIT - rateLimit.count
  };
}

/**
 * Amazonのレート制限情報を取得（カウントせずに現在の状態のみ）
 */
async function getAmazonRateLimitInfo() {
  const data = await chrome.storage.local.get(['amazonRateLimit']);
  const today = new Date().toDateString();
  let rateLimit = data.amazonRateLimit || { count: 0, date: today };

  // 日付が変わったらリセット
  if (rateLimit.date !== today) {
    rateLimit = { count: 0, date: today };
    await chrome.storage.local.set({ amazonRateLimit: rateLimit });
  }

  return {
    count: rateLimit.count,
    remaining: AMAZON_DAILY_LIMIT - rateLimit.count,
    limit: AMAZON_DAILY_LIMIT
  };
}

// Amazonページ遷移時の収集再開リスナー（グローバル）
// 注意: Service Workerが再起動するとactiveCollectionTabsがリセットされるため、
// collectionState.isRunningをメインの条件として使用
//
// 重要: このリスナーは以下の2つのケースを処理する
// 1. レビューページ間の遷移（ページ送り、フィルター切り替え）
// 2. キュー処理からの開始（商品ページまたはレビューページ）
chrome.tabs.onUpdated.addListener(async (tabId, changeInfo, tab) => {
  // ページ読み込み完了時のみ処理
  if (changeInfo.status !== 'complete') return;

  const url = tab.url || '';

  // Amazonページかどうか確認（商品ページまたはレビューページ）
  const isAmazonProductPage = url.includes('amazon.co.jp/dp/') || url.includes('amazon.co.jp/gp/product/');
  const isAmazonReviewPage = url.includes('amazon.co.jp/product-reviews/');
  const isAmazonPage = isAmazonProductPage || isAmazonReviewPage;

  if (!isAmazonPage) return;

  // 収集状態を確認（Service Worker再起動後もストレージは永続化される）
  const result = await chrome.storage.local.get(['collectionState']);
  const state = result.collectionState;

  console.log('[background] Amazonページ検出:', {
    tabId,
    url: url.substring(0, 100),
    isProductPage: isAmazonProductPage,
    isReviewPage: isAmazonReviewPage,
    isActiveTab: activeCollectionTabs.has(tabId),
    activeTabsCount: activeCollectionTabs.size,
    state: state ? {
      isRunning: state.isRunning,
      source: state.source,
      startedFromQueue: state.startedFromQueue,
      lastProcessedPage: state.lastProcessedPage
    } : null
  });

  // 収集中でAmazonの場合のみ処理
  // 重要: activeCollectionTabs.has(tabId) を追加して、実際に収集中のタブのみを対象にする
  // これにより、手動で開いたAmazonレビューページに干渉しない
  // さらに: activeTabId でもチェックして、別タブでの操作を完全にブロック
  if (state && state.isRunning && state.source === 'amazon' && activeCollectionTabs.has(tabId)) {
    // activeTabIdが設定されていて、このタブと異なる場合はスキップ
    if (state.activeTabId && state.activeTabId !== tabId) {
      console.log('[background] activeTabId不一致のためスキップ:', {
        activeTabId: state.activeTabId,
        thisTabId: tabId
      });
      return;
    }

    // キュー処理から開始された場合の特別処理（商品ページでもレビューページでも発火）
    // processNextInQueue内のローカルリスナーを使わず、このグローバルリスナーで統一処理
    if (state.startedFromQueue) {
      console.log('[background] キュー処理からの開始 - startCollectionを送信', {
        isProductPage: isAmazonProductPage,
        isReviewPage: isAmazonReviewPage
      });
      // フラグをクリア（次回のフィルター遷移等では通常処理を行う）
      state.startedFromQueue = false;
      await chrome.storage.local.set({ collectionState: state });

      // 重複送信防止
      const now = Date.now();
      lastResumeSentTime.set(tabId, now);

      // startCollectionを送信
      // バックグラウンドタブでもcontent scriptが確実に初期化されるように、リトライ機能付きで送信
      sendMessageWithRetry(tabId, {
        action: 'startCollection',
        incrementalOnly: state.incrementalOnly || false,
        lastCollectedDate: state.lastCollectedDate || null,
        queueName: state.queueName || null,
        productId: state.productId || null
      }, 'startCollection（キュー開始）');
      return;
    }

    // 以下はレビューページでのみ処理（ページ遷移、フィルター切り替え）
    if (!isAmazonReviewPage) return;

    // URLからページ番号を取得
    const urlObj = new URL(url);
    const currentPage = parseInt(urlObj.searchParams.get('pageNumber') || '1', 10);
    const lastPage = state.lastProcessedPage || 0;

    // フィルター遷移フラグをチェック
    const filterTransitionPending = state.filterTransitionPending === true;

    // フィルター遷移中の場合は無条件で再開処理を実行（デバウンスをスキップ）
    if (filterTransitionPending) {
      console.log('[background] フィルター遷移を検出、収集を再開します:', {
        currentPage,
        filterIndex: state.currentStarFilterIndex
      });
      // フラグをクリア
      state.filterTransitionPending = false;
      await chrome.storage.local.set({ collectionState: state });

      // デバウンスをスキップして直接resumeCollectionを送信
      sendMessageWithRetry(tabId, {
        action: 'resumeCollection',
        incrementalOnly: state.incrementalOnly || false,
        lastCollectedDate: state.lastCollectedDate || null,
        queueName: state.queueName || null,
        productId: state.productId || null
      }, 'resumeCollection（フィルター遷移）');
      return;
    }

    // 同じページなら重複実行しない
    if (currentPage <= lastPage) {
      console.log('[background] 同じまたは前のページのため再開スキップ:', { currentPage, lastPage });
      return;
    }

    // 重複送信防止: 同じタブに対して5秒以内の再送信を防ぐ
    const now = Date.now();
    const lastSent = lastResumeSentTime.get(tabId) || 0;
    if (now - lastSent < RESUME_DEBOUNCE_MS) {
      console.log('[background] 重複送信防止: 前回送信から5秒未満のためスキップ:', {
        tabId,
        elapsed: now - lastSent,
        threshold: RESUME_DEBOUNCE_MS
      });
      return;
    }
    lastResumeSentTime.set(tabId, now);

    console.log('[background] Amazonレビューページ遷移検出、収集再開メッセージを送信:', { currentPage, lastPage });

    // 注意: 171行目の条件でactiveCollectionTabs.has(tabId)をチェック済みなので、
    // ここに到達した時点でタブは既にactiveCollectionTabsに含まれている
    // 以下のチェックは防御的プログラミングとして残す
    if (!activeCollectionTabs.has(tabId)) {
      // 本来ここには到達しないはず（171行目でフィルタされる）
      console.log('[background] 警告: 収集中タブではないためスキップ:', tabId);
      return;
    }

    // バックグラウンドタブでもcontent scriptが確実に初期化されるように、リトライ機能付きで送信
    console.log('[background] resumeCollectionメッセージ送信準備...', {
      queueName: state.queueName,
      productId: state.productId
    });
    sendMessageWithRetry(tabId, {
      action: 'resumeCollection',
      incrementalOnly: state.incrementalOnly || false,
      lastCollectedDate: state.lastCollectedDate || null,
      queueName: state.queueName || null,
      productId: state.productId || null
    }, 'resumeCollection');
  }
});

/**
 * タブにメッセージを送信（リトライ機能付き）
 * バックグラウンドタブでは content script の初期化が遅れる可能性があるため、
 * 失敗した場合は最大10回までリトライする
 *
 * 重要: setTimeoutはバックグラウンドタブで最小1秒にスロットリングされるため、
 * setIntervalベースで500msごとにポーリングしてリトライする
 *
 * @param {number} tabId - タブID
 * @param {Object} message - 送信するメッセージ
 * @param {string} description - ログ出力用の説明
 */
/**
 * タブにメッセージを送信（即座にリトライ、最大10回）
 * Service WorkerではsetTimeout/setIntervalがスロットリングされるため、
 * 再帰的に即座にリトライする方式を採用
 *
 * 注意: content scriptが準備できていない場合（ページ読み込み中など）は
 * 失敗するため、適切な遅延後に呼び出すこと
 */
async function sendMessageWithRetry(tabId, message, description = '', retryCount = 0) {
  const MAX_RETRIES = 10;

  try {
    console.log(`[background] ${description}メッセージ送信（試行${retryCount + 1}/${MAX_RETRIES}）`);
    const response = await chrome.tabs.sendMessage(tabId, message);
    console.log(`[background] ${description}応答:`, response);
    return response;
  } catch (err) {
    console.log(`[background] ${description}送信エラー（試行${retryCount + 1}）:`, err.message);

    if (retryCount < MAX_RETRIES - 1) {
      // 即座にリトライ（Service Workerでは遅延なしで再帰呼び出し）
      // ただし、少しの遅延を入れるためにPromise.resolve()を挟む
      await Promise.resolve();
      return sendMessageWithRetry(tabId, message, description, retryCount + 1);
    } else {
      console.log(`[background] ${description}最大リトライ回数に達しました`);
      throw err;
    }
  }
}

/**
 * URLが楽天かAmazonかを判定
 * @param {string} url - URL
 * @returns {string} 'rakuten' | 'amazon' | 'unknown'
 */
function detectSource(url) {
  if (!url) return 'unknown';
  if (url.includes('rakuten.co.jp')) return 'rakuten';
  if (url.includes('amazon.co.jp')) return 'amazon';
  return 'unknown';
}

/**
 * URLからASINを抽出（Amazon用）
 * @param {string} url - Amazon URL
 * @returns {string} ASIN または空文字
 */
function extractASIN(url) {
  if (!url) return '';
  // /dp/ASIN パターン
  const dpMatch = url.match(/\/dp\/([A-Z0-9]{10})/i);
  if (dpMatch) return dpMatch[1].toUpperCase();
  // /product-reviews/ASIN パターン
  const reviewMatch = url.match(/\/product-reviews\/([A-Z0-9]{10})/i);
  if (reviewMatch) return reviewMatch[1].toUpperCase();
  // /gp/product/ASIN パターン
  const gpMatch = url.match(/\/gp\/product\/([A-Z0-9]{10})/i);
  if (gpMatch) return gpMatch[1].toUpperCase();
  return '';
}

/**
 * ===== 認証機能 =====
 */

/**
 * Googleログインを実行
 * @returns {Promise<Object>} ユーザー情報 {email, name, picture}
 */
async function googleLogin() {
  try {
    // OAuth2トークンを取得
    const token = await new Promise((resolve, reject) => {
      chrome.identity.getAuthToken({ interactive: true }, (token) => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          resolve(token);
        }
      });
    });

    // ユーザー情報を取得
    const response = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!response.ok) {
      throw new Error('ユーザー情報の取得に失敗しました');
    }

    const userInfo = await response.json();

    return {
      email: userInfo.email,
      name: userInfo.name || userInfo.email,
      picture: userInfo.picture || null
    };
  } catch (error) {
    console.error('Googleログインエラー:', error);
    throw error;
  }
}


/**
 * 認証フロー全体を実行（ログイン + 保存）
 * @returns {Promise<Object>} 認証結果
 */
async function authenticate() {
  try {
    // 1. Googleログイン
    const userInfo = await googleLogin();

    // 2. 認証情報を保存
    await chrome.storage.local.set({
      authUser: {
        email: userInfo.email,
        name: userInfo.name,
        picture: userInfo.picture,
        authenticatedAt: Date.now()
      }
    });

    return {
      success: true,
      authenticated: true,
      user: userInfo,
      message: '認証成功'
    };
  } catch (error) {
    console.error('認証エラー:', error);
    return {
      success: false,
      authenticated: false,
      message: error.message
    };
  }
}

/**
 * 認証状態を確認
 * @returns {Promise<Object>} 認証状態
 */
async function checkAuthStatus() {
  const data = await chrome.storage.local.get(['authUser']);
  if (data.authUser && data.authUser.email) {
    return {
      authenticated: true,
      user: data.authUser
    };
  }
  return {
    authenticated: false,
    user: null
  };
}

/**
 * ログアウト（トークン無効化 + 保存データ削除）
 */
async function logout() {
  try {
    await revokeToken();
    await chrome.storage.local.remove(['authUser']);
    return { success: true, message: 'ログアウトしました' };
  } catch (error) {
    console.error('ログアウトエラー:', error);
    return { success: false, message: error.message };
  }
}

/**
 * OAuthトークンを無効化
 */
async function revokeToken() {
  return new Promise((resolve) => {
    chrome.identity.getAuthToken({ interactive: false }, (token) => {
      if (token) {
        chrome.identity.removeCachedAuthToken({ token }, () => {
          // トークンをGoogleサーバーからも無効化
          fetch(`https://accounts.google.com/o/oauth2/revoke?token=${token}`)
            .finally(() => resolve());
        });
      } else {
        resolve();
      }
    });
  });
}

/**
 * OAuthトークンを取得
 */
async function getAuthToken() {
  return new Promise((resolve) => {
    chrome.identity.getAuthToken({ interactive: false }, (token) => {
      if (chrome.runtime.lastError || !token) {
        resolve(null);
      } else {
        resolve(token);
      }
    });
  });
}

/**
 * スプレッドシートのタイトルを取得
 */
async function getSpreadsheetTitle(spreadsheetId) {
  const token = await getAuthToken();
  if (!token) {
    throw new Error('認証が必要です');
  }

  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=properties.title`,
    {
      headers: { 'Authorization': `Bearer ${token}` }
    }
  );

  if (!response.ok) {
    if (response.status === 404) {
      throw new Error('スプレッドシートが見つかりません');
    }
    if (response.status === 403) {
      throw new Error('アクセス権限がありません');
    }
    throw new Error('タイトルの取得に失敗しました');
  }

  const data = await response.json();
  return data.properties?.title || '';
}

/**
 * ===== 認証機能ここまで =====
 */

/**
 * デスクトップ通知を表示
 */
async function showNotification(title, message, productId = null) {
  // 通知設定を確認
  const settings = await chrome.storage.sync.get(['enableNotification', 'notifyPerProduct']);

  // 通知が無効の場合は何もしない
  if (settings.enableNotification === false) {
    return;
  }

  // 商品ごとの通知で、設定がOFFの場合は何もしない
  if (productId && settings.notifyPerProduct !== true) {
    return;
  }

  // 通知を作成
  const notificationId = `rakuten-review-${Date.now()}`;
  chrome.notifications.create(notificationId, {
    type: 'basic',
    iconUrl: 'icons/icon128.png',
    title: title,
    message: message,
    priority: 2
  });
}

// メッセージリスナー
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  switch (message.action) {
    // ===== 認証関連 =====
    case 'authenticate':
      authenticate()
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, message: error.message }));
      return true;

    case 'checkAuthStatus':
      checkAuthStatus()
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ authenticated: false, error: error.message }));
      return true;

    case 'logout':
      logout()
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, message: error.message }));
      return true;

    case 'getSpreadsheetTitle':
      getSpreadsheetTitle(message.spreadsheetId)
        .then(title => sendResponse({ success: true, title }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    // ===== タブID取得（content scriptが自分のタブIDを知るため） =====
    case 'getMyTabId':
      sendResponse({ tabId: sender.tab?.id || null });
      return false; // 同期レスポンス

    // ===== 既存の機能 =====
    case 'saveReviews':
      handleSaveReviews(message.reviews, sender.tab?.id, message.source || 'rakuten')
        .then(() => sendResponse({ success: true }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'downloadCSV':
      handleDownloadCSV()
        .then(() => sendResponse({ success: true }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'collectionComplete':
      // 非同期関数を正しく処理（awaitして完了を待つ）
      handleCollectionComplete(sender.tab?.id)
        .then(() => {
          console.log('[collectionComplete] 処理完了');
          sendResponse({ success: true });
        })
        .catch(error => {
          console.error('[collectionComplete] エラー:', error);
          sendResponse({ success: false, error: error.message });
        });
      return true; // 非同期レスポンスを示す

    case 'collectionStopped':
      handleCollectionStopped(sender.tab?.id)
        .then(() => sendResponse({ success: true }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'startQueueCollection':
      startQueueCollection()
        .then(() => sendResponse({ success: true }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'stopQueueCollection':
      stopQueueCollection()
        .then(() => sendResponse({ success: true }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'fetchRanking':
      fetchRankingProducts(message.url, message.count)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'updateScheduledAlarm':
      updateScheduledAlarm(message.settings)
        .then(() => sendResponse({ success: true }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'startSingleCollection':
      startSingleCollection(message.productInfo, message.tabId)
        .then(() => sendResponse({ success: true }))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'updateProgress':
    case 'log':
      forwardToAll(message);
      break;

    case 'updateProductTitle':
      // 収集開始時に商品名を更新（ASINのみだった場合）
      (async () => {
        try {
          const { productId, title, url } = message;
          if (!productId || !title) {
            sendResponse({ success: false, error: 'productId and title are required' });
            return;
          }

          console.log(`[Background] updateProductTitle: ${productId} → "${title.substring(0, 50)}..."`);

          // URLからproductIdを抽出するヘルパー関数
          const extractProductIdFromItemUrl = (itemUrl) => {
            if (!itemUrl) return null;
            // Amazon ASIN
            const amazonMatch = itemUrl.match(/(?:\/dp\/|\/gp\/product\/|\/product-reviews\/)([A-Z0-9]{10})/i);
            if (amazonMatch) return amazonMatch[1];
            // 楽天商品コード
            const rakutenMatch = itemUrl.match(/item\.rakuten\.co\.jp\/[^\/]+\/([^\/\?]+)/);
            if (rakutenMatch) return rakutenMatch[1];
            return null;
          };

          // アイテムがproductIdと一致するか判定（複数の方法で比較）
          const isMatchingItem = (item, targetProductId) => {
            // 直接比較
            if (item.productId === targetProductId) return true;
            if (item.id === targetProductId) return true;
            // titleがASINの場合（旧形式対応）
            if (item.title === targetProductId) return true;
            // URLから抽出して比較（フォールバック）
            const urlProductId = extractProductIdFromItemUrl(item.url);
            if (urlProductId === targetProductId) return true;
            return false;
          };

          // キューの更新
          const { queue = [] } = await chrome.storage.local.get('queue');
          let queueUpdated = false;

          for (const item of queue) {
            if (isMatchingItem(item, productId)) {
              // タイトルがASINのみ、または空の場合に更新
              const isAsinOnly = !item.title || item.title === productId || /^[A-Z0-9]{10}$/.test(item.title);
              if (isAsinOnly) {
                item.title = title;
                item.productId = productId; // productIdも設定（なかった場合のため）
                // 注意: URLは変更しない（既に正しく設定されているため）
                queueUpdated = true;
                console.log(`[Background] キュー更新成功: "${title.substring(0, 50)}..."`);
              }
              break;
            }
          }

          if (queueUpdated) {
            await chrome.storage.local.set({ queue });
            forwardToAll({ action: 'queueUpdated' });
          }

          // collectingItemsも更新
          const { collectingItems = [] } = await chrome.storage.local.get('collectingItems');
          let collectingUpdated = false;

          for (const item of collectingItems) {
            if (isMatchingItem(item, productId)) {
              const isAsinOnly = !item.title || item.title === productId || /^[A-Z0-9]{10}$/.test(item.title);
              if (isAsinOnly) {
                item.title = title;
                item.productId = productId; // productIdも設定（なかった場合のため）
                // 注意: URLは変更しない（既に正しく設定されているため）
                collectingUpdated = true;
                console.log(`[Background] collectingItems更新成功: "${title.substring(0, 50)}..."`);
              }
              break;
            }
          }

          if (collectingUpdated) {
            await chrome.storage.local.set({ collectingItems });
            forwardToAll({ action: 'queueUpdated' }); // UIを更新
          }

          sendResponse({ success: true, queueUpdated, collectingUpdated });
        } catch (error) {
          console.error('[Background] updateProductTitle error:', error);
          sendResponse({ success: false, error: error.message });
        }
      })();
      return true;

    case 'openOptions':
      chrome.runtime.openOptionsPage();
      sendResponse({ success: true });
      break;

    // ===== Amazonボット対策 =====
    case 'isTabActive':
      // タブがアクティブかどうかをチェック
      (async () => {
        try {
          const tab = await chrome.tabs.get(sender.tab.id);
          const window = await chrome.windows.get(tab.windowId);
          const isActive = tab.active && window.focused;
          sendResponse({ active: isActive });
        } catch (e) {
          sendResponse({ active: false });
        }
      })();
      return true;

    case 'checkAmazonRateLimit':
      checkAmazonRateLimit()
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ allowed: false, error: error.message }));
      return true;

    case 'incrementAmazonRateLimit':
      incrementAmazonRateLimit()
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'getAmazonRateLimitInfo':
      getAmazonRateLimitInfo()
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ error: error.message }));
      return true;
  }
});

/**
 * レビューを保存
 * @param {Array} reviews - レビューデータ
 * @param {number} tabId - 送信元タブID（定期収集用）
 * @param {string} source - 販路 ('rakuten' | 'amazon')
 */
async function handleSaveReviews(reviews, tabId = null, source = 'rakuten') {
  if (!reviews || reviews.length === 0) {
    return;
  }

  // 販路をレビューから判定（source パラメータより優先）
  const detectedSource = reviews[0]?.source || source;

  // 販路別のスプレッドシートURLを取得
  const syncSettings = await chrome.storage.sync.get([
    'separateSheets',
    'spreadsheetUrl',           // 楽天用
    'amazonSpreadsheetUrl'      // Amazon用
  ]);

  // タブ固有のスプレッドシートURLを取得（定期収集用）
  let spreadsheetUrl = null;
  let currentItem = null;

  if (tabId) {
    const collectingResult = await chrome.storage.local.get(['collectingItems']);
    const collectingItems = collectingResult.collectingItems || [];
    currentItem = collectingItems.find(item => item.tabId === tabId);
    if (currentItem) {
      spreadsheetUrl = currentItem.spreadsheetUrl || null;
    }
  }

  // 定期収集かどうかを判定（queueNameがあれば定期収集）
  const isScheduled = !!(currentItem?.queueName);

  // 定期収集は個別スプレッドシート必須、通常収集はグローバル設定を使用
  if (!spreadsheetUrl && !isScheduled) {
    if (detectedSource === 'amazon') {
      spreadsheetUrl = syncSettings.amazonSpreadsheetUrl;
    } else {
      spreadsheetUrl = syncSettings.spreadsheetUrl;
    }
  }

  // ログ用のプレフィックス
  const productId = reviews[0]?.productId || '';
  let prefix = productId ? `[${productId}] ` : '';
  if (isScheduled) {
    prefix = `[${currentItem.queueName}・${productId}] `;
  }

  // 定期収集・通常収集共通: ローカルストレージに蓄積のみ
  // スプレッドシートへの書き込みは商品収集完了時にバッチ処理で行う
  const newReviews = await saveToLocalStorage(reviews, detectedSource);

  if (newReviews.length === 0) {
    return;
  }

  // スプレッドシートURLを収集中アイテムに保存（完了時に使用）
  if (spreadsheetUrl && currentItem) {
    currentItem.spreadsheetUrl = spreadsheetUrl;
    const collectingResult = await chrome.storage.local.get(['collectingItems']);
    const collectingItems = collectingResult.collectingItems || [];
    const index = collectingItems.findIndex(item => item.tabId === tabId);
    if (index >= 0) {
      collectingItems[index] = currentItem;
      await chrome.storage.local.set({ collectingItems });
    }
  }

  // スプレッドシート未設定の警告（定期収集の場合のみ）
  if (!spreadsheetUrl && isScheduled) {
    log(prefix + '定期収集用のスプレッドシートが設定されていません', 'error');
  }
}

/**
 * スプレッドシートへの保存（商品収集完了時にバッチ処理）
 * @param {Object} completedItem - 完了した収集アイテム
 * @param {string} logPrefix - ログ出力用のプレフィックス
 */
async function saveReviewsToSpreadsheet(completedItem, logPrefix) {
  // 定期収集かどうかを判定
  const isScheduled = !!(completedItem?.queueName);

  // URLから販路と商品IDを判定
  const isAmazon = completedItem?.url?.includes('amazon.co.jp');
  const productId = extractProductIdFromUrl(completedItem?.url || '');

  log(`${logPrefix} [DEBUG] バッチ保存開始 isAmazon=${isAmazon}`);

  // スプレッドシートURLを取得
  let spreadsheetUrl = completedItem?.spreadsheetUrl || null;

  // URLが収集アイテムにない場合は設定から取得
  if (!spreadsheetUrl && !isScheduled) {
    const syncSettings = await chrome.storage.sync.get(['spreadsheetUrl', 'amazonSpreadsheetUrl']);
    spreadsheetUrl = isAmazon ? syncSettings.amazonSpreadsheetUrl : syncSettings.spreadsheetUrl;
    log(`${logPrefix} [DEBUG] 設定から取得 URL=${spreadsheetUrl ? '設定済み' : '未設定'}`);
  }

  if (!spreadsheetUrl) {
    log(`${logPrefix} [DEBUG] スプレッドシートURL未設定で終了`, 'error');
    return;
  }

  try {
    // ローカルストレージから全レビューを取得
    const stateResult = await chrome.storage.local.get(['collectionState']);
    const allReviews = stateResult.collectionState?.reviews || [];

    log(`${logPrefix} [DEBUG] ローカルストレージのレビュー数: ${allReviews.length}件`);

    if (allReviews.length === 0) {
      log(`${logPrefix} [DEBUG] レビュー0件で終了`);
      return;
    }

    // この商品のレビューのみフィルタリング
    // 複数商品を並列収集している場合に他の商品のデータを含めないため
    const productReviews = productId
      ? allReviews.filter(r => r.productId === productId)
      : allReviews;

    log(`${logPrefix} [DEBUG] フィルタ後: ${productReviews.length}件`);

    if (productReviews.length === 0) {
      log(`${logPrefix} [DEBUG] フィルタ後0件で終了`);
      return;
    }

    // 販路を判定
    const source = productReviews[0]?.source || (isAmazon ? 'amazon' : 'rakuten');

    // 設定を取得
    const syncSettings = await chrome.storage.sync.get(['separateSheets']);

    // スプレッドシートに保存（この商品のレビューのみ）
    // デフォルトはオフ（false）
    await sendToSheets(spreadsheetUrl, productReviews, syncSettings.separateSheets === true, isScheduled, source);
    log(logPrefix + ` スプレッドシートに保存しました（${productReviews.length}件）`, 'success');

    // 保存済みのレビューをローカルストレージから削除
    // 並列収集時に他の商品のレビューは残す
    const remainingReviews = productId
      ? allReviews.filter(r => r.productId !== productId)
      : [];

    const state = stateResult.collectionState || {};
    state.reviews = remainingReviews;
    await chrome.storage.local.set({ collectionState: state });
  } catch (error) {
    log(`${logPrefix} スプレッドシートへの保存に失敗: ${error.message}`, 'error');
  }
}

/**
 * レビューの一意キーを生成
 */
function getReviewKey(review) {
  // 商品ID + レビュー日 + 投稿者 + タイトル + 本文の最初の100文字で一意性を判断
  // 本文はtrim()して空白や改行を除去してから取得
  const bodySnippet = (review.body || '').trim().substring(0, 100);
  const titleSnippet = (review.title || '').trim().substring(0, 50);
  return `${review.productId || ''}_${review.reviewDate || ''}_${review.author || ''}_${titleSnippet}_${bodySnippet}`;
}

/**
 * ローカルストレージに保存（重複削除付き）
 * @param {Array} reviews - レビュー配列
 * @param {string} source - 販路 ('rakuten' | 'amazon')
 * @returns {Promise<Array>} 新規追加されたレビューの配列
 */
async function saveToLocalStorage(reviews, source = 'rakuten') {
  return new Promise((resolve, reject) => {
    chrome.storage.local.get(['collectionState'], (result) => {
      const state = result.collectionState || {
        isRunning: true,
        reviewCount: 0,
        pageCount: 0,
        totalPages: 0,
        reviews: [],
        logs: [],
        source: source
      };

      // 販路を記録
      state.source = source;

      const existingReviews = state.reviews || [];

      // 既存レビューのキーをセットに格納
      const existingKeys = new Set(existingReviews.map(r => getReviewKey(r)));

      // 重複を除外して新規レビューのみ追加
      const newReviews = reviews.filter(review => {
        const key = getReviewKey(review);
        if (existingKeys.has(key)) {
          return false; // 重複
        }
        existingKeys.add(key); // 新規追加分も重複チェック対象に
        return true;
      });

      const duplicateCount = reviews.length - newReviews.length;
      if (duplicateCount > 0) {
        if (newReviews.length === 0) {
          // 全て重複の場合
          log(`全て収集済み（${duplicateCount}件の重複をスキップ）`);
        } else {
          // 一部重複の場合
          log(`${newReviews.length}件追加、${duplicateCount}件の重複をスキップ`);
        }
        // ポップアップにも通知
        forwardToAll({
          action: 'duplicatesSkipped',
          count: duplicateCount,
          newCount: newReviews.length
        });
      }

      state.reviews = existingReviews.concat(newReviews);
      // レビュー数も更新（content.jsのupdateStateとのレースコンディション回避）
      state.reviewCount = state.reviews.length;

      chrome.storage.local.set({ collectionState: state }, () => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          // 進捗を通知
          forwardToAll({
            action: 'updateProgress',
            state: {
              isRunning: state.isRunning,
              reviewCount: state.reviewCount,
              pageCount: state.pageCount,
              totalPages: state.totalPages,
              source: state.source
            }
          });
          resolve(newReviews); // 新規追加されたレビューを返す
        }
      });
    });
  });
}

/**
 * ===== Google Sheets API 直接書き込み =====
 */

/**
 * スプレッドシートURLからIDを抽出
 */
function extractSpreadsheetId(url) {
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

/**
 * レビューを商品ごとにグループ化
 */
function groupReviewsByProduct(reviews) {
  return reviews.reduce((acc, review) => {
    const key = review.productId || 'その他';
    if (!acc[key]) acc[key] = [];
    acc[key].push(review);
    return acc;
  }, {});
}

/**
 * シートが空かどうか確認（データがなければtrue）
 */
async function isSheetEmpty(token, spreadsheetId, sheetTitle) {
  try {
    const encodedTitle = encodeURIComponent(sheetTitle);
    const response = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodedTitle}'!A1:Z10`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    if (!response.ok) {
      return false; // エラーの場合は空ではないと判断
    }

    const data = await response.json();
    // valuesがないか、空配列なら空シート
    return !data.values || data.values.length === 0;
  } catch (e) {
    return false;
  }
}

/**
 * シートをリネーム
 */
async function renameSheet(token, spreadsheetId, sheetId, newTitle) {
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        requests: [{
          updateSheetProperties: {
            properties: {
              sheetId: sheetId,
              title: newTitle
            },
            fields: 'title'
          }
        }]
      })
    }
  );

  return response.ok;
}

/**
 * シートが存在するか確認し、なければ作成
 * - 同じシート名がある → そのまま使用
 * - 同じシート名がない & 空シートがある → 空シートをリネームして使用
 * - 同じシート名がない & 空シートもない → 新規作成
 */
async function ensureSheetExists(token, spreadsheetId, sheetName) {
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'スプレッドシートの情報取得に失敗しました');
  }

  const data = await response.json();
  const sheets = data.sheets || [];
  const targetSheet = sheets.find(sheet => sheet.properties.title === sheetName);

  // 同じシート名がある → そのまま使用
  if (targetSheet) {
    return sheetName;
  }

  // 空のシートを探す（「シート1」「Sheet1」などのデフォルト名を優先）
  const defaultSheetNames = ['シート1', 'Sheet1', 'シート2', 'Sheet2', 'シート3', 'Sheet3'];
  let emptySheet = null;

  // まずデフォルト名のシートから空シートを探す
  for (const defaultName of defaultSheetNames) {
    const sheet = sheets.find(s => s.properties.title === defaultName);
    if (sheet && await isSheetEmpty(token, spreadsheetId, defaultName)) {
      emptySheet = sheet;
      break;
    }
  }

  // デフォルト名で見つからなければ、全シートから空シートを探す
  if (!emptySheet) {
    for (const sheet of sheets) {
      const title = sheet.properties.title;
      if (await isSheetEmpty(token, spreadsheetId, title)) {
        emptySheet = sheet;
        break;
      }
    }
  }

  // 空シートがあればリネームして使用
  if (emptySheet) {
    const renamed = await renameSheet(token, spreadsheetId, emptySheet.properties.sheetId, sheetName);
    if (renamed) {
      return sheetName;
    }
    // リネーム失敗時は新規作成にフォールバック
  }

  // 空シートがない → 新規作成
  const createResponse = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        requests: [{
          addSheet: {
            properties: { title: sheetName }
          }
        }]
      })
    }
  );

  if (!createResponse.ok) {
    const error = await createResponse.json();
    throw new Error(error.error?.message || 'シートの作成に失敗しました');
  }

  return sheetName;
}

/**
 * シートIDを取得
 */
async function getSheetId(token, spreadsheetId, sheetName) {
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const data = await response.json();
  const sheet = data.sheets?.find(s => s.properties.title === sheetName);
  return sheet?.properties?.sheetId || 0;
}

/**
 * データ行に書式を適用（白背景・黒文字・垂直中央揃え）
 * @param {string} source - 販路 ('rakuten' | 'amazon')
 */
async function formatDataRows(token, spreadsheetId, sheetId, startRow, endRow, source = 'rakuten') {
  // 販路別の列数
  const columnCount = source === 'amazon' ? 16 : 22;

  const requests = [
    {
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: startRow,
          endRowIndex: endRow,
          startColumnIndex: 0,
          endColumnIndex: columnCount  // 販路別の列数
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: {
              red: 1,
              green: 1,
              blue: 1
            },
            textFormat: {
              foregroundColor: {
                red: 0,
                green: 0,
                blue: 0
              }
            },
            verticalAlignment: 'MIDDLE'
          }
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)'
      }
    }
  ];

  await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ requests })
    }
  );
}

/**
 * URL列にクリック可能なリンク書式を適用
 * 楽天: D列（商品URL）、S列（レビュー掲載URL）
 * Amazon: D列（商品URL）、O列（レビュー掲載URL）
 * @param {string} token - OAuth token
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {number} sheetId - シートID
 * @param {Array} reviews - レビューデータ配列
 * @param {string} source - 販路 ('rakuten' | 'amazon')
 */
async function formatUrlColumns(token, spreadsheetId, sheetId, reviews, source = 'rakuten') {
  if (!reviews || reviews.length === 0) return;

  // 販路別のレビュー掲載URL列インデックス
  // 楽天: S列（19番目、インデックス18）
  // Amazon: O列（15番目、インデックス14）
  const pageUrlColIndex = source === 'amazon' ? 14 : 18;

  const requests = [];

  // 各レビュー行のURL列にリンク書式を適用
  reviews.forEach((review, index) => {
    const rowIndex = index + 1; // ヘッダー行の次から

    // D列（商品URL、列インデックス3）
    if (review.productUrl) {
      requests.push({
        updateCells: {
          range: {
            sheetId: sheetId,
            startRowIndex: rowIndex,
            endRowIndex: rowIndex + 1,
            startColumnIndex: 3,
            endColumnIndex: 4
          },
          rows: [{
            values: [{
              userEnteredValue: { stringValue: review.productUrl },
              textFormatRuns: [{
                startIndex: 0,
                format: {
                  link: { uri: review.productUrl }
                }
              }]
            }]
          }],
          fields: 'userEnteredValue,textFormatRuns'
        }
      });
    }

    // レビュー掲載URL列（販路別）
    if (review.pageUrl) {
      requests.push({
        updateCells: {
          range: {
            sheetId: sheetId,
            startRowIndex: rowIndex,
            endRowIndex: rowIndex + 1,
            startColumnIndex: pageUrlColIndex,
            endColumnIndex: pageUrlColIndex + 1
          },
          rows: [{
            values: [{
              userEnteredValue: { stringValue: review.pageUrl },
              textFormatRuns: [{
                startIndex: 0,
                format: {
                  link: { uri: review.pageUrl }
                }
              }]
            }]
          }],
          fields: 'userEnteredValue,textFormatRuns'
        }
      });
    }
  });

  // リクエストがない場合はスキップ
  if (requests.length === 0) return;

  // Google Sheets APIは一度に最大100リクエストまでのため、分割して送信
  const chunkSize = 100;
  for (let i = 0; i < requests.length; i += chunkSize) {
    const chunk = requests.slice(i, i + chunkSize);
    const response = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ requests: chunk })
      }
    );

    if (!response.ok) {
      console.error('[background] URL列のリンク書式適用エラー:', await response.text());
    }
  }
}

/**
 * ヘッダー行に書式を適用
 * 楽天: 赤背景・白テキスト
 * Amazon: 黒背景・オレンジテキスト (#ff9900)
 * @param {string} source - 販路 ('rakuten' | 'amazon')
 */
async function formatHeaderRow(token, spreadsheetId, sheetId, source = 'rakuten') {
  // 販路別の色設定と列数
  // 楽天: 22列（A-V）、Amazon: 16列（A-P）
  let backgroundColor, textColor;
  const columnCount = source === 'amazon' ? 16 : 22;

  if (source === 'amazon') {
    // Amazon: 黒背景・オレンジテキスト (#ff9900)
    backgroundColor = { red: 0, green: 0, blue: 0 };
    textColor = { red: 255/255, green: 153/255, blue: 0 }; // #ff9900
  } else {
    // 楽天: 赤背景・白テキスト (#BF0000)
    backgroundColor = { red: 191/255, green: 0, blue: 0 };
    textColor = { red: 1, green: 1, blue: 1 };
  }

  const requests = [
    // ヘッダー行の書式設定
    {
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: 0,
          endRowIndex: 1,
          startColumnIndex: 0,
          endColumnIndex: columnCount  // 販路別の列数
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: backgroundColor,
            textFormat: {
              foregroundColor: textColor,
              bold: true
            },
            horizontalAlignment: 'CENTER',
            verticalAlignment: 'MIDDLE'
          }
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
      }
    },
    // 1行目を固定
    {
      updateSheetProperties: {
        properties: {
          sheetId: sheetId,
          gridProperties: {
            frozenRowCount: 1
          }
        },
        fields: 'gridProperties.frozenRowCount'
      }
    },
    // 余分な列を削除（販路別の列数以降）
    {
      deleteDimension: {
        range: {
          sheetId: sheetId,
          dimension: 'COLUMNS',
          startIndex: columnCount,  // 販路別の列数から
          endIndex: 26              // Z列まで（デフォルトの列数）
        }
      }
    }
  ];

  await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ requests })
    }
  );
}

/**
 * シートにデータを書き込み（既存データは上書き）
 * @param {string} source - 販路 ('rakuten' | 'amazon')
 */
async function appendToSheet(token, spreadsheetId, sheetName, reviews, source = 'rakuten') {
  // シートが存在しなければ作成
  const actualSheetName = await ensureSheetExists(token, spreadsheetId, sheetName);

  const encodedSheetName = encodeURIComponent(actualSheetName);
  const sheetId = await getSheetId(token, spreadsheetId, actualSheetName);

  // テキストが=で始まる場合はエスケープ（数式として解釈されないように）
  const escapeFormula = (text) => {
    if (!text) return '';
    const str = String(text);
    if (str.startsWith('=') || str.startsWith('+') || str.startsWith('-') || str.startsWith('@')) {
      return "'" + str;
    }
    return str;
  };

  // 販路別のヘッダーとデータマッピング
  let headers;
  let dataValues;

  if (source === 'amazon') {
    // Amazon用ヘッダー（16項目）
    headers = [
      'レビュー日', 'ASIN', '商品名', '商品URL', '評価', 'タイトル', '本文',
      '投稿者', 'バリエーション', '参考になった数', '国', '認証購入', 'Vineレビュー',
      '画像あり', 'レビュー掲載URL', '収集日時'
    ];
    dataValues = reviews.map(review => [
      escapeFormula(review.reviewDate || ''),
      escapeFormula(review.productId || ''),
      escapeFormula(review.productName || ''),
      review.productUrl || '',
      review.rating || '',
      escapeFormula(review.title || ''),
      escapeFormula(review.body || ''),
      escapeFormula(review.author || ''),
      escapeFormula(review.variation || ''),
      review.helpfulCount || 0,
      escapeFormula(review.country || ''),
      review.isVerified ? '○' : '',
      review.isVine ? '○' : '',
      review.hasImage ? '○' : '',
      review.pageUrl || '',
      escapeFormula(review.collectedAt || '')
    ]);
  } else {
    // 楽天用ヘッダー（22項目）
    headers = [
      'レビュー日', '商品管理番号', '商品名', '商品URL', '評価', 'タイトル', '本文',
      '投稿者', '年代', '性別', '注文日', 'バリエーション', '用途', '贈り先',
      '購入回数', '参考になった数', 'ショップからの返信', 'ショップ名', 'レビュー掲載URL', '収集日時',
      '販路', '国'
    ];
    dataValues = reviews.map(review => [
      escapeFormula(review.reviewDate || ''),
      escapeFormula(review.productId || ''),
      escapeFormula(review.productName || ''),
      review.productUrl || '',
      review.rating || '',
      escapeFormula(review.title || ''),
      escapeFormula(review.body || ''),
      escapeFormula(review.author || ''),
      escapeFormula(review.age || ''),
      escapeFormula(review.gender || ''),
      escapeFormula(review.orderDate || ''),
      escapeFormula(review.variation || ''),
      escapeFormula(review.usage || ''),
      escapeFormula(review.recipient || ''),
      escapeFormula(review.purchaseCount || ''),
      review.helpfulCount || 0,
      escapeFormula(review.shopReply || ''),
      escapeFormula(review.shopName || ''),
      review.pageUrl || '',
      escapeFormula(review.collectedAt || ''),
      '楽天',
      escapeFormula(review.country || '日本')
    ]);
  }

  // ヘッダー + データを結合
  const allValues = [headers, ...dataValues];
  const totalRows = allValues.length;

  // 販路別の最終列（楽天: V、Amazon: P）
  const lastColumn = source === 'amazon' ? 'P' : 'V';

  // 1. シートの全データをクリア（販路別の列範囲）
  await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodedSheetName}'!A:${lastColumn}:clear`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}` }
    }
  );

  // 2. ヘッダーとデータを書き込み（販路別の列範囲）
  // USER_ENTERED: HYPERLINK関数を評価してクリック可能なリンクにする
  const writeResponse = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodedSheetName}'!A1:${lastColumn}${totalRows}?valueInputOption=USER_ENTERED`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ values: allValues })
    }
  );

  if (!writeResponse.ok) {
    const error = await writeResponse.json();
    throw new Error(error.error?.message || 'スプレッドシートへの書き込みに失敗しました');
  }

  // 3. シートの行数をデータ行数に一致させる（余分な行を削除）
  const sheetPropsResponse = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets(properties)`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const sheetPropsData = await sheetPropsResponse.json();
  const sheetProps = sheetPropsData.sheets?.find(s => s.properties.sheetId === sheetId);
  const currentRowCount = sheetProps?.properties?.gridProperties?.rowCount || 1000;

  // 余分な行がある場合は削除
  if (currentRowCount > totalRows) {
    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          requests: [{
            deleteDimension: {
              range: {
                sheetId: sheetId,
                dimension: 'ROWS',
                startIndex: totalRows,
                endIndex: currentRowCount
              }
            }
          }]
        })
      }
    );
  }

  // 4. ヘッダー書式を適用（販路別の色）
  await formatHeaderRow(token, spreadsheetId, sheetId, source);

  // 5. データ行に書式を適用（白背景・黒テキスト・垂直中央揃え）
  if (dataValues.length > 0) {
    await formatDataRows(token, spreadsheetId, sheetId, 1, totalRows, source);
  }

  // 6. URL列にクリック可能なリンク書式を適用（販路別）
  if (reviews.length > 0) {
    await formatUrlColumns(token, spreadsheetId, sheetId, reviews, source);
  }

  // 7. C列（商品名）の列幅をデータに合わせて自動調整
  await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        requests: [{
          autoResizeDimensions: {
            dimensions: {
              sheetId: sheetId,
              dimension: 'COLUMNS',
              startIndex: 2,  // C列（0-indexed で2）
              endIndex: 3     // C列のみ
            }
          }
        }]
      })
    }
  );

  // 8. 空白行を削除（シートの行数を調整）
  await trimEmptyRows(token, spreadsheetId, sheetId, encodedSheetName);
}

/**
 * Sheets APIを使ってスプレッドシートに直接書き込み
 * @param {string} source - 販路 ('rakuten' | 'amazon')
 */
async function sendToSheets(spreadsheetUrl, reviews, separateSheets = true, isScheduled = false, source = 'rakuten') {
  const spreadsheetId = extractSpreadsheetId(spreadsheetUrl);
  if (!spreadsheetId) {
    throw new Error('無効なスプレッドシートURLです');
  }

  // OAuthトークンを取得（まずinteractive: falseで試行し、失敗したらinteractive: trueで再試行）
  let token = await new Promise((resolve) => {
    chrome.identity.getAuthToken({ interactive: false }, (token) => {
      if (chrome.runtime.lastError || !token) {
        resolve(null);
      } else {
        resolve(token);
      }
    });
  });

  // トークンが取得できなかった場合、interactive: trueで再試行
  if (!token) {
    token = await new Promise((resolve, reject) => {
      chrome.identity.getAuthToken({ interactive: true }, (token) => {
        if (chrome.runtime.lastError) {
          reject(new Error('Googleアカウントでログインしてください: ' + chrome.runtime.lastError.message));
        } else if (!token) {
          reject(new Error('認証がキャンセルされました'));
        } else {
          resolve(token);
        }
      });
    });
  }

  // レビューを日付順（古い順 = 昇順）にソート
  const sortedReviews = sortReviewsByDate(reviews);

  if (separateSheets) {
    // 商品ごとにシートを分ける（シートをクリアして書き込み）
    const reviewsByProduct = groupReviewsByProduct(sortedReviews);
    for (const [productId, productReviews] of Object.entries(reviewsByProduct)) {
      // 「楽・商品管理番号」または「Ama・ASIN」形式（全収集で統一）
      // 商品の販路はレビューから判定
      const productSource = productReviews[0]?.source || source;
      const prefix = productSource === 'amazon' ? 'Ama' : '楽';
      const sheetName = `${prefix}・${productId}`;
      await appendToSheet(token, spreadsheetId, sheetName, productReviews, productSource);
    }
  } else {
    // 全て同じシートに保存（販路別）- 追記モード
    const sheetName = source === 'amazon' ? 'Amazon' : '楽天';
    await appendToSheetWithoutClear(token, spreadsheetId, sheetName, sortedReviews, source);
  }
}

/**
 * レビューを日付順（古い順 = 昇順）にソート
 * スプレッドシートで上から古い順に並ぶようにする
 */
function sortReviewsByDate(reviews) {
  return [...reviews].sort((a, b) => {
    const dateA = a.reviewDate || '';
    const dateB = b.reviewDate || '';
    // 日付文字列を比較（YYYY-MM-DD形式を想定）
    // 日付がない場合は末尾に配置
    if (!dateA && !dateB) return 0;
    if (!dateA) return 1;
    if (!dateB) return -1;
    return dateA.localeCompare(dateB);
  });
}

/**
 * シートに追記（既存データを保持）
 * シートを分けない設定の場合に使用
 */
async function appendToSheetWithoutClear(token, spreadsheetId, sheetName, reviews, source = 'rakuten') {
  // シートが存在しなければ作成
  const actualSheetName = await ensureSheetExists(token, spreadsheetId, sheetName);
  const encodedSheetName = encodeURIComponent(actualSheetName);
  const sheetId = await getSheetId(token, spreadsheetId, actualSheetName);

  // テキストが=で始まる場合はエスケープ
  const escapeFormula = (text) => {
    if (!text) return '';
    const str = String(text);
    if (str.startsWith('=') || str.startsWith('+') || str.startsWith('-') || str.startsWith('@')) {
      return "'" + str;
    }
    return str;
  };

  // 販路別のヘッダーとデータマッピング
  let headers;
  let dataValues;

  if (source === 'amazon') {
    headers = [
      'レビュー日', 'ASIN', '商品名', '商品URL', '評価', 'タイトル', '本文',
      '投稿者', 'バリエーション', '参考になった数', '国', '認証購入', 'Vineレビュー',
      '画像あり', 'レビュー掲載URL', '収集日時'
    ];
    dataValues = reviews.map(review => [
      escapeFormula(review.reviewDate || ''),
      escapeFormula(review.productId || ''),
      escapeFormula(review.productName || ''),
      review.productUrl || '',
      review.rating || '',
      escapeFormula(review.title || ''),
      escapeFormula(review.body || ''),
      escapeFormula(review.author || ''),
      escapeFormula(review.variation || ''),
      review.helpfulCount || 0,
      escapeFormula(review.country || ''),
      review.isVerified ? '○' : '',
      review.isVine ? '○' : '',
      review.hasImage ? '○' : '',
      review.pageUrl || '',
      escapeFormula(review.collectedAt || '')
    ]);
  } else {
    headers = [
      'レビュー日', '商品管理番号', '商品名', '商品URL', '評価', 'タイトル', '本文',
      '投稿者', '年代', '性別', '注文日', 'バリエーション', '用途', '贈り先',
      '購入回数', '参考になった数', 'ショップからの返信', 'ショップ名', 'レビュー掲載URL', '収集日時',
      '販路', '国'
    ];
    dataValues = reviews.map(review => [
      escapeFormula(review.reviewDate || ''),
      escapeFormula(review.productId || ''),
      escapeFormula(review.productName || ''),
      review.productUrl || '',
      review.rating || '',
      escapeFormula(review.title || ''),
      escapeFormula(review.body || ''),
      escapeFormula(review.author || ''),
      escapeFormula(review.age || ''),
      escapeFormula(review.gender || ''),
      escapeFormula(review.orderDate || ''),
      escapeFormula(review.variation || ''),
      escapeFormula(review.usage || ''),
      escapeFormula(review.recipient || ''),
      escapeFormula(review.purchaseCount || ''),
      review.helpfulCount || 0,
      escapeFormula(review.shopReply || ''),
      escapeFormula(review.shopName || ''),
      review.pageUrl || '',
      escapeFormula(review.collectedAt || ''),
      '楽天',
      escapeFormula(review.country || '日本')
    ]);
  }

  const lastColumn = source === 'amazon' ? 'P' : 'V';

  // 現在のシートの行数を取得
  const getResponse = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodedSheetName}'!A:A`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const getData = await getResponse.json();
  const existingRows = getData.values?.length || 0;

  if (existingRows === 0) {
    // シートが空の場合はヘッダー + データを書き込み
    const allValues = [headers, ...dataValues];
    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodedSheetName}'!A1:${lastColumn}${allValues.length}?valueInputOption=USER_ENTERED`,
      {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values: allValues })
      }
    );
    // ヘッダー書式を適用
    await formatHeaderRow(token, spreadsheetId, sheetId, source);
    // データ行に書式を適用
    if (dataValues.length > 0) {
      await formatDataRows(token, spreadsheetId, sheetId, 1, allValues.length, source);
    }
  } else {
    // 既存データがある場合は最後の行の次に追記
    const startRow = existingRows + 1;
    const endRow = startRow + dataValues.length - 1;

    const appendResponse = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodedSheetName}'!A${startRow}:${lastColumn}${endRow}?valueInputOption=USER_ENTERED`,
      {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values: dataValues })
      }
    );

    if (!appendResponse.ok) {
      const error = await appendResponse.json();
      throw new Error(error.error?.message || 'スプレッドシートへの追記に失敗しました');
    }

    // 追記した行にデータ書式を適用
    if (dataValues.length > 0) {
      await formatDataRows(token, spreadsheetId, sheetId, startRow - 1, endRow, source);
    }
  }

  // URL列にクリック可能なリンク書式を適用
  if (reviews.length > 0) {
    await formatUrlColumns(token, spreadsheetId, sheetId, reviews, source);
  }

  // シートの余分な行を削除（空白行をなくす）
  await trimEmptyRows(token, spreadsheetId, sheetId, encodedSheetName);
}

/**
 * シートの空白行を削除
 */
async function trimEmptyRows(token, spreadsheetId, sheetId, encodedSheetName) {
  try {
    // A列のデータを取得して最後のデータ行を特定
    const getResponse = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodedSheetName}'!A:A`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const getData = await getResponse.json();
    const values = getData.values || [];

    // 最後のデータがある行を特定（空白行をスキップ）
    let lastDataRow = 0;
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i] && values[i][0] && values[i][0].trim() !== '') {
        lastDataRow = i + 1; // 1-indexed
        break;
      }
    }

    if (lastDataRow === 0) return; // データがない場合は何もしない

    // シートの現在の行数を取得
    const sheetPropsResponse = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets(properties)`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const sheetPropsData = await sheetPropsResponse.json();
    const sheetProps = sheetPropsData.sheets?.find(s => s.properties.sheetId === sheetId);
    const currentRowCount = sheetProps?.properties?.gridProperties?.rowCount || 1000;

    // 余分な行がある場合は削除
    if (currentRowCount > lastDataRow) {
      await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
        {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            requests: [{
              deleteDimension: {
                range: {
                  sheetId: sheetId,
                  dimension: 'ROWS',
                  startIndex: lastDataRow,
                  endIndex: currentRowCount
                }
              }
            }]
          })
        }
      );
    }
  } catch (e) {
    console.error('空白行削除エラー:', e);
    // エラーが発生しても処理を続行
  }
}

/**
 * CSVダウンロード（販路別ファイル名）
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

        // 販路別のファイル名
        const source = state.source || 'rakuten';
        const prefix = source === 'amazon' ? 'amazon' : 'rakuten';
        const filename = `${prefix}_reviews_${formatDate(new Date())}.csv`;

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
 * レビューデータをCSV形式に変換（販路別ヘッダー）
 */
function convertToCSV(reviews) {
  const escapeCSV = (value) => {
    if (value === null || value === undefined) return '';
    const str = String(value);
    if (str.includes('"') || str.includes(',') || str.includes('\n') || str.includes('\r')) {
      return '"' + str.replace(/"/g, '""') + '"';
    }
    return str;
  };

  // 販路を判定（最初のレビューのsourceから判断）
  const source = reviews[0]?.source || 'rakuten';

  let headers;
  let rows;

  if (source === 'amazon') {
    // Amazon用（16項目）
    headers = [
      'レビュー日', 'ASIN', '商品名', '商品URL', '評価', 'タイトル', '本文',
      '投稿者', 'バリエーション', '参考になった数', '国', '認証購入', 'Vineレビュー',
      '画像あり', 'レビュー掲載URL', '収集日時'
    ];
    rows = reviews.map(review => [
      review.reviewDate || '', review.productId || '', review.productName || '',
      review.productUrl || '', review.rating || '', review.title || '', review.body || '',
      review.author || '', review.variation || '', review.helpfulCount || 0,
      review.country || '', review.isVerified ? '○' : '', review.isVine ? '○' : '',
      review.hasImage ? '○' : '', review.pageUrl || '', review.collectedAt || ''
    ]);
  } else {
    // 楽天用（22項目）
    headers = [
      'レビュー日', '商品管理番号', '商品名', '商品URL', '評価', 'タイトル', '本文',
      '投稿者', '年代', '性別', '注文日', 'バリエーション', '用途', '贈り先',
      '購入回数', '参考になった数', 'ショップからの返信', 'ショップ名', 'レビュー掲載URL', '収集日時',
      '販路', '国'
    ];
    rows = reviews.map(review => [
      review.reviewDate || '', review.productId || '', review.productName || '',
      review.productUrl || '', review.rating || '', review.title || '', review.body || '',
      review.author || '', review.age || '', review.gender || '', review.orderDate || '',
      review.variation || '', review.usage || '', review.recipient || '',
      review.purchaseCount || '', review.helpfulCount || 0, review.shopReply || '',
      review.shopName || '', review.pageUrl || '', review.collectedAt || '',
      '楽天', review.country || '日本'
    ]);
  }

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
  console.log('[handleCollectionComplete] 開始 tabId:', tabId);

  try {
    // 収集中アイテムと収集状態を取得
    const initialResult = await chrome.storage.local.get(['collectingItems', 'isQueueCollecting', 'expectedReviewTotal', 'collectionState']);
    const isQueueCollecting = initialResult.isQueueCollecting || false;
    const expectedTotal = initialResult.expectedReviewTotal || 0;
    const currentState = initialResult.collectionState || {};
    const actualCount = currentState.reviewCount || 0;
    const collectingItems = initialResult.collectingItems || [];

    console.log('[handleCollectionComplete] isQueueCollecting:', isQueueCollecting);
    console.log('[handleCollectionComplete] collectingItems:', JSON.stringify(collectingItems));

  // 収集中アイテムから商品情報を取得（ログ出力用）
  let completedItem = null;
  if (tabId) {
    completedItem = collectingItems.find(item => item.tabId === tabId);
    console.log('[handleCollectionComplete] tabIdで検索した結果:', completedItem ? 'found' : 'not found');
  }

  // tabIdで見つからない場合、収集中アイテムが1件だけならそれを使う（Amazon収集時のフォールバック）
  if (!completedItem && collectingItems.length === 1) {
    completedItem = collectingItems[0];
    console.log('[handleCollectionComplete] フォールバック: 収集中アイテムが1件のため使用:', completedItem?.url);
    // tabIdを更新（後続の処理で使用）
    if (tabId && completedItem) {
      completedItem.tabId = tabId;
    }
  }

  // それでも見つからない場合、Amazon収集中のアイテムを探す
  if (!completedItem && collectingItems.length > 0) {
    completedItem = collectingItems.find(item => item.url?.includes('amazon.co.jp'));
    if (completedItem) {
      console.log('[handleCollectionComplete] フォールバック: Amazon URLで検索して発見:', completedItem?.url);
      if (tabId) {
        completedItem.tabId = tabId;
      }
    }
  }

  console.log('[handleCollectionComplete] 最終的なcompletedItem:', completedItem ? completedItem.url : 'null');

  // レビュー件数の検証（期待値との比較）
  const verifyReviewCount = () => {
    if (expectedTotal > 0 && actualCount > 0) {
      const diff = expectedTotal - actualCount;
      const diffPercent = Math.round((diff / expectedTotal) * 100);

      if (diff > 0 && diffPercent > 5) {
        // 5%以上の不足がある場合は警告
        log(`⚠️ 取得件数が期待値より少ない可能性があります（取得: ${actualCount.toLocaleString()}件 / 期待: ${expectedTotal.toLocaleString()}件、差分: ${diff.toLocaleString()}件）`, 'error');
        return false;
      } else if (actualCount >= expectedTotal) {
        log(`✅ 全${actualCount.toLocaleString()}件のレビューを取得完了`, 'success');
        return true;
      } else {
        // 5%以内の差は許容（ページ表示と実際の件数の誤差）
        log(`✅ ${actualCount.toLocaleString()}件のレビューを取得完了（期待: ${expectedTotal.toLocaleString()}件）`, 'success');
        return true;
      }
    }
    return true; // 検証できない場合は成功扱い
  };

  // アクティブタブとスプレッドシートURL、収集中リストの処理
  // completedItemがある場合は処理を続行（tabIdがなくてもフォールバックで見つかった場合）
  const effectiveTabId = tabId || completedItem?.tabId;
  console.log('[handleCollectionComplete] effectiveTabId:', effectiveTabId, 'completedItem:', completedItem ? 'あり' : 'なし');

  if (effectiveTabId || completedItem) {
    // キュー収集の場合: activeCollectionTabsからの削除とタブを閉じる処理は、
    // chrome.alarms予約の直前で行う（onRemovedリスナーとの競合を避けるため）
    // ここではスプレッドシートURLのクリアのみ行う
    if (isQueueCollecting && effectiveTabId) {
      // スプレッドシートURLはクリア（次の商品用に新しく設定される）
      tabSpreadsheetUrls.delete(effectiveTabId);
    } else if (effectiveTabId) {
      // 単一収集の場合は通常通り削除
      activeCollectionTabs.delete(effectiveTabId);
      tabSpreadsheetUrls.delete(effectiveTabId);
    }

    // タブごとの状態をクリーンアップ
    if (effectiveTabId) {
      const stateKey = `collectionState_${effectiveTabId}`;
      await chrome.storage.local.remove(stateKey);
    }

    // 収集中リストから削除（URLベースでも削除できるように）
    const collectingResult = await chrome.storage.local.get(['collectingItems']);
    let updatedCollectingItems = collectingResult.collectingItems || [];
    const beforeCount = updatedCollectingItems.length;
    if (effectiveTabId) {
      updatedCollectingItems = updatedCollectingItems.filter(item => item.tabId !== effectiveTabId);
    }
    // tabIdで削除できなかった場合、URLベースで削除
    if (updatedCollectingItems.length === beforeCount && completedItem?.url) {
      updatedCollectingItems = updatedCollectingItems.filter(item => item.url !== completedItem.url);
    }
    console.log('[handleCollectionComplete] collectingItems更新: before=', beforeCount, 'after=', updatedCollectingItems.length);
    await chrome.storage.local.set({ collectingItems: updatedCollectingItems });

    // 商品の収集完了ログを出力
    if (completedItem) {
      const productId = extractProductIdFromUrl(completedItem.url);
      // 定期収集の場合は[キュー名・商品ID]形式
      const logPrefix = completedItem.queueName ? `[${completedItem.queueName}・${productId}]` : `[${productId}]`;
      log(`${logPrefix} 収集が完了しました`, 'success');
      console.log('[handleCollectionComplete] スプレッドシート保存を開始:', logPrefix);
      // 商品ごとの通知（設定で有効な場合のみ）
      showNotification('楽天レビュー収集', `${logPrefix} 収集が完了しました`, productId);

      // 最終収集日を保存（差分取得用）
      const lcResult = await chrome.storage.local.get(['productLastCollected']);
      const productLastCollected = lcResult.productLastCollected || {};
      productLastCollected[productId] = new Date().toISOString().split('T')[0]; // YYYY-MM-DD形式
      await chrome.storage.local.set({ productLastCollected });

      // スプレッドシートへの保存（バッチ処理：商品収集完了時に一括書き込み）
      try {
        await saveReviewsToSpreadsheet(completedItem, logPrefix);
        console.log('[handleCollectionComplete] スプレッドシート保存完了');
      } catch (saveError) {
        console.error('[handleCollectionComplete] スプレッドシート保存エラー:', saveError);
        log(`${logPrefix} スプレッドシート保存でエラーが発生しました: ${saveError.message}`, 'error');
      }
    } else {
      console.log('[handleCollectionComplete] completedItemがnullのためスプレッドシート保存をスキップ');
    }
  } else {
    console.log('[handleCollectionComplete] tabIdもcompletedItemもないため、スプレッドシート保存をスキップ');
  }

  // 状態を更新
  const result = await chrome.storage.local.get(['collectionState', 'queue']);
  const state = result.collectionState || {};
  const queue = result.queue || [];

  // 条件判定用の詳細ログ
  console.log('[handleCollectionComplete] ==== 条件チェック ====');
  console.log('[handleCollectionComplete] queue.length:', queue.length);
  console.log('[handleCollectionComplete] isQueueCollecting:', isQueueCollecting);
  console.log('[handleCollectionComplete] activeCollectionTabs.size:', activeCollectionTabs.size);
  console.log('[handleCollectionComplete] effectiveTabId:', effectiveTabId);
  console.log('[handleCollectionComplete] 条件1 (queue.length > 0 && isQueueCollecting):', queue.length > 0 && isQueueCollecting);
  console.log('[handleCollectionComplete] 条件2 (queue.length === 0 && isQueueCollecting):', queue.length === 0 && isQueueCollecting);
  console.log('[handleCollectionComplete] ========================');

  state.isRunning = activeCollectionTabs.size > 0;
  await chrome.storage.local.set({ collectionState: state });

  // ポップアップに完了を通知
  forwardToAll({
    action: 'collectionComplete',
    state: state
  });

  // キューに次の商品がある場合、次を処理
  if (queue.length > 0 && isQueueCollecting) {
    console.log('[handleCollectionComplete] >>> 条件1に入った: 次の商品を処理');
    log('次の商品の収集を開始します...');

    // タブを再利用するため、閉じずにlastUsedTabIdに保存
    if (effectiveTabId) {
      // activeCollectionTabsから先に削除（次の商品用に再登録するため）
      activeCollectionTabs.delete(effectiveTabId);
      await persistActiveCollectionTabs();  // 永続化
      console.log('[handleCollectionComplete] activeCollectionTabsから削除:', effectiveTabId);

      // タブは閉じずに再利用用に保持
      lastUsedTabId = effectiveTabId;
      console.log('[handleCollectionComplete] タブを再利用用に保持:', effectiveTabId);
    }

    console.log('[handleCollectionComplete] 次のキュー処理をchrome.alarmsで予約');
    // 重要: Service WorkerではsetTimeout/setIntervalがスロットリングされるため、
    // chrome.alarms APIを使用して確実に次の処理を実行する
    // delayInMinutesの最小値は約0.5秒（Chrome 120以降）
    // 参考: https://developer.chrome.com/docs/extensions/reference/alarms/
    await chrome.alarms.create(QUEUE_NEXT_ALARM_NAME, {
      delayInMinutes: 0.05  // 約3秒後（0.05分 = 3秒）
    });
    console.log('[handleCollectionComplete] アラーム設定完了: processNextInQueue');

  } else if (queue.length === 0 && isQueueCollecting) {
    // 重要: activeCollectionTabs.size === 0 ではなく queue.length === 0 で判定
    // 理由: Service Worker再起動時にactiveCollectionTabsはクリアされるため、
    //       activeCollectionTabs.size で判定すると誤って「完了」と判断してしまう
    console.log('[handleCollectionComplete] すべてのキュー収集が完了（queue.length === 0）');
    // すべて完了（キュー収集中の場合）
    await chrome.storage.local.set({ isQueueCollecting: false, collectingItems: [] });
    await persistActiveCollectionTabs();  // 永続化

    // 再利用したタブを閉じる
    if (lastUsedTabId) {
      try {
        await chrome.tabs.remove(lastUsedTabId);
      } catch (e) {
        // タブが既に閉じられている場合は無視
      }
      lastUsedTabId = null;
    }

    // 収集用ウィンドウを閉じる
    if (collectionWindowId) {
      try {
        await chrome.windows.remove(collectionWindowId);
      } catch (e) {
        // ウィンドウが既に閉じられている場合は無視
      }
      collectionWindowId = null;
    }

    // 収集結果のサマリーを取得
    const stateResult = await chrome.storage.local.get(['collectionState']);
    const finalState = stateResult.collectionState || {};
    const reviewCount = finalState.reviewCount || 0;
    log(`すべての収集が完了しました（${reviewCount}件のレビュー）`, 'success');
    // 全体完了の通知
    showNotification('楽天レビュー収集', `すべての収集が完了しました（${reviewCount}件のレビュー）`);
  } else if (activeCollectionTabs.size === 0 && !isQueueCollecting && !completedItem) {
    // 単一収集完了時の検証と通知
    verifyReviewCount();

    // expectedReviewTotalをクリア
    await chrome.storage.local.remove('expectedReviewTotal');

    const reviewCount = state.reviewCount || 0;
    showNotification('楽天レビュー収集', `収集が完了しました（${reviewCount}件のレビュー）`);
  }

  } catch (error) {
    // handleCollectionComplete内の全エラーをキャッチ
    console.error('[handleCollectionComplete] 致命的エラー:', error);
    log(`収集完了処理でエラーが発生しました: ${error.message}`, 'error');
    throw error; // 呼び出し元に再スロー
  }
}

/**
 * 収集停止時の処理
 */
async function handleCollectionStopped(tabId) {
  if (!tabId) return;

  // 収集中アイテムから商品情報を取得（ログ出力用）
  const collectingResult = await chrome.storage.local.get(['collectingItems']);
  let collectingItems = collectingResult.collectingItems || [];
  const stoppedItem = collectingItems.find(item => item.tabId === tabId);

  // アクティブタブから削除
  activeCollectionTabs.delete(tabId);
  await persistActiveCollectionTabs();  // 永続化
  tabSpreadsheetUrls.delete(tabId);

  // タブごとの状態をクリーンアップ
  const stateKey = `collectionState_${tabId}`;
  await chrome.storage.local.remove(stateKey);

  // 収集中リストから削除
  collectingItems = collectingItems.filter(item => item.tabId !== tabId);
  await chrome.storage.local.set({ collectingItems });

  // ログを出力
  if (stoppedItem) {
    const productId = extractProductIdFromUrl(stoppedItem.url);
    log(`[${productId}] 収集を停止しました`, 'error');
  }

  // 状態を更新
  const result = await chrome.storage.local.get(['collectionState']);
  const state = result.collectionState || {};
  state.isRunning = activeCollectionTabs.size > 0;
  await chrome.storage.local.set({ collectionState: state });

  // UIを更新
  forwardToAll({
    action: 'collectionComplete',
    state: state
  });
  forwardToAll({ action: 'queueUpdated' });
}

/**
 * URLから商品ID（楽天: 商品管理番号、Amazon: ASIN）を抽出
 */
function extractProductIdFromUrl(url) {
  if (!url) return 'unknown';

  // 楽天の場合
  const rakutenMatch = url.match(/item\.rakuten\.co\.jp\/[^\/]+\/([^\/\?]+)/);
  if (rakutenMatch) return rakutenMatch[1];

  // Amazonの場合
  const asin = extractASIN(url);
  if (asin) return asin;

  return 'unknown';
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

  // 最初の商品がAmazonかどうかで同時収集数を決定
  const firstItem = queue[0];
  const isAmazonFirst = firstItem?.url?.includes('amazon.co.jp');
  const maxConcurrent = isAmazonFirst ? 1 : 3;

  // 収集中フラグを立てる
  await chrome.storage.local.set({ isQueueCollecting: true, collectingItems: [] });

  // 収集用ウィンドウを新規作成（大きいサイズ、フォーカスあり）
  // Amazonボット対策: アクティブウィンドウで操作することで検出を回避
  try {
    const window = await chrome.windows.create({
      url: 'about:blank',
      width: 1280,
      height: 800,
      focused: true  // ボット対策: フォーカスを当てる
    });
    collectionWindowId = window.id;

    // 楽天のみウィンドウを自動最小化（Amazonは最小化すると正常に動作しない）
    if (!isAmazonFirst) {
      setTimeout(() => {
        chrome.windows.update(collectionWindowId, { state: 'minimized' }).catch(() => {});
      }, 500);
    }

    // about:blankタブは後で閉じる
    if (window.tabs && window.tabs[0]) {
      setTimeout(() => {
        chrome.tabs.remove(window.tabs[0].id).catch(() => {});
      }, 1000);
    }
  } catch (e) {
    console.error('ウィンドウ作成エラー:', e);
    collectionWindowId = null;
  }

  log(`${queue.length}件のキュー収集を開始します（同時${maxConcurrent}件）`, 'success');

  // 同時収集数分だけ開始
  for (let i = 0; i < maxConcurrent && i < queue.length; i++) {
    await processNextInQueue();
  }
}

/**
 * キュー収集を停止
 */
async function stopQueueCollection() {
  // すべてのアクティブタブに停止メッセージを送信
  for (const tabId of activeCollectionTabs) {
    try {
      await chrome.tabs.sendMessage(tabId, { action: 'stopCollection' });
      // タブを閉じる
      await chrome.tabs.remove(tabId);
    } catch (e) {
      // タブが既に閉じられている場合は無視
    }
  }

  // 収集用ウィンドウを閉じる
  if (collectionWindowId) {
    try {
      await chrome.windows.remove(collectionWindowId);
    } catch (e) {
      // ウィンドウが既に閉じられている場合は無視
    }
    collectionWindowId = null;
  }

  // 状態をクリア
  activeCollectionTabs.clear();
  tabSpreadsheetUrls.clear();
  await chrome.storage.local.set({
    isQueueCollecting: false,
    collectingItems: []
  });

  log('すべての収集を停止しました', 'error');
  forwardToAll({ action: 'queueUpdated' });
}

/**
 * キューの次の商品を処理
 */
async function processNextInQueue() {
  const result = await chrome.storage.local.get(['queue', 'collectingItems']);
  const queue = result.queue || [];
  const collectingItems = result.collectingItems || [];

  if (queue.length === 0) {
    return;
  }

  // 次の商品がAmazonかどうかで同時収集数を決定
  const nextItem = queue[0]; // まだshiftしない
  const isAmazonNext = nextItem?.url?.includes('amazon.co.jp');

  // 同時収集数: Amazonは1件、楽天は3件
  const maxConcurrent = isAmazonNext ? 1 : 3;

  // 現在のアクティブタブ数をチェック
  if (activeCollectionTabs.size >= maxConcurrent) {
    if (isAmazonNext) {
      // Amazonは1件ずつなのでログは出さない（頻繁に表示されるため）
    } else {
      log(`同時収集数上限（${maxConcurrent}）に達しています`, 'info');
    }
    return;
  }

  // キューから取り出し（上で queue[0] で確認済み）
  queue.shift(); // nextItemは上で取得済み

  // デバッグログ
  console.log('[processNextInQueue] nextItem:', JSON.stringify(nextItem, null, 2));

  // 収集中リストに追加
  nextItem.tabId = null; // 後で設定
  collectingItems.push(nextItem);

  await chrome.storage.local.set({ queue, collectingItems });

  forwardToAll({ action: 'queueUpdated' });

  // 定期収集の場合は[キュー名・商品ID]形式
  const productId = extractProductIdFromUrl(nextItem.url);
  const queuePrefix = nextItem.queueName ? `[${nextItem.queueName}・${productId}]` : '';
  console.log('[processNextInQueue] 処理URL:', nextItem.url, 'isAmazon:', nextItem.url.includes('amazon.co.jp'));

  // タブ再利用の実装
  let tab;
  const isAmazonUrl = nextItem.url.includes('amazon.co.jp');

  // 既存のタブを再利用するか、新規タブを作成
  if (lastUsedTabId) {
    try {
      // 既存タブが存在するか確認
      const existingTab = await chrome.tabs.get(lastUsedTabId);
      if (existingTab) {
        // タブを再利用してURLを更新
        await chrome.tabs.update(lastUsedTabId, { url: nextItem.url });
        tab = { id: lastUsedTabId };
        console.log('[processNextInQueue] タブを再利用:', lastUsedTabId);
      }
    } catch (e) {
      // タブが存在しない場合は新規作成
      console.log('[processNextInQueue] 再利用タブが存在しないため新規作成');
      lastUsedTabId = null;
    }
  }

  // 再利用できなかった場合は新規タブを作成
  if (!tab) {
    if (collectionWindowId) {
      try {
        tab = await chrome.tabs.create({
          url: nextItem.url,
          windowId: collectionWindowId,
          active: false
        });
        console.log('[processNextInQueue] 新規タブを作成:', tab.id);
      } catch (e) {
        // ウィンドウが閉じられている場合は通常のタブで開く
        console.error('収集用ウィンドウにタブ作成失敗:', e);
        tab = await chrome.tabs.create({ url: nextItem.url, active: false });
      }
    } else {
      // フォールバック: 通常のタブ
      tab = await chrome.tabs.create({ url: nextItem.url, active: false });
      console.log('[processNextInQueue] 新規タブを作成（通常）:', tab.id);
    }
    // 新規作成したタブを再利用用に保持
    lastUsedTabId = tab.id;
  }

  // tabIdを収集中アイテムに設定
  nextItem.tabId = tab.id;
  await chrome.storage.local.set({ collectingItems });

  // アクティブタブとして追跡
  activeCollectionTabs.add(tab.id);
  await persistActiveCollectionTabs();  // 永続化

  // タブのアクティブ化は無効化（バックグラウンドで動作させるため）

  // キュー固有のスプレッドシートURLがあれば保存
  if (nextItem.spreadsheetUrl) {
    tabSpreadsheetUrls.set(tab.id, nextItem.spreadsheetUrl);
  }

  // 収集状態をセット（タブIDごとに管理）
  const stateKey = `collectionState_${tab.id}`;
  await chrome.storage.local.set({
    [stateKey]: {
      isRunning: true,
      reviewCount: 0,
      pageCount: 0,
      totalPages: 0,
      reviews: [],
      tabId: tab.id,
      queueName: nextItem.queueName || null,
      incrementalOnly: nextItem.incrementalOnly || false
    }
  });

  // 共通のcollectionStateもリセット（レビュー蓄積を防ぐ）
  // queueNameも含めて自動再開時に正しく復元できるようにする
  // 重要: すべての状態を明示的に初期化（以前の収集の状態が残らないようにする）
  // 新しい商品の収集開始時は、星フィルター状態もリセット（★1から開始）
  await chrome.storage.local.set({
    collectionState: {
      isRunning: true,
      reviewCount: 0,
      pageCount: 0,
      totalPages: 0,
      reviews: [],
      collectedReviewKeys: [],       // 明示的にクリア（以前の収集のキーを引き継がない）
      consecutiveSkipPages: 0,       // 明示的にクリア
      sessionPageCount: 0,           // 明示的にクリア
      lastProcessedPage: 0,          // 明示的にクリア
      productId: extractProductIdFromUrl(nextItem.url), // 商品IDも設定
      queueName: nextItem.queueName || null,
      incrementalOnly: nextItem.incrementalOnly || false,
      source: isAmazonUrl ? 'amazon' : 'rakuten', // 販路を設定（自動再開に必要）
      startedFromQueue: true,        // キュー処理からの開始フラグ（競合防止用）
      // 新しい商品の収集なので、星フィルター状態をリセット（★1から開始）
      useStarFilter: true,
      currentStarFilterIndex: 0,     // ★1から開始
      pagesCollectedInCurrentFilter: 0,
      activeTabId: tab.id            // 収集中のタブID（他タブでの操作を防止）
    }
  });

  // Amazonの場合: グローバルリスナー（chrome.tabs.onUpdated）がstartedFromQueueフラグを見て
  // startCollectionを送信するので、ここでは何もしない
  //
  // 楽天の場合: グローバルリスナーはAmazonのみ対応しているため、ここでローカルリスナーを使用
  if (!isAmazonUrl) {
    // 楽天用のローカルリスナー
    let startCollectionSent = false;

    chrome.tabs.onUpdated.addListener(function listener(tabId, info, tabInfo) {
      if (tabId === tab.id && info.status === 'complete') {
        if (startCollectionSent) {
          console.log('[キュー処理・楽天] startCollection既に送信済みのためスキップ');
          return;
        }
        startCollectionSent = true;
        chrome.tabs.onUpdated.removeListener(listener);

        (async () => {
          let lastCollectedDate = null;
          if (nextItem.incrementalOnly) {
            const productId = extractProductIdFromUrl(nextItem.url);
            const lcResult = await chrome.storage.local.get(['productLastCollected']);
            const productLastCollected = lcResult.productLastCollected || {};
            lastCollectedDate = productLastCollected[productId] || null;
          }

          const productIdForMessage = extractProductIdFromUrl(nextItem.url);

          console.log('[キュー処理・楽天] startCollectionメッセージ送信');
          sendMessageWithRetry(tab.id, {
            action: 'startCollection',
            incrementalOnly: nextItem.incrementalOnly || false,
            lastCollectedDate: lastCollectedDate,
            queueName: nextItem.queueName || nextItem.title || productIdForMessage,
            productId: productIdForMessage
          }, 'startCollection（楽天キュー処理）');
        })();
      }
    });
  } else {
    // Amazonの場合はグローバルリスナーに任せる
    // startedFromQueue フラグが設定されているので、グローバルリスナーが処理する
    console.log('[キュー処理・Amazon] グローバルリスナーに処理を委譲（startedFromQueue=true）');
  }
}

/**
 * 単一ページの収集を開始（ポップアップからの収集開始ボタン用）
 * キューに追加して収集中状態にし、完了時に削除する
 */
async function startSingleCollection(productInfo, tabId) {
  const result = await chrome.storage.local.get(['queue', 'collectingItems']);
  let queue = result.queue || [];
  let collectingItems = result.collectingItems || [];

  // キューに同じURLがあるか確認
  const existingInQueue = queue.find(item => item.url === productInfo.url);
  const existingInCollecting = collectingItems.find(item => item.url === productInfo.url);

  if (existingInCollecting) {
    // 既に収集中の場合はエラー
    throw new Error('この商品は既に収集中です');
  }

  if (!existingInQueue) {
    // キューにない場合は追加
    queue.push(productInfo);
  }

  // 収集中リストに追加
  const collectingItem = {
    ...productInfo,
    tabId: tabId
  };
  collectingItems.push(collectingItem);

  // キューから削除（収集中リストに移動したため）
  queue = queue.filter(item => item.url !== productInfo.url);

  // ストレージを更新
  await chrome.storage.local.set({ queue, collectingItems });

  // アクティブタブとして追跡
  activeCollectionTabs.add(tabId);
  await persistActiveCollectionTabs();  // 永続化

  // 収集状態をセット（タブIDごとに管理）
  const stateKey = `collectionState_${tabId}`;
  await chrome.storage.local.set({
    [stateKey]: {
      isRunning: true,
      reviewCount: 0,
      pageCount: 0,
      totalPages: 0,
      reviews: [],
      tabId: tabId
    }
  });

  const productId = extractProductIdFromUrl(productInfo.url);
  const queueName = productInfo.title || productId;

  // 共通のcollectionStateもリセット（レビュー蓄積を防ぐ）
  // queueName, productIdも設定（フィルター遷移後の再開に必要）
  // 重要: processNextInQueueと同様の形式に統一
  // 新しい商品の収集開始時は、星フィルター状態もリセット（★1から開始）
  await chrome.storage.local.set({
    collectionState: {
      isRunning: true,
      reviewCount: 0,
      pageCount: 0,
      totalPages: 0,
      reviews: [],
      collectedReviewKeys: [],       // 明示的にクリア
      consecutiveSkipPages: 0,       // 明示的にクリア
      sessionPageCount: 0,           // 明示的にクリア
      lastProcessedPage: 0,          // 明示的にクリア
      source: productInfo.source || (productInfo.url.includes('amazon') ? 'amazon' : 'rakuten'),
      queueName: queueName,
      productId: productId,
      startedFromQueue: false,       // 拡張ウィンドウからの開始
      // 新しい商品の収集なので、星フィルター状態をリセット（★1から開始）
      useStarFilter: true,
      currentStarFilterIndex: 0,     // ★1から開始
      pagesCollectedInCurrentFilter: 0,
      activeTabId: tabId             // 収集中のタブID（他タブでの操作を防止）
    }
  });

  // UIを更新
  forwardToAll({ action: 'queueUpdated' });

  log(`[${productId}] 収集を開始しました`);

  // content.jsに収集開始を指示（queueName, productIdも送信）
  chrome.tabs.sendMessage(tabId, {
    action: 'startCollection',
    queueName: queueName,
    productId: productId
  }).catch((error) => {
    log(`収集開始に失敗: ${error.message}`, 'error');
  });
}

/**
 * ランキングから商品を取得してキューに追加
 * URLに応じて楽天またはAmazonのパーサーを呼び出す
 */
async function fetchRankingProducts(url, count) {
  if (url.includes('amazon.co.jp')) {
    return fetchAmazonRankingProducts(url, count);
  } else {
    return fetchRakutenRankingProducts(url, count);
  }
}

/**
 * 楽天ランキングから商品を取得
 * Service WorkerではDOMParserが使えないため、正規表現でパース
 */
async function fetchRakutenRankingProducts(url, count) {
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
          addedAt: new Date().toISOString(),
          source: 'rakuten'
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

    if (addedCount > 0) {
      log(`楽天ランキングから${addedCount}件の商品をキューに追加しました`, 'success');
    } else {
      log('追加する商品がありません（全て重複）', 'info');
    }

    return { success: true, addedCount };
  } catch (error) {
    log(`楽天ランキング取得エラー: ${error.message}`, 'error');
    return { success: false, error: error.message };
  }
}

/**
 * Amazonランキングから商品を取得
 * data-asin属性からASINを抽出
 */
async function fetchAmazonRankingProducts(url, count) {
  try {
    // ランキングページをfetchで取得
    const response = await fetch(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
        'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8'
      }
    });
    const html = await response.text();

    const products = [];
    const seenAsins = new Set();

    // 方法1: data-asin属性からASINを抽出
    // パターン: data-asin="B0XXXXXXXX"
    const asinPattern = /data-asin="([A-Z0-9]{10})"/g;
    let match;

    while ((match = asinPattern.exec(html)) !== null && products.length < count) {
      const asin = match[1].toUpperCase();

      // 空のASINや重複をスキップ
      if (!asin || seenAsins.has(asin)) continue;
      seenAsins.add(asin);

      // 商品URLを構築
      const productUrl = `https://www.amazon.co.jp/dp/${asin}`;

      products.push({
        url: productUrl,
        title: asin,
        addedAt: new Date().toISOString(),
        source: 'amazon'
      });
    }

    // 方法2: /dp/ASINパターンからも抽出（フォールバック）
    if (products.length < count) {
      const dpPattern = /\/dp\/([A-Z0-9]{10})/gi;
      while ((match = dpPattern.exec(html)) !== null && products.length < count) {
        const asin = match[1].toUpperCase();

        if (!asin || seenAsins.has(asin)) continue;
        seenAsins.add(asin);

        const productUrl = `https://www.amazon.co.jp/dp/${asin}`;

        products.push({
          url: productUrl,
          title: asin,
          addedAt: new Date().toISOString(),
          source: 'amazon'
        });
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

    if (addedCount > 0) {
      log(`Amazonランキングから${addedCount}件の商品をキューに追加しました`, 'success');
    } else {
      log('追加する商品がありません（全て重複）', 'info');
    }

    return { success: true, addedCount };
  } catch (error) {
    log(`Amazonランキング取得エラー: ${error.message}`, 'error');
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
 * ストレージへの保存はoptions.jsのaddLog()が行うため、ここでは転送のみ
 */
function log(text, type = '') {
  console.log(`[レビュー収集] ${text}`);

  forwardToAll({
    action: 'log',
    text: text,
    type: type
  });
}

// 拡張機能インストール時の初期化
chrome.runtime.onInstalled.addListener(() => {
  console.log('レビュー収集拡張機能がインストールされました（楽天・Amazon対応）');

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

// キーボードショートカットのリスナー
chrome.commands.onCommand.addListener((command) => {
  if (command === 'open_options') {
    chrome.runtime.openOptionsPage();
  }
});

// タブが閉じられた時のクリーンアップ
chrome.tabs.onRemoved.addListener(async (tabId) => {
  if (activeCollectionTabs.has(tabId)) {
    activeCollectionTabs.delete(tabId);
    await persistActiveCollectionTabs();  // 永続化
    tabSpreadsheetUrls.delete(tabId);
    // 注: このログは収集完了後の正常なタブクローズでも表示される
    // collectionCompleteメッセージの後にタブが閉じられた場合は正常終了
    console.log(`[tabs.onRemoved] タブ ${tabId} がactiveCollectionTabsから削除されました`);

    // タブごとの状態をクリーンアップ
    const stateKey = `collectionState_${tabId}`;
    chrome.storage.local.remove(stateKey);

    // キューに次がある場合は処理を続行（chrome.alarmsで1秒後）
    console.log('[tabs.onRemoved] キュー処理続行をアラームで予約');
    chrome.alarms.create(QUEUE_NEXT_ALARM_NAME, {
      delayInMinutes: 0.017  // 約1秒後（0.017分 ≈ 1秒）
    });
  }
});

// 収集用ウィンドウが閉じられた時のクリーンアップ
chrome.windows.onRemoved.addListener((windowId) => {
  // collectionWindowIdが明示的に設定されていて、かつそのウィンドウが閉じられた場合のみ処理
  // 注意: タブ再利用時やClaude in Chrome使用時はcollectionWindowIdが未設定(null)なので、
  // この条件に入らないようにする
  if (collectionWindowId && windowId === collectionWindowId) {
    console.log('[windows.onRemoved] 収集用ウィンドウが閉じられました windowId:', windowId);
    collectionWindowId = null;
    // ウィンドウが手動で閉じられた場合、収集を停止
    // ただし、activeCollectionTabsが空でない場合のみ（実際に収集中の場合）
    if (activeCollectionTabs.size > 0) {
      // 収集タブがまだ存在するか確認（ウィンドウが閉じられてもタブが別ウィンドウに移動している可能性）
      Promise.all([...activeCollectionTabs].map(tabId =>
        chrome.tabs.get(tabId).catch(() => null)
      )).then(tabs => {
        const existingTabs = tabs.filter(t => t !== null);
        if (existingTabs.length === 0) {
          // 本当にタブがなくなった場合のみ停止
          log('収集用ウィンドウが閉じられました。収集を停止します', 'error');
          activeCollectionTabs.clear();
          persistActiveCollectionTabs();  // 永続化
          chrome.storage.local.set({
            isQueueCollecting: false,
            collectingItems: []
          });
          forwardToAll({ action: 'queueUpdated' });
        } else {
          console.log('[windows.onRemoved] 収集タブはまだ存在します:', existingTabs.length);
        }
      });
    }
  }
});

// ========================================
// 定期収集機能
// ========================================

const SCHEDULED_ALARM_NAME = 'scheduledCollection';

/**
 * 定期収集アラームを設定/更新
 * @param {Object} settings - 定期収集設定
 */
async function updateScheduledAlarm(settings) {
  // 既存のアラームをクリア
  await chrome.alarms.clear(SCHEDULED_ALARM_NAME);

  if (!settings.hasEnabledQueues) {
    console.log('定期収集アラームを無効化（有効なキューなし）');
    return;
  }

  // 次の実行時刻を計算
  const [hours, minutes] = (settings.time || '07:00').split(':').map(Number);
  const now = new Date();
  const nextRun = new Date(now);
  nextRun.setHours(hours, minutes, 0, 0);

  // 既に今日の時刻を過ぎていたら明日に設定
  if (nextRun <= now) {
    nextRun.setDate(nextRun.getDate() + 1);
  }

  const delayInMinutes = (nextRun.getTime() - now.getTime()) / 1000 / 60;

  // アラームを設定（毎日繰り返し）
  await chrome.alarms.create(SCHEDULED_ALARM_NAME, {
    delayInMinutes: delayInMinutes,
    periodInMinutes: 24 * 60  // 24時間ごと
  });

  console.log(`定期収集アラームを設定: 次回 ${nextRun.toLocaleString('ja-JP')}`);
}

/**
 * 定期収集を実行（複数キュー対応）
 */
async function runScheduledCollection() {
  console.log('定期収集を開始');

  const result = await chrome.storage.local.get(['scheduledCollection', 'scheduledQueues']);
  const settings = result.scheduledCollection || {};
  const scheduledQueues = result.scheduledQueues || [];

  // 有効なキューを取得
  const enabledQueues = scheduledQueues.filter(q => q.enabled);

  if (enabledQueues.length === 0) {
    console.log('定期収集が有効なキューがありません');
    return;
  }

  // キューに追加
  const queueResult = await chrome.storage.local.get(['queue']);
  const currentQueue = queueResult.queue || [];

  let totalAdded = 0;
  const processedQueues = [];

  enabledQueues.forEach(targetQueue => {
    if (!targetQueue.items || targetQueue.items.length === 0) return;

    let addedCount = 0;
    targetQueue.items.forEach(item => {
      const exists = currentQueue.some(q => q.url === item.url);
      if (!exists) {
        currentQueue.push({
          url: item.url,
          title: item.title,
          addedAt: new Date().toISOString(),
          scheduledRun: true,
          incrementalOnly: targetQueue.incrementalOnly !== false,
          spreadsheetUrl: targetQueue.spreadsheetUrl || null,
          queueName: targetQueue.name
        });
        addedCount++;
      }
    });

    if (addedCount > 0) {
      totalAdded += addedCount;
      processedQueues.push({ name: targetQueue.name, count: addedCount });

      // キューごとの最終実行時刻を更新
      targetQueue.lastRun = new Date().toISOString();
    }
  });

  if (totalAdded === 0) {
    console.log('追加するアイテムがありません（全て重複）');
    return;
  }

  // キューと最終実行時刻を保存
  await chrome.storage.local.set({ queue: currentQueue, scheduledQueues });

  // 収集開始
  const queueSummary = processedQueues.map(q => `「${q.name}」${q.count}件`).join('、');
  log(`定期収集開始: ${queueSummary}`, 'success');
  forwardToAll({ action: 'queueUpdated' });

  await startQueueCollection();

  // グローバルな実行記録を更新
  settings.lastRun = new Date().toISOString();
  await chrome.storage.local.set({ scheduledCollection: settings });
}

// アラームリスナー
chrome.alarms.onAlarm.addListener((alarm) => {
  console.log('[alarms.onAlarm] アラーム発火:', alarm.name);

  if (alarm.name === SCHEDULED_ALARM_NAME) {
    runScheduledCollection().catch(error => {
      console.error('定期収集エラー:', error);
      log('定期収集でエラーが発生しました: ' + error.message, 'error');
    });
  } else if (alarm.name === QUEUE_NEXT_ALARM_NAME) {
    // キュー処理の次の商品を開始
    console.log('[alarms.onAlarm] processNextInQueueを実行');
    processNextInQueue().catch(error => {
      console.error('キュー処理エラー:', error);
      log('キュー処理でエラーが発生しました: ' + error.message, 'error');
    });
  }
});

// 拡張機能起動時に定期収集設定を復元
chrome.storage.local.get(['scheduledCollection'], (result) => {
  const settings = result.scheduledCollection;
  if (settings && settings.enabled) {
    updateScheduledAlarm(settings).catch(console.error);
  }
});
