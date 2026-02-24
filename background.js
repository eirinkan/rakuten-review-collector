/**
 * バックグラウンドサービスワーカー
 * レビューデータの保存、CSVダウンロード、キュー管理を処理
 * 楽天市場・Amazon両対応
 */

// PDF生成ライブラリの読み込み
importScripts('pdf-lib.min.js', 'pdf-generator.js');

// ===== 収集項目の定義 =====
// 楽天のフィールド定義（順序固定）
const RAKUTEN_FIELD_DEFINITIONS = [
  { key: 'reviewDate', header: 'レビュー日', getValue: (r, esc) => esc(r.reviewDate || '') },
  { key: 'productId', header: '商品管理番号', getValue: (r, esc) => esc(r.productId || '') },
  { key: 'productName', header: '商品名', getValue: (r, esc) => esc(r.productName || '') },
  { key: 'productUrl', header: '商品URL', getValue: (r) => r.productUrl || '' },
  { key: 'rating', header: '評価', getValue: (r) => r.rating || '' },
  { key: 'title', header: 'タイトル', getValue: (r, esc) => esc(r.title || '') },
  { key: 'body', header: '本文', getValue: (r, esc) => esc(r.body || '') },
  { key: 'author', header: '投稿者', getValue: (r, esc) => esc(r.author || '') },
  { key: 'age', header: '年代', getValue: (r, esc) => esc(r.age || '') },
  { key: 'gender', header: '性別', getValue: (r, esc) => esc(r.gender || '') },
  { key: 'orderDate', header: '注文日', getValue: (r, esc) => esc(r.orderDate || '') },
  { key: 'variation', header: 'バリエーション', getValue: (r, esc) => esc(r.variation || '') },
  { key: 'usage', header: '用途', getValue: (r, esc) => esc(r.usage || '') },
  { key: 'recipient', header: '贈り先', getValue: (r, esc) => esc(r.recipient || '') },
  { key: 'purchaseCount', header: '購入回数', getValue: (r, esc) => esc(r.purchaseCount || '') },
  { key: 'helpfulCount', header: '参考になった数', getValue: (r) => r.helpfulCount || 0 },
  { key: 'shopReply', header: 'ショップからの返信', getValue: (r, esc) => esc(r.shopReply || '') },
  { key: 'shopName', header: 'ショップ名', getValue: (r, esc) => esc(r.shopName || '') },
  { key: 'pageUrl', header: 'レビュー掲載URL', getValue: (r) => r.pageUrl || '' },
  { key: 'collectedAt', header: '収集日時', getValue: (r, esc) => esc(r.collectedAt || '') }
];

// Amazonのフィールド定義（順序固定）
const AMAZON_FIELD_DEFINITIONS = [
  { key: 'reviewDate', header: 'レビュー日', getValue: (r, esc) => esc(r.reviewDate || '') },
  { key: 'productId', header: 'ASIN', getValue: (r, esc) => esc(r.productId || '') },
  { key: 'productName', header: '商品名', getValue: (r, esc) => esc(r.productName || '') },
  { key: 'productUrl', header: '商品URL', getValue: (r) => r.productUrl || '' },
  { key: 'rating', header: '評価', getValue: (r) => r.rating || '' },
  { key: 'title', header: 'タイトル', getValue: (r, esc) => esc(r.title || '') },
  { key: 'body', header: '本文', getValue: (r, esc) => esc(r.body || '') },
  { key: 'author', header: '投稿者', getValue: (r, esc) => esc(r.author || '') },
  { key: 'variation', header: 'バリエーション', getValue: (r, esc) => esc(r.variation || '') },
  { key: 'helpfulCount', header: '参考になった数', getValue: (r) => r.helpfulCount || 0 },
  { key: 'country', header: '国', getValue: (r, esc) => esc(r.country || '') },
  { key: 'isVerified', header: '認証購入', getValue: (r) => r.isVerified ? '○' : '' },
  { key: 'isVine', header: 'Vineレビュー', getValue: (r) => r.isVine ? '○' : '' },
  { key: 'hasImage', header: '画像あり', getValue: (r) => r.hasImage ? '○' : '' },
  { key: 'pageUrl', header: 'レビュー掲載URL', getValue: (r) => r.pageUrl || '' },
  { key: 'collectedAt', header: '収集日時', getValue: (r, esc) => esc(r.collectedAt || '') }
];

// デフォルトの収集項目
const DEFAULT_RAKUTEN_FIELDS = ['rating', 'title', 'body', 'productUrl'];
const DEFAULT_AMAZON_FIELDS = ['rating', 'title', 'body', 'productUrl'];

/**
 * 選択されたフィールドからヘッダーとデータを生成
 */
function getSelectedFieldsData(reviews, source, selectedFields, escapeFormula) {
  const definitions = source === 'amazon' ? AMAZON_FIELD_DEFINITIONS : RAKUTEN_FIELD_DEFINITIONS;
  const defaultFields = source === 'amazon' ? DEFAULT_AMAZON_FIELDS : DEFAULT_RAKUTEN_FIELDS;
  const fields = selectedFields && selectedFields.length > 0 ? selectedFields : defaultFields;

  // 選択されたフィールドのみをフィルタリング（定義順を維持）
  const activeDefinitions = definitions.filter(def => fields.includes(def.key));

  // ヘッダー生成
  const headers = activeDefinitions.map(def => def.header);

  // データ行生成
  const dataValues = reviews.map(review =>
    activeDefinitions.map(def => def.getValue(review, escapeFormula))
  );

  return { headers, dataValues, columnCount: activeDefinitions.length };
}

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
  // 注意: Amazon商品URLは amazon.co.jp/商品名スラッグ/dp/ASIN の形式が多い
  // amazon.co.jp/dp/ だけでなく /dp/ を含むかで判定する
  const isAmazonProductPage = url.includes('amazon.co.jp') && (url.includes('/dp/') || url.includes('/gp/product/'));
  const isAmazonReviewPage = url.includes('amazon.co.jp') && url.includes('/product-reviews/');
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

// Google DriveフォルダIDからフォルダ名を取得
async function getDriveFolderName(folderId) {
  const token = await getAuthToken();
  if (!token) {
    throw new Error('認証が必要です');
  }

  const response = await fetch(
    `https://www.googleapis.com/drive/v3/files/${folderId}?fields=id,name&supportsAllDrives=true`,
    {
      headers: { 'Authorization': `Bearer ${token}` }
    }
  );

  if (!response.ok) {
    if (response.status === 404) {
      throw new Error('フォルダが見つかりません');
    }
    if (response.status === 403) {
      throw new Error('アクセス権限がありません');
    }
    throw new Error('フォルダ名の取得に失敗しました');
  }

  const data = await response.json();
  return data.name || '';
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

    case 'getDriveFolderName':
      getDriveFolderName(message.folderId)
        .then(name => sendResponse({ success: true, name }))
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

    case 'addRankingToProductQueue':
      addRankingToProductQueue(message.url, message.count)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'startBatchProductFromPopup':
      startBatchProductFromPopup()
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
      forwardToAll(message);
      break;

    case 'log':
      // content scriptからのログ → ストレージに保存 + UIに転送
      log(message.text, message.type || '', 'review');
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

    // ===== 商品情報収集（Google Drive保存） =====
    case 'collectAndSaveProductInfo':
      collectProductInfoWithMobile(message.tabId)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'startBatchProductCollection':
      startBatchProductCollection(message.items || message.asins)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'getBatchProductProgress':
      sendResponse({ progress: batchProductProgress });
      return false;

    case 'cancelBatchProductCollection':
      batchProductCancelled = true;
      sendResponse({ success: true });
      return false;

    // ===== 商品情報JSONダウンロード =====
    case 'downloadProductInfoJson':
      downloadProductInfoAsJson(message.tabId)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'startBatchJsonDownload':
      startBatchJsonDownload(message.items)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'getBatchJsonProgress':
      sendResponse({ progress: batchJsonProgress });
      return false;

    case 'cancelBatchJsonDownload':
      batchJsonCancelled = true;
      sendResponse({ success: true });
      return false;

    // ===== フォルダピッカー（Google Drive） =====
    case 'getDriveFolders':
      getDriveFolders(message.parentId, message.driveId)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'searchDriveFolders':
      searchDriveFolders(message.query, message.driveId)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'getSharedDrives':
      getSharedDrives()
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'getDriveSpreadsheets':
      getDriveSpreadsheets(message.parentId, message.driveId)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'createDriveFolder':
      createDriveFolder(message.name, message.parentId)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true;

    case 'getDriveFolderPath':
      getDriveFolderPath(message.folderId)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
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

  const sourceName = isAmazon ? 'Amazon' : '楽天';
  log(`${logPrefix} スプレッドシートへの保存を開始（販路: ${sourceName}）`);

  // スプレッドシートURLを取得
  let spreadsheetUrl = completedItem?.spreadsheetUrl || null;

  // URLが収集アイテムにない場合は設定から取得
  if (!spreadsheetUrl && !isScheduled) {
    const syncSettings = await chrome.storage.sync.get(['spreadsheetUrl', 'amazonSpreadsheetUrl']);
    spreadsheetUrl = isAmazon ? syncSettings.amazonSpreadsheetUrl : syncSettings.spreadsheetUrl;
  }

  if (!spreadsheetUrl) {
    log(`${logPrefix} スプレッドシートURLが未設定のため保存をスキップ`, 'error');
    return;
  }

  try {
    // ローカルストレージから全レビューを取得
    const stateResult = await chrome.storage.local.get(['collectionState']);
    const allReviews = stateResult.collectionState?.reviews || [];

    if (allReviews.length === 0) {
      log(`${logPrefix} 保存するレビューがありません`);
      return;
    }

    // この商品のレビューのみフィルタリング
    // 複数商品を並列収集している場合に他の商品のデータを含めないため
    const productReviews = productId
      ? allReviews.filter(r => r.productId === productId)
      : allReviews;

    log(`${logPrefix} 保存対象: ${productReviews.length}件`);

    if (productReviews.length === 0) {
      log(`${logPrefix} この商品のレビューが見つかりません`);
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
async function formatDataRows(token, spreadsheetId, sheetId, startRow, endRow, source = 'rakuten', columnCount = null) {
  // 列数が指定されていない場合は販路別のデフォルト値を使用
  const cols = columnCount || (source === 'amazon' ? 16 : 22);

  const requests = [
    {
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: startRow,
          endRowIndex: endRow,
          startColumnIndex: 0,
          endColumnIndex: cols
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
 * 特定の列に「書式なしテキスト」を適用（日付自動変換を防止）
 * @param {string} token - OAuth token
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {number} sheetId - シートID
 * @param {number} columnIndex - 列インデックス（0始まり）
 * @param {number} startRow - 開始行インデックス（0始まり）
 * @param {number} endRow - 終了行インデックス（0始まり、この行は含まない）
 */
async function formatColumnAsPlainText(token, spreadsheetId, sheetId, columnIndex, startRow, endRow) {
  const requests = [
    {
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: startRow,
          endRowIndex: endRow,
          startColumnIndex: columnIndex,
          endColumnIndex: columnIndex + 1
        },
        cell: {
          userEnteredFormat: {
            numberFormat: {
              type: 'TEXT'
            }
          }
        },
        fields: 'userEnteredFormat.numberFormat'
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
 * 選択されたフィールドから特定のフィールドの列インデックスを取得
 * @param {Array} selectedFields - 選択されたフィールドのキー配列
 * @param {string} fieldKey - 探すフィールドのキー
 * @param {string} source - 販路 ('rakuten' | 'amazon')
 * @returns {number} 列インデックス（見つからない場合は-1）
 */
function getFieldColumnIndex(selectedFields, fieldKey, source) {
  const definitions = source === 'amazon' ? AMAZON_FIELD_DEFINITIONS : RAKUTEN_FIELD_DEFINITIONS;
  const defaultFields = source === 'amazon' ? DEFAULT_AMAZON_FIELDS : DEFAULT_RAKUTEN_FIELDS;
  const fields = selectedFields && selectedFields.length > 0 ? selectedFields : defaultFields;

  // 選択されたフィールドのみをフィルタリング（定義順を維持）
  const activeFields = definitions.filter(def => fields.includes(def.key)).map(def => def.key);

  return activeFields.indexOf(fieldKey);
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
async function formatHeaderRow(token, spreadsheetId, sheetId, source = 'rakuten', columnCount = null) {
  // 販路別の色設定
  let backgroundColor, textColor;
  // 列数が指定されていない場合はデフォルト値を使用
  if (!columnCount) {
    columnCount = source === 'amazon' ? 16 : 22;
  }

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

  // 設定から選択されたフィールドを取得
  const settings = await chrome.storage.sync.get(['rakutenFields', 'amazonFields']);
  const selectedFields = source === 'amazon' ? settings.amazonFields : settings.rakutenFields;

  // 選択されたフィールドからヘッダーとデータを生成
  const { headers, dataValues, columnCount } = getSelectedFieldsData(reviews, source, selectedFields, escapeFormula);

  // ヘッダー + データを結合
  const allValues = [headers, ...dataValues];
  const totalRows = allValues.length;

  // 列数から最終列を計算（A=1, B=2, ..., Z=26）
  const getColumnLetter = (num) => {
    let letter = '';
    while (num > 0) {
      const remainder = (num - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      num = Math.floor((num - 1) / 26);
    }
    return letter;
  };
  const lastColumn = getColumnLetter(columnCount);

  // 1. シートの全データをクリア
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
  await formatHeaderRow(token, spreadsheetId, sheetId, source, columnCount);

  // 5. データ行に書式を適用（白背景・黒テキスト・垂直中央揃え）
  if (dataValues.length > 0) {
    await formatDataRows(token, spreadsheetId, sheetId, 1, totalRows, source, columnCount);
  }

  // 6. 評価列に「書式なしテキスト」を適用（日付自動変換を防止）
  const ratingColumnIndex = getFieldColumnIndex(selectedFields, 'rating', source);
  if (ratingColumnIndex >= 0) {
    await formatColumnAsPlainText(token, spreadsheetId, sheetId, ratingColumnIndex, 1, totalRows);
    console.log(`[appendToSheet] 評価列（${ratingColumnIndex}列目）に書式なしテキストを適用`);
  }

  // 7. URL列にクリック可能なリンク書式を適用（販路別）
  if (reviews.length > 0) {
    await formatUrlColumns(token, spreadsheetId, sheetId, reviews, source);
  }

  // 8. C列（商品名）の列幅をデータに合わせて自動調整
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

  // 9. 空白行を削除（シートの行数を調整）
  await trimEmptyRows(token, spreadsheetId, sheetId, encodedSheetName);

  // 10. シートの列数を設定した列数に調整（余分な列を削除）
  await adjustSheetColumns(token, spreadsheetId, sheetId, columnCount);
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

  // 設定から選択されたフィールドを取得
  const settings = await chrome.storage.sync.get(['rakutenFields', 'amazonFields']);
  const selectedFields = source === 'amazon' ? settings.amazonFields : settings.rakutenFields;
  console.log(`[appendToSheetWithoutClear] source: ${source}, selectedFields:`, selectedFields);

  // 選択されたフィールドからヘッダーとデータを生成
  const { headers, dataValues, columnCount } = getSelectedFieldsData(reviews, source, selectedFields, escapeFormula);
  console.log(`[appendToSheetWithoutClear] headers:`, headers);
  console.log(`[appendToSheetWithoutClear] 最初のデータ行:`, dataValues[0]);

  // 列数から最終列を計算
  const getColumnLetter = (num) => {
    let letter = '';
    while (num > 0) {
      const remainder = (num - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      num = Math.floor((num - 1) / 26);
    }
    return letter;
  };
  const lastColumn = getColumnLetter(columnCount);

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
    await formatHeaderRow(token, spreadsheetId, sheetId, source, columnCount);
    // データ行に書式を適用
    if (dataValues.length > 0) {
      await formatDataRows(token, spreadsheetId, sheetId, 1, allValues.length, source, columnCount);
    }
    // 評価列に「書式なしテキスト」を適用（日付自動変換を防止）
    const ratingColumnIndexEmpty = getFieldColumnIndex(selectedFields, 'rating', source);
    if (ratingColumnIndexEmpty >= 0) {
      await formatColumnAsPlainText(token, spreadsheetId, sheetId, ratingColumnIndexEmpty, 1, allValues.length);
      console.log(`[appendToSheetWithoutClear] 評価列（${ratingColumnIndexEmpty}列目）に書式なしテキストを適用`);
    }
  } else {
    // 既存データがある場合
    // 1. 既存のヘッダーを取得して比較
    const existingHeaderResponse = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodedSheetName}'!1:1`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const existingHeaderData = await existingHeaderResponse.json();
    const existingHeaders = existingHeaderData.values?.[0] || [];

    // ヘッダーが異なる場合は警告
    const headersMatch = headers.length === existingHeaders.length &&
      headers.every((h, i) => h === existingHeaders[i]);

    if (!headersMatch && existingHeaders.length > 0) {
      console.warn('[appendToSheetWithoutClear] 警告: 収集項目が変更されています。');
      console.warn('  既存ヘッダー:', existingHeaders);
      console.warn('  新しいヘッダー:', headers);
      console.warn('  既存データとの整合性が取れなくなる可能性があります。');
    }

    // ヘッダー行を上書き（書式も適用）
    await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodedSheetName}'!A1:${lastColumn}1?valueInputOption=USER_ENTERED`,
      {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values: [headers] })
      }
    );
    // ヘッダー書式を適用
    await formatHeaderRow(token, spreadsheetId, sheetId, source, columnCount);

    // 2. データを最後の行の次に追記
    const startRow = existingRows + 1;
    const endRow = startRow + dataValues.length - 1;

    // 必要な行数を確保（行が足りない場合は自動追加）
    await ensureSheetRows(token, spreadsheetId, sheetId, endRow);

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
      await formatDataRows(token, spreadsheetId, sheetId, startRow - 1, endRow, source, columnCount);
    }
  }

  // 評価列に「書式なしテキスト」を適用（日付自動変換を防止）
  const ratingColumnIndex = getFieldColumnIndex(selectedFields, 'rating', source);
  if (ratingColumnIndex >= 0) {
    // シート全体の評価列に適用（ヘッダー除く）
    const totalRows = existingRows + dataValues.length;
    await formatColumnAsPlainText(token, spreadsheetId, sheetId, ratingColumnIndex, 1, totalRows);
    console.log(`[appendToSheetWithoutClear] 評価列（${ratingColumnIndex}列目）に書式なしテキストを適用`);
  }

  // URL列にクリック可能なリンク書式を適用
  if (reviews.length > 0) {
    await formatUrlColumns(token, spreadsheetId, sheetId, reviews, source);
  }

  // シートの余分な行を削除（空白行をなくす）
  await trimEmptyRows(token, spreadsheetId, sheetId, encodedSheetName);

  // シートの列数を設定した列数に調整（余分な列を削除）
  await adjustSheetColumns(token, spreadsheetId, sheetId, columnCount);
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
 * シートの行数を確保（必要に応じて行を追加）
 */
async function ensureSheetRows(token, spreadsheetId, sheetId, requiredRows) {
  try {
    // シートの現在の行数を取得
    const sheetPropsResponse = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets(properties)`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const sheetPropsData = await sheetPropsResponse.json();
    const sheetProps = sheetPropsData.sheets?.find(s => s.properties.sheetId === sheetId);
    const currentRowCount = sheetProps?.properties?.gridProperties?.rowCount || 1000;

    console.log(`[ensureSheetRows] 現在の行数: ${currentRowCount}, 必要な行数: ${requiredRows}`);

    // 必要な行数が現在の行数より多い場合、行を追加
    if (requiredRows > currentRowCount) {
      const rowsToAdd = requiredRows - currentRowCount + 100; // 余裕を持って100行追加
      console.log(`[ensureSheetRows] ${rowsToAdd}行を追加します`);

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
              appendDimension: {
                sheetId: sheetId,
                dimension: 'ROWS',
                length: rowsToAdd
              }
            }]
          })
        }
      );
      console.log(`[ensureSheetRows] 行追加完了`);
    }
  } catch (e) {
    console.error('[ensureSheetRows] エラー:', e);
    throw e;
  }
}

/**
 * シートの列数を設定した列数に調整（余分な列を削除）
 */
async function adjustSheetColumns(token, spreadsheetId, sheetId, requiredColumns) {
  try {
    // シートの現在の列数を取得
    const sheetPropsResponse = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets(properties)`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const sheetPropsData = await sheetPropsResponse.json();
    const sheetProps = sheetPropsData.sheets?.find(s => s.properties.sheetId === sheetId);
    const currentColumnCount = sheetProps?.properties?.gridProperties?.columnCount || 26;

    console.log(`[adjustSheetColumns] 現在の列数: ${currentColumnCount}, 必要な列数: ${requiredColumns}`);

    // 余分な列がある場合は削除
    if (currentColumnCount > requiredColumns) {
      console.log(`[adjustSheetColumns] ${currentColumnCount - requiredColumns}列を削除します`);

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
                  dimension: 'COLUMNS',
                  startIndex: requiredColumns,
                  endIndex: currentColumnCount
                }
              }
            }]
          })
        }
      );
      console.log(`[adjustSheetColumns] 列削除完了`);
    }
    // 列が足りない場合は追加
    else if (currentColumnCount < requiredColumns) {
      const columnsToAdd = requiredColumns - currentColumnCount;
      console.log(`[adjustSheetColumns] ${columnsToAdd}列を追加します`);

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
              appendDimension: {
                sheetId: sheetId,
                dimension: 'COLUMNS',
                length: columnsToAdd
              }
            }]
          })
        }
      );
      console.log(`[adjustSheetColumns] 列追加完了`);
    }
  } catch (e) {
    console.error('[adjustSheetColumns] エラー:', e);
    // エラーが発生しても処理を続行（列調整は必須ではない）
  }
}

/**
 * CSVダウンロード（販路別ファイル名）
 */
async function handleDownloadCSV() {
  return new Promise((resolve, reject) => {
    // 収集状態と選択フィールドを取得
    chrome.storage.local.get(['collectionState'], async (localResult) => {
      const state = localResult.collectionState;

      if (!state || !state.reviews || state.reviews.length === 0) {
        reject(new Error('ダウンロードするデータがありません'));
        return;
      }

      try {
        // 選択されたフィールドを取得
        const syncResult = await chrome.storage.sync.get(['rakutenFields', 'amazonFields']);
        const source = state.source || 'rakuten';
        const selectedFields = source === 'amazon'
          ? (syncResult.amazonFields || DEFAULT_AMAZON_FIELDS)
          : (syncResult.rakutenFields || DEFAULT_RAKUTEN_FIELDS);

        const csv = convertToCSV(state.reviews, selectedFields);
        const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);

        // 販路別のファイル名
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
 * レビューデータをCSV形式に変換（選択項目のみ）
 */
function convertToCSV(reviews, selectedFields) {
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

  // getSelectedFieldsDataを使用して動的にヘッダーとデータを生成
  // CSVではescapeCSVをそのまま使用（数式エスケープ不要）
  const { headers, dataValues } = getSelectedFieldsData(reviews, source, selectedFields, escapeCSV);

  return [
    headers.map(escapeCSV).join(','),
    ...dataValues.map(row => row.map(escapeCSV).join(','))
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
    // UIにキュー収集完了を通知（ポップアップのボタン状態リセット用）
    forwardToAll({ action: 'queueCollectionComplete' });
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
 * 楽天ランキングHTMLから商品リストを抽出（キュー操作なし）
 * @returns {Array} [{url, title, source: 'rakuten'}]
 */
async function parseRakutenRankingPage(url, count) {
  const response = await fetch(url);
  const html = await response.text();

  const products = [];
  const seenUrls = new Set();

  const linkPattern = /href="(https?:\/\/item\.rakuten\.co\.jp\/[^"]+)"/g;
  let match;

  while ((match = linkPattern.exec(html)) !== null && products.length < count) {
    try {
      const urlObj = new URL(match[1]);
      const cleanUrl = `${urlObj.origin}${urlObj.pathname}`;
      if (seenUrls.has(cleanUrl)) continue;
      seenUrls.add(cleanUrl);

      const pathMatch = cleanUrl.match(/item\.rakuten\.co\.jp\/([^\/]+)\/([^\/]+)/);
      const title = pathMatch ? `${pathMatch[1]} - ${pathMatch[2]}` : '商品';

      products.push({
        url: cleanUrl,
        title: title.substring(0, 100),
        addedAt: new Date().toISOString(),
        source: 'rakuten'
      });
    } catch (e) {
      continue;
    }
  }

  return products;
}

/**
 * AmazonランキングHTMLからASIN/商品リストを抽出（キュー操作なし）
 * @returns {Array} [{url, title, asin, source: 'amazon'}]
 */
async function parseAmazonRankingPage(url, count) {
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

  // data-asin属性からASINを抽出
  const asinPattern = /data-asin="([A-Z0-9]{10})"/g;
  let match;
  while ((match = asinPattern.exec(html)) !== null && products.length < count) {
    const asin = match[1].toUpperCase();
    if (!asin || seenAsins.has(asin)) continue;
    seenAsins.add(asin);
    products.push({
      url: `https://www.amazon.co.jp/dp/${asin}`,
      title: asin,
      asin,
      addedAt: new Date().toISOString(),
      source: 'amazon'
    });
  }

  // フォールバック: /dp/ASINパターン
  if (products.length < count) {
    const dpPattern = /\/dp\/([A-Z0-9]{10})/gi;
    while ((match = dpPattern.exec(html)) !== null && products.length < count) {
      const asin = match[1].toUpperCase();
      if (!asin || seenAsins.has(asin)) continue;
      seenAsins.add(asin);
      products.push({
        url: `https://www.amazon.co.jp/dp/${asin}`,
        title: asin,
        asin,
        addedAt: new Date().toISOString(),
        source: 'amazon'
      });
    }
  }

  return products;
}

/**
 * 楽天ランキングからレビューキューに追加
 */
async function fetchRakutenRankingProducts(url, count) {
  try {
    const products = await parseRakutenRankingPage(url, count);
    if (products.length === 0) {
      return { success: false, error: '商品が見つかりませんでした' };
    }

    const result = await chrome.storage.local.get(['queue']);
    const queue = result.queue || [];
    let addedCount = 0;
    for (const product of products) {
      if (!queue.some(item => item.url === product.url)) {
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
 * Amazonランキングからレビューキューに追加
 */
async function fetchAmazonRankingProducts(url, count) {
  try {
    const products = await parseAmazonRankingPage(url, count);
    if (products.length === 0) {
      return { success: false, error: '商品が見つかりませんでした' };
    }

    const result = await chrome.storage.local.get(['queue']);
    const queue = result.queue || [];
    let addedCount = 0;
    for (const product of products) {
      if (!queue.some(item => item.url === product.url)) {
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
 * ランキングから商品キュー（batchProductQueue）にASINを追加
 * Amazon: ASINを直接追加
 * 楽天: 商品URLからASINは取得不可のため、URLベースで追加（商品情報収集はスキップ）
 */
async function addRankingToProductQueue(url, count) {
  const isAmazon = url.includes('amazon.co.jp');
  const isRakuten = url.includes('ranking.rakuten.co.jp');

  if (!isAmazon && !isRakuten) {
    return { success: false, error: 'Amazon・楽天ランキングページのみ対応しています' };
  }

  try {
    if (isAmazon) {
      return await addAmazonRankingToProductQueue(url, count);
    } else {
      return await addRakutenRankingToProductQueue(url, count);
    }
  } catch (error) {
    log(`商品キュー追加エラー: ${error.message}`, 'error', 'product');
    return { success: false, error: error.message };
  }
}

/**
 * Amazonランキングから商品キュー(batchProductQueue)にASINを追加
 * ※ レビューキュー(queue)には触れない
 */
async function addAmazonRankingToProductQueue(url, count) {
  const products = await parseAmazonRankingPage(url, count);
  if (products.length === 0) {
    return { success: false, error: '商品が見つかりませんでした' };
  }

  const result = await chrome.storage.local.get(['batchProductQueue']);
  const queue = result.batchProductQueue || [];
  let addedCount = 0;
  const addedItems = [];
  for (const product of products) {
    const asin = product.asin;
    if (!queue.includes(asin)) {
      queue.push(asin);
      addedItems.push(asin);
      addedCount++;
    }
  }

  await chrome.storage.local.set({ batchProductQueue: queue });
  forwardToAll({ action: 'batchProductQueueUpdated' });

  if (addedCount > 0) {
    log(`Amazonランキングから${addedCount}件を商品キューに追加`, 'success', 'product');
  } else {
    log('追加する商品がありません（全て重複）', 'info', 'product');
  }

  return { success: true, addedCount, addedItems, totalCount: queue.length };
}

/**
 * 楽天ランキングから商品キュー(batchProductQueue)に楽天URLを追加
 * ※ レビューキュー(queue)には触れない
 */
async function addRakutenRankingToProductQueue(url, count) {
  const products = await parseRakutenRankingPage(url, count);
  if (products.length === 0) {
    return { success: false, error: '商品が見つかりませんでした' };
  }

  const result = await chrome.storage.local.get(['batchProductQueue']);
  const queue = result.batchProductQueue || [];
  const existingUrls = new Set(queue);
  let addedCount = 0;
  const addedItems = [];
  for (const product of products) {
    if (!existingUrls.has(product.url)) {
      queue.push(product.url);
      existingUrls.add(product.url);
      addedItems.push(product.url);
      addedCount++;
    }
  }

  await chrome.storage.local.set({ batchProductQueue: queue });
  forwardToAll({ action: 'batchProductQueueUpdated' });

  if (addedCount > 0) {
    log(`楽天ランキングから${addedCount}件を商品キューに追加`, 'success', 'product');
  } else {
    log('追加する商品がありません（全て重複）', 'info', 'product');
  }

  return { success: true, addedCount, addedItems, totalCount: queue.length };
}

/**
 * ポップアップから商品キューの商品情報バッチ収集を開始
 * batchProductQueueに入っているASIN・楽天URLを使ってバッチ収集を実行
 */
async function startBatchProductFromPopup() {
  const result = await chrome.storage.local.get(['batchProductQueue']);
  const queue = result.batchProductQueue || [];

  if (queue.length === 0) {
    return { success: false, error: '商品キューが空です' };
  }

  // キューからASINまたはURLを抽出
  const items = queue.map(item => {
    if (typeof item === 'string') return item;
    return item.url || item.asin || item;
  }).filter(a => {
    if (!a) return false;
    // 楽天URL
    if (a.includes('item.rakuten.co.jp')) return true;
    // Amazon ASIN
    if (/^[A-Z0-9]{10}$/i.test(a)) return true;
    return false;
  });

  if (items.length === 0) {
    return { success: false, error: '有効な商品が見つかりません' };
  }

  return startBatchProductCollection(items);
}

/**
 * すべてのページにメッセージを転送
 */
function forwardToAll(message) {
  chrome.runtime.sendMessage(message).catch(() => {});
}

/**
 * ログをストレージに保存
 * バッファ方式: 複数のlog()呼び出しを200msごとにまとめて1回で書き込み
 * → get→push→setの競合を構造的に排除
 */
const _logBuffer = [];
let _logFlushTimer = null;

function log(text, type = '', category = 'review') {
  console.log(`[収集] ${text}`);

  const time = new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  const storageKey = category === 'product' ? 'productLogs' : 'logs';

  _logBuffer.push({ entry: { time, text, type }, storageKey });

  // バッファを200ms後にまとめてフラッシュ（初回即時実行）
  if (!_logFlushTimer) {
    _logFlushTimer = setTimeout(flushLogBuffer, 200);
  }

  // UIに直接送信（即時表示用 — ストレージ書込みを待たない）
  forwardToAll({ action: 'appendLog', entry: { time, text, type }, category });
}

function flushLogBuffer() {
  _logFlushTimer = null;
  if (_logBuffer.length === 0) return;

  // バッファからエントリーを取り出し、カテゴリ別にグループ化
  const entries = _logBuffer.splice(0);
  const grouped = {};
  for (const item of entries) {
    if (!grouped[item.storageKey]) grouped[item.storageKey] = [];
    grouped[item.storageKey].push(item.entry);
  }

  // カテゴリごとに1回だけget→push→set（競合なし）
  for (const [storageKey, newEntries] of Object.entries(grouped)) {
    chrome.storage.local.get([storageKey], (result) => {
      const logs = result[storageKey] || [];
      logs.push(...newEntries);
      chrome.storage.local.set({ [storageKey]: logs });
    });
  }
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
    logs: [],
    productLogs: []
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

// ========================================
// 商品情報収集機能（Google Drive保存）
// ========================================

// バッチ処理の進捗
let batchProductProgress = null;
let batchProductCancelled = false;

/**
 * 画像URLをfetchしてbase64に変換
 * @param {string} imageUrl - 画像URL
 * @returns {Promise<{base64: string, mimeType: string, size: number}>}
 */
async function fetchImageAsBase64(imageUrl) {
  try {
    // 画像URLのドメインに応じたRefererを設定
    let referer = 'https://www.amazon.co.jp/';
    if (imageUrl.includes('rakuten.co.jp') || imageUrl.includes('r10s.jp')) {
      referer = 'https://item.rakuten.co.jp/';
    }

    const response = await fetch(imageUrl, {
      headers: {
        'Accept': 'image/*',
        'Referer': referer
      }
    });

    if (!response.ok) {
      console.warn(`[商品情報] 画像取得失敗: ${response.status} - ${imageUrl.substring(0, 80)}`);
      return null;
    }

    const blob = await response.blob();
    const mimeType = blob.type || 'image/jpeg';

    // BlobをArrayBufferに変換してbase64化
    const arrayBuffer = await blob.arrayBuffer();
    const uint8Array = new Uint8Array(arrayBuffer);
    let binary = '';
    for (let i = 0; i < uint8Array.length; i++) {
      binary += String.fromCharCode(uint8Array[i]);
    }
    const base64 = btoa(binary);

    return {
      base64,
      mimeType,
      size: blob.size
    };
  } catch (error) {
    console.warn(`[商品情報] 画像取得エラー: ${error.message} - ${imageUrl.substring(0, 80)}`);
    return null;
  }
}

/**
 * 画像/動画URLからBlobを取得（base64変換なし）
 * @param {string} mediaUrl - メディアURL
 * @returns {Promise<{blob: Blob, mimeType: string, size: number}|null>}
 */
async function fetchMediaAsBlob(mediaUrl) {
  try {
    // blob: URLはbackground.jsからアクセスできない
    if (mediaUrl.startsWith('blob:')) {
      console.warn(`[商品情報] blob URLはダウンロード不可: ${mediaUrl.substring(0, 80)}`);
      return null;
    }

    let referer = 'https://www.amazon.co.jp/';
    if (mediaUrl.includes('rakuten.co.jp') || mediaUrl.includes('r10s.jp') || mediaUrl.includes('rakuten.ne.jp')) {
      referer = 'https://item.rakuten.co.jp/';
    }

    const response = await fetch(mediaUrl, {
      headers: {
        'Accept': 'image/*,video/*,*/*',
        'Referer': referer
      }
    });

    if (!response.ok) {
      console.warn(`[商品情報] メディア取得失敗: ${response.status} - ${mediaUrl.substring(0, 80)}`);
      return null;
    }

    const blob = await response.blob();
    const mimeType = blob.type || 'application/octet-stream';

    return { blob, mimeType, size: blob.size };
  } catch (error) {
    console.warn(`[商品情報] メディア取得エラー: ${error.message} - ${mediaUrl.substring(0, 80)}`);
    return null;
  }
}

/**
 * MIMEタイプからファイル拡張子を推定
 */
function getExtensionFromMimeType(mimeType, url = '') {
  const mimeMap = {
    'image/jpeg': 'jpg',
    'image/png': 'png',
    'image/gif': 'gif',
    'image/webp': 'webp',
    'image/svg+xml': 'svg',
    'video/mp4': 'mp4',
    'video/webm': 'webm',
    'video/quicktime': 'mov'
  };
  if (mimeMap[mimeType]) return mimeMap[mimeType];

  // URLから拡張子を推定
  const urlMatch = url.match(/\.([a-zA-Z0-9]+)(?:\?|$)/);
  if (urlMatch) {
    const ext = urlMatch[1].toLowerCase();
    if (['jpg', 'jpeg', 'png', 'gif', 'webp', 'mp4', 'webm', 'mov'].includes(ext)) {
      return ext === 'jpeg' ? 'jpg' : ext;
    }
  }

  return 'jpg'; // デフォルト
}

/**
 * バイナリメディアファイルをGoogle Driveにアップロード
 * @param {string} token - OAuthトークン
 * @param {string} folderId - 保存先フォルダID
 * @param {string} fileName - ファイル名
 * @param {Blob} blob - バイナリデータ
 * @param {string} mimeType - MIMEタイプ
 * @returns {Promise<Object>} {id, name}
 */
async function uploadMediaToDrive(token, folderId, fileName, blob, mimeType) {
  // 5MB未満: multipart upload, 5MB以上: resumable upload
  if (blob.size < 5 * 1024 * 1024) {
    const metadata = JSON.stringify({
      name: fileName,
      mimeType: mimeType,
      parents: [folderId]
    });

    const boundary = '-------mediaboundary' + Date.now();
    const encoder = new TextEncoder();

    // マルチパートボディをバイナリで構築
    const metadataPart = encoder.encode(
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${metadata}\r\n--${boundary}\r\nContent-Type: ${mimeType}\r\nContent-Transfer-Encoding: binary\r\n\r\n`
    );
    const closingPart = encoder.encode(`\r\n--${boundary}--`);
    const blobBuffer = await blob.arrayBuffer();

    const body = new Uint8Array(metadataPart.length + blobBuffer.byteLength + closingPart.length);
    body.set(metadataPart, 0);
    body.set(new Uint8Array(blobBuffer), metadataPart.length);
    body.set(closingPart, metadataPart.length + blobBuffer.byteLength);

    const response = await fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name&supportsAllDrives=true',
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': `multipart/related; boundary=${boundary}`
        },
        body: body.buffer
      }
    );

    if (!response.ok) {
      const error = await response.json().catch(() => ({}));
      throw new Error(error.error?.message || `メディアアップロードエラー: ${response.status}`);
    }

    return await response.json();
  } else {
    // 大きいファイル: resumable upload
    const initResponse = await fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id,name&supportsAllDrives=true',
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
          'X-Upload-Content-Type': mimeType,
          'X-Upload-Content-Length': blob.size
        },
        body: JSON.stringify({
          name: fileName,
          mimeType: mimeType,
          parents: [folderId]
        })
      }
    );

    if (!initResponse.ok) {
      throw new Error(`メディアアップロード開始エラー: ${initResponse.status}`);
    }

    const uploadUrl = initResponse.headers.get('Location');
    const uploadResponse = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Type': mimeType,
        'Content-Length': blob.size
      },
      body: blob
    });

    if (!uploadResponse.ok) {
      throw new Error(`メディアアップロードエラー: ${uploadResponse.status}`);
    }

    return await uploadResponse.json();
  }
}

/**
 * Google DriveフォルダURLからフォルダIDを抽出
 */
function extractDriveFolderId(url) {
  if (!url) return null;
  // https://drive.google.com/drive/folders/FOLDER_ID
  const match = url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Google DriveにJSONファイルをアップロード
 * @param {string} token - OAuthトークン
 * @param {string} folderId - 保存先フォルダID
 * @param {string} fileName - ファイル名
 * @param {Object} jsonData - 保存するデータ
 * @returns {Promise<Object>} アップロード結果
 */
async function uploadJsonToDrive(token, folderId, fileName, jsonData) {
  const jsonString = JSON.stringify(jsonData, null, 2);
  const blob = new Blob([jsonString], { type: 'application/json' });

  // ファイルサイズが5MB未満の場合はシンプルアップロード
  if (blob.size < 5 * 1024 * 1024) {
    return await simpleUploadToDrive(token, folderId, fileName, blob);
  } else {
    return await resumableUploadToDrive(token, folderId, fileName, blob);
  }
}

/**
 * シンプルアップロード（5MB未満）
 */
async function simpleUploadToDrive(token, folderId, fileName, blob) {
  // multipart/related でメタデータとファイルを一緒に送信
  const metadata = {
    name: fileName,
    mimeType: 'application/json',
    parents: [folderId]
  };

  const boundary = '-------314159265358979323846';
  const delimiter = `\r\n--${boundary}\r\n`;
  const closeDelimiter = `\r\n--${boundary}--`;

  // ArrayBufferからテキストを作成
  const fileContent = await blob.text();

  const multipartBody =
    delimiter +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(metadata) +
    delimiter +
    'Content-Type: application/json\r\n\r\n' +
    fileContent +
    closeDelimiter;

  const response = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true',
    {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': `multipart/related; boundary=${boundary}`
      },
      body: multipartBody
    }
  );

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.error?.message || `Drive APIエラー: ${response.status}`);
  }

  return await response.json();
}

/**
 * 分割アップロード（5MB以上）
 */
async function resumableUploadToDrive(token, folderId, fileName, blob) {
  // 1. アップロードセッションを開始
  const metadata = {
    name: fileName,
    mimeType: 'application/json',
    parents: [folderId]
  };

  const initResponse = await fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true',
    {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'X-Upload-Content-Type': 'application/json',
        'X-Upload-Content-Length': blob.size
      },
      body: JSON.stringify(metadata)
    }
  );

  if (!initResponse.ok) {
    throw new Error(`分割アップロード開始エラー: ${initResponse.status}`);
  }

  const uploadUrl = initResponse.headers.get('Location');

  // 2. ファイルデータをアップロード
  const uploadResponse = await fetch(uploadUrl, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/json',
      'Content-Length': blob.size
    },
    body: blob
  });

  if (!uploadResponse.ok) {
    throw new Error(`分割アップロードエラー: ${uploadResponse.status}`);
  }

  return await uploadResponse.json();
}

// ===== PDF生成用: offscreenドキュメント管理 =====

let offscreenCreated = false;

/**
 * offscreenドキュメントを作成（既に存在する場合はスキップ）
 */
async function ensureOffscreenDocument() {
  if (offscreenCreated) return;
  try {
    // 既存のoffscreenドキュメントがあるか確認
    const existingContexts = await chrome.runtime.getContexts({
      contextTypes: ['OFFSCREEN_DOCUMENT']
    });
    if (existingContexts.length > 0) {
      offscreenCreated = true;
      return;
    }
  } catch (e) {
    // getContexts未対応の古いChromeの場合は無視
  }

  try {
    await chrome.offscreen.createDocument({
      url: 'offscreen.html',
      reasons: ['CANVAS'],
      justification: 'WebP→JPEG変換と画像リサイズのためCanvasが必要'
    });
    offscreenCreated = true;
  } catch (e) {
    if (!e.message.includes('already exists')) {
      throw e;
    }
    offscreenCreated = true;
  }
}

/**
 * offscreenドキュメントを閉じる
 */
async function closeOffscreenDocument() {
  if (!offscreenCreated) return;
  try {
    await chrome.offscreen.closeDocument();
  } catch (e) {
    // 既に閉じている場合は無視
  }
  offscreenCreated = false;
}

/**
 * 画像をPDF用に処理（WebP変換 / リサイズ）
 * offscreenドキュメント経由でCanvas APIを使用
 * @param {string} base64 - Base64エンコードされた画像データ
 * @param {string} mimeType - 元のMIMEタイプ
 * @returns {Promise<{base64: string, mimeType: string, width: number, height: number}>}
 */
async function processImageForPdf(base64, mimeType) {
  await ensureOffscreenDocument();
  const response = await chrome.runtime.sendMessage({
    action: 'processImageForPdf',
    imageBase64: base64,
    mimeType: mimeType
  });
  if (response?.error) {
    throw new Error(`画像処理エラー: ${response.error}`);
  }
  return response;
}

/**
 * 画像URLをフェッチしてPDF埋め込み用のUint8Arrayに変換
 * WebP画像は自動でJPEGに変換、幅1500px超は1200pxにリサイズ
 * @param {string} imageUrl - 画像URL
 * @returns {Promise<{data: Uint8Array, mimeType: string}|null>}
 */
async function fetchImageForPdf(imageUrl) {
  try {
    // 画像をBlobとして取得
    const mediaResult = await fetchMediaAsBlob(imageUrl);
    if (!mediaResult) return null;

    const isWebP = mediaResult.mimeType === 'image/webp';
    const isJpeg = mediaResult.mimeType === 'image/jpeg' || mediaResult.mimeType === 'image/jpg';
    const isPng = mediaResult.mimeType === 'image/png';

    // JPG/PNGでサイズが小さい場合はそのまま使用（offscreen不要）
    // 500KB以下ならリサイズ不要と判断（ヒューリスティック）
    if ((isJpeg || isPng) && mediaResult.size <= 500 * 1024) {
      const arrayBuffer = await mediaResult.blob.arrayBuffer();
      return { data: new Uint8Array(arrayBuffer), mimeType: mediaResult.mimeType };
    }

    // WebP or 大きい画像 → offscreenで処理
    // Blob → base64に変換
    const arrayBuffer = await mediaResult.blob.arrayBuffer();
    const bytes = new Uint8Array(arrayBuffer);

    // 大きい画像（JPG/PNG 500KB超）もoffscreenで処理（リサイズ判定のため）
    if (isWebP || mediaResult.size > 500 * 1024) {
      let base64 = '';
      // Uint8Array → base64（チャンク分割でスタックオーバーフロー回避）
      const chunkSize = 8192;
      for (let i = 0; i < bytes.length; i += chunkSize) {
        const chunk = bytes.subarray(i, i + chunkSize);
        base64 += String.fromCharCode.apply(null, chunk);
      }
      base64 = btoa(base64);

      const processed = await processImageForPdf(base64, mediaResult.mimeType);

      // base64 → Uint8Array
      const processedBinary = atob(processed.base64);
      const processedBytes = new Uint8Array(processedBinary.length);
      for (let i = 0; i < processedBinary.length; i++) {
        processedBytes[i] = processedBinary.charCodeAt(i);
      }
      return { data: processedBytes, mimeType: processed.mimeType };
    }

    // その他（GIF等のサポート外形式）→ そのまま返す（pdf-libで埋め込めないがエラーページで処理）
    return { data: bytes, mimeType: mediaResult.mimeType };
  } catch (error) {
    console.warn(`[PDF] 画像取得/処理エラー: ${error.message} - ${imageUrl}`);
    return null;
  }
}

/**
 * 商品情報を収集してGoogle Driveに保存（PDF形式）
 * - 画像はPDFにまとめて保存（個別ファイルはアップロードしない）
 * - 動画のサムネイルもPDFに含める
 * - JSONにはテキスト + メディアメタデータのみ保存
 * @param {number} tabId - 対象タブのID
 */
async function collectAndSaveProductInfo(tabId, mode = 'desktop', mobileOptions = null) {
  const isMobile = mode === 'mobile';
  // modeLabelはAmazon/楽天判定後に設定（Amazon: ラベルなし、楽天: PC版/スマホ版）
  let modeLabel = '';
  log('収集を開始します...', '', 'product');

  // 1. 設定チェック（スマホ版は既存の商品フォルダを使うため親フォルダ不要）
  let parentFolderId = null;
  if (!isMobile || !mobileOptions?.parentFolderId) {
    const settings = await chrome.storage.sync.get(['productInfoFolderUrl']);
    const folderUrl = settings.productInfoFolderUrl;
    if (!folderUrl) {
      throw new Error('Google Driveの保存先フォルダが設定されていません。管理画面の設定で「商品情報の保存先フォルダURL」を設定してください。');
    }
    parentFolderId = extractDriveFolderId(folderUrl);
    if (!parentFolderId) {
      throw new Error('Google DriveのフォルダURLが正しくありません。');
    }
  }

  // 2. OAuthトークンを取得（未認証なら対話型ダイアログを表示）
  const token = await getAuthTokenWithFallback();

  // 3. content scriptから商品情報を取得（メッセージチャンネル切断時は1回リトライ）
  log('ページから情報を読み取り中...', '', 'product');
  let productData;
  const maxRetries = 2;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = await chrome.tabs.sendMessage(tabId, { action: 'collectProductInfo' });
      if (!response || !response.success) {
        throw new Error(response?.error || '商品情報の取得に失敗しました');
      }
      productData = response.data;
      break;
    } catch (error) {
      if (attempt < maxRetries && error.message.includes('message channel closed')) {
        console.warn(`[商品情報] メッセージチャンネル切断、${attempt}回目のリトライ...`);
        await sleep(2000);
        continue;
      }
      throw new Error(`商品情報の取得に失敗: ${error.message}`);
    }
  }

  // Amazon / 楽天の判定
  const isRakuten = productData.source === 'rakuten';
  // AmazonはPC/SP同一データのためラベルなし、楽天はPC版/スマホ版を表示
  modeLabel = isRakuten ? (isMobile ? '[スマホ版]' : '[PC版]') : '';
  const productId = isRakuten
    ? (productData.itemSlug || '???')
    : (productData.asin || '???');

  const dateStr = formatDate(new Date());
  const baseName = (isMobile && mobileOptions?.baseName)
    ? mobileOptions.baseName
    : (isRakuten
      ? `rakuten_${productData.itemSlug}_${dateStr}`
      : `amazon_${productData.asin}_${dateStr}`);

  // 4. 商品フォルダを作成（サブフォルダなし、フラット構造）
  let productFolderId; // 商品フォルダID（JSONとPDFの保存先）
  if (isMobile && mobileOptions?.parentFolderId) {
    // 楽天スマホ版: 商品フォルダは既にPC版で作成済み → 同じフォルダを使用
    productFolderId = mobileOptions.parentFolderId;
  } else {
    // 新規: 商品フォルダを作成
    log(`[${productId}]${modeLabel} 保存フォルダを作成中...`, '', 'product');
    const productFolderResult = await createDriveFolder(baseName, parentFolderId);
    productFolderId = productFolderResult.folder.id;
  }

  // 5. 画像をダウンロード → PDF生成 → Driveにアップロード
  const aplusImgUrls = isRakuten ? [] : (productData.aplusImages || []);
  const rawVideos = productData.videos || [];
  const videoThumbs = productData.videoThumbnails || [];
  const totalImages = productData.images.length + aplusImgUrls.length;
  log(`[${productId}]${modeLabel} 画像を取得中...（${totalImages}枚）`, '', 'product');

  // --- 画像をフェッチしてPDF用データに変換（2件並列） ---
  const pdfImages = []; // PDF埋め込み用: {data, mimeType, section, order, originalUrl}
  const imageMetadata = []; // JSONメタデータ用

  // 商品画像
  const imgTasks = productData.images.map((img, i) => ({
    order: i + 1,
    url: typeof img === 'string' ? img : img.url,
    section: typeof img === 'string' ? 'product' : (img.type || 'product')
  }));

  for (let i = 0; i < imgTasks.length; i += 2) {
    const chunk = imgTasks.slice(i, i + 2);
    const results = await Promise.all(chunk.map(async (task) => {
      const imgData = await fetchImageForPdf(task.url);
      if (imgData) {
        return {
          data: imgData.data,
          mimeType: imgData.mimeType,
          section: task.section,
          order: task.order,
          originalUrl: task.url
        };
      }
      return null;
    }));

    for (const r of results) {
      if (r) {
        pdfImages.push(r);
        imageMetadata.push({
          order: r.order,
          section: r.section,
          description: getSectionDescription(r.section, r.order, isRakuten),
          originalUrl: r.originalUrl,
          mimeType: r.mimeType
        });
      }
    }
    const imgDone = Math.min(i + 2, imgTasks.length);
    log(`[${productId}]${modeLabel} 画像 ${imgDone}/${imgTasks.length} 取得中...`, '', 'product');
    forwardToAll({ action: 'productInfoProgress', progress: { phase: 'images', current: imgDone, total: totalImages, productId } });
  }

  // A+コンテンツ画像（Amazonのみ）
  const aplusImageMetadata = [];
  if (!isRakuten && aplusImgUrls.length > 0) {
    for (let i = 0; i < aplusImgUrls.length; i += 2) {
      const chunk = aplusImgUrls.slice(i, i + 2);
      const results = await Promise.all(chunk.map(async (imgUrl, j) => {
        const order = i + j + 1;
        const imgData = await fetchImageForPdf(imgUrl);
        if (imgData) {
          return {
            data: imgData.data,
            mimeType: imgData.mimeType,
            section: 'aplus',
            order,
            originalUrl: imgUrl
          };
        }
        return null;
      }));

      for (const r of results) {
        if (r) {
          pdfImages.push(r);
          aplusImageMetadata.push({
            order: r.order,
            description: `A+コンテンツ画像${r.order}枚目`,
            originalUrl: r.originalUrl,
            mimeType: r.mimeType
          });
        }
      }
      const aplusDone = Math.min(i + 2, aplusImgUrls.length);
      log(`[${productId}]${modeLabel} A+画像 ${aplusDone}/${aplusImgUrls.length} 取得中...`, '', 'product');
      forwardToAll({ action: 'productInfoProgress', progress: { phase: 'images', current: imgTasks.length + aplusDone, total: totalImages, productId } });
    }
  }

  // --- PDF生成 ---
  let pdfFileName = null;
  let pdfDriveFileId = null;
  if (pdfImages.length > 0) {
    log(`[${productId}]${modeLabel} PDF生成中...（${pdfImages.length}枚）`, '', 'product');
    forwardToAll({ action: 'productInfoProgress', progress: { phase: 'pdf', productId } });

    try {
      const pdfBytes = await generateProductImagesPDF(pdfImages, {
        productId,
        source: isRakuten ? 'rakuten' : 'amazon',
        baseName,
        mode
      });

      // PDFファイル名: 楽天はPC/SP別、Amazonは_images
      if (isRakuten) {
        pdfFileName = `${baseName}_${isMobile ? 'sp' : 'pc'}.pdf`;
      } else {
        pdfFileName = `${baseName}_images.pdf`;
      }

      const pdfSizeMB = (pdfBytes.length / 1024 / 1024).toFixed(1);
      log(`[${productId}]${modeLabel} PDFをアップロード中...（${pdfSizeMB}MB）`, '', 'product');
      forwardToAll({ action: 'productInfoProgress', progress: { phase: 'upload', productId } });

      const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
      const pdfUploaded = await uploadMediaToDrive(token, productFolderId, pdfFileName, pdfBlob, 'application/pdf');
      pdfDriveFileId = pdfUploaded.id;

      log(`[${productId}]${modeLabel} PDF保存完了（${pdfSizeMB}MB, ${pdfImages.length}枚）`, '', 'product');
    } catch (pdfError) {
      console.warn(`[商品情報] PDF生成/アップロードエラー: ${pdfError.message}`);
      log(`[${productId}]${modeLabel} PDF生成に失敗（収集は続行）: ${pdfError.message}`, 'warning', 'product');
    }
  }

  // offscreenドキュメントを閉じる（リソース解放）
  await closeOffscreenDocument();

  // --- 動画情報の記録（アップロードはせず、URLのみ記録） ---
  const videoMetadata = [];
  if (rawVideos.length > 0) {
    let videoOrder = 0;
    for (const video of rawVideos) {
      videoOrder++;
      videoMetadata.push({
        order: videoOrder,
        description: '商品紹介動画',
        originalUrl: video.url,
        type: video.type,
        saved: false,
        reason: video.type === 'hls' ? 'HLSストリーミングのためダウンロード不可' :
                video.url.startsWith('blob:') ? 'ストリーミング配信のためダウンロード不可' :
                'PDF形式のため動画は保存不可'
      });
    }
  }

  // 6. 軽量JSONを生成（画像バイナリなし、メタデータのみ）
  log(`[${productId}]${modeLabel} JSONを保存中...`, '', 'product');

  // content scriptから取得した生データからimages/aplusImages/videosを除去してテキスト情報のみ残す
  const { images: _img, aplusImages: _aplus, videos: _vid, videoThumbnails: _vt, ...textData } = productData;

  const jsonData = {
    ...textData,
    viewType: mode,
    productFolderId,
    pdfFileName: pdfFileName || undefined,
    pdfDriveFileId: pdfDriveFileId || undefined,
    images: imageMetadata,
    aplusImages: aplusImageMetadata.length > 0 ? aplusImageMetadata : undefined,
    videos: videoMetadata.length > 0 ? videoMetadata : undefined,
    sectionDescriptions: isRakuten
      ? {
          main: 'サムネイル - 商品ページで最初に大きく表示される代表画像',
          gallery: 'フリック画像 - サムネイルの下に並ぶ画像群',
          product: '商品画像 - ショップが掲載した商品関連画像',
          description: 'LP - ページ下部のランディングページ部分に掲載された画像'
        }
      : {
          main: 'メイン画像 - 商品ページで最初に大きく表示される代表画像',
          gallery: 'サブ画像 - メイン画像の下に並ぶ画像群',
          product: '商品画像',
          description: '商品説明画像',
          aplus: 'A+コンテンツ - 拡張商品説明エリアの画像'
        }
  };

  // undefinedキーを除去
  const cleanJsonData = JSON.parse(JSON.stringify(jsonData));

  const jsonFileName = `${baseName}.json`;
  const uploadResult = await uploadJsonToDrive(token, productFolderId, jsonFileName, cleanJsonData);

  const totalImageCount = imageMetadata.length + aplusImageMetadata.length;
  log(`[${productId}]${modeLabel} 保存完了（画像PDF: ${totalImageCount}枚${pdfFileName ? '' : '（PDF生成失敗）'}）`, 'success', 'product');

  return {
    success: true,
    fileName: jsonFileName,
    fileId: uploadResult.id,
    productFolderId,
    parentFolderId: productFolderId,
    baseName,
    productId,
    source: isRakuten ? 'rakuten' : 'amazon',
    title: productData.title,
    imageCount: imageMetadata.length,
    aplusImageCount: aplusImageMetadata.length,
    videoCount: videoMetadata.length,
    pdfFileName,
    mode
  };
}

/**
 * セクション名から説明文を生成
 */
function getSectionDescription(section, order, isRakuten = true) {
  const descriptions = isRakuten
    ? {
        main: 'サムネイル（代表画像）',
        gallery: `フリック画像${order}枚目`,
        product: `商品画像${order}枚目`,
        description: `LP画像${order}枚目`
      }
    : {
        main: 'メイン画像（代表画像）',
        gallery: `サブ画像${order}枚目`,
        product: `商品画像${order}枚目`,
        description: `商品説明画像${order}枚目`,
        aplus: `A+コンテンツ画像${order}枚目`
      };
  return descriptions[section] || `画像${order}枚目`;
}

/**
 * 商品をbatchProductQueueから削除（キューに入っている元の値で完全一致検索）
 */
async function removeFromBatchProductQueue(originalItem) {
  const result = await chrome.storage.local.get(['batchProductQueue']);
  const queue = result.batchProductQueue || [];
  const idx = queue.indexOf(originalItem);
  if (idx !== -1) {
    queue.splice(idx, 1);
    await chrome.storage.local.set({ batchProductQueue: queue });
    forwardToAll({ action: 'batchProductQueueUpdated' });
  }
}

/**
 * 商品リストからバッチで商品情報を収集（Amazon ASIN / 楽天URLの両方に対応）
 * @param {string[]} items - ASINまたは商品URLのリスト
 */
async function startBatchProductCollection(items) {
  if (!items || items.length === 0) {
    throw new Error('商品が入力されていません');
  }

  // 設定を事前チェック
  const settings = await chrome.storage.sync.get(['productInfoFolderUrl']);
  if (!settings.productInfoFolderUrl) {
    throw new Error('Google Driveの保存先フォルダが設定されていません。');
  }

  // 認証トークン取得（未認証なら対話型ダイアログを表示）
  const token = await getAuthTokenWithFallback();

  batchProductCancelled = false;
  batchProductProgress = {
    total: items.length,
    current: 0,
    completed: [],
    failed: [],
    isRunning: true
  };

  log(`${items.length}件の商品情報収集を開始します`, '', 'product');
  forwardToAll({ action: 'batchProductProgressUpdate', progress: batchProductProgress });

  // バッチ処理を非同期で実行（即座にレスポンスを返す）
  (async () => {
    // バッチ収集用のタブを作成
    let batchTab;
    try {
      batchTab = await chrome.tabs.create({
        url: 'about:blank',
        active: false
      });
    } catch (e) {
      log('タブ作成に失敗しました', 'error', 'product');
      batchProductProgress.isRunning = false;
      return;
    }

    for (let i = 0; i < items.length; i++) {
      if (batchProductCancelled) {
        log('キャンセルされました', '', 'product');
        break;
      }

      const item = items[i].trim();
      const isRakutenUrl = item.includes('item.rakuten.co.jp');
      let productUrl, displayId;

      if (isRakutenUrl) {
        // 楽天URL: そのまま使用
        productUrl = item.startsWith('http') ? item : `https://${item}`;
        // URLからitemSlugを抽出（/shop/item/ → item）
        const slugMatch = productUrl.match(/item\.rakuten\.co\.jp\/[^/]+\/([^/?]+)/);
        displayId = slugMatch ? slugMatch[1] : '楽天商品';
      } else {
        // Amazon: ASINとして処理
        const asin = item.toUpperCase();
        if (!/^[A-Z0-9]{10}$/.test(asin)) {
          batchProductProgress.failed.push({ id: item, error: '無効なASINまたはURL形式' });
          batchProductProgress.current = i + 1;
          forwardToAll({ action: 'batchProductProgressUpdate', progress: batchProductProgress });
          continue;
        }
        productUrl = `https://www.amazon.co.jp/dp/${asin}`;
        displayId = asin;
      }

      log(`[${displayId}] (${i + 1}/${items.length}) ${isRakutenUrl ? 'PC版を' : ''}収集中...`, '', 'product');

      try {
        // --- PC版の収集 ---
        await chrome.tabs.update(batchTab.id, { url: productUrl });
        await waitForTabComplete(batchTab.id, 15000);
        await sleep(1500);
        const desktopResult = await collectAndSaveProductInfo(batchTab.id, 'desktop');

        // --- スマホ版の収集（楽天のみ、AmazonはPC/SP同一データのためスキップ） ---
        if (!batchProductCancelled && isRakutenUrl) {
          log(`[${displayId}] (${i + 1}/${items.length}) スマホ版を収集中...`, '', 'product');
          try {
            await enableMobileUA(batchTab.id);
            await chrome.tabs.update(batchTab.id, { url: productUrl });
            await waitForTabComplete(batchTab.id, 15000);
            await sleep(1500);
            await collectAndSaveProductInfo(batchTab.id, 'mobile', {
              parentFolderId: desktopResult.parentFolderId,
              baseName: desktopResult.baseName
            });
          } catch (mobileError) {
            console.warn(`[商品情報] ${displayId} スマホ版エラー:`, mobileError);
            log(`[${displayId}] スマホ版の収集に失敗（PC版は保存済み）: ${mobileError.message}`, 'warning', 'product');
          } finally {
            await disableMobileUA();
          }
        }

        batchProductProgress.completed.push({
          id: displayId,
          source: desktopResult.source,
          title: desktopResult.title,
          fileName: desktopResult.fileName
        });

        // 成功した商品をキューから即時削除（元のキュー値で完全一致）
        await removeFromBatchProductQueue(item);
      } catch (error) {
        console.error(`[商品情報] ${displayId} エラー:`, error);
        batchProductProgress.failed.push({ id: displayId, error: error.message });
        log(`[${displayId}] ${error.message}`, 'error', 'product');
        // 失敗した商品もキューから削除（ログに記録済みなので再試行はユーザー判断）
        await removeFromBatchProductQueue(item);
      }

      batchProductProgress.current = i + 1;
      forwardToAll({ action: 'batchProductProgressUpdate', progress: batchProductProgress });

      // 商品間のランダム待機（ボット対策: 3〜8秒）
      if (i < items.length - 1 && !batchProductCancelled) {
        const waitMs = 3000 + Math.random() * 5000;
        await sleep(waitMs);
      }
    }

    // バッチタブを閉じる
    try {
      await chrome.tabs.remove(batchTab.id);
    } catch (e) {
      // 既に閉じられている場合は無視
    }

    batchProductProgress.isRunning = false;
    forwardToAll({ action: 'batchProductProgressUpdate', progress: batchProductProgress });

    const successCount = batchProductProgress.completed.length;
    const failCount = batchProductProgress.failed.length;
    log(`完了: 成功 ${successCount}件、失敗 ${failCount}件`, successCount > 0 ? 'success' : 'error', 'product');
  })();

  return { success: true, message: `${items.length}件のバッチ収集を開始しました` };
}

/**
 * タブのページ読み込み完了を待つ
 */
function waitForTabComplete(tabId, timeoutMs = 15000) {
  return new Promise((resolve, reject) => {
    const startTime = Date.now();

    function listener(updatedTabId, changeInfo) {
      if (updatedTabId === tabId && changeInfo.status === 'complete') {
        chrome.tabs.onUpdated.removeListener(listener);
        resolve();
      }
    }

    chrome.tabs.onUpdated.addListener(listener);

    // タイムアウト
    const checkInterval = setInterval(async () => {
      if (Date.now() - startTime > timeoutMs) {
        clearInterval(checkInterval);
        chrome.tabs.onUpdated.removeListener(listener);
        // タイムアウトしてもページが読み込まれている可能性がある
        try {
          const tab = await chrome.tabs.get(tabId);
          if (tab.status === 'complete') {
            resolve();
          } else {
            reject(new Error('ページの読み込みがタイムアウトしました'));
          }
        } catch (e) {
          reject(new Error('タブが閉じられました'));
        }
      }
    }, 1000);
  });
}

/**
 * 指定ミリ秒待機
 */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// ===== スマホUA切替機能 =====
const MOBILE_USER_AGENT = 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1';
const MOBILE_UA_RULE_ID = 99999;

/**
 * 指定タブのリクエストをスマホUAで送信するよう設定
 */
async function enableMobileUA(tabId) {
  await chrome.declarativeNetRequest.updateSessionRules({
    addRules: [{
      id: MOBILE_UA_RULE_ID,
      priority: 1,
      action: {
        type: 'modifyHeaders',
        requestHeaders: [{
          header: 'User-Agent',
          operation: 'set',
          value: MOBILE_USER_AGENT
        }]
      },
      condition: {
        tabIds: [tabId],
        resourceTypes: ['main_frame', 'sub_frame', 'image', 'xmlhttprequest', 'script', 'stylesheet']
      }
    }],
    removeRuleIds: [MOBILE_UA_RULE_ID]
  });
  console.log(`[商品情報] スマホUA有効化 (tabId: ${tabId})`);
}

/**
 * スマホUA切替を解除して元に戻す
 */
async function disableMobileUA() {
  await chrome.declarativeNetRequest.updateSessionRules({
    removeRuleIds: [MOBILE_UA_RULE_ID]
  });
  console.log('[商品情報] スマホUA無効化');
}

/**
 * PC版とスマホ版の両方を収集する（単品収集用）
 * PC版は指定タブで収集、スマホ版はバックグラウンドタブで収集
 */
async function collectProductInfoWithMobile(tabId) {
  // 現在のタブのURLを取得
  const tab = await chrome.tabs.get(tabId);
  const productUrl = tab.url;
  const isAmazonUrl = productUrl.includes('amazon.co.jp');

  // PC版を収集
  const desktopResult = await collectAndSaveProductInfo(tabId, 'desktop');

  // AmazonはPC/SPで同じデータのためスマホ版収集をスキップ
  if (isAmazonUrl) {
    log('収集完了', 'success', 'product');
    return desktopResult;
  }

  // 楽天のみ: スマホ版をバックグラウンドタブで収集
  let mobileTab;
  try {
    mobileTab = await chrome.tabs.create({ url: 'about:blank', active: false });
    await enableMobileUA(mobileTab.id);
    await chrome.tabs.update(mobileTab.id, { url: productUrl });
    await waitForTabComplete(mobileTab.id, 15000);
    await sleep(1500);
    await collectAndSaveProductInfo(mobileTab.id, 'mobile', {
      parentFolderId: desktopResult.parentFolderId,
      baseName: desktopResult.baseName
    });
    log('PC版・スマホ版の収集完了', 'success', 'product');
  } catch (mobileError) {
    console.warn('[商品情報] スマホ版収集エラー:', mobileError);
    log(`スマホ版の収集に失敗（PC版は保存済み）: ${mobileError.message}`, 'warning', 'product');
  } finally {
    await disableMobileUA();
    if (mobileTab) {
      try { await chrome.tabs.remove(mobileTab.id); } catch (e) { /* 既に閉じられている場合は無視 */ }
    }
  }

  return desktopResult;
}

// ===== フォルダピッカー用 Google Drive API =====

/**
 * 認証トークンを取得（対話型フォールバック付き）
 * フォルダピッカーでは、未認証でも対話型ダイアログで認証できるようにする
 */
async function getAuthTokenWithFallback() {
  // まず非対話型で試す
  let token = await getAuthToken();
  if (token) return token;

  // 非対話型で取得できなければ対話型で試す
  return getAuthTokenInteractive();
}

/**
 * 対話型で認証トークンを取得
 */
function getAuthTokenInteractive() {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive: true }, (token) => {
      if (chrome.runtime.lastError || !token) {
        reject(new Error('Google認証が必要です。ログインしてから再度お試しください。'));
      } else {
        resolve(token);
      }
    });
  });
}

/**
 * キャッシュされたトークンを削除して新しいトークンを対話型で取得
 * スコープが更新された場合に使用
 */
async function refreshAuthToken() {
  const oldToken = await getAuthToken();
  if (oldToken) {
    await new Promise(resolve => {
      chrome.identity.removeCachedAuthToken({ token: oldToken }, resolve);
    });
  }
  return getAuthTokenInteractive();
}

/**
 * Drive APIリクエストを実行（スコープ不足時は自動リトライ）
 */
async function driveApiFetch(url, options = {}) {
  const token = await getAuthTokenWithFallback();
  const response = await fetch(url, {
    ...options,
    headers: { ...options.headers, Authorization: `Bearer ${token}` }
  });

  // スコープ不足エラー: トークンを更新してリトライ
  if (response.status === 403) {
    const errorData = await response.json().catch(() => ({}));
    if (errorData.error?.message?.includes('insufficient authentication scopes')) {
      const newToken = await refreshAuthToken();
      return fetch(url, {
        ...options,
        headers: { ...options.headers, Authorization: `Bearer ${newToken}` }
      });
    }
    throw new Error(errorData.error?.message || `Drive API エラー (403)`);
  }

  return response;
}

/**
 * 指定フォルダ内のサブフォルダ一覧を取得
 * @param {string} parentId - 親フォルダID（'root'でマイドライブ直下）
 * @param {string} driveId - 共有ドライブID（省略時はマイドライブ）
 * @returns {Promise<Object>} { success: true, folders: [...] }
 */
async function getDriveFolders(parentId = 'root', driveId = null) {
  const query = `'${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`;
  const params = new URLSearchParams({
    q: query,
    fields: 'files(id,name)',
    orderBy: 'name',
    pageSize: '100'
  });

  if (driveId) {
    params.set('includeItemsFromAllDrives', 'true');
    params.set('supportsAllDrives', 'true');
    params.set('corpora', 'drive');
    params.set('driveId', driveId);
  } else {
    params.set('supportsAllDrives', 'true');
  }

  console.log('[getDriveFolders]', { parentId, driveId });
  const response = await driveApiFetch(`https://www.googleapis.com/drive/v3/files?${params}`);

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    console.error('[getDriveFolders] エラー:', errorData);
    throw new Error(errorData.error?.message || `Drive API エラー (${response.status})`);
  }

  const data = await response.json();
  console.log('[getDriveFolders] 結果:', data.files?.length, '件');
  return {
    success: true,
    folders: (data.files || []).map(f => ({ id: f.id, name: f.name }))
  };
}

/**
 * 指定フォルダ内のスプレッドシート一覧を取得
 * @param {string} parentId - 親フォルダID（'root'でマイドライブ直下）
 * @param {string} driveId - 共有ドライブID（省略時はマイドライブ）
 * @returns {Promise<Object>} { success: true, files: [...] }
 */
async function getDriveSpreadsheets(parentId = 'root', driveId = null) {
  const query = `'${parentId}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`;
  const params = new URLSearchParams({
    q: query,
    fields: 'files(id,name)',
    orderBy: 'name',
    pageSize: '100'
  });

  if (driveId) {
    params.set('includeItemsFromAllDrives', 'true');
    params.set('supportsAllDrives', 'true');
    params.set('corpora', 'drive');
    params.set('driveId', driveId);
  } else {
    params.set('supportsAllDrives', 'true');
  }

  console.log('[getDriveSpreadsheets]', { parentId, driveId });
  const response = await driveApiFetch(`https://www.googleapis.com/drive/v3/files?${params}`);

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    console.error('[getDriveSpreadsheets] エラー:', errorData);
    throw new Error(errorData.error?.message || `Drive API エラー (${response.status})`);
  }

  const data = await response.json();
  console.log('[getDriveSpreadsheets] 結果:', data.files?.length, '件');
  return {
    success: true,
    files: (data.files || []).map(f => ({ id: f.id, name: f.name }))
  };
}

/**
 * フォルダ名で検索
 * @param {string} query - 検索クエリ
 * @param {string} driveId - 共有ドライブID（省略時はマイドライブ）
 * @returns {Promise<Object>} { success: true, folders: [...] }
 */
async function searchDriveFolders(query, driveId = null) {
  if (!query || query.trim().length === 0) {
    return { success: true, folders: [] };
  }

  const escaped = query.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  const q = `mimeType='application/vnd.google-apps.folder' and name contains '${escaped}' and trashed=false`;
  const params = new URLSearchParams({
    q: q,
    fields: 'files(id,name,parents)',
    orderBy: 'name',
    pageSize: '50'
  });

  // 共有ドライブ内検索とマイドライブ検索の両方をサポート
  if (driveId) {
    params.set('includeItemsFromAllDrives', 'true');
    params.set('supportsAllDrives', 'true');
    params.set('corpora', 'drive');
    params.set('driveId', driveId);
  } else {
    // マイドライブでもsupportsAllDrivesを付与（共有されたフォルダも検索可能に）
    params.set('supportsAllDrives', 'true');
    params.set('includeItemsFromAllDrives', 'true');
  }

  console.log('[searchDriveFolders]', { query, driveId, q: params.get('q') });
  const response = await driveApiFetch(`https://www.googleapis.com/drive/v3/files?${params}`);

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.error?.message || `Drive API エラー (${response.status})`);
  }

  const data = await response.json();
  return {
    success: true,
    folders: (data.files || []).map(f => ({
      id: f.id,
      name: f.name,
      parentId: f.parents ? f.parents[0] : null
    }))
  };
}

/**
 * 共有ドライブ一覧を取得
 * @returns {Promise<Object>} { success: true, drives: [...] }
 */
async function getSharedDrives() {
  const params = new URLSearchParams({
    pageSize: '100',
    fields: 'drives(id,name)'
  });

  const response = await driveApiFetch(`https://www.googleapis.com/drive/v3/drives?${params}`);

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.error?.message || `共有ドライブの取得エラー (${response.status})`);
  }

  const data = await response.json();
  return {
    success: true,
    drives: (data.drives || []).map(d => ({ id: d.id, name: d.name }))
  };
}

/**
 * 新規フォルダを作成
 * @param {string} name - フォルダ名
 * @param {string} parentId - 親フォルダID
 * @returns {Promise<Object>} { success: true, folder: { id, name } }
 */
async function createDriveFolder(name, parentId = 'root') {
  if (!name || name.trim().length === 0) {
    throw new Error('フォルダ名を入力してください');
  }

  // 共有ドライブでもマイドライブでも動作するようsupportsAllDrivesを常に付与
  const response = await driveApiFetch('https://www.googleapis.com/drive/v3/files?supportsAllDrives=true', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      name: name.trim(),
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentId]
    })
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    console.error('[createDriveFolder] エラー:', { name, parentId, status: response.status, error: errorData });
    throw new Error(errorData.error?.message || `フォルダ作成エラー (${response.status})`);
  }

  const folder = await response.json();
  console.log('[createDriveFolder] 成功:', { id: folder.id, name: folder.name, parentId });
  return {
    success: true,
    folder: { id: folder.id, name: folder.name }
  };
}

/**
 * フォルダのパス（パンくずリスト）を取得
 * @param {string} folderId - フォルダID
 * @returns {Promise<Object>} { success: true, path: [{ id, name }, ...] }
 */
async function getDriveFolderPath(folderId) {
  if (!folderId || folderId === 'root') {
    return { success: true, path: [{ id: 'root', name: 'マイドライブ' }] };
  }

  const path = [];
  let currentId = folderId;

  for (let i = 0; i < 10; i++) {
    const params = new URLSearchParams({ fields: 'id,name,parents', supportsAllDrives: 'true' });
    const response = await driveApiFetch(
      `https://www.googleapis.com/drive/v3/files/${currentId}?${params}`
    );

    if (!response.ok) break;

    const file = await response.json();
    path.unshift({ id: file.id, name: file.name });

    if (!file.parents || file.parents.length === 0) break;
    currentId = file.parents[0];
  }

  if (path.length === 0 || path[0].id !== 'root') {
    path.unshift({ id: 'root', name: 'マイドライブ' });
  }

  return { success: true, path };
}

// ===== 商品情報JSONダウンロード（base64画像埋め込み） =====

// バッチJSONダウンロード用の状態管理
let batchJsonProgress = { total: 0, current: 0, currentAsin: '' };
let batchJsonCancelled = false;

/**
 * 価格文字列から数値を抽出
 */
function parsePriceToNumber(priceStr) {
  if (!priceStr) return null;
  const match = priceStr.replace(/,/g, '').match(/(\d+)/);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * 評価文字列から数値を抽出
 */
function parseRatingToNumber(ratingStr) {
  if (!ratingStr) return null;
  const match = ratingStr.match(/([\d.]+)/);
  return match ? parseFloat(match[1]) : null;
}

/**
 * レビュー件数文字列から数値を抽出
 */
function parseReviewCountToNumber(countStr) {
  if (!countStr) return null;
  const match = countStr.replace(/,/g, '').match(/(\d+)/);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * 商品情報をJSONファイル（base64画像埋め込み）としてダウンロード
 * @param {number} tabId - 商品ページのタブID
 * @returns {Promise<{success: boolean, asin: string, fileName: string}>}
 */
async function downloadProductInfoAsJson(tabId) {
  // 1. content scriptから商品情報を取得
  log('ページから情報を読み取り中...', '', 'product');
  forwardToAll({ action: 'productJsonProgress', progress: { phase: 'collecting' } });

  let productData;
  const maxRetries = 2;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = await chrome.tabs.sendMessage(tabId, { action: 'collectProductInfo' });
      if (!response || !response.success) {
        throw new Error(response?.error || '商品情報の取得に失敗しました');
      }
      productData = response.data;
      break;
    } catch (error) {
      if (attempt < maxRetries && error.message.includes('message channel closed')) {
        console.warn('[JSON DL] メッセージチャンネル切断、リトライ...');
        await sleep(2000);
        continue;
      }
      throw error;
    }
  }

  const asin = productData.asin;
  if (!asin) throw new Error('ASINを取得できませんでした');

  // 2. 画像をbase64に変換
  const images = productData.images || [];
  const aplusImageUrls = productData.aplusImages || [];
  const totalImages = images.length + aplusImageUrls.length;
  let processedImages = 0;

  let mainImageBase64 = '';
  const subImagesBase64 = [];
  const aplusImagesBase64 = [];

  // 商品画像（メイン + サブ）
  for (const img of images) {
    const url = typeof img === 'string' ? img : img.url;
    const type = typeof img === 'string' ? 'gallery' : (img.type || 'gallery');
    processedImages++;

    forwardToAll({
      action: 'productJsonProgress',
      progress: { phase: 'images', current: processedImages, total: totalImages, productId: asin }
    });

    const result = await fetchImageAsBase64(url);
    if (result) {
      const dataUri = `data:${result.mimeType};base64,${result.base64}`;
      if (type === 'main' && !mainImageBase64) {
        mainImageBase64 = dataUri;
      } else {
        subImagesBase64.push(dataUri);
      }
    }
  }

  // A+コンテンツ画像
  for (const url of aplusImageUrls) {
    processedImages++;
    forwardToAll({
      action: 'productJsonProgress',
      progress: { phase: 'images', current: processedImages, total: totalImages, productId: asin }
    });

    const result = await fetchImageAsBase64(url);
    if (result) {
      aplusImagesBase64.push(`data:${result.mimeType};base64,${result.base64}`);
    }
  }

  // 3. 仕様通りのJSON構造を構築
  const jsonData = {
    asin: asin,
    title: productData.title || '',
    brand: productData.brand || '',
    price: parsePriceToNumber(productData.price),
    original_price: parsePriceToNumber(productData.listPrice),
    rating: parseRatingToNumber(productData.rating),
    review_count: parseReviewCountToNumber(productData.reviewCount),
    bullet_points: productData.bullets || [],
    description: productData.description || '',
    aplus_text: productData.aplusText || '',
    categories: productData.categories || [],
    ranking: productData.ranking || '',
    variations: productData.structuredVariations || [],
    images: {
      main: mainImageBase64,
      sub: subImagesBase64,
      aplus: aplusImagesBase64
    },
    url: productData.url || `https://www.amazon.co.jp/dp/${asin}`,
    collected_at: new Date().toISOString()
  };

  // 4. JSONファイルをダウンロード
  const now = new Date();
  const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
  const fileName = `${asin}_${dateStr}.json`;

  forwardToAll({
    action: 'productJsonProgress',
    progress: { phase: 'download', productId: asin }
  });

  const jsonString = JSON.stringify(jsonData, null, 2);
  const blob = new Blob([jsonString], { type: 'application/json;charset=utf-8' });
  const blobUrl = URL.createObjectURL(blob);

  await chrome.downloads.download({
    url: blobUrl,
    filename: fileName,
    saveAs: false
  });

  // Blob URLの後片付け
  setTimeout(() => URL.revokeObjectURL(blobUrl), 60000);

  forwardToAll({
    action: 'productJsonProgress',
    progress: { phase: 'done', productId: asin, fileName }
  });

  log(`[${asin}] JSONダウンロード完了: ${fileName} (画像: ${totalImages}枚)`, '', 'product');

  return { success: true, asin, fileName };
}

/**
 * 複数商品のJSONを一括ダウンロード
 * @param {string[]} items - ASINまたはURLのリスト
 */
async function startBatchJsonDownload(items) {
  if (!items || items.length === 0) {
    throw new Error('商品が入力されていません');
  }

  batchJsonCancelled = false;
  batchJsonProgress = { total: items.length, current: 0, currentAsin: '' };

  forwardToAll({
    action: 'batchJsonProgress',
    progress: batchJsonProgress
  });

  // 収集用ウィンドウを作成
  const win = await chrome.windows.create({
    url: 'about:blank',
    width: 1280,
    height: 800,
    focused: true
  });
  const windowId = win.id;
  const tabId = win.tabs[0].id;

  try {
    for (let i = 0; i < items.length; i++) {
      if (batchJsonCancelled) {
        log('一括JSON収集がキャンセルされました', '', 'product');
        break;
      }

      const item = items[i];
      // ASINからURLを生成（URLでない場合）
      const url = item.includes('http') ? item : `https://www.amazon.co.jp/dp/${item}`;
      const asinMatch = url.match(/\/dp\/([A-Z0-9]{10})/i) || url.match(/\/gp\/product\/([A-Z0-9]{10})/i);
      const asin = asinMatch ? asinMatch[1].toUpperCase() : item;

      batchJsonProgress.current = i + 1;
      batchJsonProgress.currentAsin = asin;
      forwardToAll({
        action: 'batchJsonProgress',
        progress: batchJsonProgress
      });

      log(`[${i + 1}/${items.length}] ${asin} を処理中...`, '', 'product');

      try {
        // タブでページを開く
        await chrome.tabs.update(tabId, { url });

        // ページ読み込みを待機
        await new Promise((resolve) => {
          const listener = (updatedTabId, info) => {
            if (updatedTabId === tabId && info.status === 'complete') {
              chrome.tabs.onUpdated.removeListener(listener);
              resolve();
            }
          };
          chrome.tabs.onUpdated.addListener(listener);
          // タイムアウト
          setTimeout(() => {
            chrome.tabs.onUpdated.removeListener(listener);
            resolve();
          }, 30000);
        });

        // ボット対策: ページ遷移前のランダムウェイト
        const waitTime = Math.max(1500, Math.min(-1500 * Math.log(1 - Math.random()), 8000));
        await sleep(waitTime);

        // JSONダウンロード
        await downloadProductInfoAsJson(tabId);

      } catch (error) {
        log(`[${asin}] エラー: ${error.message}`, 'error', 'product');
      }

      // 次の商品への待機（ボット対策）
      if (i < items.length - 1) {
        const interval = Math.max(2000, Math.min(-2000 * Math.log(1 - Math.random()), 10000));
        await sleep(interval);
      }
    }
  } finally {
    // ウィンドウを閉じる
    try {
      await chrome.windows.remove(windowId);
    } catch (e) {
      // ウィンドウが既に閉じられている場合
    }
  }

  batchJsonProgress.currentAsin = '';
  forwardToAll({
    action: 'batchJsonProgress',
    progress: { ...batchJsonProgress, phase: 'done' }
  });

  log(`一括JSON収集完了: ${batchJsonProgress.current}/${batchJsonProgress.total}件`, '', 'product');
  return { success: true, total: items.length, completed: batchJsonProgress.current };
}
