/**
 * バックグラウンドサービスワーカー
 * レビューデータの保存、CSVダウンロード、キュー管理を処理
 * 楽天市場・Amazon両対応
 */

// アクティブな収集タブを追跡
let activeCollectionTabs = new Set();
// 収集用ウィンドウのID
let collectionWindowId = null;
// タブごとのスプレッドシートURL（定期収集用）
const tabSpreadsheetUrls = new Map();

// Amazonページ遷移時の収集再開リスナー（グローバル）
// 注意: Service Workerが再起動するとactiveCollectionTabsがリセットされるため、
// collectionState.isRunningをメインの条件として使用
chrome.tabs.onUpdated.addListener(async (tabId, changeInfo, tab) => {
  // ページ読み込み完了時のみ処理
  if (changeInfo.status !== 'complete') return;

  // Amazonレビューページかどうか確認
  const url = tab.url || '';
  if (!url.includes('amazon.co.jp/product-reviews/')) return;

  // 収集状態を確認（Service Worker再起動後もストレージは永続化される）
  const result = await chrome.storage.local.get(['collectionState']);
  const state = result.collectionState;

  console.log('[background] Amazonレビューページ検出:', {
    tabId,
    url: url.substring(0, 100),
    isActiveTab: activeCollectionTabs.has(tabId),
    activeTabsCount: activeCollectionTabs.size,
    state: state ? {
      isRunning: state.isRunning,
      source: state.source,
      lastProcessedPage: state.lastProcessedPage
    } : null
  });

  // 収集中でAmazonの場合のみ処理
  if (state && state.isRunning && state.source === 'amazon') {
    console.log('[background] Amazonレビューページ遷移検出、収集再開メッセージを送信');

    // Service Workerが再起動している可能性があるので、タブをアクティブに追加
    if (!activeCollectionTabs.has(tabId)) {
      console.log('[background] タブをactiveCollectionTabsに追加:', tabId);
      activeCollectionTabs.add(tabId);
    }

    // 少し待ってからメッセージを送信（DOM読み込み完了を待つ）
    setTimeout(() => {
      console.log('[background] resumeCollectionメッセージ送信...');
      chrome.tabs.sendMessage(tabId, {
        action: 'resumeCollection',
        incrementalOnly: state.incrementalOnly || false,
        lastCollectedDate: state.lastCollectedDate || null,
        queueName: state.queueName || null
      }).then((response) => {
        console.log('[background] resumeCollection応答:', response);
      }).catch((err) => {
        console.log('[background] 収集再開メッセージ送信エラー:', err);
      });
    }, 3000);
  }
});

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
      handleCollectionComplete(sender.tab?.id);
      sendResponse({ success: true });
      break;

    case 'collectionStopped':
      handleCollectionStopped(sender.tab?.id);
      sendResponse({ success: true });
      break;

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

    case 'openOptions':
      chrome.runtime.openOptionsPage();
      sendResponse({ success: true });
      break;
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

  if (isScheduled) {
    // 定期収集: スプレッドシートのみに保存（CSVには保存しない）
    if (spreadsheetUrl) {
      try {
        await sendToSheets(spreadsheetUrl, reviews, syncSettings.separateSheets !== false, true, detectedSource);
        log(prefix + 'スプレッドシートに保存しました');
      } catch (error) {
        log(`スプレッドシートへの保存に失敗: ${error.message}`, 'error');
      }
    } else {
      log(prefix + '定期収集用のスプレッドシートが設定されていません', 'error');
    }
  } else {
    // 通常収集: ローカルストレージ（CSV用）とスプレッドシートに保存
    const newReviews = await saveToLocalStorage(reviews, detectedSource);

    if (newReviews.length === 0) {
      return;
    }

    if (spreadsheetUrl) {
      try {
        // ローカルストレージから全レビューを取得
        const stateResult = await chrome.storage.local.get(['collectionState']);
        const allReviews = stateResult.collectionState?.reviews || [];

        if (allReviews.length > 0) {
          await sendToSheets(spreadsheetUrl, allReviews, syncSettings.separateSheets !== false, false, detectedSource);
          log(prefix + 'スプレッドシートに保存しました');
        }
      } catch (error) {
        log(`スプレッドシートへの保存に失敗: ${error.message}`, 'error');
      }
    }
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
 */
async function formatDataRows(token, spreadsheetId, sheetId, startRow, endRow) {
  const requests = [
    {
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: startRow,
          endRowIndex: endRow,
          startColumnIndex: 0,
          endColumnIndex: 22  // A-V列（22項目に拡張）
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
 * URL列（D列、S列）にクリック可能なリンク書式を適用
 * @param {string} token - OAuth token
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {number} sheetId - シートID
 * @param {Array} reviews - レビューデータ配列
 */
async function formatUrlColumns(token, spreadsheetId, sheetId, reviews) {
  if (!reviews || reviews.length === 0) return;

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

    // S列（レビュー掲載URL、列インデックス18）
    if (review.pageUrl) {
      requests.push({
        updateCells: {
          range: {
            sheetId: sheetId,
            startRowIndex: rowIndex,
            endRowIndex: rowIndex + 1,
            startColumnIndex: 18,
            endColumnIndex: 19
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
  // 販路別の色設定
  let backgroundColor, textColor;
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
          endColumnIndex: 22  // A-V列（22項目に拡張）
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
    // W列以降を削除（23列目以降）
    {
      deleteDimension: {
        range: {
          sheetId: sheetId,
          dimension: 'COLUMNS',
          startIndex: 22,  // W列（0-indexed で22）
          endIndex: 26     // Z列まで（デフォルトの列数）
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

  // ヘッダー（22項目: 販路と国を追加）
  const headers = [
    'レビュー日', '商品管理番号', '商品名', '商品URL', '評価', 'タイトル', '本文',
    '投稿者', '年代', '性別', '注文日', 'バリエーション', '用途', '贈り先',
    '購入回数', '参考になった数', 'ショップからの返信', 'ショップ名', 'レビュー掲載URL', '収集日時',
    '販路', '国'
  ];

  // テキストが=で始まる場合はエスケープ（数式として解釈されないように）
  const escapeFormula = (text) => {
    if (!text) return '';
    const str = String(text);
    if (str.startsWith('=') || str.startsWith('+') || str.startsWith('-') || str.startsWith('@')) {
      return "'" + str;
    }
    return str;
  };

  // データを行形式に変換（22項目）
  const dataValues = reviews.map(review => [
    escapeFormula(review.reviewDate || ''),
    escapeFormula(review.productId || ''),
    escapeFormula(review.productName || ''),
    review.productUrl || '',  // URLは後でformatUrlColumnsでリンク化
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
    review.pageUrl || '',  // URLは後でformatUrlColumnsでリンク化
    escapeFormula(review.collectedAt || ''),
    review.source === 'amazon' ? 'Amazon' : '楽天',  // 販路
    escapeFormula(review.country || (review.source === 'amazon' ? '' : '日本'))  // 国
  ]);

  // ヘッダー + データを結合
  const allValues = [headers, ...dataValues];
  const totalRows = allValues.length;

  // 1. シートの全データをクリア（A-V列に拡張）
  await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodedSheetName}'!A:V:clear`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}` }
    }
  );

  // 2. ヘッダーとデータを書き込み（A1:V${totalRows}に拡張）
  // USER_ENTERED: HYPERLINK関数を評価してクリック可能なリンクにする
  const writeResponse = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodedSheetName}'!A1:V${totalRows}?valueInputOption=USER_ENTERED`,
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
    await formatDataRows(token, spreadsheetId, sheetId, 1, totalRows);
  }

  // 6. URL列（D列・S列）にクリック可能なリンク書式を適用
  if (reviews.length > 0) {
    await formatUrlColumns(token, spreadsheetId, sheetId, reviews);
  }
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

  if (separateSheets) {
    // 商品ごとにシートを分ける
    const reviewsByProduct = groupReviewsByProduct(reviews);
    for (const [productId, productReviews] of Object.entries(reviewsByProduct)) {
      // 定期収集の場合は「楽・商品管理番号」または「Ama・ASIN」形式
      // 商品の販路はレビューから判定
      const productSource = productReviews[0]?.source || source;
      const prefix = productSource === 'amazon' ? 'Ama' : '楽';
      const sheetName = isScheduled ? `${prefix}・${productId}` : productId;
      await appendToSheet(token, spreadsheetId, sheetName, productReviews, productSource);
    }
  } else {
    // 全て同じシートに保存
    const sheetName = isScheduled ? '定期・レビュー' : 'レビュー';
    await appendToSheet(token, spreadsheetId, sheetName, reviews, source);
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
 * レビューデータをCSV形式に変換（22項目: 販路と国を追加）
 */
function convertToCSV(reviews) {
  const headers = [
    'レビュー日', '商品管理番号', '商品名', '商品URL', '評価', 'タイトル', '本文',
    '投稿者', '年代', '性別', '注文日', 'バリエーション', '用途', '贈り先',
    '購入回数', '参考になった数', 'ショップからの返信', 'ショップ名', 'レビュー掲載URL', '収集日時',
    '販路', '国'
  ];

  const rows = reviews.map(review => [
    review.reviewDate || '', review.productId || '', review.productName || '',
    review.productUrl || '', review.rating || '', review.title || '', review.body || '',
    review.author || '', review.age || '', review.gender || '', review.orderDate || '',
    review.variation || '', review.usage || '', review.recipient || '',
    review.purchaseCount || '', review.helpfulCount || 0, review.shopReply || '',
    review.shopName || '', review.pageUrl || '', review.collectedAt || '',
    review.source === 'amazon' ? 'Amazon' : '楽天',  // 販路
    review.country || (review.source === 'amazon' ? '' : '日本')  // 国
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
  // 収集中アイテムと収集状態を取得
  const initialResult = await chrome.storage.local.get(['collectingItems', 'isQueueCollecting', 'expectedReviewTotal', 'collectionState']);
  const isQueueCollecting = initialResult.isQueueCollecting || false;
  const expectedTotal = initialResult.expectedReviewTotal || 0;
  const currentState = initialResult.collectionState || {};
  const actualCount = currentState.reviewCount || 0;

  // 収集中アイテムから商品情報を取得（ログ出力用）
  let completedItem = null;
  if (tabId) {
    const collectingItems = initialResult.collectingItems || [];
    completedItem = collectingItems.find(item => item.tabId === tabId);
  }

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

  // アクティブタブから削除
  if (tabId) {
    activeCollectionTabs.delete(tabId);
    tabSpreadsheetUrls.delete(tabId);
    // タブごとの状態をクリーンアップ
    const stateKey = `collectionState_${tabId}`;
    await chrome.storage.local.remove(stateKey);

    // 収集中リストから削除
    const collectingResult = await chrome.storage.local.get(['collectingItems']);
    let collectingItems = collectingResult.collectingItems || [];
    collectingItems = collectingItems.filter(item => item.tabId !== tabId);
    await chrome.storage.local.set({ collectingItems });

    // 商品の収集完了ログを出力
    if (completedItem) {
      const productId = extractProductIdFromUrl(completedItem.url);
      // 定期収集の場合は[キュー名・商品ID]形式
      const logPrefix = completedItem.queueName ? `[${completedItem.queueName}・${productId}]` : `[${productId}]`;
      log(`${logPrefix} 収集が完了しました`, 'success');
      // 商品ごとの通知（設定で有効な場合のみ）
      showNotification('楽天レビュー収集', `${logPrefix} 収集が完了しました`, productId);

      // 最終収集日を保存（差分取得用）
      const lcResult = await chrome.storage.local.get(['productLastCollected']);
      const productLastCollected = lcResult.productLastCollected || {};
      productLastCollected[productId] = new Date().toISOString().split('T')[0]; // YYYY-MM-DD形式
      await chrome.storage.local.set({ productLastCollected });
    }

    // キュー収集の場合のみタブを閉じる（単一収集ではユーザーがページに留まる）
    if (isQueueCollecting) {
      try {
        await chrome.tabs.remove(tabId);
      } catch (e) {
        // タブが既に閉じられている場合は無視
      }
    }
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
  if (queue.length > 0 && isQueueCollecting) {
    log('次の商品の収集を開始します...');
    setTimeout(() => {
      processNextInQueue();
    }, 3000);
  } else if (activeCollectionTabs.size === 0 && isQueueCollecting) {
    // すべて完了（キュー収集中の場合）
    await chrome.storage.local.set({ isQueueCollecting: false, collectingItems: [] });

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

  // 同時収集数は固定値3
  const maxConcurrent = 3;

  // 収集中フラグを立てる
  await chrome.storage.local.set({ isQueueCollecting: true, collectingItems: [] });

  // 収集用ウィンドウを新規作成（大きいサイズ、最小化しない）
  try {
    const window = await chrome.windows.create({
      url: 'about:blank',
      width: 1280,
      height: 800,
      focused: false
    });
    collectionWindowId = window.id;

    // 最小化しない（最小化するとページ読み込みが正常に動作しない場合がある）
    // await chrome.windows.update(collectionWindowId, { state: 'minimized' });

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
  // 同時収集数は固定値3
  const maxConcurrent = 3;

  // 現在のアクティブタブ数をチェック
  if (activeCollectionTabs.size >= maxConcurrent) {
    log(`同時収集数上限（${maxConcurrent}）に達しています`, 'info');
    return;
  }

  const result = await chrome.storage.local.get(['queue', 'collectingItems']);
  const queue = result.queue || [];
  const collectingItems = result.collectingItems || [];

  if (queue.length === 0) {
    // キューが空の場合は何もしない（完了ログはhandleCollectionCompleteで出す）
    return;
  }

  const nextItem = queue.shift();

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
  log(`${queuePrefix ? queuePrefix + ' ' : ''}収集中: ${nextItem.title || nextItem.url}`);
  console.log('[processNextInQueue] 処理URL:', nextItem.url, 'isAmazon:', nextItem.url.includes('amazon.co.jp'));

  // 収集用ウィンドウにタブを作成
  let tab;
  const isAmazonUrl = nextItem.url.includes('amazon.co.jp');

  // 直接URLにアクセス（トップページリダイレクトはchrome.tabs.updateもブロックされるため廃止）
  // active: trueにしないとページが正しく読み込まれない場合がある
  if (collectionWindowId) {
    try {
      tab = await chrome.tabs.create({
        url: nextItem.url,
        windowId: collectionWindowId,
        active: true
      });
    } catch (e) {
      // ウィンドウが閉じられている場合は通常のタブで開く
      console.error('収集用ウィンドウにタブ作成失敗:', e);
      tab = await chrome.tabs.create({ url: nextItem.url, active: true });
    }
  } else {
    // フォールバック: 通常のタブ
    tab = await chrome.tabs.create({ url: nextItem.url, active: true });
  }

  // tabIdを収集中アイテムに設定
  nextItem.tabId = tab.id;
  await chrome.storage.local.set({ collectingItems });

  // アクティブタブとして追跡
  activeCollectionTabs.add(tab.id);

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
  await chrome.storage.local.set({
    collectionState: {
      isRunning: true,
      reviewCount: 0,
      pageCount: 0,
      totalPages: 0,
      reviews: [],
      queueName: nextItem.queueName || null,
      incrementalOnly: nextItem.incrementalOnly || false,
      source: isAmazonUrl ? 'amazon' : 'rakuten' // 販路を設定（自動再開に必要）
    }
  });

  // ページ読み込み完了後に収集開始を指示
  let startCollectionSent = false; // 重複送信防止フラグ

  chrome.tabs.onUpdated.addListener(function listener(tabId, info, tabInfo) {
    if (tabId === tab.id && info.status === 'complete') {
      // 重複送信防止
      if (startCollectionSent) {
        console.log('[キュー処理] startCollection既に送信済みのためスキップ');
        return;
      }
      startCollectionSent = true;
      chrome.tabs.onUpdated.removeListener(listener);

      setTimeout(async () => {
        // 差分取得の設定を取得
        let lastCollectedDate = null;
        if (nextItem.incrementalOnly) {
          const productId = extractProductIdFromUrl(nextItem.url);
          const lcResult = await chrome.storage.local.get(['productLastCollected']);
          const productLastCollected = lcResult.productLastCollected || {};
          lastCollectedDate = productLastCollected[productId] || null;
        }

        console.log('[キュー処理] startCollectionメッセージ送信');
        chrome.tabs.sendMessage(tab.id, {
          action: 'startCollection',
          incrementalOnly: nextItem.incrementalOnly || false,
          lastCollectedDate: lastCollectedDate,
          queueName: nextItem.queueName || null
        }).catch(() => {});
      }, 2000);
    }
  });
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

  // 共通のcollectionStateもリセット（レビュー蓄積を防ぐ）
  await chrome.storage.local.set({
    collectionState: {
      isRunning: true,
      reviewCount: 0,
      pageCount: 0,
      totalPages: 0,
      reviews: []
    }
  });

  // UIを更新
  forwardToAll({ action: 'queueUpdated' });

  const productId = extractProductIdFromUrl(productInfo.url);
  log(`[${productId}] 収集を開始しました`);

  // content.jsに収集開始を指示
  chrome.tabs.sendMessage(tabId, { action: 'startCollection' }).catch((error) => {
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
chrome.tabs.onRemoved.addListener((tabId) => {
  if (activeCollectionTabs.has(tabId)) {
    activeCollectionTabs.delete(tabId);
    tabSpreadsheetUrls.delete(tabId);
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

// 収集用ウィンドウが閉じられた時のクリーンアップ
chrome.windows.onRemoved.addListener((windowId) => {
  if (windowId === collectionWindowId) {
    collectionWindowId = null;
    // ウィンドウが手動で閉じられた場合、収集を停止
    if (activeCollectionTabs.size > 0) {
      log('収集用ウィンドウが閉じられました。収集を停止します', 'error');
      activeCollectionTabs.clear();
      chrome.storage.local.set({
        isQueueCollecting: false,
        collectingItems: []
      });
      forwardToAll({ action: 'queueUpdated' });
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
  if (alarm.name === SCHEDULED_ALARM_NAME) {
    runScheduledCollection().catch(error => {
      console.error('定期収集エラー:', error);
      log('定期収集でエラーが発生しました: ' + error.message, 'error');
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
