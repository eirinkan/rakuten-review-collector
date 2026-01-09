/**
 * バックグラウンドサービスワーカー
 * レビューデータの保存、GASへの送信、CSVダウンロード、キュー管理を処理
 */

// アクティブな収集タブを追跡
let activeCollectionTabs = new Set();
// 収集用ウィンドウのID
let collectionWindowId = null;
// タブごとのGAS URL（キュー固有のURL用）
const tabGasUrls = new Map();
// タブごとのスプレッドシートURL（定期収集用）
const tabSpreadsheetUrls = new Map();

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
 * GASで許可チェック
 * @param {string} email - チェックするメールアドレス
 * @returns {Promise<Object>} 許可チェック結果
 */
async function checkUserPermission(email) {
  try {
    // GAS URLを取得
    const settings = await chrome.storage.sync.get(['gasUrl']);
    if (!settings.gasUrl) {
      // GAS URLが設定されていない場合は許可（後でGAS設定が必要になる）
      return { allowed: true, message: 'GAS未設定のため許可' };
    }

    // GASに許可チェックリクエスト
    const response = await fetch(settings.gasUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'checkAuth', email: email })
    });

    if (!response.ok) {
      throw new Error('許可チェックに失敗しました');
    }

    const result = await response.json();
    return result;
  } catch (error) {
    console.error('許可チェックエラー:', error);
    // エラー時は安全のため許可しない
    return { allowed: false, message: error.message };
  }
}

/**
 * 認証フロー全体を実行（ログイン + 許可チェック + 保存）
 * @returns {Promise<Object>} 認証結果
 */
async function authenticate() {
  try {
    // 1. Googleログイン
    const userInfo = await googleLogin();

    // 2. 許可チェック
    const permissionResult = await checkUserPermission(userInfo.email);

    if (!permissionResult.allowed) {
      // 許可されていない場合、トークンを無効化
      await revokeToken();
      return {
        success: false,
        authenticated: false,
        message: 'このメールアドレスは許可されていません: ' + userInfo.email
      };
    }

    // 3. 認証情報を保存
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

    // ===== 既存の機能 =====
    case 'saveReviews':
      handleSaveReviews(message.reviews, sender.tab?.id)
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
 * @param {number} tabId - 送信元タブID（キュー固有のGAS URL取得用）
 */
async function handleSaveReviews(reviews, tabId = null) {
  if (!reviews || reviews.length === 0) {
    return;
  }

  // 重複を除外してローカルストレージに保存
  const newReviews = await saveToLocalStorage(reviews);

  // 新規レビューがない場合はスプレッドシートへの送信をスキップ
  if (newReviews.length === 0) {
    return;
  }

  const { separateSheets, spreadsheetUrl: globalSpreadsheetUrl, gasUrl: globalGasUrl } = await chrome.storage.sync.get(['separateSheets', 'spreadsheetUrl', 'gasUrl']);

  // タブ固有のスプレッドシートURL/GAS URLを取得（定期収集用）
  // collectingItemsから取得（Service Workerのメモリがクリアされても対応）
  let spreadsheetUrl = null;
  let gasUrl = null;

  if (tabId) {
    const collectingResult = await chrome.storage.local.get(['collectingItems']);
    const collectingItems = collectingResult.collectingItems || [];
    const currentItem = collectingItems.find(item => item.tabId === tabId);
    if (currentItem) {
      spreadsheetUrl = currentItem.spreadsheetUrl || null;
      gasUrl = currentItem.gasUrl || null;
    }
  }

  // 定期収集用がなければグローバル設定を使用
  if (!spreadsheetUrl) {
    spreadsheetUrl = globalSpreadsheetUrl;
  }
  if (!gasUrl) {
    gasUrl = globalGasUrl;
  }

  // ログ用のプレフィックスを決定（定期収集の場合はキュー名、それ以外は商品管理番号）
  const productId = newReviews[0]?.productId || '';
  let prefix = productId ? `[${productId}] ` : '';

  // 定期収集の場合は[キュー名・商品ID]形式
  if (tabId) {
    const collectingResult = await chrome.storage.local.get(['collectingItems']);
    const collectingItems = collectingResult.collectingItems || [];
    const currentItem = collectingItems.find(item => item.tabId === tabId);
    if (currentItem?.queueName) {
      prefix = `[${currentItem.queueName}・${productId}] `;
    }
  }

  // 優先順位: スプレッドシートURL（Sheets API直接） > GAS URL
  if (spreadsheetUrl) {
    try {
      await sendToSheets(spreadsheetUrl, newReviews, separateSheets !== false);
      log(prefix + 'スプレッドシートに保存しました');
    } catch (error) {
      log(`スプレッドシートへの保存に失敗: ${error.message}`, 'error');
    }
  } else if (gasUrl) {
    try {
      await sendToGas(gasUrl, newReviews, separateSheets !== false);
      log(prefix + 'スプレッドシートに保存しました');
    } catch (error) {
      log(`スプレッドシートへの保存に失敗: ${error.message}`, 'error');
    }
  }
}

/**
 * レビューの一意キーを生成
 */
function getReviewKey(review) {
  // 商品ID + レビュー日 + 投稿者 + 本文の最初の50文字で一意性を判断
  const bodySnippet = (review.body || '').substring(0, 50);
  return `${review.productId || ''}_${review.reviewDate || ''}_${review.author || ''}_${bodySnippet}`;
}

/**
 * ローカルストレージに保存（重複削除付き）
 * @returns {Promise<Array>} 新規追加されたレビューの配列
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
              totalPages: state.totalPages
            }
          });
          resolve(newReviews); // 新規追加されたレビューを返す
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
 * シートのデータ行数を取得
 */
async function getSheetDataRowCount(token, spreadsheetId, sheetTitle) {
  const valuesResponse = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodeURIComponent(sheetTitle)}'!A:A`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const valuesData = await valuesResponse.json();
  return valuesData.values?.length || 0;
}

/**
 * シートが存在するか確認し、なければ作成
 * 既存シートにデータがある場合は新しいシートを作成（_2, _3...）
 * 戻り値: 使用するシート名
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
  const existingSheetNames = sheets.map(sheet => sheet.properties.title);
  const targetSheet = sheets.find(sheet => sheet.properties.title === sheetName);

  // 対象シートが存在する場合、データがあるか確認
  if (targetSheet) {
    const actualRows = await getSheetDataRowCount(token, spreadsheetId, sheetName);

    // データがある場合（ヘッダー含め2行以上）、新しいシートを作成
    if (actualRows >= 2) {
      // 連番で空いているシート名を探す（_2, _3, ...）
      let newSheetName = sheetName;
      let counter = 2;
      while (existingSheetNames.includes(newSheetName)) {
        // 既存シートにデータがあるか確認
        const existingRows = await getSheetDataRowCount(token, spreadsheetId, newSheetName);
        if (existingRows <= 1) {
          // 空のシートがあればそれを使用
          return newSheetName;
        }
        newSheetName = `${sheetName}_${counter}`;
        counter++;
      }

      // 新しいシートを作成
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
                properties: { title: newSheetName }
              }
            }]
          })
        }
      );

      if (!createResponse.ok) {
        const error = await createResponse.json();
        throw new Error(error.error?.message || 'シートの作成に失敗しました');
      }

      return newSheetName;
    }

    // データがない場合（空またはヘッダーのみ）、そのシートを使用
    return sheetName;
  }

  // 対象シートが存在しない場合
  // 空のシートを探す（1行以下 = 空またはヘッダーのみ）
  let emptySheet = null;
  for (const sheet of sheets) {
    const sheetTitle = sheet.properties.title;
    const actualRows = await getSheetDataRowCount(token, spreadsheetId, sheetTitle);

    if (actualRows <= 1) {
      emptySheet = sheet;
      break;
    }
  }

  if (emptySheet) {
    // 空シートの名前を変更
    const renameResponse = await fetch(
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
                sheetId: emptySheet.properties.sheetId,
                title: sheetName
              },
              fields: 'title'
            }
          }]
        })
      }
    );

    if (!renameResponse.ok) {
      const error = await renameResponse.json();
      throw new Error(error.error?.message || 'シート名の変更に失敗しました');
    }
  } else {
    // 空シートがなければ新規作成
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
          endColumnIndex: 20  // A-T列
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
 * ヘッダー行に書式を適用（赤背景・白テキスト・太字・行固定）
 */
async function formatHeaderRow(token, spreadsheetId, sheetId) {
  const requests = [
    // ヘッダー行の書式設定（赤背景・白テキスト・太字・中央揃え）
    {
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: 0,
          endRowIndex: 1,
          startColumnIndex: 0,
          endColumnIndex: 20  // A-T列
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: {
              red: 191/255,  // #BF0000
              green: 0,
              blue: 0
            },
            textFormat: {
              foregroundColor: {
                red: 1,
                green: 1,
                blue: 1
              },
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
    // U列以降を削除（21列目以降）
    {
      deleteDimension: {
        range: {
          sheetId: sheetId,
          dimension: 'COLUMNS',
          startIndex: 20,  // U列（0-indexed で20）
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
 * シートにデータを追加
 */
async function appendToSheet(token, spreadsheetId, sheetName, reviews) {
  // ensureSheetExistsは実際に使用するシート名を返す（データがある場合は _2, _3 など）
  const actualSheetName = await ensureSheetExists(token, spreadsheetId, sheetName);

  const encodedSheetName = encodeURIComponent(actualSheetName);
  const sheetId = await getSheetId(token, spreadsheetId, actualSheetName);

  // ヘッダー
  const headers = [
    'レビュー日', '商品管理番号', '商品名', '商品URL', '評価', 'タイトル', '本文',
    '投稿者', '年代', '性別', '注文日', 'バリエーション', '用途', '贈り先',
    '購入回数', '参考になった数', 'ショップからの返信', 'ショップ名', 'レビュー掲載URL', '収集日時'
  ];

  // データを行形式に変換
  const dataValues = reviews.map(review => [
    review.reviewDate || '', review.productId || '', review.productName || '',
    review.productUrl || '', review.rating || '', review.title || '', review.body || '',
    review.author || '', review.age || '', review.gender || '', review.orderDate || '',
    review.variation || '', review.usage || '', review.recipient || '',
    review.purchaseCount || '', review.helpfulCount || 0, review.shopReply || '',
    review.shopName || '', review.pageUrl || '', review.collectedAt || ''
  ]);

  // ヘッダー + データを結合
  const allValues = [headers, ...dataValues];
  const totalRows = allValues.length;

  // 1. シートの全データをクリア
  await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodedSheetName}'!A:T:clear`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}` }
    }
  );

  // 2. ヘッダーとデータを書き込み
  const writeResponse = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/'${encodedSheetName}'!A1:T${totalRows}?valueInputOption=RAW`,
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
  // まずシートのプロパティを取得して現在の行数を確認
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

  // 4. ヘッダー書式を適用（赤背景・白テキスト・太字・行固定・U列以降削除）
  await formatHeaderRow(token, spreadsheetId, sheetId);

  // 5. データ行に書式を適用（白背景・黒テキスト・垂直中央揃え）
  if (dataValues.length > 0) {
    await formatDataRows(token, spreadsheetId, sheetId, 1, totalRows);
  }
}

/**
 * Sheets APIを使ってスプレッドシートに直接書き込み
 */
async function sendToSheets(spreadsheetUrl, reviews, separateSheets = true) {
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
      await appendToSheet(token, spreadsheetId, productId, productReviews);
    }
  } else {
    // 全て同じシートに保存
    await appendToSheet(token, spreadsheetId, 'レビュー', reviews);
  }
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
    '購入回数', '参考になった数', 'ショップからの返信', 'ショップ名', 'レビュー掲載URL', '収集日時'
  ];

  const rows = reviews.map(review => [
    review.reviewDate || '', review.productId || '', review.productName || '',
    review.productUrl || '', review.rating || '', review.title || '', review.body || '',
    review.author || '', review.age || '', review.gender || '', review.orderDate || '',
    review.variation || '', review.usage || '', review.recipient || '',
    review.purchaseCount || '', review.helpfulCount || 0, review.shopReply || '',
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
    // タブ固有のGAS URLをクリーンアップ
    tabGasUrls.delete(tabId);
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
  // タブ固有のGAS URLをクリーンアップ
  tabGasUrls.delete(tabId);
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
 * URLから商品管理番号を抽出
 */
function extractProductIdFromUrl(url) {
  if (!url) return 'unknown';
  const match = url.match(/item\.rakuten\.co\.jp\/[^\/]+\/([^\/\?]+)/);
  return match ? match[1] : 'unknown';
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

  // 収集用ウィンドウを新規作成して最小化
  try {
    const window = await chrome.windows.create({
      url: 'about:blank',
      width: 400,
      height: 300,
      focused: false
    });
    collectionWindowId = window.id;

    // 作成後に最小化を試行
    await chrome.windows.update(collectionWindowId, { state: 'minimized' });

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
  tabGasUrls.clear();
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

  // 収集中リストに追加
  nextItem.tabId = null; // 後で設定
  collectingItems.push(nextItem);

  await chrome.storage.local.set({ queue, collectingItems });

  forwardToAll({ action: 'queueUpdated' });

  // 定期収集の場合は[キュー名・商品ID]形式
  const productId = extractProductIdFromUrl(nextItem.url);
  const queuePrefix = nextItem.queueName ? `[${nextItem.queueName}・${productId}]` : '';
  log(`${queuePrefix ? queuePrefix + ' ' : ''}収集中: ${nextItem.title || nextItem.url}`);

  // 収集用ウィンドウにタブを作成（最小化ウィンドウ内）
  let tab;
  if (collectionWindowId) {
    try {
      tab = await chrome.tabs.create({
        url: nextItem.url,
        windowId: collectionWindowId,
        active: false
      });
    } catch (e) {
      // ウィンドウが閉じられている場合は通常のタブで開く
      console.error('収集用ウィンドウにタブ作成失敗:', e);
      tab = await chrome.tabs.create({ url: nextItem.url, active: false });
    }
  } else {
    // フォールバック: 通常のバックグラウンドタブ
    tab = await chrome.tabs.create({ url: nextItem.url, active: false });
  }

  // tabIdを収集中アイテムに設定
  nextItem.tabId = tab.id;
  await chrome.storage.local.set({ collectingItems });

  // アクティブタブとして追跡
  activeCollectionTabs.add(tab.id);

  // キュー固有のGAS URLがあれば保存
  if (nextItem.gasUrl) {
    tabGasUrls.set(tab.id, nextItem.gasUrl);
  }

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
      incrementalOnly: nextItem.incrementalOnly || false
    }
  });

  // ページ読み込み完了後に収集開始を指示
  chrome.tabs.onUpdated.addListener(function listener(tabId, info) {
    if (tabId === tab.id && info.status === 'complete') {
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
 * ストレージへの保存はoptions.jsのaddLog()が行うため、ここでは転送のみ
 */
function log(text, type = '') {
  console.log(`[楽天レビュー収集] ${text}`);

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
    // タブ固有のGAS URLをクリーンアップ
    tabGasUrls.delete(tabId);
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
          gasUrl: targetQueue.gasUrl || null,
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
