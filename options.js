/**
 * 設定画面のスクリプト
 * キュー管理、ランキング追加、設定、ログ表示
 */

/**
 * テーマ管理クラス
 * ダークモード/ライトモードの切り替えを管理
 */
class ThemeManager {
  constructor() {
    this.storageKey = 'rakuten-review-theme';
    this.init();
  }

  init() {
    // 保存された設定を読み込み、なければシステム設定に従う
    const savedTheme = localStorage.getItem(this.storageKey);

    if (savedTheme) {
      this.setTheme(savedTheme);
    } else {
      // デフォルトはライトモード
      this.setTheme('light');
    }

    // トグルボタンのイベント
    this.bindToggle();
  }

  setTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    const toggle = document.getElementById('themeToggle');
    if (toggle) {
      toggle.checked = theme === 'dark';
    }
  }

  bindToggle() {
    const toggle = document.getElementById('themeToggle');
    if (toggle) {
      toggle.addEventListener('change', (e) => {
        const newTheme = e.target.checked ? 'dark' : 'light';
        this.setTheme(newTheme);
        localStorage.setItem(this.storageKey, newTheme);
      });
    }
  }
}

document.addEventListener('DOMContentLoaded', () => {
  // テーマ管理を初期化
  new ThemeManager();
  // DOM要素
  const queueRemaining = document.getElementById('queueRemaining');
  const spreadsheetLink = document.getElementById('spreadsheetLink');
  const downloadBtn = document.getElementById('downloadBtn');
  const clearDataBtn = document.getElementById('clearDataBtn');
  const dataButtons = document.getElementById('dataButtons');

  const gasUrlInput = document.getElementById('gasUrl');
  const separateSheetsCheckbox = document.getElementById('separateSheets');
  const separateCsvFilesCheckbox = document.getElementById('separateCsvFiles');
  const enableNotificationCheckbox = document.getElementById('enableNotification');
  const notifyPerProductCheckbox = document.getElementById('notifyPerProduct');

  const queueList = document.getElementById('queueList');
  const startQueueBtn = document.getElementById('startQueueBtn');
  const stopQueueBtn = document.getElementById('stopQueueBtn');
  const clearQueueBtn = document.getElementById('clearQueueBtn');
  const copyLogBtn = document.getElementById('copyLogBtn');

  const productUrl = document.getElementById('productUrl');
  const rankingCount = document.getElementById('rankingCount');
  const rankingCountWrapper = document.getElementById('rankingCountWrapper');
  const addToQueueBtn = document.getElementById('addToQueueBtn');
  const addStatus = document.getElementById('addStatus');
  const urlCountLabel = document.getElementById('urlCountLabel');

  const logCard = document.getElementById('logCard');
  const logContainer = document.getElementById('logContainer');
  const clearLogBtn = document.getElementById('clearLogBtn');

  // ヘッダーボタン
  const settingsToggleBtn = document.getElementById('settingsToggleBtn');
  const helpToggleBtn = document.getElementById('helpToggleBtn');
  const gasCodeArea = document.getElementById('gasCodeArea');
  const copyGasCodeBtn = document.getElementById('copyGasCodeBtn');
  const spreadsheetUrlForCode = document.getElementById('spreadsheetUrlForCode');
  const spreadsheetIdStatus = document.getElementById('spreadsheetIdStatus');

  // キュー保存関連（ヘッダーアイコン方式）
  const saveQueueBtn = document.getElementById('saveQueueBtn');
  const loadSavedQueuesBtn = document.getElementById('loadSavedQueuesBtn');
  const savedQueuesDropdown = document.getElementById('savedQueuesDropdown');
  const savedQueuesDropdownList = document.getElementById('savedQueuesDropdownList');

  // ビュー切り替え
  const mainView = document.getElementById('main-view');
  const scheduledView = document.getElementById('scheduled-view');
  const settingsView = document.getElementById('settings-view');
  const helpView = document.getElementById('help-view');
  const scheduledViewBtn = document.getElementById('scheduledViewBtn');
  const backToMainBtn = document.getElementById('backToMainBtn');

  // 戻るボタン
  const settingsBackBtn = document.getElementById('settingsBackBtn');
  const helpBackBtn = document.getElementById('helpBackBtn');

  // 現在のビュー状態
  let currentView = 'main';

  // 定期収集関連
  const scheduledQueuesList = document.getElementById('scheduledQueuesList');
  const addScheduledQueueBtn = document.getElementById('addScheduledQueueBtn');
  const addScheduledQueueDropdown = document.getElementById('addScheduledQueueDropdown');
  const addScheduledQueueList = document.getElementById('addScheduledQueueList');

  // 定期収集ログ関連
  const scheduledLogCard = document.getElementById('scheduledLogCard');
  const scheduledLogContainer = document.getElementById('scheduledLogContainer');
  const copyScheduledLogBtn = document.getElementById('copyScheduledLogBtn');
  const clearScheduledLogBtn = document.getElementById('clearScheduledLogBtn');

  // 現在のスプレッドシートID
  let currentSpreadsheetId = '';

  // GASコードテンプレート（__SPREADSHEET_ID__がプレースホルダー）
  const GAS_CODE_TEMPLATE = `/**
 * 楽天レビュー収集 - Google Apps Script
 * Chrome拡張機能から送信されたレビューデータをスプレッドシートに保存する
 */

const SPREADSHEET_ID = '__SPREADSHEET_ID__';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = getSpreadsheet();
    const spreadsheetUrl = ss.getUrl();

    if (data.test) {
      return createResponse({ success: true, message: '接続テスト成功', spreadsheetUrl: spreadsheetUrl });
    }

    if (!data.reviews || data.reviews.length === 0) {
      return createResponse({ success: false, error: 'レビューデータがありません', spreadsheetUrl: spreadsheetUrl });
    }

    const separateSheets = data.separateSheets !== false;
    const savedCount = saveReviews(data.reviews, separateSheets);

    return createResponse({ success: true, message: savedCount + '件のレビューを保存しました', savedCount: savedCount, spreadsheetUrl: spreadsheetUrl });
  } catch (error) {
    console.error('エラー:', error);
    return createResponse({ success: false, error: error.message });
  }
}

function doGet(e) {
  const ss = getSpreadsheet();
  return createResponse({ success: true, message: '楽天レビュー収集 GAS API は正常に動作しています', timestamp: new Date().toISOString(), spreadsheetUrl: ss.getUrl() });
}

function saveReviews(reviews, separateSheets) {
  const ss = getSpreadsheet();
  if (separateSheets) {
    return saveReviewsByProduct(ss, reviews);
  } else {
    return saveReviewsToSingleSheet(ss, reviews);
  }
}

function saveReviewsByProduct(ss, reviews) {
  let totalSaved = 0;
  const reviewsByProduct = {};
  reviews.forEach(review => {
    const productId = review.productId || extractProductId(review.productUrl) || '不明な商品';
    if (!reviewsByProduct[productId]) reviewsByProduct[productId] = [];
    reviewsByProduct[productId].push(review);
  });

  for (const productId in reviewsByProduct) {
    const productReviews = reviewsByProduct[productId];
    let sheetName = sanitizeSheetName(productId);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      const defaultSheet = ss.getSheetByName('レビュー');
      if (defaultSheet && defaultSheet.getLastRow() <= 1) {
        defaultSheet.setName(sheetName);
        sheet = defaultSheet;
        addHeader(sheet);
      } else {
        sheet = ss.insertSheet(sheetName);
        addHeader(sheet);
      }
    }

    const rows = productReviews.map(review => [
      review.reviewDate || '', review.productId || extractProductId(review.productUrl) || '',
      review.productName || '', review.productUrl || '', review.rating || '',
      review.title || '', review.body || '', review.author || '',
      review.age || '', review.gender || '', review.orderDate || '',
      review.variation || '', review.usage || '', review.recipient || '',
      review.purchaseCount || '', review.helpfulCount || 0, review.shopReply || '',
      review.shopName || '', review.pageUrl || '', review.collectedAt || new Date().toISOString()
    ]);

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

function saveReviewsToSingleSheet(ss, reviews) {
  let sheet = ss.getSheetByName('レビュー');
  if (!sheet) {
    sheet = ss.insertSheet('レビュー');
    addHeader(sheet);
  }
  if (sheet.getLastRow() === 0) addHeader(sheet);

  const rows = reviews.map(review => [
    review.reviewDate || '', review.productId || extractProductId(review.productUrl) || '',
    review.productName || '', review.productUrl || '', review.rating || '',
    review.title || '', review.body || '', review.author || '',
    review.age || '', review.gender || '', review.orderDate || '',
    review.variation || '', review.usage || '', review.recipient || '',
    review.purchaseCount || '', review.helpfulCount || 0, review.shopReply || '',
    review.shopName || '', review.pageUrl || '', review.collectedAt || new Date().toISOString()
  ]);

  if (rows.length > 0) {
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length);
    dataRange.setValues(rows);
    dataRange.setVerticalAlignment('middle');
  }
  return rows.length;
}

function extractProductId(productUrl) {
  if (!productUrl) return null;
  try {
    const match = productUrl.match(/item\\.rakuten\\.co\\.jp\\/[^\\/]+\\/([^\\/\\?]+)/);
    if (match && match[1]) return match[1];
    const reviewMatch = productUrl.match(/review\\.rakuten\\.co\\.jp\\/item\\/\\d+\\/[^\\/]+\\/([^\\/\\?]+)/);
    if (reviewMatch && reviewMatch[1]) return reviewMatch[1];
    return null;
  } catch (e) { return null; }
}

function sanitizeSheetName(name) {
  let sanitized = name.replace(/[*?:\\\\/\\[\\]]/g, '');
  if (sanitized.length > 31) sanitized = sanitized.substring(0, 31);
  if (!sanitized.trim()) sanitized = '不明な商品';
  return sanitized;
}

function addHeader(sheet) {
  const headers = ['レビュー日', '商品管理番号', '商品名', '商品URL', '評価', 'タイトル', '本文', '投稿者', '年代', '性別', '注文日', 'バリエーション', '用途', '贈り先', '購入回数', '参考になった数', 'ショップからの返信', 'ショップ名', 'レビュー掲載URL', '収集日時'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#BF0000');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setVerticalAlignment('middle');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

function initializeSheet(sheet) {
  const headers = ['レビュー日', '商品管理番号', '商品名', '商品URL', '評価', 'タイトル', '本文', '投稿者', '年代', '性別', '注文日', 'バリエーション', '用途', '贈り先', '購入回数', '参考になった数', 'ショップからの返信', 'ショップ名', 'レビュー掲載URL', '収集日時'];
  sheet.clear();
  // 行数を調整（ヘッダー1行 + データ用1行 = 最低2行必要）
  const maxRows = sheet.getMaxRows();
  if (maxRows > 2) {
    sheet.deleteRows(3, maxRows - 2);
  } else if (maxRows < 2) {
    sheet.insertRows(2, 2 - maxRows);
  }
  // 余分な列を削除（ヘッダー列より後）
  const maxCols = sheet.getMaxColumns();
  if (maxCols > headers.length) {
    sheet.deleteColumns(headers.length + 1, maxCols - headers.length);
  }
  // ヘッダーを設定
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#BF0000');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setVerticalAlignment('middle');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

function createResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🛠️ レビュー管理')
    .addItem('📊 スプレッドシートを初期化', 'initializeSpreadsheet')
    .addItem('🔄 重複レビューを削除', 'removeDuplicates')
    .addToUi();
}

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
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).setVerticalAlignment('middle');
    }
    fixedCount++;
  });
  SpreadsheetApp.getUi().alert(fixedCount + '個のシートのヘッダーを赤色に修正しました');
}

function initializeSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('⚠️ スプレッドシートの初期化', 'すべてのシートとデータが削除されます。\\nこの操作は取り消せません。\\n\\n本当に初期化しますか？', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) { ui.alert('初期化をキャンセルしました'); return; }
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  // 既存の「レビュー」シートがあれば使用、なければ新規作成
  let reviewSheet = ss.getSheetByName('レビュー');
  if (!reviewSheet) {
    reviewSheet = ss.insertSheet('レビュー');
  }
  initializeSheet(reviewSheet);
  let deletedCount = 0;
  sheets.forEach(sheet => { if (sheet.getName() !== 'レビュー') { ss.deleteSheet(sheet); deletedCount++; } });
  ui.alert('✅ 初期化完了', deletedCount + '個のシートを削除しました。', ui.ButtonSet.OK);
}

function deleteEmptySheets() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  let deletedCount = 0;

  // まず空シートを特定
  const emptySheets = sheets.filter(sheet => sheet.getLastRow() <= 1);
  const nonEmptySheets = sheets.filter(sheet => sheet.getLastRow() > 1);

  // 空シートを削除（最低1シートは残す）
  emptySheets.forEach(sheet => {
    if (ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
      deletedCount++;
    }
  });

  // レビューが入っているシートがない場合、初期化シートを作成
  if (nonEmptySheets.length === 0) {
    let reviewSheet = ss.getSheetByName('レビュー');
    if (!reviewSheet) {
      // 残っているシートがあれば名前を変更、なければ新規作成
      const remaining = ss.getSheets();
      if (remaining.length > 0 && remaining[0].getLastRow() <= 1) {
        reviewSheet = remaining[0];
        reviewSheet.setName('レビュー');
      } else {
        reviewSheet = ss.insertSheet('レビュー');
      }
    }
    initializeSheet(reviewSheet);
    SpreadsheetApp.getUi().alert(deletedCount + '個の空シートを削除しました。\\n初期化済みの「レビュー」シートを作成しました。');
  } else {
    SpreadsheetApp.getUi().alert(deletedCount + '個の空シートを削除しました');
  }
}

function removeDuplicates() {
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  let totalRemoved = 0;
  sheets.forEach(sheet => {
    if (sheet.getLastRow() <= 1) return;
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    const seen = new Set();
    const uniqueRows = [];
    rows.forEach(row => {
      const key = (row[6] || '').substring(0, 100) + '|' + (row[7] || '');
      if (!seen.has(key)) { seen.add(key); uniqueRows.push(row); }
    });
    const removedCount = rows.length - uniqueRows.length;
    if (removedCount > 0) {
      sheet.clear();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      if (uniqueRows.length > 0) sheet.getRange(2, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
      addHeader(sheet);
      totalRemoved += removedCount;
    }
  });
  SpreadsheetApp.getUi().alert(totalRemoved + '件の重複を削除しました');
}`;

  // 初期化
  init();

  function init() {
    loadSettings();
    loadState();
    loadQueue();
    loadLogs();
    loadGasCode();
    loadSavedQueues();
    loadScheduledSettings();

    // イベントリスナー
    downloadBtn.addEventListener('click', downloadCSV);
    clearDataBtn.addEventListener('click', clearData);
    startQueueBtn.addEventListener('click', startQueueCollection);
    stopQueueBtn.addEventListener('click', stopQueueCollection);
    clearQueueBtn.addEventListener('click', clearQueue);
    addToQueueBtn.addEventListener('click', addToQueue);
    clearLogBtn.addEventListener('click', clearLogs);
    copyLogBtn.addEventListener('click', copyLogs);

    // キュー保存イベント（ヘッダーアイコン）
    if (saveQueueBtn) {
      saveQueueBtn.addEventListener('click', saveCurrentQueue);
    }
    if (loadSavedQueuesBtn) {
      loadSavedQueuesBtn.addEventListener('click', toggleSavedQueuesDropdown);
    }
    // ドロップダウン外クリックで閉じる
    document.addEventListener('click', (e) => {
      if (savedQueuesDropdown && savedQueuesDropdown.style.display !== 'none') {
        if (!savedQueuesDropdown.contains(e.target) && !loadSavedQueuesBtn.contains(e.target)) {
          savedQueuesDropdown.style.display = 'none';
        }
      }
    });

    // ビュー切り替えイベント（トグル動作）
    if (scheduledViewBtn) {
      scheduledViewBtn.addEventListener('click', () => {
        if (currentView === 'scheduled') {
          showMainView();
        } else {
          showScheduledView();
        }
      });
    }
    if (backToMainBtn) {
      backToMainBtn.addEventListener('click', showMainView);
    }

    // 定期収集キュー追加ドロップダウン
    if (addScheduledQueueBtn) {
      addScheduledQueueBtn.addEventListener('click', toggleAddScheduledQueueDropdown);
    }
    // ドロップダウン外クリックで閉じる
    document.addEventListener('click', (e) => {
      if (addScheduledQueueDropdown && addScheduledQueueDropdown.style.display !== 'none') {
        if (!addScheduledQueueDropdown.contains(e.target) && !addScheduledQueueBtn.contains(e.target)) {
          addScheduledQueueDropdown.style.display = 'none';
        }
      }
    });
    if (copyScheduledLogBtn) {
      copyScheduledLogBtn.addEventListener('click', copyScheduledLogs);
    }
    if (clearScheduledLogBtn) {
      clearScheduledLogBtn.addEventListener('click', clearScheduledLogs);
    }

    // ヘッダーボタンのイベント（トグル動作）
    settingsToggleBtn.addEventListener('click', () => {
      if (currentView === 'settings') {
        showMainView();
      } else {
        showSettingsView();
      }
    });
    helpToggleBtn.addEventListener('click', () => {
      if (currentView === 'help') {
        showMainView();
      } else {
        showHelpView();
      }
    });

    // 戻るボタンのイベント
    if (settingsBackBtn) {
      settingsBackBtn.addEventListener('click', showMainView);
    }
    if (helpBackBtn) {
      helpBackBtn.addEventListener('click', showMainView);
    }

    // URL入力時にランキングかどうか判定して件数入力の表示を切り替え、URLカウントを表示
    productUrl.addEventListener('input', () => {
      // 高さを自動調整
      productUrl.style.height = '38px';
      productUrl.style.height = Math.min(productUrl.scrollHeight, 120) + 'px';

      const text = productUrl.value.trim();
      const urls = text.split('\n').map(u => u.trim()).filter(u => u.length > 0);

      // ランキングURLチェック
      const hasRankingUrl = urls.some(u => u.includes('ranking.rakuten.co.jp'));
      if (hasRankingUrl && urls.length === 1) {
        rankingCountWrapper.style.display = 'flex';
      } else {
        rankingCountWrapper.style.display = 'none';
      }

      // URLカウント表示
      const validUrls = urls.filter(u =>
        u.includes('item.rakuten.co.jp') ||
        u.includes('review.rakuten.co.jp') ||
        u.includes('ranking.rakuten.co.jp')
      );

      if (urlCountLabel) {
        if (validUrls.length > 0) {
          urlCountLabel.textContent = `${validUrls.length}件のURL`;
          urlCountLabel.className = 'url-count-label has-urls';
        } else if (urls.length > 0) {
          urlCountLabel.textContent = '有効なURLがありません';
          urlCountLabel.className = 'url-count-label';
        } else {
          urlCountLabel.textContent = '';
          urlCountLabel.className = 'url-count-label';
        }
      }
    });

    // 通知設定のチェックボックス変更時に自動保存
    if (enableNotificationCheckbox) {
      enableNotificationCheckbox.addEventListener('change', saveNotificationSettings);
    }
    if (notifyPerProductCheckbox) {
      notifyPerProductCheckbox.addEventListener('change', saveNotificationSettings);
    }

    // GASコードコピーボタン
    if (copyGasCodeBtn) {
      copyGasCodeBtn.addEventListener('click', copyGasCode);
    }

    // スプレッドシートURL入力（自動保存）
    if (spreadsheetUrlForCode) {
      spreadsheetUrlForCode.addEventListener('input', handleSpreadsheetUrlInput);
    }

    // ウェブアプリURL入力（自動保存）
    if (gasUrlInput) {
      let gasUrlSaveTimeout = null;
      gasUrlInput.addEventListener('input', () => {
        if (gasUrlSaveTimeout) clearTimeout(gasUrlSaveTimeout);
        gasUrlSaveTimeout = setTimeout(() => {
          saveGasUrlAuto();
        }, 500);
      });
    }

    // バックグラウンドからのメッセージ
    chrome.runtime.onMessage.addListener(handleMessage);

    // 定期更新
    setInterval(() => {
      loadState();
      loadQueue();
    }, 2000);
  }

  function loadSettings() {
    chrome.storage.sync.get(['gasUrl', 'separateSheets', 'separateCsvFiles', 'spreadsheetUrl', 'enableNotification', 'notifyPerProduct'], (result) => {
      if (result.gasUrl) {
        gasUrlInput.value = result.gasUrl;
      }
      // CSV機能は常に表示（スプレッドシートと併用可能）
      dataButtons.style.display = 'flex';
      if (separateSheetsCheckbox) {
        separateSheetsCheckbox.checked = result.separateSheets !== false;
      }
      if (separateCsvFilesCheckbox) {
        separateCsvFilesCheckbox.checked = result.separateCsvFiles !== false;
      }
      // 通知設定（デフォルト: 通知ON、商品ごとOFF）
      if (enableNotificationCheckbox) {
        enableNotificationCheckbox.checked = result.enableNotification !== false;
      }
      if (notifyPerProductCheckbox) {
        notifyPerProductCheckbox.checked = result.notifyPerProduct === true;
      }

      if (result.spreadsheetUrl) {
        spreadsheetLink.href = result.spreadsheetUrl;
        spreadsheetLink.style.display = 'inline-flex';
      }
    });
  }

  function loadState() {
    chrome.storage.local.get(['collectionState', 'isQueueCollecting', 'collectingItems'], (result) => {
      const state = result.collectionState || {};
      const isQueueCollecting = result.isQueueCollecting || false;

      const hasData = (state.reviewCount || 0) > 0;
      downloadBtn.disabled = !hasData;
      clearDataBtn.disabled = !hasData;

      // 収集中かどうかでボタンを切り替え
      updateQueueButtons(isQueueCollecting);
    });
  }

  function loadQueue() {
    chrome.storage.local.get(['queue', 'collectingItems'], (result) => {
      const queue = result.queue || [];
      const collectingItems = result.collectingItems || [];
      const totalCount = queue.length + collectingItems.length;
      queueRemaining.textContent = `${totalCount}件`;
      startQueueBtn.disabled = totalCount === 0;

      if (totalCount === 0) {
        queueList.innerHTML = '';
        return;
      }

      // 収集中アイテムを先頭に表示
      const collectingHtml = collectingItems.map(item => `
        <div class="queue-item collecting">
          <div class="queue-item-info">
            <div class="queue-item-title">
              <span class="collecting-badge">収集中</span>
              ${escapeHtml(item.title || '商品')}
            </div>
            <div class="queue-item-url">${escapeHtml(item.url)}</div>
          </div>
        </div>
      `).join('');

      // 待機中アイテム
      const waitingHtml = queue.map((item, index) => `
        <div class="queue-item">
          <div class="queue-item-info">
            <div class="queue-item-title">${escapeHtml(item.title || '商品')}</div>
            <div class="queue-item-url">${escapeHtml(item.url)}</div>
          </div>
          <button class="queue-item-remove" data-index="${index}">×</button>
        </div>
      `).join('');

      queueList.innerHTML = collectingHtml + waitingHtml;

      // 削除ボタンのイベント
      queueList.querySelectorAll('.queue-item-remove').forEach(btn => {
        btn.addEventListener('click', (e) => {
          removeFromQueue(parseInt(e.target.dataset.index));
        });
      });
    });
  }

  function loadLogs() {
    chrome.storage.local.get(['logs'], (result) => {
      const logs = result.logs || [];
      if (logs.length === 0) {
        logCard.style.display = 'none';
        logContainer.innerHTML = '';
        return;
      }

      logCard.style.display = 'block';
      logContainer.innerHTML = logs.map(log => {
        const typeClass = log.type ? ` ${log.type}` : '';
        return `<div class="log-entry${typeClass}"><span class="time">[${log.time}]</span> ${escapeHtml(log.text)}</div>`;
      }).join('');

      logContainer.scrollTop = logContainer.scrollHeight;
    });
  }

  // ウェブアプリURLの自動保存
  async function saveGasUrlAuto() {
    const gasUrl = gasUrlInput.value.trim();
    const settingsStatus = document.getElementById('settingsStatus');

    if (gasUrl && !isValidGasUrl(gasUrl)) {
      showStatus(settingsStatus, 'error', 'URLの形式が正しくありません');
      return;
    }

    chrome.storage.sync.set({ gasUrl }, async () => {
      if (chrome.runtime.lastError) {
        showStatus(settingsStatus, 'error', '保存に失敗しました');
        return;
      }

      if (gasUrl) {
        // 接続テスト
        showStatus(settingsStatus, 'info', '接続テスト中...');
        try {
          const response = await fetch(gasUrl, { method: 'GET', mode: 'cors' });
          const data = await response.json();

          if (data.success) {
            showStatus(settingsStatus, 'success', '✓ 接続成功・保存完了');
            if (data.spreadsheetUrl) {
              chrome.storage.sync.set({ spreadsheetUrl: data.spreadsheetUrl });
              spreadsheetLink.href = data.spreadsheetUrl;
              spreadsheetLink.style.display = 'inline-flex';
            }
          } else {
            showStatus(settingsStatus, 'error', '接続失敗');
          }
        } catch (e) {
          showStatus(settingsStatus, 'success', '✓ 保存しました');
        }
      } else {
        spreadsheetLink.style.display = 'none';
        showStatus(settingsStatus, 'info', 'URLが空のため、CSVモードになります');
      }
    });
  }

  function isValidGasUrl(url) {
    return url.startsWith('https://script.google.com/macros/s/') && url.includes('/exec');
  }

  // スプレッドシートURLからIDを抽出
  function extractSpreadsheetId(url) {
    if (!url) return '';
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : '';
  }

  // GASコードを生成（スプレッドシートIDを埋め込み）
  function generateGasCode() {
    if (currentSpreadsheetId) {
      return GAS_CODE_TEMPLATE.replace('__SPREADSHEET_ID__', currentSpreadsheetId);
    } else {
      return GAS_CODE_TEMPLATE.replace('__SPREADSHEET_ID__', 'ここにスプレッドシートURLを入力してください');
    }
  }

  // GASコードをテキストエリアに表示
  function loadGasCode() {
    if (gasCodeArea) {
      gasCodeArea.value = generateGasCode();
    }
  }

  // スプレッドシートURL入力時の処理
  function handleSpreadsheetUrlInput() {
    const url = spreadsheetUrlForCode.value.trim();
    const id = extractSpreadsheetId(url);

    if (id) {
      currentSpreadsheetId = id;
      spreadsheetIdStatus.innerHTML = '<span style="color: #28a745;">✓ ID検出: ' + id.substring(0, 20) + '...</span>';
      // コードを更新
      loadGasCode();
    } else if (url) {
      currentSpreadsheetId = '';
      spreadsheetIdStatus.innerHTML = '<span style="color: #dc3545;">✗ 正しいスプレッドシートURLを入力してください</span>';
    } else {
      currentSpreadsheetId = '';
      spreadsheetIdStatus.innerHTML = '';
      loadGasCode();
    }
  }

  // GASコードをクリップボードにコピー
  function copyGasCode() {
    if (!gasCodeArea) return;

    if (!currentSpreadsheetId) {
      copyGasCodeBtn.textContent = 'URLを入力してください';
      copyGasCodeBtn.style.background = '#dc3545';
      setTimeout(() => {
        copyGasCodeBtn.textContent = '📋 コードをコピー';
        copyGasCodeBtn.style.background = '';
      }, 2000);
      return;
    }

    navigator.clipboard.writeText(generateGasCode()).then(() => {
      copyGasCodeBtn.textContent = 'コピーしました!';
      copyGasCodeBtn.style.background = '#28a745';
      setTimeout(() => {
        copyGasCodeBtn.textContent = '📋 コードをコピー';
        copyGasCodeBtn.style.background = '';
      }, 2000);
    }).catch(err => {
      console.error('コピー失敗:', err);
    });
  }

  // 通知設定のみを保存（チェックボックス変更時）
  function saveNotificationSettings() {
    const enableNotification = enableNotificationCheckbox ? enableNotificationCheckbox.checked : true;
    const notifyPerProduct = notifyPerProductCheckbox ? notifyPerProductCheckbox.checked : false;
    chrome.storage.sync.set({ enableNotification, notifyPerProduct });
  }

  async function downloadCSV() {
    console.log('downloadCSV called');
    // 設定を取得してからダウンロード処理
    chrome.storage.sync.get(['separateCsvFiles'], (syncResult) => {
      console.log('syncResult:', syncResult);
      const separateCsvFiles = syncResult.separateCsvFiles !== false;
      console.log('separateCsvFiles:', separateCsvFiles);

      chrome.storage.local.get(['collectionState'], async (result) => {
        console.log('collectionState result:', result);
        const state = result.collectionState;

        if (!state || !state.reviews || state.reviews.length === 0) {
          addLog('ダウンロードするデータがありません', 'error');
          console.log('No data to download');
          return;
        }

        console.log('Reviews count:', state.reviews.length);
        console.log('JSZip available:', typeof JSZip !== 'undefined');

        try {
          // 分割設定がOFFの場合、または商品が1つの場合は単一CSVをダウンロード
          if (!separateCsvFiles) {
            const csv = convertToCSV(state.reviews);
            downloadSingleCSV(csv, 'rakuten_reviews');
            addLog('CSVダウンロード完了', 'success');
            return;
          }

          // 商品ごとにレビューをグループ化
          const reviewsByProduct = {};
          state.reviews.forEach(review => {
            const productId = review.productId || 'unknown';
            if (!reviewsByProduct[productId]) {
              reviewsByProduct[productId] = [];
            }
            reviewsByProduct[productId].push(review);
          });

          const productIds = Object.keys(reviewsByProduct);

          // 商品が1つだけの場合は単一CSVをダウンロード
          if (productIds.length === 1) {
            const csv = convertToCSV(state.reviews);
            downloadSingleCSV(csv, productIds[0]);
            addLog('CSVダウンロード完了', 'success');
            return;
          }

          // 複数商品の場合はZIPでダウンロード
          // JSZipが利用できない場合は単一CSVにフォールバック
          if (typeof JSZip === 'undefined') {
            console.log('JSZip not available, falling back to single CSV');
            const csv = convertToCSV(state.reviews);
            downloadSingleCSV(csv, 'rakuten_reviews_all');
            addLog('CSVダウンロード完了（全商品統合）', 'success');
            return;
          }

          const zip = new JSZip();

          productIds.forEach(productId => {
            const reviews = reviewsByProduct[productId];
            const csv = convertToCSV(reviews);
            const filename = `${sanitizeFilename(productId)}.csv`;
            zip.file(filename, '\uFEFF' + csv);
          });

          const blob = await zip.generateAsync({ type: 'blob' });
          const url = URL.createObjectURL(blob);

          const now = new Date();
          const pad = (n) => String(n).padStart(2, '0');
          const zipFilename = `rakuten_reviews_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}.zip`;

          const a = document.createElement('a');
          a.href = url;
          a.download = zipFilename;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);

          addLog(`${productIds.length}商品分のCSVをZIPでダウンロード完了`, 'success');
        } catch (error) {
          console.error('CSV download error:', error);
          addLog('CSVダウンロード失敗: ' + error.message, 'error');
          // エラー時も単一CSVでフォールバック
          try {
            const csv = convertToCSV(state.reviews);
            downloadSingleCSV(csv, 'rakuten_reviews_fallback');
            addLog('フォールバック: 単一CSVとしてダウンロード完了', 'success');
          } catch (fallbackError) {
            console.error('Fallback download error:', fallbackError);
            addLog('CSVダウンロード完全失敗: ' + fallbackError.message, 'error');
          }
        }
      });
    });
  }

  function downloadSingleCSV(csv, productId) {
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);

    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const filename = `${sanitizeFilename(productId)}_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}.csv`;

    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function sanitizeFilename(name) {
    // ファイル名に使えない文字を置換
    return name.replace(/[\\/:*?"<>|]/g, '_').substring(0, 100);
  }

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

  function clearData() {
    if (!confirm('収集したデータをすべて削除しますか？')) return;

    chrome.storage.local.set({
      collectionState: {
        isRunning: false,
        reviewCount: 0,
        pageCount: 0,
        totalPages: 0,
        reviews: [],
        logs: []
      }
    }, () => {
      loadState();
      addLog('データをクリアしました', 'success');
    });
  }

  function removeFromQueue(index) {
    chrome.storage.local.get(['queue'], (result) => {
      const queue = result.queue || [];
      queue.splice(index, 1);
      chrome.storage.local.set({ queue }, () => {
        loadQueue();
      });
    });
  }

  function clearQueue() {
    // 収集中かチェック
    chrome.storage.local.get(['isQueueCollecting', 'collectingItems'], (result) => {
      const isCollecting = result.isQueueCollecting || (result.collectingItems && result.collectingItems.length > 0);

      const message = isCollecting
        ? 'キューをクリアし、収集中の処理も全て中止しますか？'
        : 'キューをクリアしますか？';

      if (!confirm(message)) return;

      // 収集中の場合は中止
      if (isCollecting) {
        chrome.runtime.sendMessage({ action: 'stopQueueCollection' }, () => {
          // キューをクリア
          chrome.storage.local.set({ queue: [], collectingItems: [] }, () => {
            loadQueue();
            addLog('収集を中止し、キューをクリアしました', 'error');
            updateQueueButtons(false);
          });
        });
      } else {
        // キューのみクリア
        chrome.storage.local.set({ queue: [] }, () => {
          loadQueue();
          addLog('キューをクリアしました');
        });
      }
    });
  }

  function startQueueCollection() {
    chrome.runtime.sendMessage({ action: 'startQueueCollection' }, (response) => {
      if (response && response.success) {
        addLog('キュー一括収集を開始しました', 'success');
        updateQueueButtons(true);
      } else {
        addLog('開始に失敗: ' + (response?.error || ''), 'error');
      }
    });
  }

  function stopQueueCollection() {
    chrome.runtime.sendMessage({ action: 'stopQueueCollection' }, (response) => {
      if (response && response.success) {
        addLog('収集を中止しました', 'error');
        updateQueueButtons(false);
      } else {
        addLog('中止に失敗: ' + (response?.error || ''), 'error');
      }
    });
  }

  function copyLogs() {
    chrome.storage.local.get(['logs'], (result) => {
      const logs = result.logs || [];
      if (logs.length === 0) {
        return;
      }

      const logText = logs.map(log => `[${log.time}] ${log.text}`).join('\n');
      navigator.clipboard.writeText(logText).then(() => {
        // コピー成功のフィードバック（色変化）
        copyLogBtn.style.background = '#28a745';
        copyLogBtn.style.color = 'white';
        copyLogBtn.title = 'コピーしました!';
        setTimeout(() => {
          copyLogBtn.style.background = '';
          copyLogBtn.style.color = '';
          copyLogBtn.title = 'ログをコピー';
        }, 1500);
      }).catch(err => {
        console.error('コピーに失敗:', err);
      });
    });
  }

  function updateQueueButtons(isRunning) {
    if (isRunning) {
      startQueueBtn.style.display = 'none';
      stopQueueBtn.style.display = 'block';
    } else {
      startQueueBtn.style.display = 'block';
      stopQueueBtn.style.display = 'none';
    }
  }

  async function addToQueue() {
    const text = productUrl.value.trim();

    if (!text) {
      showStatus(addStatus, 'error', 'URLを入力してください');
      return;
    }

    // 改行で分割して複数URLを取得
    const urls = text.split('\n').map(u => u.trim()).filter(u => u.length > 0);

    // ランキングURLの場合（1件のみ対応）
    const rankingUrl = urls.find(u => u.includes('ranking.rakuten.co.jp'));
    if (rankingUrl && urls.length === 1) {
      const count = parseInt(rankingCount.value) || 10;
      showStatus(addStatus, 'info', 'ランキングを取得中...');
      addToQueueBtn.disabled = true;

      try {
        chrome.runtime.sendMessage({
          action: 'fetchRanking',
          url: rankingUrl,
          count: count
        }, (response) => {
          addToQueueBtn.disabled = false;
          if (response && response.success) {
            showStatus(addStatus, 'success', `${response.addedCount}件追加しました`);
            loadQueue();
            addLog(`ランキングから${response.addedCount}件をキューに追加`, 'success');
            productUrl.value = '';
            rankingCountWrapper.style.display = 'none';
            if (urlCountLabel) {
              urlCountLabel.textContent = '';
              urlCountLabel.className = 'url-count-label';
            }
          } else {
            showStatus(addStatus, 'error', response?.error || '取得に失敗しました');
          }
        });
      } catch (e) {
        addToQueueBtn.disabled = false;
        showStatus(addStatus, 'error', '取得に失敗しました');
      }
      return;
    }

    // 商品URLの場合（複数対応）
    const productUrls = urls.filter(u =>
      u.includes('item.rakuten.co.jp') || u.includes('review.rakuten.co.jp')
    );

    if (productUrls.length === 0) {
      showStatus(addStatus, 'error', '楽天の商品ページまたはランキングURLを入力してください');
      return;
    }

    chrome.storage.local.get(['queue'], (result) => {
      const queue = result.queue || [];
      let addedCount = 0;
      let skippedCount = 0;

      productUrls.forEach(url => {
        // 重複チェック
        const exists = queue.some(item => item.url === url);
        if (exists) {
          skippedCount++;
          return;
        }

        // URLからタイトルを生成
        let productTitle = '商品';
        const pathMatch = url.match(/item\.rakuten\.co\.jp\/([^\/]+)\/([^\/\?]+)/);
        if (pathMatch) {
          productTitle = `${pathMatch[1]} - ${pathMatch[2]}`;
        }

        queue.push({
          url: url,
          title: productTitle.substring(0, 100),
          addedAt: new Date().toISOString()
        });
        addedCount++;
      });

      if (addedCount === 0 && skippedCount > 0) {
        showStatus(addStatus, 'error', `${skippedCount}件は既に追加済みです`);
        return;
      }

      chrome.storage.local.set({ queue }, () => {
        let message = `${addedCount}件追加しました`;
        if (skippedCount > 0) {
          message += `（${skippedCount}件は重複のためスキップ）`;
        }
        showStatus(addStatus, 'success', message);
        loadQueue();
        addLog(`${addedCount}件の商品をキューに追加`, 'success');
        productUrl.value = '';
        if (urlCountLabel) {
          urlCountLabel.textContent = '';
          urlCountLabel.className = 'url-count-label';
        }
      });
    });
  }

  function clearLogs() {
    // クリア成功のフィードバック（色変化）
    clearLogBtn.style.background = '#dc3545';
    clearLogBtn.style.color = 'white';
    clearLogBtn.title = 'クリアしました!';

    chrome.storage.local.set({ logs: [] }, () => {
      loadLogs();
      setTimeout(() => {
        clearLogBtn.style.background = '';
        clearLogBtn.style.color = '';
        clearLogBtn.title = 'クリア';
      }, 1500);
    });
  }

  function handleMessage(msg) {
    if (!msg || !msg.action) return;

    switch (msg.action) {
      case 'updateProgress':
        loadState();
        break;
      case 'collectionComplete':
        loadState();
        loadQueue();
        break;
      case 'queueUpdated':
        loadQueue();
        break;
      case 'log':
        addLog(msg.text, msg.type);
        break;
    }
  }

  function addLog(text, type = '') {
    const time = new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

    chrome.storage.local.get(['logs'], (result) => {
      const logs = result.logs || [];
      logs.push({ time, text, type });

      chrome.storage.local.set({ logs }, () => {
        loadLogs();
      });
    });
  }

  function showStatus(element, type, message) {
    element.textContent = message;
    element.className = 'status-message ' + type;

    if (type === 'success') {
      setTimeout(() => {
        element.className = 'status-message';
      }, 3000);
    }
  }

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // ========================================
  // キュー保存機能
  // ========================================

  function loadSavedQueues() {
    chrome.storage.local.get(['savedQueues'], (result) => {
      const savedQueues = result.savedQueues || [];
      renderSavedQueuesDropdown(savedQueues);
      renderScheduledQueues(savedQueues);
    });
  }

  // ドロップダウン表示/非表示
  function toggleSavedQueuesDropdown() {
    if (!savedQueuesDropdown) return;
    const isVisible = savedQueuesDropdown.style.display !== 'none';
    savedQueuesDropdown.style.display = isVisible ? 'none' : 'block';
  }

  // ドロップダウン内のキュー一覧をレンダリング
  function renderSavedQueuesDropdown(savedQueues) {
    if (!savedQueuesDropdownList) return;

    if (savedQueues.length === 0) {
      savedQueuesDropdownList.innerHTML = '<div class="saved-queues-empty">保存済みキューはありません</div>';
      return;
    }

    savedQueuesDropdownList.innerHTML = savedQueues.map(queue => `
      <div class="saved-queue-item" data-id="${queue.id}">
        <div class="saved-queue-info" data-id="${queue.id}">
          <span class="saved-queue-name">${escapeHtml(queue.name)}</span>
          <span class="saved-queue-count">${queue.items.length}件</span>
        </div>
        <div class="saved-queue-actions">
          <button class="icon-btn edit-queue-btn" data-id="${queue.id}" title="名前を変更">✏️</button>
          <button class="icon-btn delete-queue-btn" data-id="${queue.id}" title="削除">🗑️</button>
        </div>
      </div>
    `).join('');

    // イベントリスナー
    savedQueuesDropdownList.querySelectorAll('.saved-queue-info').forEach(el => {
      el.addEventListener('click', (e) => {
        loadSavedQueue(e.currentTarget.dataset.id);
        savedQueuesDropdown.style.display = 'none';
      });
    });
    savedQueuesDropdownList.querySelectorAll('.edit-queue-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        editSavedQueueName(e.target.dataset.id);
      });
    });
    savedQueuesDropdownList.querySelectorAll('.delete-queue-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        deleteSavedQueue(e.target.dataset.id);
      });
    });
  }

  // キューを保存（プロンプトで名前入力）
  function saveCurrentQueue() {
    chrome.storage.local.get(['queue', 'savedQueues'], (result) => {
      const currentQueue = result.queue || [];
      if (currentQueue.length === 0) {
        alert('キューが空です');
        return;
      }

      const name = prompt('保存するキューの名前を入力してください');
      if (!name || name.trim() === '') return;

      const savedQueues = result.savedQueues || [];
      const newQueue = {
        id: 'queue_' + Date.now(),
        name: name.trim(),
        createdAt: new Date().toISOString(),
        items: currentQueue.map(item => ({
          url: item.url,
          title: item.title
        }))
      };

      savedQueues.push(newQueue);

      chrome.storage.local.set({ savedQueues }, () => {
        loadSavedQueues();
        addLog(`キュー「${name}」を保存（${newQueue.items.length}件）`, 'success');
      });
    });
  }

  function loadSavedQueue(queueId) {
    chrome.storage.local.get(['queue', 'savedQueues'], (result) => {
      const savedQueues = result.savedQueues || [];
      const savedQueue = savedQueues.find(q => q.id === queueId);
      if (!savedQueue) return;

      const currentQueue = result.queue || [];
      let addedCount = 0;

      savedQueue.items.forEach(item => {
        const exists = currentQueue.some(q => q.url === item.url);
        if (!exists) {
          currentQueue.push({
            url: item.url,
            title: item.title,
            addedAt: new Date().toISOString()
          });
          addedCount++;
        }
      });

      chrome.storage.local.set({ queue: currentQueue }, () => {
        loadQueue();
        addLog(`「${savedQueue.name}」から${addedCount}件をキューに追加`, 'success');
      });
    });
  }

  function editSavedQueueName(queueId) {
    chrome.storage.local.get(['savedQueues'], (result) => {
      const savedQueues = result.savedQueues || [];
      const queue = savedQueues.find(q => q.id === queueId);
      if (!queue) return;

      const newName = prompt('新しいキュー名を入力', queue.name);
      if (!newName || newName.trim() === '') return;

      queue.name = newName.trim();

      chrome.storage.local.set({ savedQueues }, () => {
        loadSavedQueues();
        addLog(`キュー名を「${newName}」に変更`, 'success');
      });
    });
  }

  function deleteSavedQueue(queueId) {
    chrome.storage.local.get(['savedQueues', 'scheduledCollection'], (result) => {
      const savedQueues = result.savedQueues || [];
      const queue = savedQueues.find(q => q.id === queueId);
      if (!queue) return;

      if (!confirm(`「${queue.name}」を削除しますか？`)) return;

      const newQueues = savedQueues.filter(q => q.id !== queueId);

      // 定期収集の対象だった場合はクリア
      const scheduled = result.scheduledCollection || {};
      if (scheduled.targetQueueId === queueId) {
        scheduled.targetQueueId = '';
      }

      chrome.storage.local.set({ savedQueues: newQueues, scheduledCollection: scheduled }, () => {
        loadSavedQueues();
        loadScheduledSettings();
        addLog(`キュー「${queue.name}」を削除`, 'success');
      });
    });
  }

  // ========================================
  // ビュー切り替え機能
  // ========================================

  function hideAllViews() {
    if (mainView) mainView.classList.remove('active');
    if (scheduledView) scheduledView.classList.remove('active');
    if (settingsView) settingsView.classList.remove('active');
    if (helpView) helpView.classList.remove('active');
  }

  function showMainView() {
    hideAllViews();
    if (mainView) mainView.classList.add('active');
    currentView = 'main';
  }

  function showScheduledView() {
    hideAllViews();
    if (scheduledView) scheduledView.classList.add('active');
    currentView = 'scheduled';
    loadScheduledSettings();
    updateScheduledButtonsState();
  }

  function showSettingsView() {
    hideAllViews();
    if (settingsView) settingsView.classList.add('active');
    currentView = 'settings';
  }

  function showHelpView() {
    hideAllViews();
    if (helpView) helpView.classList.add('active');
    currentView = 'help';
  }

  // 定期収集ボタンのグレーアウト状態を更新
  function updateScheduledButtonsState() {
    chrome.storage.sync.get(['gasUrl'], (result) => {
      const globalGasUrl = result.gasUrl && result.gasUrl.trim() !== '';

      // キューカードのグレーアウト処理
      const queueCards = document.querySelectorAll('.scheduled-queue-card');
      queueCards.forEach(card => {
        const toggle = card.querySelector('.scheduled-queue-toggle');
        const runBtn = card.querySelector('.scheduled-queue-run-btn');
        const urlInput = card.querySelector('.scheduled-queue-url-input');

        // キュー個別のGAS URLまたはグローバルのGAS URLがあるか
        const queueGasUrl = urlInput ? urlInput.value.trim() : '';
        const hasValidUrl = queueGasUrl !== '' || globalGasUrl;

        if (!hasValidUrl) {
          card.classList.add('disabled');
          if (toggle) toggle.disabled = true;
          if (runBtn) runBtn.disabled = true;
        } else {
          card.classList.remove('disabled');
          if (toggle) toggle.disabled = false;
          if (runBtn) runBtn.disabled = false;
        }
      });
    });
  }

  // ========================================
  // 定期収集機能
  // ========================================

  function loadScheduledSettings() {
    chrome.storage.local.get(['scheduledQueues', 'savedQueues'], (result) => {
      const scheduledQueues = result.scheduledQueues || [];
      const savedQueues = result.savedQueues || [];

      renderScheduledQueues(scheduledQueues);
      renderAddScheduledQueueList(savedQueues, scheduledQueues);
    });
  }

  // 追加ドロップダウンの表示切り替え
  function toggleAddScheduledQueueDropdown() {
    if (!addScheduledQueueDropdown) return;
    const isVisible = addScheduledQueueDropdown.style.display !== 'none';
    addScheduledQueueDropdown.style.display = isVisible ? 'none' : 'block';
    if (!isVisible) {
      // ドロップダウンを開いたら最新のリストを表示
      loadScheduledSettings();
    }
  }

  // 追加用ドロップダウンのリストをレンダリング
  function renderAddScheduledQueueList(savedQueues, scheduledQueues) {
    if (!addScheduledQueueList) return;

    // 既に追加済みのキューを除外
    const addedIds = scheduledQueues.map(q => q.sourceQueueId);
    const availableQueues = savedQueues.filter(q => !addedIds.includes(q.id));

    if (availableQueues.length === 0) {
      addScheduledQueueList.innerHTML = '<div class="saved-queues-empty">追加できるキューがありません</div>';
      return;
    }

    addScheduledQueueList.innerHTML = availableQueues.map(queue => `
      <div class="saved-queue-item" data-id="${queue.id}">
        <div class="saved-queue-info">
          <span class="saved-queue-name">${escapeHtml(queue.name)}</span>
          <span class="saved-queue-count">${queue.items.length}件</span>
        </div>
      </div>
    `).join('');

    // クリックで追加
    addScheduledQueueList.querySelectorAll('.saved-queue-item').forEach(el => {
      el.addEventListener('click', () => {
        addToScheduledQueues(el.dataset.id);
        addScheduledQueueDropdown.style.display = 'none';
      });
    });
  }

  // 定期収集にキューを追加
  function addToScheduledQueues(savedQueueId) {
    chrome.storage.local.get(['savedQueues', 'scheduledQueues'], (result) => {
      const savedQueues = result.savedQueues || [];
      const scheduledQueues = result.scheduledQueues || [];
      const sourceQueue = savedQueues.find(q => q.id === savedQueueId);

      if (!sourceQueue) return;

      // 新しい定期収集キューを作成
      const newScheduledQueue = {
        id: 'sched_' + Date.now(),
        sourceQueueId: savedQueueId,
        name: sourceQueue.name,
        items: sourceQueue.items.slice(), // コピー
        time: '07:00',
        incrementalOnly: true,
        enabled: true,
        gasUrl: '',
        lastRun: null
      };

      scheduledQueues.push(newScheduledQueue);

      chrome.storage.local.set({ scheduledQueues }, () => {
        loadScheduledSettings();
        addScheduledLog(`「${sourceQueue.name}」を定期収集に追加`, 'success');
        updateScheduledAlarm();
      });
    });
  }

  // 定期収集画面のキュー一覧をレンダリング
  function renderScheduledQueues(scheduledQueues) {
    if (!scheduledQueuesList) return;

    if (scheduledQueues.length === 0) {
      scheduledQueuesList.innerHTML = `
        <div class="scheduled-queues-empty">
          定期収集キューがありません。<br>
          「保存キューから追加」で追加してください。
        </div>
      `;
      return;
    }

    // 時刻選択のHTML生成ヘルパー
    const generateHourOptions = (selected) => {
      let html = '';
      for (let i = 0; i < 24; i++) {
        const val = String(i).padStart(2, '0');
        html += `<option value="${val}" ${val === selected ? 'selected' : ''}>${i}</option>`;
      }
      return html;
    };

    const generateMinuteOptions = (selected) => {
      return ['00', '15', '30', '45'].map(val =>
        `<option value="${val}" ${val === selected ? 'selected' : ''}>${val}</option>`
      ).join('');
    };

    scheduledQueuesList.innerHTML = scheduledQueues.map(queue => {
      const [hours, minutes] = (queue.time || '07:00').split(':');
      const lastRun = queue.lastRun ? new Date(queue.lastRun) : null;
      const lastRunText = lastRun
        ? `${lastRun.toLocaleDateString('ja-JP')} ${lastRun.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' })}`
        : 'なし';

      return `
        <div class="scheduled-queue-card ${queue.enabled ? 'enabled' : ''}" data-id="${queue.id}">
          <div class="scheduled-queue-header">
            <div class="scheduled-queue-title">
              <span class="scheduled-queue-name">${escapeHtml(queue.name)}</span>
              <span class="scheduled-queue-count">${queue.items.length}件</span>
            </div>
            <div class="scheduled-queue-actions">
              <button class="scheduled-queue-run-btn" data-queue-id="${queue.id}">今すぐ実行</button>
              <label class="toggle-switch">
                <input type="checkbox" class="scheduled-queue-toggle" data-queue-id="${queue.id}" ${queue.enabled ? 'checked' : ''}>
                <span class="toggle-slider"></span>
              </label>
              <button class="scheduled-queue-delete-btn" data-queue-id="${queue.id}" title="削除">🗑️</button>
            </div>
          </div>
          <div class="scheduled-queue-settings">
            <div class="scheduled-queue-time-row">
              <span class="scheduled-queue-time-label">実行時刻:</span>
              <div class="time-picker">
                <select class="time-select scheduled-queue-hour" data-queue-id="${queue.id}">
                  ${generateHourOptions(hours)}
                </select>
                <span class="time-separator">:</span>
                <select class="time-select scheduled-queue-minute" data-queue-id="${queue.id}">
                  ${generateMinuteOptions(minutes)}
                </select>
              </div>
            </div>
            <div class="scheduled-queue-checkbox-row">
              <label class="checkbox-label">
                <input type="checkbox" class="scheduled-queue-incremental" data-queue-id="${queue.id}" ${queue.incrementalOnly ? 'checked' : ''}>
                <span>差分取得（新着のみ）</span>
              </label>
            </div>
            <div class="scheduled-queue-url-row">
              <span class="scheduled-queue-url-label">GAS URL:</span>
              <input type="text" class="scheduled-queue-url-input" data-queue-id="${queue.id}"
                     value="${escapeHtml(queue.gasUrl || '')}" placeholder="（通常設定を使用）">
            </div>
            <div class="scheduled-queue-last-run">
              最終実行: ${lastRunText}
            </div>
          </div>
        </div>
      `;
    }).join('');

    // イベントリスナー
    scheduledQueuesList.querySelectorAll('.scheduled-queue-toggle').forEach(toggle => {
      toggle.addEventListener('change', (e) => {
        updateScheduledQueueProperty(e.target.dataset.queueId, 'enabled', e.target.checked);
      });
    });

    scheduledQueuesList.querySelectorAll('.scheduled-queue-run-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        runScheduledQueueNow(e.target.dataset.queueId);
      });
    });

    scheduledQueuesList.querySelectorAll('.scheduled-queue-delete-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        deleteScheduledQueue(e.target.dataset.queueId);
      });
    });

    scheduledQueuesList.querySelectorAll('.scheduled-queue-hour').forEach(select => {
      select.addEventListener('change', (e) => {
        updateScheduledQueueTime(e.target.dataset.queueId);
      });
    });

    scheduledQueuesList.querySelectorAll('.scheduled-queue-minute').forEach(select => {
      select.addEventListener('change', (e) => {
        updateScheduledQueueTime(e.target.dataset.queueId);
      });
    });

    scheduledQueuesList.querySelectorAll('.scheduled-queue-incremental').forEach(checkbox => {
      checkbox.addEventListener('change', (e) => {
        updateScheduledQueueProperty(e.target.dataset.queueId, 'incrementalOnly', e.target.checked);
      });
    });

    scheduledQueuesList.querySelectorAll('.scheduled-queue-url-input').forEach(input => {
      let saveTimeout = null;
      input.addEventListener('input', (e) => {
        if (saveTimeout) clearTimeout(saveTimeout);
        saveTimeout = setTimeout(() => {
          updateScheduledQueueProperty(e.target.dataset.queueId, 'gasUrl', e.target.value.trim(), e.target);
        }, 500);
      });
    });

    updateScheduledButtonsState();
  }

  // 定期収集キューのプロパティを更新
  function updateScheduledQueueProperty(queueId, property, value, inputElement = null) {
    chrome.storage.local.get(['scheduledQueues'], (result) => {
      const scheduledQueues = result.scheduledQueues || [];
      const queue = scheduledQueues.find(q => q.id === queueId);
      if (queue) {
        queue[property] = value;
        chrome.storage.local.set({ scheduledQueues }, () => {
          if (property === 'enabled') {
            const card = scheduledQueuesList.querySelector(`.scheduled-queue-card[data-id="${queueId}"]`);
            if (card) card.classList.toggle('enabled', value);
          }
          if (inputElement) {
            showAutoSaveIndicator(inputElement);
          }
          updateScheduledAlarm();
        });
      }
    });
  }

  // 定期収集キューの時刻を更新
  function updateScheduledQueueTime(queueId) {
    const hourSelect = scheduledQueuesList.querySelector(`.scheduled-queue-hour[data-queue-id="${queueId}"]`);
    const minuteSelect = scheduledQueuesList.querySelector(`.scheduled-queue-minute[data-queue-id="${queueId}"]`);
    if (hourSelect && minuteSelect) {
      const time = `${hourSelect.value}:${minuteSelect.value}`;
      updateScheduledQueueProperty(queueId, 'time', time);
    }
  }

  // 定期収集キューを削除
  function deleteScheduledQueue(queueId) {
    chrome.storage.local.get(['scheduledQueues'], (result) => {
      const scheduledQueues = result.scheduledQueues || [];
      const queue = scheduledQueues.find(q => q.id === queueId);
      if (!queue) return;

      if (!confirm(`「${queue.name}」を定期収集から削除しますか？`)) return;

      const newQueues = scheduledQueues.filter(q => q.id !== queueId);
      chrome.storage.local.set({ scheduledQueues: newQueues }, () => {
        loadScheduledSettings();
        addScheduledLog(`「${queue.name}」を定期収集から削除`, 'success');
        updateScheduledAlarm();
      });
    });
  }

  // 定期収集キューを今すぐ実行
  function runScheduledQueueNow(queueId) {
    chrome.storage.local.get(['scheduledQueues'], (result) => {
      const scheduledQueues = result.scheduledQueues || [];
      const targetQueue = scheduledQueues.find(q => q.id === queueId);

      if (!targetQueue || targetQueue.items.length === 0) {
        addScheduledLog('キューが見つからないか、空です', 'error');
        return;
      }

      chrome.storage.local.get(['queue'], (queueResult) => {
        const currentQueue = queueResult.queue || [];
        let addedCount = 0;

        targetQueue.items.forEach(item => {
          const exists = currentQueue.some(q => q.url === item.url);
          if (!exists) {
            currentQueue.push({
              url: item.url,
              title: item.title,
              addedAt: new Date().toISOString(),
              scheduledRun: true,
              incrementalOnly: targetQueue.incrementalOnly,
              gasUrl: targetQueue.gasUrl || null
            });
            addedCount++;
          }
        });

        if (addedCount === 0) {
          addScheduledLog(`「${targetQueue.name}」は全て収集済みまたはキューに追加済みです`, 'error');
          return;
        }

        chrome.storage.local.set({ queue: currentQueue }, () => {
          loadQueue();
          addScheduledLog(`「${targetQueue.name}」の収集を開始（${addedCount}件）`, 'success');
          chrome.runtime.sendMessage({ action: 'startQueueCollection' });
        });
      });
    });
  }

  // 自動保存インジケーターを表示
  function showAutoSaveIndicator(inputElement) {
    const existingIndicator = inputElement.parentNode.querySelector('.auto-save-indicator');
    if (existingIndicator) existingIndicator.remove();

    const indicator = document.createElement('span');
    indicator.className = 'auto-save-indicator';
    indicator.innerHTML = '✓ 保存';
    inputElement.parentNode.appendChild(indicator);

    setTimeout(() => indicator.remove(), 2000);
  }

  // アラームを更新
  function updateScheduledAlarm() {
    chrome.storage.local.get(['scheduledQueues'], (result) => {
      const scheduledQueues = result.scheduledQueues || [];
      const enabledQueues = scheduledQueues.filter(q => q.enabled);

      chrome.runtime.sendMessage({
        action: 'updateScheduledAlarm',
        settings: { queues: enabledQueues }
      });
    });
  }

  // ========================================
  // 定期収集ログ機能
  // ========================================

  function addScheduledLog(text, type = '') {
    const timestamp = new Date().toLocaleTimeString('ja-JP', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    });

    const logEntry = document.createElement('div');
    logEntry.className = 'log-entry' + (type ? ` log-${type}` : '');
    logEntry.innerHTML = `<span class="log-time">${timestamp}</span><span class="log-text">${escapeHtml(text)}</span>`;

    if (scheduledLogContainer) {
      scheduledLogContainer.appendChild(logEntry);
      scheduledLogContainer.scrollTop = scheduledLogContainer.scrollHeight;
    }

    // ストレージにも保存
    chrome.storage.local.get(['scheduledLogs'], (result) => {
      const logs = result.scheduledLogs || [];
      logs.push({ timestamp: new Date().toISOString(), text, type });
      // 最新500件のみ保持
      const trimmedLogs = logs.slice(-500);
      chrome.storage.local.set({ scheduledLogs: trimmedLogs });
    });
  }

  function loadScheduledLogs() {
    chrome.storage.local.get(['scheduledLogs'], (result) => {
      const logs = result.scheduledLogs || [];
      if (scheduledLogContainer) {
        scheduledLogContainer.innerHTML = '';
        logs.forEach(log => {
          const timestamp = new Date(log.timestamp).toLocaleTimeString('ja-JP', {
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
          });
          const logEntry = document.createElement('div');
          logEntry.className = 'log-entry' + (log.type ? ` log-${log.type}` : '');
          logEntry.innerHTML = `<span class="log-time">${timestamp}</span><span class="log-text">${escapeHtml(log.text)}</span>`;
          scheduledLogContainer.appendChild(logEntry);
        });
        scheduledLogContainer.scrollTop = scheduledLogContainer.scrollHeight;
      }
    });
  }

  function clearScheduledLogs() {
    chrome.storage.local.set({ scheduledLogs: [] }, () => {
      if (scheduledLogContainer) {
        scheduledLogContainer.innerHTML = '';
      }
    });
  }

  function copyScheduledLogs() {
    chrome.storage.local.get(['scheduledLogs'], (result) => {
      const logs = result.scheduledLogs || [];
      const text = logs.map(log => {
        const timestamp = new Date(log.timestamp).toLocaleString('ja-JP');
        return `[${timestamp}] ${log.text}`;
      }).join('\n');

      navigator.clipboard.writeText(text).then(() => {
        addScheduledLog('ログをクリップボードにコピーしました', 'success');
      });
    });
  }
});
