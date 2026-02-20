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

  // バージョン表示
  const versionDisplay = document.getElementById('versionDisplay');
  if (versionDisplay) {
    const manifest = chrome.runtime.getManifest();
    versionDisplay.textContent = `v${manifest.version}`;
  }

  // クイックスタートガイド（初回表示）
  initQuickStartGuide();

  // DOM要素
  const queueRemaining = document.getElementById('queueRemaining');
  const spreadsheetLinkRakutenEl = document.getElementById('spreadsheetLinkRakuten');
  const spreadsheetLinkAmazonEl = document.getElementById('spreadsheetLinkAmazon');
  const spreadsheetLinkBtn = document.getElementById('spreadsheetLinkBtn');
  const spreadsheetLinkDropdown = document.getElementById('spreadsheetLinkDropdown');
  const downloadBtn = document.getElementById('downloadBtn');
  const clearDataBtn = document.getElementById('clearDataBtn');
  const dataButtons = document.getElementById('dataButtons');

  const spreadsheetUrlInput = document.getElementById('spreadsheetUrl');
  const spreadsheetUrlStatus = document.getElementById('spreadsheetUrlStatus');
  const spreadsheetTitleEl = document.getElementById('spreadsheetTitle');
  const amazonSpreadsheetUrlInput = document.getElementById('amazonSpreadsheetUrl');
  const amazonSpreadsheetUrlStatus = document.getElementById('amazonSpreadsheetUrlStatus');
  const amazonSpreadsheetTitleEl = document.getElementById('amazonSpreadsheetTitle');
  const separateSheetsCheckbox = document.getElementById('separateSheets');
  const separateCsvFilesCheckbox = document.getElementById('separateCsvFiles');
  const enableNotificationCheckbox = document.getElementById('enableNotification');
  const notifyPerProductCheckbox = document.getElementById('notifyPerProduct');
  const showScheduledCollectionCheckbox = document.getElementById('showScheduledCollection');
  const scheduledCollectionSection = document.getElementById('scheduledCollectionSection');

  // 期間指定フィルター
  const dateFilterFromInput = document.getElementById('dateFilterFrom');
  const dateFilterToInput = document.getElementById('dateFilterTo');

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

  const logSection = document.getElementById('logSection');
  const logContainer = document.getElementById('logContainer');
  const clearLogBtn = document.getElementById('clearLogBtn');

  // 商品情報ログ
  const productLogSection = document.getElementById('productLogSection');
  const productLogContainer = document.getElementById('productLogContainer');
  const copyProductLogBtn = document.getElementById('copyProductLogBtn');
  const clearProductLogBtn = document.getElementById('clearProductLogBtn');

  // ヘッダーボタン
  const headerTitle = document.getElementById('headerTitle');
  const settingsToggleBtn = document.getElementById('settingsToggleBtn');

  // キュー保存関連（ヘッダーアイコン方式）
  const saveQueueBtn = document.getElementById('saveQueueBtn');
  const loadSavedQueuesBtn = document.getElementById('loadSavedQueuesBtn');
  const savedQueuesDropdown = document.getElementById('savedQueuesDropdown');
  const savedQueuesDropdownList = document.getElementById('savedQueuesDropdownList');

  // ビュー切り替え
  const mainView = document.getElementById('main-view');
  const settingsView = document.getElementById('settings-view');

  // 戻るボタン
  const settingsBackBtn = document.getElementById('settingsBackBtn');

  // 現在のビュー状態
  let currentView = 'main';

  // 商品情報収集（Drive）関連
  const productInfoFolderUrlInput = document.getElementById('productInfoFolderUrl');
  const productInfoFolderUrlStatus = document.getElementById('productInfoFolderUrlStatus');
  const batchProductAsins = document.getElementById('batchProductAsins');
  const startBatchProductBtn = document.getElementById('startBatchProductBtn');  // 「追加」ボタン
  const startBatchProductRunBtn = document.getElementById('startBatchProductRunBtn');  // 「収集開始」ボタン
  const cancelBatchProductBtn = document.getElementById('cancelBatchProductBtn');
  const batchProductStatus = document.getElementById('batchProductStatus');
  const batchProductList = document.getElementById('batchProductList');
  const batchProductCountEl = document.getElementById('batchProductCount');
  // 商品情報収集のキューリスト（メモリ上で管理）
  let batchProductQueue = [];

  // 商品キュー保存関連
  const saveProductQueueBtn = document.getElementById('saveProductQueueBtn');
  const loadProductQueuesBtn = document.getElementById('loadProductQueuesBtn');
  const productQueuesDropdown = document.getElementById('productQueuesDropdown');
  const productQueuesDropdownList = document.getElementById('productQueuesDropdownList');
  const clearProductQueueBtn = document.getElementById('clearProductQueueBtn');

  // 定期収集関連
  const scheduledQueuesList = document.getElementById('scheduledQueuesList');
  const addScheduledQueueBtn = document.getElementById('addScheduledQueueBtn');
  const addScheduledQueueDropdown = document.getElementById('addScheduledQueueDropdown');
  const addScheduledQueueList = document.getElementById('addScheduledQueueList');

  // スプレッドシートボタンの有効/無効を更新
  function updateSpreadsheetBtnState() {
    if (!spreadsheetLinkBtn) return;
    const rakutenHasUrl = spreadsheetLinkRakutenEl && !spreadsheetLinkRakutenEl.classList.contains('disabled');
    const amazonHasUrl = spreadsheetLinkAmazonEl && !spreadsheetLinkAmazonEl.classList.contains('disabled');
    if (rakutenHasUrl || amazonHasUrl) {
      spreadsheetLinkBtn.classList.remove('disabled');
    } else {
      spreadsheetLinkBtn.classList.add('disabled');
    }
  }

  // Google Driveフォルダリンクの状態更新
  function updateDriveFolderLink(url) {
    const link = document.getElementById('driveFolderLink');
    if (!link) return;
    if (url && url.trim()) {
      link.href = url.trim();
      link.classList.remove('disabled');
    } else {
      link.href = '#';
      link.classList.add('disabled');
    }
  }

  // 初期化
  init();

  function init() {
    loadSettings();
    loadState();
    loadQueue();
    loadBatchProductQueue();
    initLogsFromStorage();
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
    if (clearProductLogBtn) clearProductLogBtn.addEventListener('click', clearProductLogs);
    if (copyProductLogBtn) copyProductLogBtn.addEventListener('click', copyProductLogs);

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

    // ヘッダーのスプレッドシートリンクドロップダウン
    if (spreadsheetLinkBtn && spreadsheetLinkDropdown) {
      spreadsheetLinkBtn.addEventListener('click', (e) => {
        if (spreadsheetLinkBtn.classList.contains('disabled')) return;
        e.stopPropagation();
        spreadsheetLinkDropdown.classList.toggle('show');
      });

      // ドロップダウン外クリックで閉じる
      document.addEventListener('click', (e) => {
        if (!spreadsheetLinkDropdown.contains(e.target) && !spreadsheetLinkBtn.contains(e.target)) {
          spreadsheetLinkDropdown.classList.remove('show');
        }
      });
    }

    // ヘッダータイトルクリックで収集画面に遷移
    if (headerTitle) {
      headerTitle.addEventListener('click', showMainView);
    }

    // ヘッダーボタンのイベント（トグル動作）
    settingsToggleBtn.addEventListener('click', () => {
      if (currentView === 'settings') {
        showMainView();
      } else {
        showSettingsView();
      }
    });

    // 戻るボタンのイベント
    if (settingsBackBtn) {
      settingsBackBtn.addEventListener('click', showMainView);
    }

    // URL入力時にランキングかどうか判定して件数入力の表示を切り替え、URLカウントを表示
    productUrl.addEventListener('input', () => {
      // 高さを自動調整（空なら初期サイズに戻す）
      autoResizeTextarea(productUrl);

      const text = productUrl.value.trim();
      const urls = text.split('\n').map(u => u.trim()).filter(u => u.length > 0);

      // ランキングURLチェック（楽天 + Amazon）
      const isRakutenRanking = urls.some(u => u.includes('ranking.rakuten.co.jp'));
      const isAmazonRanking = urls.some(u =>
        u.includes('amazon.co.jp') && (u.includes('/bestsellers/') || u.includes('/ranking/'))
      );
      const hasRankingUrl = isRakutenRanking || isAmazonRanking;
      if (hasRankingUrl && urls.length === 1) {
        rankingCountWrapper.style.display = 'flex';
      } else {
        rankingCountWrapper.style.display = 'none';
      }

      // URLカウント表示（楽天 + Amazon）
      const validUrls = urls.filter(u =>
        u.includes('item.rakuten.co.jp') ||
        u.includes('review.rakuten.co.jp') ||
        u.includes('ranking.rakuten.co.jp') ||
        (u.includes('amazon.co.jp') && (
          u.includes('/dp/') ||
          u.includes('/gp/product/') ||
          u.includes('/product-reviews/') ||
          u.includes('/bestsellers/') ||
          u.includes('/ranking/')
        ))
      );


      // 追加ボタンの色を変更
      if (addToQueueBtn) {
        if (validUrls.length > 0) {
          addToQueueBtn.classList.remove('btn-secondary');
          addToQueueBtn.classList.add('btn-primary');
          // 有効なURLが入力されたらエラーメッセージをクリア
          if (addStatus) {
            addStatus.textContent = '';
            addStatus.className = 'status-message';
          }
        } else {
          addToQueueBtn.classList.remove('btn-primary');
          addToQueueBtn.classList.add('btn-secondary');
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

    // シート分割・CSV分割設定のチェックボックス変更時に自動保存
    if (separateSheetsCheckbox) {
      separateSheetsCheckbox.addEventListener('change', saveSheetSettings);
    }
    if (separateCsvFilesCheckbox) {
      separateCsvFilesCheckbox.addEventListener('change', saveSheetSettings);
    }

    // 定期収集表示設定のチェックボックス変更時に自動保存
    if (showScheduledCollectionCheckbox) {
      showScheduledCollectionCheckbox.addEventListener('change', saveScheduledCollectionVisibility);
    }

    // 期間指定フィルターの変更時に自動保存
    if (dateFilterFromInput) {
      dateFilterFromInput.addEventListener('change', saveDateFilter);
    }
    if (dateFilterToInput) {
      dateFilterToInput.addEventListener('change', saveDateFilter);
    }

    // スプレッドシートURL入力（自動保存 - Sheets API直接連携）
    if (spreadsheetUrlInput) {
      let spreadsheetUrlSaveTimeout = null;
      spreadsheetUrlInput.addEventListener('input', () => {
        if (spreadsheetUrlSaveTimeout) clearTimeout(spreadsheetUrlSaveTimeout);
        spreadsheetUrlSaveTimeout = setTimeout(() => {
          saveSpreadsheetUrlAuto();
        }, 500);
      });
    }

    // Amazon用スプレッドシートURL入力（自動保存）
    if (amazonSpreadsheetUrlInput) {
      let amazonSpreadsheetUrlSaveTimeout = null;
      amazonSpreadsheetUrlInput.addEventListener('input', () => {
        if (amazonSpreadsheetUrlSaveTimeout) clearTimeout(amazonSpreadsheetUrlSaveTimeout);
        amazonSpreadsheetUrlSaveTimeout = setTimeout(() => {
          saveAmazonSpreadsheetUrlAuto();
        }, 500);
      });
    }

    // 商品情報収集（Drive）関連
    if (productInfoFolderUrlInput) {
      let folderUrlSaveTimeout = null;
      productInfoFolderUrlInput.addEventListener('input', () => {
        if (folderUrlSaveTimeout) clearTimeout(folderUrlSaveTimeout);
        folderUrlSaveTimeout = setTimeout(() => {
          saveProductInfoFolderUrl();
        }, 500);
      });
    }
    if (startBatchProductBtn) {
      startBatchProductBtn.addEventListener('click', addToBatchProductQueue);
    }
    if (startBatchProductRunBtn) {
      startBatchProductRunBtn.addEventListener('click', startBatchProductCollection);
    }
    if (cancelBatchProductBtn) {
      cancelBatchProductBtn.addEventListener('click', cancelBatchProductCollection);
    }

    // 商品キュー保存/読み込み/クリア
    if (saveProductQueueBtn) saveProductQueueBtn.addEventListener('click', saveProductQueue);
    if (loadProductQueuesBtn) loadProductQueuesBtn.addEventListener('click', toggleProductQueuesDropdown);
    if (clearProductQueueBtn) {
      clearProductQueueBtn.addEventListener('click', () => {
        if (batchProductQueue.length === 0) return;
        batchProductQueue = [];
        chrome.storage.local.set({ batchProductQueue: [] });
        renderBatchProductQueue();
        addLog('キューをクリアしました', '', 'product');
      });
    }
    // ドロップダウン外クリックで閉じる
    document.addEventListener('click', (e) => {
      if (productQueuesDropdown && productQueuesDropdown.style.display !== 'none') {
        if (!productQueuesDropdown.contains(e.target) && !loadProductQueuesBtn.contains(e.target)) {
          productQueuesDropdown.style.display = 'none';
        }
      }
    });

    // 商品入力欄の高さ自動調整
    if (batchProductAsins) {
      batchProductAsins.addEventListener('input', () => {
        autoResizeTextarea(batchProductAsins);
      });
    }

    // フォルダピッカー初期化
    initFolderPicker();

    // バックグラウンドからのメッセージ
    chrome.runtime.onMessage.addListener(handleMessage);

    // 定期更新
    setInterval(() => {
      loadState();
      loadQueue();
    }, 2000);
  }

  function loadSettings() {
    chrome.storage.sync.get(['separateSheets', 'separateCsvFiles', 'spreadsheetUrl', 'amazonSpreadsheetUrl', 'enableNotification', 'notifyPerProduct', 'showScheduledCollection', 'dateFilterFrom', 'dateFilterTo', 'productInfoFolderUrl'], (result) => {
      // 楽天用スプレッドシートURL（Sheets API直接連携）
      if (result.spreadsheetUrl && spreadsheetUrlInput) {
        spreadsheetUrlInput.value = result.spreadsheetUrl;
        if (spreadsheetLinkRakutenEl) {
          spreadsheetLinkRakutenEl.href = result.spreadsheetUrl;
          spreadsheetLinkRakutenEl.classList.remove('disabled');
        }
        // タイトル取得
        fetchAndShowSpreadsheetTitle(result.spreadsheetUrl, spreadsheetTitleEl, spreadsheetUrlInput);
      }
      // Amazon用スプレッドシートURL
      if (result.amazonSpreadsheetUrl && amazonSpreadsheetUrlInput) {
        amazonSpreadsheetUrlInput.value = result.amazonSpreadsheetUrl;
        if (spreadsheetLinkAmazonEl) {
          spreadsheetLinkAmazonEl.href = result.amazonSpreadsheetUrl;
          spreadsheetLinkAmazonEl.classList.remove('disabled');
        }
        // タイトル取得
        fetchAndShowSpreadsheetTitle(result.amazonSpreadsheetUrl, amazonSpreadsheetTitleEl, amazonSpreadsheetUrlInput);
      }
      // CSV機能は常に表示（スプレッドシートと併用可能）
      dataButtons.style.display = 'flex';
      // スプレッドシートボタンの状態を更新
      updateSpreadsheetBtnState();
      if (separateSheetsCheckbox) {
        // デフォルトはオフ（false）
        separateSheetsCheckbox.checked = result.separateSheets === true;
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
      // 定期収集表示設定（デフォルト: 非表示 = false）
      if (showScheduledCollectionCheckbox && scheduledCollectionSection) {
        const show = result.showScheduledCollection === true;
        showScheduledCollectionCheckbox.checked = show;
        scheduledCollectionSection.style.display = show ? 'block' : 'none';
      }
      // 期間指定フィルター
      if (dateFilterFromInput && result.dateFilterFrom) {
        dateFilterFromInput.value = result.dateFilterFrom;
      }
      if (dateFilterToInput && result.dateFilterTo) {
        dateFilterToInput.value = result.dateFilterTo;
      }
      // 商品情報収集フォルダURL
      if (productInfoFolderUrlInput && result.productInfoFolderUrl) {
        productInfoFolderUrlInput.value = result.productInfoFolderUrl;
      }
      // Google Driveフォルダリンクの状態更新
      updateDriveFolderLink(result.productInfoFolderUrl);
    });

    // 収集項目の設定を読み込み
    loadCollectionFields();
  }

  // 収集項目の設定を読み込み
  function loadCollectionFields() {
    chrome.storage.sync.get(['rakutenFields', 'amazonFields'], (result) => {
      // 楽天のデフォルト項目
      const defaultRakutenFields = ['rating', 'title', 'body', 'productUrl'];
      const rakutenFields = result.rakutenFields || defaultRakutenFields;

      // Amazonのデフォルト項目
      const defaultAmazonFields = ['rating', 'title', 'body'];
      const amazonFields = result.amazonFields || defaultAmazonFields;

      // 楽天のチェックボックスを更新
      const rakutenGrid = document.getElementById('rakutenFieldsGrid');
      if (rakutenGrid) {
        rakutenGrid.querySelectorAll('input[type="checkbox"]').forEach(cb => {
          cb.checked = rakutenFields.includes(cb.dataset.field);
          cb.addEventListener('change', saveCollectionFields);
        });
      }

      // Amazonのチェックボックスを更新
      const amazonGrid = document.getElementById('amazonFieldsGrid');
      if (amazonGrid) {
        amazonGrid.querySelectorAll('input[type="checkbox"]').forEach(cb => {
          cb.checked = amazonFields.includes(cb.dataset.field);
          cb.addEventListener('change', saveCollectionFields);
        });
      }
    });
  }

  // 収集項目の設定を保存
  function saveCollectionFields() {
    const rakutenFields = [];
    const amazonFields = [];

    // 楽天のチェックされた項目を収集
    const rakutenGrid = document.getElementById('rakutenFieldsGrid');
    if (rakutenGrid) {
      rakutenGrid.querySelectorAll('input[type="checkbox"]:checked').forEach(cb => {
        rakutenFields.push(cb.dataset.field);
      });
    }

    // Amazonのチェックされた項目を収集
    const amazonGrid = document.getElementById('amazonFieldsGrid');
    if (amazonGrid) {
      amazonGrid.querySelectorAll('input[type="checkbox"]:checked').forEach(cb => {
        amazonFields.push(cb.dataset.field);
      });
    }

    chrome.storage.sync.set({ rakutenFields, amazonFields }, () => {
      console.log('[設定保存] 収集項目:', { rakutenFields, amazonFields });
    });
  }

  function loadState() {
    chrome.storage.local.get(['collectionState', 'isQueueCollecting', 'collectingItems'], (result) => {
      const state = result.collectionState || {};
      const isQueueCollecting = result.isQueueCollecting || false;
      const collectingItems = result.collectingItems || [];

      const hasData = (state.reviewCount || 0) > 0;
      downloadBtn.disabled = !hasData;
      clearDataBtn.disabled = !hasData;

      // 収集中かどうかでボタンを切り替え
      // isQueueCollectingフラグまたはcollectingItemsがあれば収集中
      const isCollecting = isQueueCollecting || collectingItems.length > 0;
      updateQueueButtons(isCollecting);
    });
  }

  // URLから販路を判定
  function detectSourceFromUrl(url) {
    if (!url) return 'unknown';
    if (url.includes('rakuten.co.jp')) return 'rakuten';
    if (url.includes('amazon.co.jp')) return 'amazon';
    return 'unknown';
  }

  // AmazonのURLからASINを抽出
  function extractAsinFromUrl(url) {
    if (!url) return '';
    // /dp/ASIN, /product-reviews/ASIN, /gp/product/ASIN パターン
    const patterns = [
      /\/dp\/([A-Z0-9]{10})/i,
      /\/product-reviews\/([A-Z0-9]{10})/i,
      /\/gp\/product\/([A-Z0-9]{10})/i
    ];
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match) return match[1].toUpperCase();
    }
    return '';
  }

  /**
   * 入力がASIN形式かどうか判定
   * ASINは10文字の英数字（先頭がB、数字、または大文字英字で始まる）
   */
  function isAsin(input) {
    if (!input) return false;
    const trimmed = input.trim().toUpperCase();
    return /^[A-Z0-9]{10}$/.test(trimmed);
  }

  /**
   * ASINからAmazonレビューページURLを生成
   */
  function asinToUrl(asin) {
    return `https://www.amazon.co.jp/product-reviews/${asin.trim().toUpperCase()}/`;
  }

  // 楽天のURLから商品管理番号を抽出
  function extractRakutenItemCodeFromUrl(url) {
    if (!url) return '';
    // item.rakuten.co.jp/shop-name/item-code/ パターン
    const itemMatch = url.match(/item\.rakuten\.co\.jp\/[^\/]+\/([^\/\?]+)/);
    if (itemMatch) return itemMatch[1];
    // review.rakuten.co.jp/item/shop-id/item-code/ パターン
    const reviewMatch = url.match(/review\.rakuten\.co\.jp\/item\/\d+\/([^\/\?]+)/);
    if (reviewMatch) return reviewMatch[1];
    return '';
  }

  // タイトルから不要なプレフィックスを除去
  function cleanTitle(title, source) {
    if (!title) return '商品';
    let cleaned = title;
    if (source === 'amazon') {
      // 「Amazon.co.jp：」「Amazon.co.jp:」を除去
      cleaned = cleaned.replace(/^Amazon\.co\.jp[：:]\s*/i, '');
      // 「カスタマーレビュー：」「カスタマーレビュー:」を除去
      cleaned = cleaned.replace(/^カスタマーレビュー[：:]\s*/i, '');
    }
    return cleaned.trim() || '商品';
  }

  // キューアイテムのタイトルを生成（Amazon: ASIN：商品名、楽天: 商品管理番号：商品名）
  function getQueueItemTitle(item) {
    const source = item.source || detectSourceFromUrl(item.url);
    let title = cleanTitle(item.title, source);
    if (source === 'amazon') {
      const asin = extractAsinFromUrl(item.url);
      // titleがASINと同じ、または「商品」の場合はASINだけ表示
      // (商品名が取得できなかった場合)
      if (!title || title === asin || title === '商品') {
        return asin || title || '商品';
      }
      // 正常な場合: ASIN：商品名
      if (asin) {
        return `${asin}：${title}`;
      }
      return title;
    } else if (source === 'rakuten') {
      const itemCode = extractRakutenItemCodeFromUrl(item.url);
      // titleが商品管理番号と同じ、または「商品」の場合は商品管理番号だけ表示
      if (!title || title === itemCode || title === '商品') {
        return itemCode || title || '商品';
      }
      // 正常な場合: 商品管理番号：商品名
      if (itemCode) {
        return `${itemCode}：${title}`;
      }
      return title;
    }
    return title || '商品';
  }

  // 販路バッジのHTMLを生成
  function getSourceBadgeHtml(source) {
    if (source === 'amazon') {
      return '<span class="source-badge source-amazon">Amazon</span>';
    } else if (source === 'rakuten') {
      return '<span class="source-badge source-rakuten">楽天</span>';
    }
    return '';
  }

  function loadQueue() {
    chrome.storage.local.get(['queue', 'collectingItems'], (result) => {
      const queue = result.queue || [];
      const collectingItems = result.collectingItems || [];
      const totalCount = queue.length + collectingItems.length;
      queueRemaining.textContent = `${totalCount}`;
      startQueueBtn.disabled = totalCount === 0;

      if (totalCount === 0) {
        queueList.innerHTML = '';
        return;
      }

      // 収集中アイテムを先頭に表示
      const collectingHtml = collectingItems.map(item => {
        const source = item.source || detectSourceFromUrl(item.url);
        const displayTitle = getQueueItemTitle(item);
        return `
        <div class="queue-item collecting">
          <div class="queue-item-info">
            <div class="queue-item-title">
              <span class="collecting-badge">収集中</span>
              ${getSourceBadgeHtml(source)}
              ${escapeHtml(displayTitle)}
            </div>
            <div class="queue-item-url">${escapeHtml(item.url)}</div>
          </div>
        </div>
      `}).join('');

      // 待機中アイテム
      const waitingHtml = queue.map((item, index) => {
        const source = item.source || detectSourceFromUrl(item.url);
        const displayTitle = getQueueItemTitle(item);
        return `
        <div class="queue-item">
          <div class="queue-item-info">
            <div class="queue-item-title">${getSourceBadgeHtml(source)}${escapeHtml(displayTitle)}</div>
            <div class="queue-item-url">${escapeHtml(item.url)}</div>
          </div>
          <button class="queue-item-remove" data-index="${index}">×</button>
        </div>
      `}).join('');

      queueList.innerHTML = collectingHtml + waitingHtml;

      // 削除ボタンのイベント
      queueList.querySelectorAll('.queue-item-remove').forEach(btn => {
        btn.addEventListener('click', (e) => {
          removeFromQueue(parseInt(e.target.dataset.index));
        });
      });
    });
  }

  // ===== インメモリログ管理 =====
  // ストレージの読み書き競合を排除するため、ログはメモリ上の配列で管理
  // DOM更新は即座に行い、ストレージ保存はデバウンスで定期的に実行
  let _memReviewLogs = [];
  let _memProductLogs = [];
  let _logSaveTimer = null;

  // ===== 商品ログのストレージポーリング =====
  // background.jsはストレージにログを書き込む。options.jsはポーリングで拾う。
  // chrome.runtime.sendMessageは信頼できないため、ストレージ経由で同期する。
  let _productLogPollTimer = null;
  let _lastPolledProductLogCount = 0;
  let _productCollectionActive = false;

  function startProductLogPolling() {
    stopProductLogPolling();
    _productCollectionActive = true;
    // 現在のストレージログ数を記録（ベースライン）
    chrome.storage.local.get(['productLogs'], (result) => {
      _lastPolledProductLogCount = (result.productLogs || []).length;
      _productLogPollTimer = setInterval(pollProductLogs, 500);
    });
  }

  function pollProductLogs() {
    chrome.storage.local.get(['productLogs'], (result) => {
      const storageLogs = result.productLogs || [];
      if (storageLogs.length > _lastPolledProductLogCount) {
        // background.jsが書いた新しいエントリーをメモリに追加
        const newEntries = storageLogs.slice(_lastPolledProductLogCount);
        _memProductLogs.push(...newEntries);
        _lastPolledProductLogCount = storageLogs.length;
        renderLogs('product');
      }
    });
  }

  function stopProductLogPolling() {
    _productCollectionActive = false;
    if (_productLogPollTimer) {
      clearInterval(_productLogPollTimer);
      _productLogPollTimer = null;
    }
  }

  // 最終同期: ストレージから残りのログを取得
  function finalSyncProductLogs(callback) {
    chrome.storage.local.get(['productLogs'], (result) => {
      const storageLogs = result.productLogs || [];
      if (storageLogs.length > _lastPolledProductLogCount) {
        const newEntries = storageLogs.slice(_lastPolledProductLogCount);
        _memProductLogs.push(...newEntries);
        _lastPolledProductLogCount = storageLogs.length;
        renderLogs('product');
      }
      if (callback) callback();
    });
  }

  // ページ読み込み時にストレージからメモリに復元
  function initLogsFromStorage() {
    chrome.storage.local.get(['logs', 'productLogs'], (result) => {
      _memReviewLogs = result.logs || [];
      _memProductLogs = result.productLogs || [];
      renderLogs('review');
      renderLogs('product');
    });
  }

  // メモリ上のログをDOMに描画
  function renderLogs(category) {
    if (category === 'review' || !category) {
      if (_memReviewLogs.length === 0) {
        logSection.style.display = 'none';
        logContainer.innerHTML = '';
      } else {
        logSection.style.display = 'block';
        logContainer.innerHTML = _memReviewLogs.map(log => {
          const typeClass = log.type ? ` ${log.type}` : '';
          return `<div class="log-entry${typeClass}"><span class="time">[${log.time}]</span> ${escapeHtml(log.text)}</div>`;
        }).join('');
        logContainer.scrollTop = logContainer.scrollHeight;
      }
    }

    if (category === 'product' || !category) {
      if (_memProductLogs.length === 0) {
        productLogSection.style.display = 'none';
        productLogContainer.innerHTML = '';
      } else {
        productLogSection.style.display = 'block';
        productLogContainer.innerHTML = _memProductLogs.map(log => {
          const typeClass = log.type ? ` ${log.type}` : '';
          return `<div class="log-entry${typeClass}"><span class="time">[${log.time}]</span> ${escapeHtml(log.text)}</div>`;
        }).join('');
        productLogContainer.scrollTop = productLogContainer.scrollHeight;
      }
    }
  }

  // メモリに追加 → 即座にDOM更新 → デバウンスでストレージ保存
  function appendToLog(entry, category) {
    const logs = category === 'product' ? _memProductLogs : _memReviewLogs;
    logs.push(entry);
    renderLogs(category);
    scheduleLogSave();
  }

  // ストレージ保存（デバウンス: 最後の書き込みから300ms後に1回だけ実行）
  // 商品収集中はproductLogsを書かない（background.jsが管轄するため競合防止）
  function scheduleLogSave() {
    if (_logSaveTimer) clearTimeout(_logSaveTimer);
    _logSaveTimer = setTimeout(() => {
      const data = { logs: _memReviewLogs };
      if (!_productCollectionActive) {
        data.productLogs = _memProductLogs;
      }
      chrome.storage.local.set(data);
    }, 300);
  }

  // 後方互換: 既存のloadLogs呼び出し箇所をカバー
  function loadLogs(category) {
    renderLogs(category);
  }

  // スプレッドシートURLの自動保存（Sheets API直接連携）
  async function saveSpreadsheetUrlAuto() {
    const url = spreadsheetUrlInput.value.trim();

    // URLが空の場合はクリア
    if (!url) {
      // 有効な定期キューをチェック（個別スプレッドシートの有無に関わらず）
      const result = await chrome.storage.local.get(['scheduledQueues']);
      const scheduledQueues = result.scheduledQueues || [];
      const enabledQueues = scheduledQueues.filter(q => q.enabled);

      if (enabledQueues.length > 0) {
        const queueNames = enabledQueues.map(q => `・${q.name}`).join('\n');
        const confirmed = confirm(
          `通常収集のスプレッドシートを削除すると、以下の定期収集キューが無効になります。\n\n${queueNames}\n\n続行しますか？`
        );

        if (!confirmed) {
          // キャンセル：元のURLに戻す
          const syncResult = await chrome.storage.sync.get(['spreadsheetUrl']);
          spreadsheetUrlInput.value = syncResult.spreadsheetUrl || '';
          return;
        }

        // 確認OK：全ての有効なキューを無効化
        const updatedQueues = scheduledQueues.map(q => {
          if (q.enabled) {
            return { ...q, enabled: false };
          }
          return q;
        });
        await chrome.storage.local.set({ scheduledQueues: updatedQueues });
        await renderScheduledQueues(); // UIを更新
        updateScheduledAlarm(); // アラームをクリア
      }

      chrome.storage.sync.set({ spreadsheetUrl: '' }, () => {
        showStatus(spreadsheetUrlStatus, 'info', '設定をクリアしました');
        if (spreadsheetLinkRakutenEl) {
          spreadsheetLinkRakutenEl.classList.add('disabled');
        }
        // タイトル表示をクリア
        if (spreadsheetTitleEl) {
          spreadsheetTitleEl.classList.remove('show', 'loading', 'error');
          spreadsheetTitleEl.innerHTML = '';
        }
        updateSpreadsheetBtnState();
      });
      return;
    }

    // URL形式チェック
    const spreadsheetIdMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!spreadsheetIdMatch) {
      showStatus(spreadsheetUrlStatus, 'error', 'スプレッドシートURLの形式が正しくありません');
      return;
    }

    // 保存
    chrome.storage.sync.set({ spreadsheetUrl: url }, () => {
      if (chrome.runtime.lastError) {
        showStatus(spreadsheetUrlStatus, 'error', '保存に失敗しました');
        return;
      }

      showStatus(spreadsheetUrlStatus, 'success', '✓ 保存しました');
      if (spreadsheetLinkRakutenEl) {
        spreadsheetLinkRakutenEl.href = url;
        spreadsheetLinkRakutenEl.classList.remove('disabled');
      }
      // タイトル取得
      fetchAndShowSpreadsheetTitle(url, spreadsheetTitleEl, spreadsheetUrlInput);
      updateSpreadsheetBtnState();
    });
  }

  // Amazon用スプレッドシートURLの自動保存
  async function saveAmazonSpreadsheetUrlAuto() {
    const url = amazonSpreadsheetUrlInput.value.trim();

    // URLが空の場合はクリア
    if (!url) {
      chrome.storage.sync.set({ amazonSpreadsheetUrl: '' }, () => {
        showStatus(amazonSpreadsheetUrlStatus, 'info', '設定をクリアしました');
        if (spreadsheetLinkAmazonEl) {
          spreadsheetLinkAmazonEl.classList.add('disabled');
        }
        // タイトル表示をクリア
        if (amazonSpreadsheetTitleEl) {
          amazonSpreadsheetTitleEl.classList.remove('show', 'loading', 'error');
          amazonSpreadsheetTitleEl.innerHTML = '';
        }
        updateSpreadsheetBtnState();
      });
      return;
    }

    // URL形式チェック
    const spreadsheetIdMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!spreadsheetIdMatch) {
      showStatus(amazonSpreadsheetUrlStatus, 'error', 'スプレッドシートURLの形式が正しくありません');
      return;
    }

    // 保存
    chrome.storage.sync.set({ amazonSpreadsheetUrl: url }, () => {
      if (chrome.runtime.lastError) {
        showStatus(amazonSpreadsheetUrlStatus, 'error', '保存に失敗しました');
        return;
      }

      showStatus(amazonSpreadsheetUrlStatus, 'success', '✓ 保存しました');
      if (spreadsheetLinkAmazonEl) {
        spreadsheetLinkAmazonEl.href = url;
        spreadsheetLinkAmazonEl.classList.remove('disabled');
      }
      // タイトル取得
      fetchAndShowSpreadsheetTitle(url, amazonSpreadsheetTitleEl, amazonSpreadsheetUrlInput);
      updateSpreadsheetBtnState();
    });
  }

  // スプレッドシートURLからIDを抽出
  function extractSpreadsheetId(url) {
    if (!url) return '';
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : '';
  }

  // Google Sheetsアイコン（SVG）
  const SHEETS_ICON_SVG = `<svg class="sheets-icon" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M14 2H6C4.9 2 4 2.9 4 4V20C4 21.1 4.9 22 6 22H18C19.1 22 20 21.1 20 20V8L14 2Z" fill="#23A566"/>
    <path d="M14 2V8H20L14 2Z" fill="#8ED1B1"/>
    <path d="M8 13H16V14H8V13ZM8 15H16V16H8V15ZM8 17H13V18H8V17Z" fill="white"/>
  </svg>`;

  // スプレッドシートのタイトルを取得して表示
  async function fetchAndShowSpreadsheetTitle(url, titleEl, inputEl) {
    if (!titleEl) return;

    // クリックでURL編集モードに切り替え（どの状態でも有効）
    const setupClickHandler = () => {
      titleEl.onclick = () => {
        titleEl.classList.remove('show');
        if (inputEl) {
          inputEl.focus();
          inputEl.select();
        }
      };
    };

    const spreadsheetId = extractSpreadsheetId(url);
    if (!spreadsheetId) {
      titleEl.classList.remove('show', 'loading', 'error');
      titleEl.innerHTML = '';
      return;
    }

    // ローディング表示
    titleEl.classList.add('show', 'loading');
    titleEl.classList.remove('error');
    titleEl.innerHTML = '読み込み中...';
    setupClickHandler(); // ローディング中もクリック可能

    try {
      const response = await chrome.runtime.sendMessage({
        action: 'getSpreadsheetTitle',
        spreadsheetId
      });

      if (response.success && response.title) {
        titleEl.classList.add('show');
        titleEl.classList.remove('loading', 'error');
        titleEl.innerHTML = `${SHEETS_ICON_SVG}<span class="title-text">${response.title}</span><span class="edit-hint">クリックで編集</span>`;
        setupClickHandler();
      } else {
        titleEl.classList.add('show', 'error');
        titleEl.classList.remove('loading');
        titleEl.innerHTML = (response.error || 'タイトル取得失敗') + '<span class="edit-hint">クリックで編集</span>';
        setupClickHandler(); // エラー時もクリック可能
      }
    } catch (error) {
      titleEl.classList.add('show', 'error');
      titleEl.classList.remove('loading');
      titleEl.innerHTML = 'タイトル取得失敗<span class="edit-hint">クリックで編集</span>';
      setupClickHandler(); // エラー時もクリック可能
    }
  }

  // 通知設定のみを保存（チェックボックス変更時）
  function saveNotificationSettings() {
    const enableNotification = enableNotificationCheckbox ? enableNotificationCheckbox.checked : true;
    const notifyPerProduct = notifyPerProductCheckbox ? notifyPerProductCheckbox.checked : false;
    chrome.storage.sync.set({ enableNotification, notifyPerProduct });
  }

  // シート分割・CSV分割設定を保存（チェックボックス変更時）
  function saveSheetSettings() {
    const separateSheets = separateSheetsCheckbox ? separateSheetsCheckbox.checked : true;
    const separateCsvFiles = separateCsvFilesCheckbox ? separateCsvFilesCheckbox.checked : true;
    chrome.storage.sync.set({ separateSheets, separateCsvFiles });
    console.log('[設定保存] separateSheets:', separateSheets, 'separateCsvFiles:', separateCsvFiles);
  }

  // 定期収集表示設定を保存（チェックボックス変更時）
  function saveScheduledCollectionVisibility() {
    const show = showScheduledCollectionCheckbox ? showScheduledCollectionCheckbox.checked : false;
    chrome.storage.sync.set({ showScheduledCollection: show });
    if (scheduledCollectionSection) {
      scheduledCollectionSection.style.display = show ? 'block' : 'none';
    }
    console.log('[設定保存] showScheduledCollection:', show);
  }

  // 期間指定フィルターを保存
  function saveDateFilter() {
    const dateFilterFrom = dateFilterFromInput?.value || '';
    const dateFilterTo = dateFilterToInput?.value || '';

    // 値が入力されている場合のみenableDateFilterをtrueに
    const enableDateFilter = !!(dateFilterFrom || dateFilterTo);

    chrome.storage.sync.set({
      enableDateFilter,
      dateFilterFrom,
      dateFilterTo
    });
    console.log('[設定保存] 期間指定:', { enableDateFilter, dateFilterFrom, dateFilterTo });
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
      review.source === 'amazon' ? 'Amazon' : (review.source === 'rakuten' ? '楽天' : review.source || ''),
      review.country || ''
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
    if (_memReviewLogs.length === 0) return;

    const logText = _memReviewLogs.map(log => `[${log.time}] ${log.text}`).join('\n');
    navigator.clipboard.writeText(logText).then(() => {
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
    let urls = text.split('\n').map(u => u.trim()).filter(u => u.length > 0);

    // ASIN形式の入力をURLに変換
    urls = urls.map(input => {
      if (isAsin(input)) {
        const convertedUrl = asinToUrl(input);
        console.log('[addToQueue] ASINをURLに変換:', input, '→', convertedUrl);
        return convertedUrl;
      }
      return input;
    });

    // ランキングURLの場合（1件のみ対応）- 楽天 + Amazon
    const rankingUrl = urls.find(u =>
      u.includes('ranking.rakuten.co.jp') ||
      (u.includes('amazon.co.jp') && (u.includes('/bestsellers/') || u.includes('/ranking/')))
    );
    if (rankingUrl && urls.length === 1) {
      const count = parseInt(rankingCount.value) || 10;
      addToQueueBtn.disabled = true;

      try {
        chrome.runtime.sendMessage({
          action: 'fetchRanking',
          url: rankingUrl,
          count: count
        }, (response) => {
          addToQueueBtn.disabled = false;
          if (response && response.success) {
            loadQueue();
            showStatus(addStatus, 'success', `${response.addedCount}件追加しました`);
            const source = rankingUrl.includes('amazon.co.jp') ? 'Amazon' : '楽天';
            addLog(`${source}ランキングから${response.addedCount}件をキューに追加`, 'success');
            productUrl.value = '';
            productUrl.style.height = '38px';
            rankingCountWrapper.style.display = 'none';
            // ボタンをデフォルトに戻す
            addToQueueBtn.classList.remove('btn-primary');
            addToQueueBtn.classList.add('btn-secondary');
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

    // 商品URLの場合（複数対応）- 楽天 + Amazon
    const rakutenUrls = urls.filter(u =>
      u.includes('item.rakuten.co.jp') || u.includes('review.rakuten.co.jp')
    );
    const amazonUrls = urls.filter(u =>
      u.includes('amazon.co.jp') && (u.includes('/dp/') || u.includes('/gp/product/') || u.includes('/product-reviews/'))
    );

    const productUrls = [...rakutenUrls, ...amazonUrls];

    if (productUrls.length === 0) {
      showStatus(addStatus, 'error', '楽天/AmazonのURLまたはASIN（10桁）を入力してください');
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

        // URLからタイトルと販路を生成
        let productTitle = '商品';
        let source = 'unknown';

        // 楽天の場合
        const rakutenPathMatch = url.match(/item\.rakuten\.co\.jp\/([^\/]+)\/([^\/\?]+)/);
        if (rakutenPathMatch) {
          productTitle = `${rakutenPathMatch[1]} - ${rakutenPathMatch[2]}`;
          source = 'rakuten';
        }

        // Amazonの場合
        let productId = null;
        const amazonAsinMatch = url.match(/(?:\/dp\/|\/gp\/product\/|\/product-reviews\/)([A-Z0-9]{10})/i);
        if (amazonAsinMatch) {
          productTitle = amazonAsinMatch[1];
          productId = amazonAsinMatch[1]; // ASINをproductIdとして設定
          source = 'amazon';
        }

        // 楽天の商品コードを抽出
        if (source === 'rakuten') {
          const rakutenMatch = url.match(/item\.rakuten\.co\.jp\/[^\/]+\/([^\/\?]+)/);
          if (rakutenMatch) {
            productId = rakutenMatch[1];
          }
        }

        queue.push({
          id: Date.now().toString() + '_' + addedCount, // 一意のID
          url: url,
          title: productTitle.substring(0, 100),
          productId: productId, // 商品識別子（Amazon: ASIN, 楽天: 商品コード）
          source: source,
          addedAt: new Date().toISOString()
        });
        addedCount++;
      });

      if (addedCount === 0 && skippedCount > 0) {
        showStatus(addStatus, 'error', `${skippedCount}件は既に追加済みです`);
        return;
      }

      chrome.storage.local.set({ queue }, () => {
        loadQueue();
        addLog(`${addedCount}件の商品をキューに追加`, 'success');
        productUrl.value = '';
        productUrl.style.height = '38px';
        // ボタンをデフォルトに戻す
        addToQueueBtn.classList.remove('btn-primary');
        addToQueueBtn.classList.add('btn-secondary');
        if (urlCountLabel) {
          urlCountLabel.textContent = '';
          urlCountLabel.className = 'url-count-label';
        }
      });
    });
  }

  function clearLogs() {
    clearLogBtn.style.background = '#dc3545';
    clearLogBtn.style.color = 'white';
    clearLogBtn.title = 'クリアしました!';

    _memReviewLogs = [];
    renderLogs('review');
    chrome.storage.local.set({ logs: [] });

    setTimeout(() => {
      clearLogBtn.style.background = '';
      clearLogBtn.style.color = '';
      clearLogBtn.title = 'クリア';
    }, 1500);
  }

  function clearProductLogs() {
    clearProductLogBtn.style.background = '#dc3545';
    clearProductLogBtn.style.color = 'white';
    clearProductLogBtn.title = 'クリアしました!';

    _memProductLogs = [];
    renderLogs('product');
    chrome.storage.local.set({ productLogs: [] });

    setTimeout(() => {
      clearProductLogBtn.style.background = '';
      clearProductLogBtn.style.color = '';
      clearProductLogBtn.title = 'クリア';
    }, 1500);
  }

  function copyProductLogs() {
    if (_memProductLogs.length === 0) return;

    const logText = _memProductLogs.map(log => `[${log.time}] ${log.text}`).join('\n');
    navigator.clipboard.writeText(logText).then(() => {
      copyProductLogBtn.style.background = '#28a745';
      copyProductLogBtn.style.color = 'white';
      copyProductLogBtn.title = 'コピーしました!';
      setTimeout(() => {
        copyProductLogBtn.style.background = '';
        copyProductLogBtn.style.color = '';
        copyProductLogBtn.title = 'ログをコピー';
      }, 1500);
    }).catch(err => {
      console.error('コピーに失敗:', err);
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
        loadState(); // ボタン状態も更新
        break;
      case 'appendLog':
        // background.jsからログデータを直接受信→即座にメモリ追加＋DOM更新
        if (msg.entry) {
          appendToLog(msg.entry, msg.category);
        }
        break;
      case 'batchProductProgressUpdate':
        updateBatchProductProgress(msg.progress);
        break;
      case 'batchProductQueueUpdated':
        loadBatchProductQueue();
        break;
    }
  }

  function addLog(text, type = '', category = 'review') {
    const time = new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    appendToLog({ time, text, type }, category);
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

  // テキストエリアの高さを内容に応じて自動調整（空なら初期サイズに戻す）
  function autoResizeTextarea(el) {
    if (!el) return;
    if (el.value === '') {
      el.style.height = '38px';
      return;
    }
    el.style.height = '0';
    el.style.height = Math.max(38, Math.min(el.scrollHeight, 120)) + 'px';
  }

  // ========================================
  // キュー保存機能
  // ========================================

  function loadSavedQueues() {
    chrome.storage.local.get(['savedQueues'], (result) => {
      const savedQueues = result.savedQueues || [];
      renderSavedQueuesDropdown(savedQueues);
      renderScheduledQueues(); // 引数なしでストレージから取得
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
          <button class="dropdown-icon-btn add-to-queue-btn" data-id="${queue.id}" title="キューに追加">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
          </button>
          <button class="dropdown-icon-btn edit-queue-btn" data-id="${queue.id}" title="名前を変更">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
          </button>
          <button class="dropdown-icon-btn delete-queue-btn" data-id="${queue.id}" title="削除">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
          </button>
        </div>
      </div>
    `).join('');

    // イベントリスナー
    // キューに追加ボタン
    savedQueuesDropdownList.querySelectorAll('.add-to-queue-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const queueId = e.currentTarget.dataset.id;
        loadSavedQueue(queueId);
        savedQueuesDropdown.style.display = 'none';
      });
    });
    // クリックでもキューに追加
    savedQueuesDropdownList.querySelectorAll('.saved-queue-info').forEach(el => {
      el.addEventListener('click', (e) => {
        loadSavedQueue(e.currentTarget.dataset.id);
        savedQueuesDropdown.style.display = 'none';
      });
    });
    savedQueuesDropdownList.querySelectorAll('.edit-queue-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        // SVG要素がクリックされた場合もボタンのdata-idを取得
        const queueId = e.currentTarget.dataset.id;
        editSavedQueueName(queueId);
      });
    });
    savedQueuesDropdownList.querySelectorAll('.delete-queue-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        // SVG要素がクリックされた場合もボタンのdata-idを取得
        const queueId = e.currentTarget.dataset.id;
        deleteSavedQueue(queueId);
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
  // 商品キュー保存機能
  // ========================================

  function toggleProductQueuesDropdown() {
    if (!productQueuesDropdown) return;
    const isVisible = productQueuesDropdown.style.display !== 'none';
    productQueuesDropdown.style.display = isVisible ? 'none' : 'block';
    if (!isVisible) loadSavedProductQueues();
  }

  function loadSavedProductQueues() {
    chrome.storage.local.get(['savedProductQueues'], (result) => {
      const queues = result.savedProductQueues || [];
      renderProductQueuesDropdown(queues);
    });
  }

  function renderProductQueuesDropdown(queues) {
    if (!productQueuesDropdownList) return;
    if (queues.length === 0) {
      productQueuesDropdownList.innerHTML = '<div class="saved-queues-empty">保存済みキューはありません</div>';
      return;
    }

    productQueuesDropdownList.innerHTML = queues.map(queue => `
      <div class="saved-queue-item" data-id="${queue.id}">
        <div class="saved-queue-info" data-id="${queue.id}">
          <span class="saved-queue-name">${escapeHtml(queue.name)}</span>
          <span class="saved-queue-count">${queue.items.length}件</span>
        </div>
        <div class="saved-queue-actions">
          <button class="dropdown-icon-btn pq-load-btn" data-id="${queue.id}" title="キューに追加">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
          </button>
          <button class="dropdown-icon-btn pq-edit-btn" data-id="${queue.id}" title="名前を変更">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
          </button>
          <button class="dropdown-icon-btn pq-delete-btn" data-id="${queue.id}" title="削除">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
          </button>
        </div>
      </div>
    `).join('');

    // イベントリスナー
    productQueuesDropdownList.querySelectorAll('.pq-load-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        loadProductQueue(e.currentTarget.dataset.id);
        productQueuesDropdown.style.display = 'none';
      });
    });
    productQueuesDropdownList.querySelectorAll('.saved-queue-info').forEach(el => {
      el.addEventListener('click', (e) => {
        loadProductQueue(e.currentTarget.dataset.id);
        productQueuesDropdown.style.display = 'none';
      });
    });
    productQueuesDropdownList.querySelectorAll('.pq-edit-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        editProductQueueName(e.currentTarget.dataset.id);
      });
    });
    productQueuesDropdownList.querySelectorAll('.pq-delete-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        deleteProductQueue(e.currentTarget.dataset.id);
      });
    });
  }

  function saveProductQueue() {
    if (batchProductQueue.length === 0) {
      alert('キューが空です');
      return;
    }
    const name = prompt('保存するキューの名前を入力してください');
    if (!name || name.trim() === '') return;

    chrome.storage.local.get(['savedProductQueues'], (result) => {
      const queues = result.savedProductQueues || [];
      queues.push({
        id: 'pq_' + Date.now(),
        name: name.trim(),
        createdAt: new Date().toISOString(),
        items: batchProductQueue.map(item => {
          if (item.includes('item.rakuten.co.jp')) {
            return { url: item, source: 'rakuten' };
          }
          return { asin: item, source: 'amazon' };
        })
      });
      chrome.storage.local.set({ savedProductQueues: queues }, () => {
        loadSavedProductQueues();
        addLog(`キュー「${name}」を保存（${batchProductQueue.length}件）`, 'success', 'product');
      });
    });
  }

  function loadProductQueue(queueId) {
    chrome.storage.local.get(['savedProductQueues'], (result) => {
      const queues = result.savedProductQueues || [];
      const queue = queues.find(q => q.id === queueId);
      if (!queue) return;
      let addedCount = 0;
      for (const item of queue.items) {
        const value = item.url || item.asin;
        if (value && !batchProductQueue.includes(value)) {
          batchProductQueue.push(value);
          addedCount++;
        }
      }
      renderBatchProductQueue();
      addLog(`「${queue.name}」から${addedCount}件をキューに追加`, 'success', 'product');
    });
  }

  function editProductQueueName(queueId) {
    chrome.storage.local.get(['savedProductQueues'], (result) => {
      const queues = result.savedProductQueues || [];
      const queue = queues.find(q => q.id === queueId);
      if (!queue) return;
      const newName = prompt('新しいキュー名を入力', queue.name);
      if (!newName || newName.trim() === '') return;
      queue.name = newName.trim();
      chrome.storage.local.set({ savedProductQueues: queues }, () => {
        loadSavedProductQueues();
        addLog(`キュー名を「${newName}」に変更`, 'success', 'product');
      });
    });
  }

  function deleteProductQueue(queueId) {
    chrome.storage.local.get(['savedProductQueues'], (result) => {
      const queues = result.savedProductQueues || [];
      const queue = queues.find(q => q.id === queueId);
      if (!queue) return;
      if (!confirm(`「${queue.name}」を削除しますか？`)) return;
      const newQueues = queues.filter(q => q.id !== queueId);
      chrome.storage.local.set({ savedProductQueues: newQueues }, () => {
        loadSavedProductQueues();
        addLog(`キュー「${queue.name}」を削除`, 'success', 'product');
      });
    });
  }

  // ========================================
  // ビュー切り替え機能
  // ========================================

  function hideAllViews() {
    if (mainView) mainView.classList.remove('active');
    if (settingsView) settingsView.classList.remove('active');
  }

  function showMainView() {
    hideAllViews();
    if (mainView) mainView.classList.add('active');
    currentView = 'main';
  }

  function showSettingsView() {
    hideAllViews();
    if (settingsView) settingsView.classList.add('active');
    currentView = 'settings';
  }

  // 定期収集ボタンの状態を更新（グレーアウトなし）
  function updateScheduledButtonsState() {
    // グレーアウト処理を削除済み
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

      // 新しい定期収集キューを作成（デフォルトは無効）
      const newScheduledQueue = {
        id: 'sched_' + Date.now(),
        sourceQueueId: savedQueueId,
        name: sourceQueue.name,
        items: sourceQueue.items.slice(), // コピー
        time: '07:00',
        incrementalOnly: true,
        enabled: false,
        lastRun: null
      };

      scheduledQueues.push(newScheduledQueue);

      chrome.storage.local.set({ scheduledQueues }, () => {
        loadScheduledSettings();
        addLog(`「${sourceQueue.name}」を定期収集に追加`, 'success');
        updateScheduledAlarm();
      });
    });
  }

  // 定期収集画面のキュー一覧をレンダリング
  async function renderScheduledQueues(scheduledQueues) {
    if (!scheduledQueuesList) return;

    // 引数がない場合はストレージから取得
    if (!scheduledQueues) {
      const result = await chrome.storage.local.get(['scheduledQueues']);
      scheduledQueues = result.scheduledQueues || [];
    }

    // 親カードとヘッダーを取得
    const parentCard = scheduledQueuesList.closest('.card');
    const queueHeader = parentCard?.querySelector('.queue-header');

    if (scheduledQueues.length === 0) {
      scheduledQueuesList.innerHTML = '';
      scheduledQueuesList.style.display = 'none';
      if (parentCard) {
        // 上下対称にして中央揃え
        parentCard.style.paddingTop = '16px';
        parentCard.style.paddingBottom = '16px';
      }
      if (queueHeader) {
        queueHeader.style.margin = '0';
        queueHeader.style.padding = '0';
        queueHeader.style.borderBottom = 'none';
      }
      return;
    }

    // キューがある場合は表示
    scheduledQueuesList.style.display = 'block';
    if (parentCard) {
      parentCard.style.paddingTop = '';
      parentCard.style.paddingBottom = '';
    }
    if (queueHeader) {
      queueHeader.style.margin = '';
      queueHeader.style.padding = '';
      queueHeader.style.borderBottom = '';
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
              <label class="toggle-switch-small">
                <input type="checkbox" class="scheduled-queue-toggle" data-queue-id="${queue.id}" ${queue.enabled ? 'checked' : ''}>
                <span class="toggle-slider-small"></span>
              </label>
              <span class="scheduled-queue-name">${escapeHtml(queue.name)}</span>
              <span class="scheduled-queue-count">${queue.items.length}件</span>
            </div>
            <div class="scheduled-queue-actions">
              <button class="scheduled-queue-run-btn" data-queue-id="${queue.id}">すぐ実行</button>
              <button class="scheduled-queue-delete-btn" data-queue-id="${queue.id}" title="削除">×</button>
            </div>
          </div>
          <div class="scheduled-queue-settings">
            <div class="scheduled-queue-row">
              <span class="scheduled-queue-label">収集時刻:</span>
              <div class="time-picker">
                <select class="time-select scheduled-queue-hour" data-queue-id="${queue.id}">
                  ${generateHourOptions(hours)}
                </select>
                <span class="time-separator">:</span>
                <select class="time-select scheduled-queue-minute" data-queue-id="${queue.id}">
                  ${generateMinuteOptions(minutes)}
                </select>
              </div>
              <label class="checkbox-label-compact">
                <input type="checkbox" class="scheduled-queue-incremental" data-queue-id="${queue.id}" ${queue.incrementalOnly ? 'checked' : ''}>
                <span>差分のみ収集</span>
              </label>
              <span class="scheduled-queue-last-run">前回: ${lastRunText}</span>
            </div>
            <div class="form-group" style="margin-bottom: 0;">
              <label>保存先スプレッドシート</label>
              <div class="spreadsheet-input-wrapper">
                <input type="text" data-queue-id="${queue.id}"
                       value="${escapeHtml(queue.spreadsheetUrl || '')}" placeholder="スプレッドシートのURLを入力（必須）">
                <div class="spreadsheet-title-overlay scheduled-queue-title" data-queue-id="${queue.id}"></div>
              </div>
            </div>
          </div>
        </div>
      `;
    }).join('');

    // スプレッドシートタイトルを取得・表示
    scheduledQueues.forEach(queue => {
      if (queue.spreadsheetUrl) {
        const titleEl = scheduledQueuesList.querySelector(`.scheduled-queue-title[data-queue-id="${queue.id}"]`);
        const inputEl = scheduledQueuesList.querySelector(`.form-group input[data-queue-id="${queue.id}"]`);
        if (titleEl && inputEl) {
          fetchAndShowSpreadsheetTitle(queue.spreadsheetUrl, titleEl, inputEl);
        }
      }
    });

    // イベントリスナー
    scheduledQueuesList.querySelectorAll('.scheduled-queue-toggle').forEach(toggle => {
      toggle.addEventListener('click', async (e) => {
        const queueId = e.target.dataset.queueId;
        const willBeEnabled = e.target.checked; // クリック後の状態（clickイベント時点で既に変更済み）

        // オンにする場合、キュー個別のスプレッドシートが設定されているかチェック
        if (willBeEnabled) {
          // 先にチェックを外しておく（バリデーション中は無効状態）
          e.target.checked = false;

          const result = await chrome.storage.local.get(['scheduledQueues']);
          const scheduledQueues = result.scheduledQueues || [];
          const queue = scheduledQueues.find(q => q.id === queueId);
          const queueSpreadsheetUrl = queue?.spreadsheetUrl || '';

          // 定期収集は個別スプレッドシート必須（通常収集用は使用不可）
          if (!queueSpreadsheetUrl) {
            alert('スプレッドシートが設定されていません。\n\n定期収集を有効にするには、このキューの「保存先スプレッドシート」欄にURLを入力してください。\n\n※定期収集では通常収集用のスプレッドシートは使用できません。');
            return;
          }

          // 検証OK、手動でオンにする
          e.target.checked = true;
          updateScheduledQueueProperty(queueId, 'enabled', true);
        } else {
          // オフにする場合はそのまま
          updateScheduledQueueProperty(queueId, 'enabled', false);
        }
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

    scheduledQueuesList.querySelectorAll('.form-group input[data-queue-id]').forEach(input => {
      let saveTimeout = null;

      // 入力時: debounceで保存 → 保存後にタイトル取得（設定欄と同じ仕様）
      input.addEventListener('input', (e) => {
        const queueId = e.target.dataset.queueId;
        const url = e.target.value.trim();

        // タイトル表示をクリア（入力中は非表示）
        const titleEl = scheduledQueuesList.querySelector(`.scheduled-queue-title[data-queue-id="${queueId}"]`);
        if (titleEl) {
          titleEl.classList.remove('show', 'loading', 'error');
          titleEl.innerHTML = '';
        }

        if (saveTimeout) clearTimeout(saveTimeout);
        saveTimeout = setTimeout(async () => {
          // 保存
          await updateScheduledQueueProperty(queueId, 'spreadsheetUrl', url, e.target);

          // 保存後にタイトル取得（URLが有効な場合）
          if (url && titleEl) {
            const spreadsheetIdMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
            if (spreadsheetIdMatch) {
              fetchAndShowSpreadsheetTitle(url, titleEl, e.target);
            }
          }
        }, 500);
      });
    });

    updateScheduledButtonsState();
  }

  // 定期収集キューのプロパティを更新
  async function updateScheduledQueueProperty(queueId, property, value, inputElement = null) {
    const result = await chrome.storage.local.get(['scheduledQueues']);
    const scheduledQueues = result.scheduledQueues || [];
    const queue = scheduledQueues.find(q => q.id === queueId);

    if (!queue) return;

    // 個別スプレッドシートURLが削除される場合のチェック
    if (property === 'spreadsheetUrl' && !value && queue.enabled) {
      const syncResult = await chrome.storage.sync.get(['spreadsheetUrl']);
      const globalSpreadsheetUrl = syncResult.spreadsheetUrl || '';

      // グローバルもない場合は確認
      if (!globalSpreadsheetUrl) {
        const confirmed = confirm(
          `「${queue.name}」の保存先スプレッドシートがなくなります。\n\nこのキューの定期収集を無効にしますか？`
        );

        if (confirmed) {
          // 確認OK：キューを無効化
          queue.spreadsheetUrl = '';
          queue.enabled = false;
          await chrome.storage.local.set({ scheduledQueues });
          await renderScheduledQueues(); // UIを更新
          updateScheduledAlarm();
          return;
        } else {
          // キャンセル：元の値に戻す
          if (inputElement) {
            inputElement.value = queue.spreadsheetUrl || '';
          }
          return;
        }
      }
    }

    queue[property] = value;
    await chrome.storage.local.set({ scheduledQueues });

    if (property === 'enabled') {
      const card = scheduledQueuesList.querySelector(`.scheduled-queue-card[data-id="${queueId}"]`);
      if (card) card.classList.toggle('enabled', value);
    }
    if (inputElement) {
      showAutoSaveIndicator(inputElement);
    }
    updateScheduledAlarm();
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
        addLog(`「${queue.name}」を定期収集から削除`, 'success');
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
        addLog('キューが見つからないか、空です', 'error');
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
              spreadsheetUrl: targetQueue.spreadsheetUrl || null,
              queueName: targetQueue.name
            });
            addedCount++;
          }
        });

        if (addedCount === 0) {
          addLog(`「${targetQueue.name}」は全て収集済みまたはキューに追加済みです`, 'error');
          return;
        }

        chrome.storage.local.set({ queue: currentQueue }, () => {
          loadQueue();
          addLog(`「${targetQueue.name}」の収集を開始（${addedCount}件）`, 'success');
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

  // クイックスタートガイドの初期化
  function initQuickStartGuide() {
    const quickStartKey = 'rakuten-review-quickstart-shown';
    const overlay = document.getElementById('quickStartOverlay');
    const closeBtn = document.getElementById('quickStartCloseBtn');

    // 楽天用
    const rakutenUrlInput = document.getElementById('quickStartSpreadsheetUrl');
    const rakutenTitleEl = document.getElementById('quickStartSpreadsheetTitle');

    // Amazon用
    const amazonUrlInput = document.getElementById('quickStartAmazonSpreadsheetUrl');
    const amazonTitleEl = document.getElementById('quickStartAmazonSpreadsheetTitle');

    if (!overlay || !closeBtn) return;

    // 初回表示チェック
    if (!localStorage.getItem(quickStartKey)) {
      overlay.style.display = 'flex';
    }

    // 楽天用スプレッドシートURL自動保存
    if (rakutenUrlInput) {
      let saveTimeout = null;
      rakutenUrlInput.addEventListener('input', () => {
        const url = rakutenUrlInput.value.trim();

        // タイトル表示をクリア
        if (rakutenTitleEl) {
          rakutenTitleEl.classList.remove('show', 'loading', 'error');
          rakutenTitleEl.innerHTML = '';
        }

        if (!url) return;

        // URL形式チェック
        const spreadsheetIdMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        if (!spreadsheetIdMatch) return;

        if (saveTimeout) clearTimeout(saveTimeout);
        saveTimeout = setTimeout(() => {
          // 保存
          chrome.storage.sync.set({ spreadsheetUrl: url }, () => {
            if (chrome.runtime.lastError) return;

            // タイトル取得・表示
            if (rakutenTitleEl) {
              fetchAndShowSpreadsheetTitle(url, rakutenTitleEl, rakutenUrlInput);
            }

            // メインの設定画面の入力欄も更新
            const mainInput = document.getElementById('spreadsheetUrl');
            if (mainInput) mainInput.value = url;

            // ヘッダーのスプレッドシートリンクも更新
            const link = document.getElementById('spreadsheetLinkRakuten');
            if (link) {
              link.href = url;
              link.classList.remove('disabled');
            }

            // ヘッダーのスプレッドシートボタン状態を更新
            updateSpreadsheetBtnState();

            // メインのタイトル表示も更新
            if (spreadsheetTitleEl) {
              fetchAndShowSpreadsheetTitle(url, spreadsheetTitleEl, spreadsheetUrlInput);
            }
          });
        }, 500);
      });
    }

    // Amazon用スプレッドシートURL自動保存
    if (amazonUrlInput) {
      let saveTimeout = null;
      amazonUrlInput.addEventListener('input', () => {
        const url = amazonUrlInput.value.trim();

        // タイトル表示をクリア
        if (amazonTitleEl) {
          amazonTitleEl.classList.remove('show', 'loading', 'error');
          amazonTitleEl.innerHTML = '';
        }

        if (!url) return;

        // URL形式チェック
        const spreadsheetIdMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        if (!spreadsheetIdMatch) return;

        if (saveTimeout) clearTimeout(saveTimeout);
        saveTimeout = setTimeout(() => {
          // 保存
          chrome.storage.sync.set({ amazonSpreadsheetUrl: url }, () => {
            if (chrome.runtime.lastError) return;

            // タイトル取得・表示
            if (amazonTitleEl) {
              fetchAndShowSpreadsheetTitle(url, amazonTitleEl, amazonUrlInput);
            }

            // メインの設定画面の入力欄も更新
            const mainInput = document.getElementById('amazonSpreadsheetUrl');
            if (mainInput) mainInput.value = url;

            // ヘッダーのスプレッドシートリンクも更新
            const link = document.getElementById('spreadsheetLinkAmazon');
            if (link) {
              link.href = url;
              link.classList.remove('disabled');
            }

            // ヘッダーのスプレッドシートボタン状態を更新
            updateSpreadsheetBtnState();

            // メインのタイトル表示も更新
            if (amazonSpreadsheetTitleEl) {
              fetchAndShowSpreadsheetTitle(url, amazonSpreadsheetTitleEl, amazonSpreadsheetUrlInput);
            }
          });
        }, 500);
      });
    }

    // 閉じるボタン
    closeBtn.addEventListener('click', () => {
      localStorage.setItem(quickStartKey, 'done');
      overlay.style.display = 'none';
    });

    // オーバーレイクリックでも閉じる
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) {
        localStorage.setItem(quickStartKey, 'done');
        overlay.style.display = 'none';
      }
    });
  }

  // ===== 統計情報の表示 =====
  function updateStats() {
    chrome.runtime.sendMessage({ action: 'getAmazonRateLimitInfo' }, (response) => {
      if (chrome.runtime.lastError || !response) return;
      const countEl = document.getElementById('amazonPageCount');
      if (countEl) {
        countEl.textContent = response.count || 0;
      }
    });
  }
  updateStats();

  // ===== 商品情報収集（Google Drive）関連 =====

  // フォルダURL保存
  function saveProductInfoFolderUrl() {
    const url = productInfoFolderUrlInput ? productInfoFolderUrlInput.value.trim() : '';
    if (url && !url.includes('drive.google.com/drive/folders/')) {
      if (productInfoFolderUrlStatus) {
        showStatus(productInfoFolderUrlStatus, 'error', 'Google DriveのフォルダURLを入力してください');
      }
      return;
    }
    chrome.storage.sync.set({ productInfoFolderUrl: url }, () => {
      if (productInfoFolderUrlStatus) {
        if (url) {
          showStatus(productInfoFolderUrlStatus, 'success', '保存しました');
        } else {
          productInfoFolderUrlStatus.className = 'status-message';
          productInfoFolderUrlStatus.textContent = '';
        }
      }
      updateDriveFolderLink(url);
    });
  }

  // 入力テキストからASINまたは楽天URLを抽出するヘルパー
  function extractProductItems(text) {
    return text.split('\n')
      .map(line => line.trim())
      .filter(line => line.length > 0)
      .map(line => {
        // 楽天商品URL（item.rakuten.co.jp/shop/item/）
        if (line.includes('item.rakuten.co.jp')) {
          // URLとして正規化（プロトコルがなければ追加）
          const url = line.startsWith('http') ? line : `https://${line}`;
          // クエリパラメータを除去
          return url.split('?')[0];
        }
        // Amazon URLからASINを抽出
        const dpMatch = line.match(/\/dp\/([A-Z0-9]{10})/i);
        if (dpMatch) return dpMatch[1].toUpperCase();
        const gpMatch = line.match(/\/gp\/product\/([A-Z0-9]{10})/i);
        if (gpMatch) return gpMatch[1].toUpperCase();
        // 直接ASIN入力
        if (/^[A-Z0-9]{10}$/i.test(line)) return line.toUpperCase();
        return null;
      })
      .filter(item => item !== null);
  }

  // 「追加」ボタン — キューに追加
  function addToBatchProductQueue() {
    const text = batchProductAsins ? batchProductAsins.value.trim() : '';
    if (!text) {
      if (batchProductStatus) showStatus(batchProductStatus, 'error', '商品URLまたはASINを入力してください');
      return;
    }

    const items = extractProductItems(text);
    if (items.length === 0) {
      if (batchProductStatus) showStatus(batchProductStatus, 'error', '有効な商品URLまたはASINが見つかりません');
      return;
    }

    // 重複を除いてキューに追加
    let addedCount = 0;
    for (const item of items) {
      if (!batchProductQueue.includes(item)) {
        batchProductQueue.push(item);
        addedCount++;
      }
    }

    if (addedCount > 0) {
      chrome.storage.local.set({ batchProductQueue: [...batchProductQueue] });
      if (batchProductStatus) showStatus(batchProductStatus, 'success', `${addedCount}件追加しました`);
      addLog(`キューに${addedCount}件追加しました（合計: ${batchProductQueue.length}件）`, 'info', 'product');
    } else {
      if (batchProductStatus) showStatus(batchProductStatus, 'info', 'すべて追加済みです');
    }

    // 入力欄をクリアしてサイズをリセット
    if (batchProductAsins) {
      batchProductAsins.value = '';
      autoResizeTextarea(batchProductAsins);
    }
    renderBatchProductQueue();
  }

  // ストレージから商品キューを読み込み（ポップアップからの追加を反映）
  function loadBatchProductQueue() {
    chrome.storage.local.get(['batchProductQueue'], (result) => {
      const stored = result.batchProductQueue || [];
      // ストレージのASINをメモリに反映（重複排除）
      let addedCount = 0;
      for (const asin of stored) {
        if (!batchProductQueue.includes(asin)) {
          batchProductQueue.push(asin);
          addedCount++;
        }
      }
      if (addedCount > 0) {
        renderBatchProductQueue();
      }
    });
  }

  // キューの表示を更新
  function renderBatchProductQueue() {
    if (batchProductCountEl) batchProductCountEl.textContent = batchProductQueue.length;
    if (startBatchProductRunBtn) startBatchProductRunBtn.disabled = batchProductQueue.length === 0;

    if (!batchProductList) return;
    if (batchProductQueue.length === 0) {
      batchProductList.innerHTML = '';
      return;
    }

    batchProductList.innerHTML = batchProductQueue.map((item, index) => {
      const isRakuten = item.includes('item.rakuten.co.jp');
      const badge = isRakuten
        ? '<span class="source-badge source-rakuten">楽天</span>'
        : '<span class="source-badge source-amazon">Amazon</span>';
      const displayTitle = isRakuten
        ? item.replace(/^https?:\/\/item\.rakuten\.co\.jp\//, '')
        : item;
      const displayUrl = isRakuten
        ? item
        : `https://www.amazon.co.jp/dp/${escapeHtml(item)}`;
      return `
      <div class="queue-item">
        <div class="queue-item-info">
          <div class="queue-item-title">${badge}${escapeHtml(displayTitle)}</div>
          <div class="queue-item-url">${escapeHtml(displayUrl)}</div>
        </div>
        <button class="queue-item-remove" data-index="${index}">×</button>
      </div>`;
    }).join('');

    // 削除ボタンのイベント
    batchProductList.querySelectorAll('.queue-item-remove').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.index, 10);
        batchProductQueue.splice(idx, 1);
        chrome.storage.local.set({ batchProductQueue: [...batchProductQueue] });
        renderBatchProductQueue();
      });
    });
  }

  // 一括収集開始
  function startBatchProductCollection() {
    if (batchProductQueue.length === 0) {
      if (batchProductStatus) showStatus(batchProductStatus, 'error', '収集する商品がありません');
      return;
    }

    const items = [...batchProductQueue];

    // ローカルで即時ログ表示（ストレージ経由ではなく直接メモリに追加）
    addLog(`${items.length}件の商品情報収集を開始します...`, 'info', 'product');

    // ストレージポーリング開始（background.jsのログをリアルタイムで拾う）
    startProductLogPolling();

    // UIを更新（ボタン切り替え）
    if (startBatchProductRunBtn) startBatchProductRunBtn.style.display = 'none';
    if (cancelBatchProductBtn) cancelBatchProductBtn.style.display = 'block';
    if (batchProductStatus) batchProductStatus.textContent = '';

    // バックグラウンドに送信（ログはbackground.jsがストレージに書き込む）
    chrome.runtime.sendMessage({
      action: 'startBatchProductCollection',
      items
    }, (response) => {
      // エラー時（設定未完了、認証失敗など）はUI・ポーリングを戻す
      if (response && !response.success) {
        stopProductLogPolling();
        addLog(`エラー: ${response.error}`, 'error', 'product');
        if (startBatchProductRunBtn) {
          startBatchProductRunBtn.style.display = 'block';
          startBatchProductRunBtn.disabled = false;
        }
        if (cancelBatchProductBtn) cancelBatchProductBtn.style.display = 'none';
      }
    });
  }

  // 一括収集中止（即時UI反映）
  function cancelBatchProductCollection() {
    chrome.runtime.sendMessage({ action: 'cancelBatchProductCollection' });

    // ポーリング停止
    stopProductLogPolling();

    // 即座にボタンを戻す
    if (startBatchProductRunBtn) {
      startBatchProductRunBtn.style.display = 'block';
      startBatchProductRunBtn.disabled = batchProductQueue.length === 0;
    }
    if (cancelBatchProductBtn) cancelBatchProductBtn.style.display = 'none';

    // 最終同期してから中止メッセージを追加
    finalSyncProductLogs(() => {
      addLog('収集を中止しました', 'warning', 'product');
    });
  }

  // 一括収集の進捗更新（ログはbackground.jsがストレージに書き込み、ポーリングで表示済み）
  function updateBatchProductProgress(progress) {
    if (!progress) return;

    const { isRunning, completed, failed } = progress;

    // 完了時のUI更新
    if (!isRunning) {
      // ポーリング停止
      stopProductLogPolling();

      if (startBatchProductRunBtn) {
        startBatchProductRunBtn.style.display = 'block';
        startBatchProductRunBtn.disabled = batchProductQueue.length === 0;
      }
      if (cancelBatchProductBtn) cancelBatchProductBtn.style.display = 'none';

      // 成功した商品をキューから削除
      if (completed) {
        completed.forEach(item => {
          // id（ASIN or itemSlug）またはasinでキューから検索
          const identifier = item.id || item.asin;
          let idx = batchProductQueue.findIndex(q => {
            if (typeof q === 'string') {
              // 楽天URLの場合はitemSlugを含むか確認
              if (q.includes('item.rakuten.co.jp') && item.source === 'rakuten') {
                return q.includes(identifier);
              }
              return q === identifier;
            }
            return false;
          });
          if (idx !== -1) batchProductQueue.splice(idx, 1);
        });
      }
      chrome.storage.local.set({ batchProductQueue: [...batchProductQueue] });
      renderBatchProductQueue();

      // 最終同期（background.jsの完了ログを拾う。バッファフラッシュ待ちのため少し遅延）
      setTimeout(() => {
        finalSyncProductLogs();
      }, 500);
    }
  }

  // ===== フォルダピッカー =====

  function initFolderPicker() {
    const pickerBtn = document.getElementById('folderPickerBtn');
    if (!pickerBtn) return;

    const overlay = document.getElementById('folderPickerOverlay');
    const searchInput = document.getElementById('fpSearch');
    const breadcrumbs = document.getElementById('fpBreadcrumbs');
    const listEl = document.getElementById('fpList');
    const selectBtn = document.getElementById('fpSelectBtn');
    const cancelBtn = document.getElementById('fpCancelBtn');
    const newFolderBtn = document.getElementById('fpNewFolderBtn');
    const newFolderRow = document.getElementById('fpNewFolderRow');
    const newFolderInput = document.getElementById('fpNewFolderInput');
    const newFolderCreate = document.getElementById('fpNewFolderCreate');
    const newFolderCancel = document.getElementById('fpNewFolderCancel');
    const tabMyDrive = document.getElementById('fpTabMyDrive');
    const tabShared = document.getElementById('fpTabShared');

    // 状態
    let currentParentId = 'root';
    let selectedFolder = null;  // nullの場合は現在のフォルダが対象
    let pathStack = [{ id: 'root', name: 'マイドライブ' }];
    let isSearchMode = false;
    let searchTimeout = null;
    let currentDriveId = null;  // nullならマイドライブ
    let currentTab = 'myDrive'; // 'myDrive' | 'shared'

    // キャッシュ（5分）
    const cache = new Map();
    const CACHE_TTL = 5 * 60 * 1000;

    function getCached(key) {
      const entry = cache.get(key);
      if (entry && Date.now() - entry.time < CACHE_TTL) return entry.data;
      cache.delete(key);
      return null;
    }

    function setCache(key, data) {
      cache.set(key, { data, time: Date.now() });
    }

    // 現在のフォルダIDを取得（選択されたフォルダ、またはブラウズ中のフォルダ）
    function getSelectedFolderId() {
      return selectedFolder ? selectedFolder.id : currentParentId;
    }

    // モーダルを開く
    pickerBtn.addEventListener('click', () => {
      overlay.classList.add('active');
      currentParentId = 'root';
      selectedFolder = null;
      pathStack = [{ id: 'root', name: 'マイドライブ' }];
      isSearchMode = false;
      searchInput.value = '';
      currentDriveId = null;
      currentTab = 'myDrive';
      tabMyDrive.classList.add('active');
      tabShared.classList.remove('active');
      selectBtn.disabled = false;  // 現在のフォルダ（マイドライブ）が常に選択可能
      newFolderRow.classList.remove('active');
      loadFolders(currentParentId);
      renderBreadcrumbs();
    });

    // タブ切り替え: マイドライブ
    tabMyDrive.addEventListener('click', () => {
      if (currentTab === 'myDrive') return;
      currentTab = 'myDrive';
      tabMyDrive.classList.add('active');
      tabShared.classList.remove('active');
      currentDriveId = null;
      currentParentId = 'root';
      selectedFolder = null;
      pathStack = [{ id: 'root', name: 'マイドライブ' }];
      isSearchMode = false;
      searchInput.value = '';
      selectBtn.disabled = false;
      renderBreadcrumbs();
      loadFolders(currentParentId);
    });

    // タブ切り替え: 共有ドライブ
    tabShared.addEventListener('click', () => {
      if (currentTab === 'shared') return;
      currentTab = 'shared';
      tabShared.classList.add('active');
      tabMyDrive.classList.remove('active');
      currentDriveId = null;
      currentParentId = null;
      selectedFolder = null;
      pathStack = [{ id: null, name: '共有ドライブ' }];
      isSearchMode = false;
      searchInput.value = '';
      selectBtn.disabled = true;  // ドライブ一覧画面では選択不可
      renderBreadcrumbs();
      loadSharedDrives();
    });

    // モーダルを閉じる
    cancelBtn.addEventListener('click', closeModal);
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) closeModal();
    });

    function closeModal() {
      overlay.classList.remove('active');
    }

    // 共有ドライブ一覧を読み込み
    function loadSharedDrives() {
      listEl.innerHTML = '<div class="fp-loading">読み込み中...</div>';

      const cached = getCached('sharedDrives');
      if (cached) {
        renderSharedDrives(cached);
        return;
      }

      chrome.runtime.sendMessage({ action: 'getSharedDrives' }, (response) => {
        if (chrome.runtime.lastError) {
          listEl.innerHTML = `<div class="fp-empty">エラー: ${escapeHtml(chrome.runtime.lastError.message)}</div>`;
          return;
        }
        if (!response || !response.success) {
          listEl.innerHTML = `<div class="fp-empty">エラー: ${escapeHtml(response?.error || '取得に失敗しました')}</div>`;
          return;
        }
        if (!response.drives || response.drives.length === 0) {
          listEl.innerHTML = '<div class="fp-empty">共有ドライブがありません</div>';
          return;
        }
        setCache('sharedDrives', response.drives);
        renderSharedDrives(response.drives);
      });
    }

    // 共有ドライブ一覧を描画
    function renderSharedDrives(drives) {
      listEl.innerHTML = '';
      drives.forEach(drive => {
        const item = document.createElement('div');
        item.className = 'fp-item';
        item.innerHTML = `<span class="fp-item-icon">🗂️</span><span class="fp-item-name">${escapeHtml(drive.name)}</span>`;

        // クリックで共有ドライブに入る
        item.addEventListener('click', () => {
          enterSharedDrive(drive);
        });

        listEl.appendChild(item);
      });
    }

    // 共有ドライブに入る
    function enterSharedDrive(drive) {
      currentDriveId = drive.id;
      currentParentId = drive.id;
      selectedFolder = null;
      pathStack = [
        { id: null, name: '共有ドライブ' },
        { id: drive.id, name: drive.name }
      ];
      selectBtn.disabled = false;  // ドライブのルートも選択可能
      isSearchMode = false;
      searchInput.value = '';
      renderBreadcrumbs();
      loadFolders(drive.id);
    }

    // フォルダ一覧を読み込み
    function loadFolders(parentId) {
      listEl.innerHTML = '<div class="fp-loading">読み込み中...</div>';
      selectedFolder = null;
      // 現在のフォルダは常に選択可能（サブフォルダ未選択でもOK）
      selectBtn.disabled = false;

      const cacheKey = `list:${parentId}:${currentDriveId || ''}`;
      const cached = getCached(cacheKey);
      if (cached) {
        renderFolders(cached);
        return;
      }

      const msg = { action: 'getDriveFolders', parentId };
      if (currentDriveId) msg.driveId = currentDriveId;

      chrome.runtime.sendMessage(msg, (response) => {
        if (chrome.runtime.lastError) {
          listEl.innerHTML = `<div class="fp-empty">エラー: ${escapeHtml(chrome.runtime.lastError.message)}</div>`;
          return;
        }
        if (!response || !response.success) {
          listEl.innerHTML = `<div class="fp-empty">エラー: ${escapeHtml(response?.error || '取得に失敗しました')}</div>`;
          return;
        }
        setCache(cacheKey, response.folders);
        renderFolders(response.folders);
      });
    }

    // フォルダ検索
    function searchFolders(query) {
      listEl.innerHTML = '<div class="fp-loading">検索中...</div>';
      selectedFolder = null;
      selectBtn.disabled = true;

      const msg = { action: 'searchDriveFolders', query };
      if (currentDriveId) msg.driveId = currentDriveId;

      chrome.runtime.sendMessage(msg, (response) => {
        if (chrome.runtime.lastError) {
          listEl.innerHTML = `<div class="fp-empty">エラー: ${escapeHtml(chrome.runtime.lastError.message)}</div>`;
          return;
        }
        if (!response || !response.success) {
          listEl.innerHTML = `<div class="fp-empty">エラー: ${escapeHtml(response?.error || '検索に失敗しました')}</div>`;
          return;
        }
        renderFolders(response.folders, true);
      });
    }

    // フォルダ一覧を描画
    function renderFolders(folders, isSearch = false) {
      if (!folders || folders.length === 0) {
        listEl.innerHTML = `<div class="fp-empty">${isSearch ? '該当するフォルダが見つかりません' : 'サブフォルダがありません'}</div>`;
        return;
      }

      listEl.innerHTML = '';
      folders.forEach(folder => {
        const item = document.createElement('div');
        item.className = 'fp-item';
        item.innerHTML = `<span class="fp-item-icon">📁</span><span class="fp-item-name">${escapeHtml(folder.name)}</span>`;

        // クリック = 選択
        item.addEventListener('click', () => {
          listEl.querySelectorAll('.fp-item.selected').forEach(el => el.classList.remove('selected'));
          item.classList.add('selected');
          selectedFolder = folder;
          selectBtn.disabled = false;
        });

        // ダブルクリック = フォルダの中に入る
        item.addEventListener('dblclick', () => {
          enterFolder(folder);
        });

        listEl.appendChild(item);
      });
    }

    // フォルダに入る
    function enterFolder(folder) {
      currentParentId = folder.id;
      selectedFolder = null;
      isSearchMode = false;
      searchInput.value = '';
      selectBtn.disabled = false;

      pathStack.push({ id: folder.id, name: folder.name });
      renderBreadcrumbs();
      loadFolders(folder.id);
    }

    // パンくずリストを描画
    function renderBreadcrumbs() {
      breadcrumbs.innerHTML = '';
      pathStack.forEach((item, index) => {
        if (index > 0) {
          const sep = document.createElement('span');
          sep.className = 'fp-separator';
          sep.textContent = ' > ';
          breadcrumbs.appendChild(sep);
        }

        const crumb = document.createElement('span');
        crumb.textContent = item.name;

        if (index === pathStack.length - 1) {
          crumb.className = 'fp-crumb current';
        } else {
          crumb.className = 'fp-crumb';
          crumb.addEventListener('click', () => {
            pathStack = pathStack.slice(0, index + 1);
            currentParentId = item.id;
            selectedFolder = null;
            isSearchMode = false;
            searchInput.value = '';

            // 共有ドライブ一覧に戻る場合
            if (currentTab === 'shared' && index === 0) {
              currentDriveId = null;
              selectBtn.disabled = true;
              renderBreadcrumbs();
              loadSharedDrives();
              return;
            }

            selectBtn.disabled = false;
            renderBreadcrumbs();
            loadFolders(currentParentId);
          });
        }

        breadcrumbs.appendChild(crumb);
      });
    }

    // 検索入力
    searchInput.addEventListener('input', () => {
      const query = searchInput.value.trim();
      if (searchTimeout) clearTimeout(searchTimeout);

      if (query.length === 0) {
        isSearchMode = false;
        renderBreadcrumbs();
        // 共有ドライブ一覧に戻る場合
        if (currentTab === 'shared' && !currentDriveId) {
          loadSharedDrives();
        } else {
          loadFolders(currentParentId);
        }
        return;
      }

      isSearchMode = true;
      breadcrumbs.innerHTML = '<span class="fp-crumb current">検索結果</span>';

      searchTimeout = setTimeout(() => {
        searchFolders(query);
      }, 400);
    });

    // 「選択」ボタン — 選択されたフォルダ、または現在のフォルダを使用
    selectBtn.addEventListener('click', () => {
      const folderId = getSelectedFolderId();
      if (!folderId) return;
      const url = `https://drive.google.com/drive/folders/${folderId}`;
      if (productInfoFolderUrlInput) {
        productInfoFolderUrlInput.value = url;
        saveProductInfoFolderUrl();
      }
      closeModal();
    });

    // 「新規フォルダ」ボタン
    newFolderBtn.addEventListener('click', () => {
      newFolderRow.classList.add('active');
      newFolderInput.value = '';
      newFolderInput.focus();
    });

    newFolderCancel.addEventListener('click', () => {
      newFolderRow.classList.remove('active');
    });

    newFolderCreate.addEventListener('click', createNewFolder);
    newFolderInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') createNewFolder();
      if (e.key === 'Escape') newFolderRow.classList.remove('active');
    });

    function createNewFolder() {
      const name = newFolderInput.value.trim();
      if (!name) return;

      newFolderCreate.disabled = true;
      newFolderCreate.textContent = '作成中...';

      chrome.runtime.sendMessage({
        action: 'createDriveFolder',
        name,
        parentId: currentParentId
      }, (response) => {
        newFolderCreate.disabled = false;
        newFolderCreate.textContent = '作成';

        if (chrome.runtime.lastError) {
          alert('エラー: ' + chrome.runtime.lastError.message);
          return;
        }
        if (!response || !response.success) {
          alert('エラー: ' + (response?.error || '作成に失敗しました'));
          return;
        }

        // キャッシュをクリアして再読み込み
        const cacheKey = `list:${currentParentId}:${currentDriveId || ''}`;
        cache.delete(cacheKey);
        newFolderRow.classList.remove('active');
        loadFolders(currentParentId);
      });
    }

    // ESCキーでモーダルを閉じる
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && overlay.classList.contains('active')) {
        closeModal();
      }
    });
  }

});
