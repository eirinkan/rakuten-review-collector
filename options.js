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

  const spreadsheetUrlInput = document.getElementById('spreadsheetUrl');
  const spreadsheetUrlStatus = document.getElementById('spreadsheetUrlStatus');
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
  const headerTitle = document.getElementById('headerTitle');
  const settingsToggleBtn = document.getElementById('settingsToggleBtn');
  const helpToggleBtn = document.getElementById('helpToggleBtn');

  // キュー保存関連（ヘッダーアイコン方式）
  const saveQueueBtn = document.getElementById('saveQueueBtn');
  const loadSavedQueuesBtn = document.getElementById('loadSavedQueuesBtn');
  const savedQueuesDropdown = document.getElementById('savedQueuesDropdown');
  const savedQueuesDropdownList = document.getElementById('savedQueuesDropdownList');

  // ビュー切り替え
  const mainView = document.getElementById('main-view');
  const settingsView = document.getElementById('settings-view');
  const helpView = document.getElementById('help-view');

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

  // 初期化
  init();

  function init() {
    loadSettings();
    loadState();
    loadQueue();
    loadLogs();
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


      // 追加ボタンの色を変更
      if (addToQueueBtn) {
        if (validUrls.length > 0) {
          addToQueueBtn.classList.remove('btn-secondary');
          addToQueueBtn.classList.add('btn-primary');
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

    // バックグラウンドからのメッセージ
    chrome.runtime.onMessage.addListener(handleMessage);

    // 定期更新
    setInterval(() => {
      loadState();
      loadQueue();
    }, 2000);
  }

  function loadSettings() {
    chrome.storage.sync.get(['separateSheets', 'separateCsvFiles', 'spreadsheetUrl', 'enableNotification', 'notifyPerProduct'], (result) => {
      // スプレッドシートURL（Sheets API直接連携）
      if (result.spreadsheetUrl && spreadsheetUrlInput) {
        spreadsheetUrlInput.value = result.spreadsheetUrl;
        spreadsheetLink.href = result.spreadsheetUrl;
        spreadsheetLink.style.display = 'inline-flex';
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
      queueRemaining.textContent = `${totalCount}`;
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
        spreadsheetLink.style.display = 'none';
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
      spreadsheetLink.href = url;
      spreadsheetLink.style.display = 'inline-flex';
    });
  }

  // スプレッドシートURLからIDを抽出
  function extractSpreadsheetId(url) {
    if (!url) return '';
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : '';
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
            loadQueue();
            addLog(`ランキングから${response.addedCount}件をキューに追加`, 'success');
            productUrl.value = '';
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
        loadQueue();
        addLog(`${addedCount}件の商品をキューに追加`, 'success');
        productUrl.value = '';
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
    if (settingsView) settingsView.classList.remove('active');
    if (helpView) helpView.classList.remove('active');
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

  function showHelpView() {
    hideAllViews();
    if (helpView) helpView.classList.add('active');
    currentView = 'help';
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
        parentCard.style.paddingBottom = '12px';
      }
      if (queueHeader) {
        queueHeader.style.marginBottom = '0';
        queueHeader.style.paddingBottom = '0';
        queueHeader.style.borderBottom = 'none';
      }
      return;
    }

    // キューがある場合は表示
    scheduledQueuesList.style.display = 'block';
    if (parentCard) {
      parentCard.style.paddingBottom = '';
    }
    if (queueHeader) {
      queueHeader.style.marginBottom = '';
      queueHeader.style.paddingBottom = '';
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
            <div class="scheduled-queue-row">
              <span class="scheduled-queue-label">保存先スプレッドシート:</span>
              <input type="text" class="scheduled-queue-url-input" data-queue-id="${queue.id}"
                     value="${escapeHtml(queue.spreadsheetUrl || '')}" placeholder="未入力で通常収集と同じスプレッドシートを使用">
            </div>
          </div>
        </div>
      `;
    }).join('');

    // イベントリスナー
    scheduledQueuesList.querySelectorAll('.scheduled-queue-toggle').forEach(toggle => {
      toggle.addEventListener('click', async (e) => {
        const queueId = e.target.dataset.queueId;
        const willBeEnabled = e.target.checked; // クリック後の状態（clickイベント時点で既に変更済み）

        // オンにする場合、スプレッドシートが設定されているかチェック
        if (willBeEnabled) {
          // 先にチェックを外しておく（バリデーション中は無効状態）
          e.target.checked = false;

          const result = await chrome.storage.local.get(['scheduledQueues']);
          const scheduledQueues = result.scheduledQueues || [];
          const queue = scheduledQueues.find(q => q.id === queueId);
          const queueSpreadsheetUrl = queue?.spreadsheetUrl || '';

          const syncResult = await chrome.storage.sync.get(['spreadsheetUrl']);
          const globalSpreadsheetUrl = syncResult.spreadsheetUrl || '';

          // どちらも設定されていない場合は警告
          if (!queueSpreadsheetUrl && !globalSpreadsheetUrl) {
            alert('スプレッドシートが設定されていません。\n\n定期収集を有効にするには、このキューの「スプレッドシート」欄にURLを入力するか、設定画面で通常収集用のスプレッドシートを設定してください。');
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

    scheduledQueuesList.querySelectorAll('.scheduled-queue-url-input').forEach(input => {
      let saveTimeout = null;
      input.addEventListener('input', (e) => {
        if (saveTimeout) clearTimeout(saveTimeout);
        saveTimeout = setTimeout(() => {
          updateScheduledQueueProperty(e.target.dataset.queueId, 'spreadsheetUrl', e.target.value.trim(), e.target);
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

});
