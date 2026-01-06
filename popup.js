/**
 * ポップアップUIのスクリプト
 * 収集の開始/停止、進捗表示、データダウンロードを制御
 */

document.addEventListener('DOMContentLoaded', () => {
  // DOM要素の取得
  const pageWarning = document.getElementById('pageWarning');
  const mainContent = document.getElementById('mainContent');
  const modeIndicator = document.getElementById('modeIndicator');
  const spreadsheetLink = document.getElementById('spreadsheetLink');
  const spreadsheetSection = document.getElementById('spreadsheetSection');
  const spreadsheetLinkBottom = document.getElementById('spreadsheetLinkBottom');
  const progressBar = document.getElementById('progressBar');
  const progressText = document.getElementById('progressText');
  const reviewCount = document.getElementById('reviewCount');
  const pageCount = document.getElementById('pageCount');
  const startBtn = document.getElementById('startBtn');
  const stopBtn = document.getElementById('stopBtn');
  const downloadBtn = document.getElementById('downloadBtn');
  const clearBtn = document.getElementById('clearBtn');
  const errorMessage = document.getElementById('errorMessage');
  const successMessage = document.getElementById('successMessage');
  const logSection = document.getElementById('logSection');
  const logContainer = document.getElementById('logContainer');

  // 初期化
  init();

  /**
   * 初期化処理
   */
  async function init() {
    // スプレッドシートリンクは常に確認・表示（どのページでも）
    checkSpreadsheetLink();

    // 現在のタブを確認
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const isReviewPage = tab.url && tab.url.includes('review.rakuten.co.jp');
    const isItemPage = tab.url && tab.url.includes('item.rakuten.co.jp');
    const isRakutenPage = isReviewPage || isItemPage;

    if (!isRakutenPage) {
      pageWarning.style.display = 'block';
      mainContent.style.display = 'none';
      return;
    }

    // 保存モードを確認
    checkSaveMode();

    // 収集状態を復元
    restoreState();

    // イベントリスナーの設定
    setupEventListeners();

    // バックグラウンドからのメッセージを受信
    chrome.runtime.onMessage.addListener(handleMessage);
  }

  /**
   * スプレッドシートリンクを確認して表示（常に実行）
   */
  function checkSpreadsheetLink() {
    chrome.storage.sync.get(['gasUrl', 'spreadsheetUrl'], (result) => {
      if (result.gasUrl && result.spreadsheetUrl) {
        spreadsheetSection.style.display = 'block';
        spreadsheetLinkBottom.href = result.spreadsheetUrl;
      } else {
        spreadsheetSection.style.display = 'none';
      }
    });
  }

  /**
   * 保存モードを確認して表示を更新
   */
  function checkSaveMode() {
    chrome.storage.sync.get(['gasUrl', 'spreadsheetUrl'], (result) => {
      if (result.gasUrl) {
        modeIndicator.className = 'mode-indicator spreadsheet';
        modeIndicator.innerHTML = '<span class="icon">📊</span><span>スプレッドシート自動保存</span>';

        // スプレッドシートリンクを表示
        if (result.spreadsheetUrl) {
          spreadsheetLink.href = result.spreadsheetUrl;
          spreadsheetLink.style.display = 'block';
        }
      } else {
        modeIndicator.className = 'mode-indicator csv';
        modeIndicator.innerHTML = '<span class="icon">📄</span><span>CSVダウンロード</span>';
        spreadsheetLink.style.display = 'none';
      }
    });
  }

  /**
   * 状態を復元
   */
  function restoreState() {
    chrome.storage.local.get(['collectionState'], (result) => {
      const state = result.collectionState || {
        isRunning: false,
        reviewCount: 0,
        pageCount: 0,
        reviews: [],
        logs: []
      };

      updateUI(state);
    });
  }

  /**
   * イベントリスナーの設定
   */
  function setupEventListeners() {
    startBtn.addEventListener('click', startCollection);
    stopBtn.addEventListener('click', stopCollection);
    downloadBtn.addEventListener('click', downloadCSV);
    clearBtn.addEventListener('click', clearData);
  }

  /**
   * 収集開始
   */
  async function startCollection() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    // コンテンツスクリプトに収集開始を指示
    chrome.tabs.sendMessage(tab.id, { action: 'startCollection' }, (response) => {
      if (chrome.runtime.lastError) {
        showError('ページとの通信に失敗しました。ページをリロードしてください。');
        return;
      }

      if (response && response.success) {
        startBtn.style.display = 'none';
        stopBtn.style.display = 'block';
        progressText.textContent = '収集中...';
        hideMessages();
        addLog('収集を開始しました');
      }
    });
  }

  /**
   * 収集停止
   */
  async function stopCollection() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    chrome.tabs.sendMessage(tab.id, { action: 'stopCollection' }, (response) => {
      if (response && response.success) {
        startBtn.style.display = 'block';
        stopBtn.style.display = 'none';
        progressText.textContent = '停止しました';
        addLog('収集を停止しました');
      }
    });
  }

  /**
   * CSVダウンロード
   */
  function downloadCSV() {
    chrome.runtime.sendMessage({ action: 'downloadCSV' }, (response) => {
      if (response && response.success) {
        showSuccess('CSVファイルをダウンロードしました');
      } else {
        showError(response?.error || 'ダウンロードに失敗しました');
      }
    });
  }

  /**
   * データクリア
   */
  function clearData() {
    if (!confirm('収集したデータをすべて削除しますか？')) {
      return;
    }

    chrome.storage.local.set({
      collectionState: {
        isRunning: false,
        reviewCount: 0,
        pageCount: 0,
        reviews: [],
        logs: []
      }
    }, () => {
      reviewCount.textContent = '0';
      pageCount.textContent = '0';
      progressBar.style.width = '0%';
      progressText.textContent = '待機中';
      downloadBtn.disabled = true;
      clearBtn.disabled = true;
      logContainer.innerHTML = '';
      showSuccess('データをクリアしました');
    });
  }

  /**
   * バックグラウンドからのメッセージを処理
   */
  function handleMessage(message, sender, sendResponse) {
    switch (message.action) {
      case 'updateProgress':
        updateUI(message.state);
        break;
      case 'collectionComplete':
        startBtn.style.display = 'block';
        stopBtn.style.display = 'none';
        progressText.textContent = '収集完了';
        showSuccess(`${message.state.reviewCount}件のレビューを収集しました`);
        updateUI(message.state);
        break;
      case 'collectionError':
        startBtn.style.display = 'block';
        stopBtn.style.display = 'none';
        showError(message.error);
        break;
      case 'log':
        addLog(message.text, message.type);
        break;
    }
  }

  /**
   * UIを更新
   */
  function updateUI(state) {
    reviewCount.textContent = state.reviewCount || 0;
    pageCount.textContent = state.pageCount || 0;

    // ボタンの有効/無効
    const hasData = (state.reviewCount || 0) > 0;
    downloadBtn.disabled = !hasData;
    clearBtn.disabled = !hasData;

    // 収集中かどうか
    if (state.isRunning) {
      startBtn.style.display = 'none';
      stopBtn.style.display = 'block';
      progressText.textContent = '収集中...';
    } else {
      startBtn.style.display = 'block';
      stopBtn.style.display = 'none';
    }

    // ログセクションの表示
    if (state.logs && state.logs.length > 0) {
      logSection.style.display = 'block';
      logContainer.innerHTML = state.logs.map(log =>
        `<div class="log-entry ${log.type || ''}">${log.text}</div>`
      ).join('');
      logContainer.scrollTop = logContainer.scrollHeight;
    }
  }

  /**
   * ログを追加
   */
  function addLog(text, type = '') {
    logSection.style.display = 'block';
    const entry = document.createElement('div');
    entry.className = `log-entry ${type}`;
    entry.textContent = `${new Date().toLocaleTimeString()} - ${text}`;
    logContainer.appendChild(entry);
    logContainer.scrollTop = logContainer.scrollHeight;

    // ストレージにも保存
    chrome.storage.local.get(['collectionState'], (result) => {
      const state = result.collectionState || { logs: [] };
      state.logs = state.logs || [];
      state.logs.push({ text: `${new Date().toLocaleTimeString()} - ${text}`, type });
      // 最新50件のみ保持
      if (state.logs.length > 50) {
        state.logs = state.logs.slice(-50);
      }
      chrome.storage.local.set({ collectionState: state });
    });
  }

  /**
   * エラーメッセージを表示
   */
  function showError(text) {
    errorMessage.textContent = text;
    errorMessage.style.display = 'block';
    successMessage.style.display = 'none';
  }

  /**
   * 成功メッセージを表示
   */
  function showSuccess(text) {
    successMessage.textContent = text;
    successMessage.style.display = 'block';
    errorMessage.style.display = 'none';

    setTimeout(() => {
      successMessage.style.display = 'none';
    }, 3000);
  }

  /**
   * メッセージを非表示
   */
  function hideMessages() {
    errorMessage.style.display = 'none';
    successMessage.style.display = 'none';
  }
});
