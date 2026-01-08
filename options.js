/**
 * 設定画面のスクリプト
 * キュー管理、ランキング追加、設定、ログ表示
 */

document.addEventListener('DOMContentLoaded', () => {
  // DOM要素
  const queueRemaining = document.getElementById('queueRemaining');
  const spreadsheetLink = document.getElementById('spreadsheetLink');
  const downloadBtn = document.getElementById('downloadBtn');
  const clearDataBtn = document.getElementById('clearDataBtn');
  const dataButtons = document.getElementById('dataButtons');

  const gasUrlInput = document.getElementById('gasUrl');
  const separateSheetsCheckbox = document.getElementById('separateSheets');
  const saveSettingsBtn = document.getElementById('saveSettingsBtn');
  const settingsStatus = document.getElementById('settingsStatus');

  const queueList = document.getElementById('queueList');
  const startQueueBtn = document.getElementById('startQueueBtn');
  const clearQueueBtn = document.getElementById('clearQueueBtn');

  const productUrl = document.getElementById('productUrl');
  const rankingCount = document.getElementById('rankingCount');
  const rankingCountWrapper = document.getElementById('rankingCountWrapper');
  const addToQueueBtn = document.getElementById('addToQueueBtn');
  const addStatus = document.getElementById('addStatus');

  const logContainer = document.getElementById('logContainer');
  const clearLogBtn = document.getElementById('clearLogBtn');

  // ヘッダーボタン
  const settingsToggleBtn = document.getElementById('settingsToggleBtn');
  const helpToggleBtn = document.getElementById('helpToggleBtn');
  const settingsCard = document.getElementById('settingsCard');
  const helpCard = document.getElementById('helpCard');
  const gasHelpToggle = document.getElementById('gasHelpToggle');
  const gasHelp = document.getElementById('gasHelp');
  const gasHelpIcon = document.getElementById('gasHelpIcon');

  // 初期化
  init();

  function init() {
    loadSettings();
    loadState();
    loadQueue();
    loadLogs();

    // イベントリスナー
    saveSettingsBtn.addEventListener('click', saveSettings);
    downloadBtn.addEventListener('click', downloadCSV);
    clearDataBtn.addEventListener('click', clearData);
    startQueueBtn.addEventListener('click', startQueueCollection);
    clearQueueBtn.addEventListener('click', clearQueue);
    addToQueueBtn.addEventListener('click', addToQueue);
    clearLogBtn.addEventListener('click', clearLogs);

    // ヘッダーボタンのイベント
    settingsToggleBtn.addEventListener('click', () => {
      settingsCard.classList.toggle('show');
    });
    helpToggleBtn.addEventListener('click', () => {
      helpCard.classList.toggle('show');
    });
    gasHelpToggle.addEventListener('click', () => {
      gasHelp.classList.toggle('show');
      gasHelpIcon.textContent = gasHelp.classList.contains('show') ? '▼' : '▶';
    });

    // URL入力時にランキングかどうか判定して件数入力の表示を切り替え
    productUrl.addEventListener('input', () => {
      const url = productUrl.value.trim();
      if (url.includes('ranking.rakuten.co.jp')) {
        rankingCountWrapper.style.display = 'flex';
      } else {
        rankingCountWrapper.style.display = 'none';
      }
    });

    // バックグラウンドからのメッセージ
    chrome.runtime.onMessage.addListener(handleMessage);

    // 定期更新
    setInterval(() => {
      loadState();
      loadQueue();
    }, 2000);
  }

  function loadSettings() {
    chrome.storage.sync.get(['gasUrl', 'separateSheets', 'spreadsheetUrl'], (result) => {
      if (result.gasUrl) {
        gasUrlInput.value = result.gasUrl;
        // スプレッドシートモードの場合、CSV/クリアボタンを非表示
        dataButtons.style.display = 'none';
      } else {
        dataButtons.style.display = 'flex';
      }
      separateSheetsCheckbox.checked = result.separateSheets !== false;

      if (result.spreadsheetUrl) {
        spreadsheetLink.href = result.spreadsheetUrl;
        spreadsheetLink.style.display = 'inline-flex';
      }
    });
  }

  function loadState() {
    chrome.storage.local.get(['collectionState'], (result) => {
      const state = result.collectionState || {};

      const hasData = (state.reviewCount || 0) > 0;
      downloadBtn.disabled = !hasData;
      clearDataBtn.disabled = !hasData;
    });
  }

  function loadQueue() {
    chrome.storage.local.get(['queue'], (result) => {
      const queue = result.queue || [];
      queueRemaining.textContent = `${queue.length}件`;
      startQueueBtn.disabled = queue.length === 0;

      if (queue.length === 0) {
        queueList.innerHTML = '<div class="queue-empty">キューは空です</div>';
        return;
      }

      queueList.innerHTML = queue.map((item, index) => `
        <div class="queue-item">
          <div class="queue-item-info">
            <div class="queue-item-title">${escapeHtml(item.title || '商品')}</div>
            <div class="queue-item-url">${escapeHtml(item.url)}</div>
          </div>
          <button class="queue-item-remove" data-index="${index}">×</button>
        </div>
      `).join('');

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
        logContainer.innerHTML = '<div class="log-entry"><span class="time">[--:--:--]</span> 待機中...</div>';
        return;
      }

      logContainer.innerHTML = logs.map(log => {
        const typeClass = log.type ? ` ${log.type}` : '';
        return `<div class="log-entry${typeClass}"><span class="time">[${log.time}]</span> ${escapeHtml(log.text)}</div>`;
      }).join('');

      logContainer.scrollTop = logContainer.scrollHeight;
    });
  }

  async function saveSettings() {
    const gasUrl = gasUrlInput.value.trim();
    const separateSheets = separateSheetsCheckbox.checked;
    const maxConcurrent = parseInt(maxConcurrentSelect.value) || 1;

    if (gasUrl && !isValidGasUrl(gasUrl)) {
      showStatus(settingsStatus, 'error', 'URLの形式が正しくありません');
      return;
    }

    chrome.storage.sync.set({ gasUrl, separateSheets, maxConcurrent }, async () => {
      if (chrome.runtime.lastError) {
        showStatus(settingsStatus, 'error', '保存に失敗しました');
        return;
      }

      // スプレッドシートモードかどうかでボタン表示を切り替え
      if (gasUrl) {
        dataButtons.style.display = 'none';
        // 接続テスト
        showStatus(settingsStatus, 'info', '接続テスト中...');
        try {
          const response = await fetch(gasUrl, { method: 'GET', mode: 'cors' });
          const data = await response.json();

          if (data.success) {
            showStatus(settingsStatus, 'success', '保存・接続成功');
            if (data.spreadsheetUrl) {
              chrome.storage.sync.set({ spreadsheetUrl: data.spreadsheetUrl });
              spreadsheetLink.href = data.spreadsheetUrl;
              spreadsheetLink.style.display = 'inline-flex';
            }
          } else {
            showStatus(settingsStatus, 'error', '接続失敗');
          }
        } catch (e) {
          showStatus(settingsStatus, 'success', '保存しました');
        }
      } else {
        dataButtons.style.display = 'flex';
        spreadsheetLink.style.display = 'none';
        showStatus(settingsStatus, 'success', '保存しました（CSVモード）');
      }
    });
  }

  function isValidGasUrl(url) {
    return url.startsWith('https://script.google.com/macros/s/') && url.includes('/exec');
  }

  function downloadCSV() {
    chrome.runtime.sendMessage({ action: 'downloadCSV' }, (response) => {
      if (response && response.success) {
        addLog('CSVダウンロード完了', 'success');
      } else {
        addLog('CSVダウンロード失敗: ' + (response?.error || ''), 'error');
      }
    });
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
    if (!confirm('キューをクリアしますか？')) return;

    chrome.storage.local.set({ queue: [] }, () => {
      loadQueue();
      addLog('キューをクリアしました');
    });
  }

  function startQueueCollection() {
    chrome.runtime.sendMessage({ action: 'startQueueCollection' }, (response) => {
      if (response && response.success) {
        addLog('キュー一括収集を開始しました', 'success');
      } else {
        addLog('開始に失敗: ' + (response?.error || ''), 'error');
      }
    });
  }

  async function addToQueue() {
    const url = productUrl.value.trim();

    if (!url) {
      showStatus(addStatus, 'error', 'URLを入力してください');
      return;
    }

    // ランキングURLの場合
    if (url.includes('ranking.rakuten.co.jp')) {
      const count = parseInt(rankingCount.value) || 10;
      showStatus(addStatus, 'info', 'ランキングを取得中...');
      addToQueueBtn.disabled = true;

      try {
        chrome.runtime.sendMessage({
          action: 'fetchRanking',
          url: url,
          count: count
        }, (response) => {
          addToQueueBtn.disabled = false;
          if (response && response.success) {
            showStatus(addStatus, 'success', `${response.addedCount}件追加しました`);
            loadQueue();
            addLog(`ランキングから${response.addedCount}件をキューに追加`, 'success');
            productUrl.value = '';
            rankingCountWrapper.style.display = 'none';
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

    // 商品URLの場合
    if (!url.includes('item.rakuten.co.jp') && !url.includes('review.rakuten.co.jp')) {
      showStatus(addStatus, 'error', '楽天の商品ページまたはランキングURLを入力してください');
      return;
    }

    chrome.storage.local.get(['queue'], (result) => {
      const queue = result.queue || [];

      // 重複チェック
      const exists = queue.some(item => item.url === url);
      if (exists) {
        showStatus(addStatus, 'error', '既にキューに追加済みです');
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

      chrome.storage.local.set({ queue }, () => {
        showStatus(addStatus, 'success', 'キューに追加しました');
        loadQueue();
        addLog(`${productTitle} をキューに追加`, 'success');
        productUrl.value = '';
      });
    });
  }

  function clearLogs() {
    chrome.storage.local.set({ logs: [] }, () => {
      loadLogs();
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

      // 最新100件のみ保持
      if (logs.length > 100) {
        logs.splice(0, logs.length - 100);
      }

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
});
