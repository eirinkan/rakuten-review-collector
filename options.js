/**
 * 設定画面のスクリプト
 * キュー管理、ランキング追加、設定、ログ表示
 */

document.addEventListener('DOMContentLoaded', () => {
  // DOM要素
  const totalReviews = document.getElementById('totalReviews');
  const currentPage = document.getElementById('currentPage');
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

  const rankingUrl = document.getElementById('rankingUrl');
  const rankingCount = document.getElementById('rankingCount');
  const addRankingBtn = document.getElementById('addRankingBtn');
  const rankingStatus = document.getElementById('rankingStatus');

  const logContainer = document.getElementById('logContainer');
  const clearLogBtn = document.getElementById('clearLogBtn');

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
    addRankingBtn.addEventListener('click', addFromRanking);
    clearLogBtn.addEventListener('click', clearLogs);

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
      totalReviews.textContent = state.reviewCount || 0;
      currentPage.textContent = `${state.pageCount || 0}/${state.totalPages || 0}`;

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

    if (gasUrl && !isValidGasUrl(gasUrl)) {
      showStatus(settingsStatus, 'error', 'URLの形式が正しくありません');
      return;
    }

    chrome.storage.sync.set({ gasUrl, separateSheets }, async () => {
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

  async function addFromRanking() {
    const url = rankingUrl.value.trim();
    const count = parseInt(rankingCount.value) || 10;

    if (!url) {
      showStatus(rankingStatus, 'error', 'URLを入力してください');
      return;
    }

    if (!url.includes('ranking.rakuten.co.jp')) {
      showStatus(rankingStatus, 'error', '楽天ランキングのURLを入力してください');
      return;
    }

    showStatus(rankingStatus, 'info', `ランキングを取得中...`);
    addRankingBtn.disabled = true;

    try {
      // バックグラウンドでランキングを取得
      chrome.runtime.sendMessage({
        action: 'fetchRanking',
        url: url,
        count: count
      }, (response) => {
        addRankingBtn.disabled = false;
        if (response && response.success) {
          showStatus(rankingStatus, 'success', `${response.addedCount}件追加しました`);
          loadQueue();
          addLog(`ランキングから${response.addedCount}件をキューに追加`, 'success');
        } else {
          showStatus(rankingStatus, 'error', response?.error || '取得に失敗しました');
        }
      });
    } catch (e) {
      addRankingBtn.disabled = false;
      showStatus(rankingStatus, 'error', '取得に失敗しました');
    }
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
        addLog('収集完了', 'success');
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
