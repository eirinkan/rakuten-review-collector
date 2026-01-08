/**
 * ポップアップUIのスクリプト（シンプル版）
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
      // システム設定に従う
      const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
      this.setTheme(prefersDark ? 'dark' : 'light');
    }

    // システム設定変更を監視
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', (e) => {
      if (!localStorage.getItem(this.storageKey)) {
        this.setTheme(e.matches ? 'dark' : 'light');
      }
    });

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

  const pageWarning = document.getElementById('pageWarning');
  const message = document.getElementById('message');
  const rankingMessage = document.getElementById('rankingMessage');
  const normalMode = document.getElementById('normalMode');
  const rankingMode = document.getElementById('rankingMode');
  const startBtn = document.getElementById('startBtn');
  const stopBtn = document.getElementById('stopBtn');
  const queueBtn = document.getElementById('queueBtn');
  const settingsBtn = document.getElementById('settingsBtn');
  const startRankingBtn = document.getElementById('startRankingBtn');
  const addRankingBtn = document.getElementById('addRankingBtn');
  const rankingCountInput = document.getElementById('rankingCount');

  init();

  async function init() {
    // 現在のタブを確認
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    const isRakutenPage = tab.url && (
      tab.url.includes('review.rakuten.co.jp') ||
      tab.url.includes('item.rakuten.co.jp')
    );
    const isRankingPage = tab.url && tab.url.includes('ranking.rakuten.co.jp');

    if (isRankingPage) {
      // ランキングページの場合
      normalMode.style.display = 'none';
      rankingMode.style.display = 'block';
    } else if (!isRakutenPage) {
      // 楽天以外のページ
      pageWarning.style.display = 'block';
      startBtn.disabled = true;
      queueBtn.disabled = true;
    }

    // 状態を復元
    restoreState();

    // イベントリスナー
    startBtn.addEventListener('click', startCollection);
    stopBtn.addEventListener('click', stopCollection);
    queueBtn.addEventListener('click', addToQueue);
    settingsBtn.addEventListener('click', openSettings);
    startRankingBtn.addEventListener('click', startRankingCollection);
    addRankingBtn.addEventListener('click', addRankingToQueue);

    // バックグラウンドからのメッセージ
    chrome.runtime.onMessage.addListener(handleMessage);
  }

  async function restoreState() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    chrome.storage.local.get(['collectingItems'], (result) => {
      const collectingItems = result.collectingItems || [];
      // 現在のタブが収集中リストにあるかチェック
      const isCurrentTabCollecting = collectingItems.some(item => item.tabId === tab.id);
      updateUI({ isRunning: isCurrentTabCollecting });
    });
  }

  function updateUI(state) {
    if (state.isRunning) {
      startBtn.style.display = 'none';
      stopBtn.style.display = 'block';
    } else {
      startBtn.style.display = 'block';
      stopBtn.style.display = 'none';
    }
  }

  async function startCollection() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    // 既に収集中かチェック
    const result = await chrome.storage.local.get(['collectingItems']);
    const collectingItems = result.collectingItems || [];
    const isAlreadyCollecting = collectingItems.some(item => item.tabId === tab.id);

    if (isAlreadyCollecting) {
      showMessage('このページは既に収集中です', 'error');
      return;
    }

    // 商品情報を取得
    chrome.tabs.sendMessage(tab.id, { action: 'getProductInfo' }, (response) => {
      if (chrome.runtime.lastError) {
        showMessage('ページをリロードしてください', 'error');
        return;
      }

      const productInfo = response?.success ? response.productInfo : {
        url: tab.url,
        title: tab.title || '商品',
        addedAt: new Date().toISOString()
      };

      // background.jsに単一収集開始を依頼（キュー管理も任せる）
      chrome.runtime.sendMessage({
        action: 'startSingleCollection',
        productInfo: productInfo,
        tabId: tab.id
      }, (res) => {
        if (res && res.success) {
          startBtn.style.display = 'none';
          stopBtn.style.display = 'block';
          showMessage('収集を開始しました', 'success');
        } else {
          showMessage(res?.error || '収集開始に失敗しました', 'error');
        }
      });
    });
  }

  async function stopCollection() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    chrome.tabs.sendMessage(tab.id, { action: 'stopCollection' }, (response) => {
      if (response && response.success) {
        startBtn.style.display = 'block';
        stopBtn.style.display = 'none';
      }
    });
  }

  async function addToQueue() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    if (!tab.url.includes('item.rakuten.co.jp') && !tab.url.includes('review.rakuten.co.jp')) {
      showMessage('楽天の商品ページを開いてください', 'error');
      return;
    }

    // 商品情報を取得してキューに追加
    chrome.tabs.sendMessage(tab.id, { action: 'getProductInfo' }, (response) => {
      if (chrome.runtime.lastError || !response) {
        // content scriptが応答しない場合はURLから情報を取得
        const productInfo = {
          url: tab.url,
          title: tab.title || 'Unknown',
          addedAt: new Date().toISOString()
        };
        addProductToQueue(productInfo);
        return;
      }

      if (response.success) {
        addProductToQueue(response.productInfo);
      }
    });
  }

  function addProductToQueue(productInfo) {
    chrome.storage.local.get(['queue'], (result) => {
      const queue = result.queue || [];

      // 重複チェック
      const exists = queue.some(item => item.url === productInfo.url);
      if (exists) {
        showMessage('既にキューに追加済みです', 'error');
        return;
      }

      queue.push(productInfo);
      chrome.storage.local.set({ queue: queue }, () => {
        showMessage('キューに追加しました', 'success');
      });
    });
  }

  async function startRankingCollection() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const count = parseInt(rankingCountInput.value) || 10;

    startRankingBtn.disabled = true;
    addRankingBtn.disabled = true;
    startRankingBtn.textContent = '準備中...';

    // まずランキングからキューに追加
    chrome.runtime.sendMessage({
      action: 'fetchRanking',
      url: tab.url,
      count: count
    }, (response) => {
      // キューの状態を確認
      chrome.storage.local.get(['queue'], (result) => {
        const queueLength = (result.queue || []).length;

        if (response && response.success && response.addedCount > 0) {
          // 新規追加があった場合
          showRankingMessage(`${response.addedCount}件追加、収集開始...`, 'success');
          startQueueCollectionFromRanking(response.addedCount);
        } else if (queueLength > 0) {
          // 新規追加はないが、キューに既存の商品がある場合
          showRankingMessage(`キューの${queueLength}件を収集開始...`, 'success');
          startQueueCollectionFromRanking(queueLength);
        } else {
          // キューが空で追加もできなかった場合
          startRankingBtn.disabled = false;
          addRankingBtn.disabled = false;
          startRankingBtn.textContent = '収集開始';
          showRankingMessage(response?.error || '商品が見つかりませんでした', 'error');
        }
      });
    });
  }

  function startQueueCollectionFromRanking(itemCount) {
    chrome.runtime.sendMessage({ action: 'startQueueCollection' }, (res) => {
      startRankingBtn.disabled = false;
      addRankingBtn.disabled = false;
      startRankingBtn.textContent = '収集開始';

      if (res && res.success) {
        showRankingMessage(`${itemCount}件の収集を開始しました`, 'success');
      } else {
        showRankingMessage(res?.error || '収集開始に失敗しました', 'error');
      }
    });
  }

  async function addRankingToQueue() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const count = parseInt(rankingCountInput.value) || 10;

    addRankingBtn.disabled = true;
    startRankingBtn.disabled = true;
    addRankingBtn.textContent = '追加中...';

    chrome.runtime.sendMessage({
      action: 'fetchRanking',
      url: tab.url,
      count: count
    }, (response) => {
      addRankingBtn.disabled = false;
      startRankingBtn.disabled = false;
      addRankingBtn.textContent = 'キューに追加';

      if (response && response.success) {
        showRankingMessage(`${response.addedCount}件追加しました`, 'success');
      } else {
        showRankingMessage(response?.error || '追加に失敗しました', 'error');
      }
    });
  }

  function openSettings() {
    chrome.runtime.openOptionsPage();
  }

  function handleMessage(msg) {
    if (!msg || !msg.action) return;

    switch (msg.action) {
      case 'updateProgress':
        restoreState(); // 現在のタブの状態を再確認
        break;
      case 'collectionComplete':
        restoreState(); // 現在のタブの状態を再確認
        break;
      case 'queueUpdated':
        restoreState(); // 現在のタブの状態を再確認
        break;
      case 'duplicatesSkipped':
        // 重複スキップの通知
        if (msg.newCount > 0) {
          showMessage(`${msg.newCount}件追加（${msg.count}件重複スキップ）`, 'success');
        } else {
          showMessage(`全て収集済み（${msg.count}件重複）`, 'error');
        }
        break;
    }
  }

  function showMessage(text, type) {
    message.textContent = text;
    message.className = 'message ' + type;
    setTimeout(() => {
      message.className = 'message';
    }, 3000);
  }

  function showRankingMessage(text, type) {
    rankingMessage.textContent = text;
    rankingMessage.className = 'message ' + type;
    setTimeout(() => {
      rankingMessage.className = 'message';
    }, 3000);
  }
});
