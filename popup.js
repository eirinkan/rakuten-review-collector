/**
 * ポップアップUIのスクリプト（シンプル版）
 */

document.addEventListener('DOMContentLoaded', () => {
  const pageWarning = document.getElementById('pageWarning');
  const message = document.getElementById('message');
  const statusText = document.getElementById('statusText');
  const reviewCount = document.getElementById('reviewCount');
  const pageCount = document.getElementById('pageCount');
  const queueCount = document.getElementById('queueCount');
  const startBtn = document.getElementById('startBtn');
  const stopBtn = document.getElementById('stopBtn');
  const queueBtn = document.getElementById('queueBtn');
  const settingsBtn = document.getElementById('settingsBtn');

  init();

  async function init() {
    // 現在のタブを確認
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const isRakutenPage = tab.url && (
      tab.url.includes('review.rakuten.co.jp') ||
      tab.url.includes('item.rakuten.co.jp')
    );

    if (!isRakutenPage) {
      pageWarning.style.display = 'block';
      startBtn.disabled = true;
      queueBtn.disabled = true;
    }

    // 状態を復元
    restoreState();

    // キュー件数を更新
    updateQueueCount();

    // イベントリスナー
    startBtn.addEventListener('click', startCollection);
    stopBtn.addEventListener('click', stopCollection);
    queueBtn.addEventListener('click', addToQueue);
    settingsBtn.addEventListener('click', openSettings);

    // バックグラウンドからのメッセージ
    chrome.runtime.onMessage.addListener(handleMessage);
  }

  function restoreState() {
    chrome.storage.local.get(['collectionState'], (result) => {
      const state = result.collectionState || {};
      updateUI(state);
    });
  }

  function updateQueueCount() {
    chrome.storage.local.get(['queue'], (result) => {
      const queue = result.queue || [];
      queueCount.textContent = `${queue.length} 件`;
    });
  }

  function updateUI(state) {
    reviewCount.textContent = state.reviewCount || 0;

    const current = state.pageCount || 0;
    const total = state.totalPages || 0;
    pageCount.textContent = `${current} / ${total}`;

    if (state.isRunning) {
      statusText.textContent = '収集中...';
      statusText.classList.add('running');
      startBtn.style.display = 'none';
      stopBtn.style.display = 'block';
    } else {
      statusText.textContent = '待機中';
      statusText.classList.remove('running');
      startBtn.style.display = 'block';
      stopBtn.style.display = 'none';
    }
  }

  async function startCollection() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    chrome.tabs.sendMessage(tab.id, { action: 'startCollection' }, (response) => {
      if (chrome.runtime.lastError) {
        showMessage('ページをリロードしてください', 'error');
        return;
      }

      if (response && response.success) {
        startBtn.style.display = 'none';
        stopBtn.style.display = 'block';
        statusText.textContent = '収集中...';
        statusText.classList.add('running');
      }
    });
  }

  async function stopCollection() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    chrome.tabs.sendMessage(tab.id, { action: 'stopCollection' }, (response) => {
      if (response && response.success) {
        startBtn.style.display = 'block';
        stopBtn.style.display = 'none';
        statusText.textContent = '停止';
        statusText.classList.remove('running');
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
        updateQueueCount();
      });
    });
  }

  function openSettings() {
    chrome.runtime.openOptionsPage();
  }

  function handleMessage(msg) {
    if (!msg || !msg.action) return;

    switch (msg.action) {
      case 'updateProgress':
        if (msg.state) updateUI(msg.state);
        break;
      case 'collectionComplete':
        statusText.textContent = '完了';
        statusText.classList.remove('running');
        startBtn.style.display = 'block';
        stopBtn.style.display = 'none';
        if (msg.state) updateUI(msg.state);
        showMessage(`${msg.state?.reviewCount || 0}件収集完了`, 'success');
        updateQueueCount();
        break;
      case 'queueUpdated':
        updateQueueCount();
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
});
