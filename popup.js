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

  const loginScreen = document.getElementById('loginScreen');
  const mainContent = document.getElementById('mainContent');
  const loginBtn = document.getElementById('loginBtn');
  const loginError = document.getElementById('loginError');
  const userInfo = document.getElementById('userInfo');
  const userEmail = document.getElementById('userEmail');
  const logoutBtn = document.getElementById('logoutBtn');

  const pageWarning = document.getElementById('pageWarning');
  const message = document.getElementById('message');
  const rankingMessage = document.getElementById('rankingMessage');
  const normalMode = document.getElementById('normalMode');
  const rankingMode = document.getElementById('rankingMode');
  const startBtn = document.getElementById('startBtn');
  const stopBtn = document.getElementById('stopBtn');
  const queueBtn = document.getElementById('queueBtn');
  const settingsBtn = document.getElementById('settingsBtn');
  const productQueueBtn = document.getElementById('productQueueBtn');
  const startBothBtn = document.getElementById('startBothBtn');
  const queueBothBtn = document.getElementById('queueBothBtn');
  const startRankingBtn = document.getElementById('startRankingBtn');
  const startRankingProductBtn = document.getElementById('startRankingProductBtn');
  const startRankingBothBtn = document.getElementById('startRankingBothBtn');
  const addRankingBtn = document.getElementById('addRankingBtn');
  const addRankingProductBtn = document.getElementById('addRankingProductBtn');
  const addRankingBothBtn = document.getElementById('addRankingBothBtn');
  const rankingCountInput = document.getElementById('rankingCount');

  // ログインボタン
  if (loginBtn) {
    loginBtn.addEventListener('click', handleLogin);
  }

  // ログアウトボタン
  if (logoutBtn) {
    logoutBtn.addEventListener('click', handleLogout);
  }

  // 認証チェックしてからinit
  checkAuth();

  /**
   * 認証状態をチェック
   */
  async function checkAuth() {
    try {
      chrome.runtime.sendMessage({ action: 'checkAuthStatus' }, (response) => {
        if (chrome.runtime.lastError) {
          showLoginScreen();
          return;
        }

        if (response && response.authenticated) {
          // 認証済み - メイン画面を表示
          showMainContent(response.user);
        } else {
          // 未認証 - ログイン画面を表示
          showLoginScreen();
        }
      });
    } catch (error) {
      console.error('認証チェックエラー:', error);
      showLoginScreen();
    }
  }

  /**
   * ログイン画面を表示
   */
  function showLoginScreen() {
    if (loginScreen) loginScreen.style.display = 'block';
    if (mainContent) mainContent.style.display = 'none';
    if (userInfo) userInfo.style.display = 'none';
    if (loginError) loginError.style.display = 'none';
  }

  /**
   * メインコンテンツを表示
   */
  function showMainContent(user) {
    // 初回ログインの場合は管理画面を開く
    const firstLoginKey = 'rakuten-review-first-login';
    if (!localStorage.getItem(firstLoginKey)) {
      localStorage.setItem(firstLoginKey, 'done');
      chrome.runtime.openOptionsPage();
      return; // ポップアップは閉じる（管理画面に誘導）
    }

    if (loginScreen) loginScreen.style.display = 'none';
    if (mainContent) mainContent.style.display = 'block';

    // ユーザー情報を表示
    if (userInfo && user) {
      userInfo.style.display = 'flex';
      if (userEmail) userEmail.textContent = user.email;
    }

    // メイン機能を初期化
    init();
  }

  /**
   * ログイン処理
   */
  async function handleLogin() {
    if (loginBtn) {
      loginBtn.disabled = true;
      loginBtn.textContent = 'ログイン中...';
    }
    if (loginError) loginError.style.display = 'none';

    try {
      chrome.runtime.sendMessage({ action: 'authenticate' }, (response) => {
        if (chrome.runtime.lastError) {
          showLoginError('認証に失敗しました');
          resetLoginButton();
          return;
        }

        if (response && response.success && response.authenticated) {
          showMainContent(response.user);
        } else {
          showLoginError(response?.message || 'このアカウントは許可されていません');
          resetLoginButton();
        }
      });
    } catch (error) {
      console.error('ログインエラー:', error);
      showLoginError('ログインに失敗しました');
      resetLoginButton();
    }
  }

  /**
   * ログアウト処理
   */
  async function handleLogout() {
    try {
      chrome.runtime.sendMessage({ action: 'logout' }, (response) => {
        if (response && response.success) {
          showLoginScreen();
        }
      });
    } catch (error) {
      console.error('ログアウトエラー:', error);
    }
  }

  /**
   * ログインエラーを表示
   */
  function showLoginError(message) {
    if (loginError) {
      loginError.textContent = message;
      loginError.style.display = 'block';
    }
  }

  /**
   * ログインボタンをリセット
   */
  function resetLoginButton() {
    if (loginBtn) {
      loginBtn.disabled = false;
      loginBtn.textContent = 'Googleでログイン';
    }
  }

  async function init() {
    // 現在のタブを確認
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    // 楽天ページの判定
    const isRakutenPage = tab.url && (
      tab.url.includes('review.rakuten.co.jp') ||
      tab.url.includes('item.rakuten.co.jp')
    );
    const isRakutenRankingPage = tab.url && tab.url.includes('ranking.rakuten.co.jp');
    const isRakutenRankingSitemap = tab.url && tab.url.includes('ranking.rakuten.co.jp/sitemap');

    // Amazonページの判定
    const isAmazonPage = tab.url && tab.url.includes('amazon.co.jp') && (
      tab.url.includes('/dp/') ||
      tab.url.includes('/gp/product/') ||
      tab.url.includes('/product-reviews/')
    );
    const isAmazonRankingPage = tab.url && tab.url.includes('amazon.co.jp') && (
      tab.url.includes('/bestsellers/') ||
      tab.url.includes('/ranking/')
    );

    // 商品ページの判定 - 商品情報収集ボタン用
    const isAmazonProductPage = tab.url && tab.url.includes('amazon.co.jp') && (
      tab.url.includes('/dp/') ||
      tab.url.includes('/gp/product/')
    ) && !tab.url.includes('/product-reviews/');
    const isRakutenProductPage = tab.url && tab.url.includes('item.rakuten.co.jp');
    const isProductPage = isAmazonProductPage || isRakutenProductPage;

    // 商品情報ボタンのイベント登録
    const productInfoBtn = document.getElementById('productInfoBtn');
    if (productInfoBtn) productInfoBtn.addEventListener('click', collectProductInfo);
    if (productQueueBtn) productQueueBtn.addEventListener('click', addToProductQueue);

    // 商品ページでない場合: 商品情報系・両方ボタンをdisabled
    if (!isProductPage) {
      if (productInfoBtn) productInfoBtn.disabled = true;
      if (productQueueBtn) productQueueBtn.disabled = true;
      if (startBothBtn) startBothBtn.disabled = true;
      if (queueBothBtn) queueBothBtn.disabled = true;
    }

    // 対応ページの判定
    const isSupportedPage = isRakutenPage || isAmazonPage;
    const isRankingPage = isRakutenRankingPage || isAmazonRankingPage;

    if (isRakutenRankingPage) {
      // 楽天ランキングページの場合
      normalMode.style.display = 'none';
      rankingMode.style.display = 'block';

      // サイトマップページの場合は全ボタン無効化
      if (isRakutenRankingSitemap) {
        startRankingBtn.disabled = true;
        if (startRankingProductBtn) startRankingProductBtn.disabled = true;
        if (startRankingBothBtn) startRankingBothBtn.disabled = true;
        addRankingBtn.disabled = true;
        if (addRankingProductBtn) addRankingProductBtn.disabled = true;
        if (addRankingBothBtn) addRankingBothBtn.disabled = true;
        showRankingMessage('任意のランキングを選んでください', 'info');
      }
    } else if (isAmazonRankingPage) {
      // Amazonランキングページの場合
      normalMode.style.display = 'none';
      rankingMode.style.display = 'block';
    } else if (!isSupportedPage) {
      // 楽天・Amazon以外のページ
      pageWarning.style.display = 'block';
      startBtn.disabled = true;
      if (productInfoBtn) productInfoBtn.disabled = true;
      if (startBothBtn) startBothBtn.disabled = true;
      queueBtn.disabled = true;
      if (productQueueBtn) productQueueBtn.disabled = true;
      if (queueBothBtn) queueBothBtn.disabled = true;
    }

    // 状態を復元
    restoreState();

    // イベントリスナー
    startBtn.addEventListener('click', startCollection);
    stopBtn.addEventListener('click', stopCollection);
    if (startBothBtn) startBothBtn.addEventListener('click', startBothCollection);
    queueBtn.addEventListener('click', addToQueue);
    if (queueBothBtn) queueBothBtn.addEventListener('click', addBothToQueue);
    settingsBtn.addEventListener('click', openSettings);
    startRankingBtn.addEventListener('click', startRankingCollection);
    if (startRankingProductBtn) startRankingProductBtn.addEventListener('click', startRankingProductCollection);
    if (startRankingBothBtn) startRankingBothBtn.addEventListener('click', startRankingBothCollection);
    addRankingBtn.addEventListener('click', addRankingToQueue);
    if (addRankingProductBtn) addRankingProductBtn.addEventListener('click', addRankingToProductQueue);
    if (addRankingBothBtn) addRankingBothBtn.addEventListener('click', addRankingBothToQueue);

    // バックグラウンドからのメッセージ
    chrome.runtime.onMessage.addListener(handleMessage);
  }

  async function restoreState() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    chrome.storage.local.get(['collectingItems', 'isQueueCollecting'], (result) => {
      const collectingItems = result.collectingItems || [];
      const isQueueCollecting = result.isQueueCollecting || false;
      // 現在のタブが収集中リストにあるかチェック
      const isCurrentTabCollecting = collectingItems.some(item => item.tabId === tab.id);
      updateUI({ isRunning: isCurrentTabCollecting });

      // ランキングページのボタン状態を復元
      if (isQueueCollecting) {
        updateRankingUI(true);
      }
    });

    // 商品バッチ収集の状態を確認
    chrome.runtime.sendMessage({ action: 'getBatchProductProgress' }, (res) => {
      if (chrome.runtime.lastError) return;
      if (res && res.progress && res.progress.isRunning) {
        updateRankingProductUI(true);
      }
    });
  }

  function updateUI(state) {
    const startGroup = document.getElementById('startGroup');
    if (state.isRunning) {
      if (startGroup) startGroup.style.display = 'none';
      stopBtn.style.display = 'block';
    } else {
      if (startGroup) startGroup.style.display = 'block';
      stopBtn.style.display = 'none';
    }
  }

  // ランキングページのレビュー収集ボタン状態を更新
  function updateRankingUI(isCollecting) {
    if (!startRankingBtn) return;
    if (isCollecting) {
      startRankingBtn.textContent = '収集中...';
      startRankingBtn.disabled = true;
      startRankingBtn.classList.remove('btn-start');
      startRankingBtn.classList.add('btn-stop');
    } else {
      startRankingBtn.textContent = 'レビュー収集開始';
      startRankingBtn.disabled = false;
      startRankingBtn.classList.remove('btn-stop');
      startRankingBtn.classList.add('btn-start');
    }
  }

  // ランキングページの商品情報収集ボタン状態を更新
  function updateRankingProductUI(isCollecting) {
    if (!startRankingProductBtn) return;
    if (isCollecting) {
      startRankingProductBtn.textContent = '商品情報収集中...';
      startRankingProductBtn.disabled = true;
    } else {
      startRankingProductBtn.textContent = '商品情報収集開始';
      startRankingProductBtn.disabled = false;
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

    // ボタンを即座に切り替え（レスポンスを待たない）
    updateUI({ isRunning: true });
    showMessage('収集を開始しています...', 'success');

    // 商品情報を取得
    chrome.tabs.sendMessage(tab.id, { action: 'getProductInfo' }, (response) => {
      if (chrome.runtime.lastError) {
        // エラー時はボタンを戻す
        updateUI({ isRunning: false });
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
          showMessage('収集を開始しました', 'success');
        } else {
          // エラー時はボタンを戻す
          updateUI({ isRunning: false });
          showMessage(res?.error || '収集開始に失敗しました', 'error');
        }
      });
    });
  }

  /**
   * 両方（レビュー＋商品情報）を同時に収集開始
   */
  async function startBothCollection() {
    // レビュー収集を開始
    await startCollection();
    // 商品情報収集も開始
    collectProductInfo();
  }

  async function stopCollection() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    chrome.tabs.sendMessage(tab.id, { action: 'stopCollection' }, (response) => {
      if (response && response.success) {
        updateUI({ isRunning: false });
      }
    });
  }

  async function addToQueue() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    // 楽天ページの判定
    const isRakutenPage = tab.url.includes('item.rakuten.co.jp') || tab.url.includes('review.rakuten.co.jp');
    // Amazonページの判定
    const isAmazonPage = tab.url.includes('amazon.co.jp') && (
      tab.url.includes('/dp/') ||
      tab.url.includes('/gp/product/') ||
      tab.url.includes('/product-reviews/')
    );

    if (!isRakutenPage && !isAmazonPage) {
      showMessage('楽天またはAmazonの商品ページを開いてください', 'error');
      return;
    }

    // 商品情報を取得してキューに追加
    chrome.tabs.sendMessage(tab.id, { action: 'getProductInfo' }, (response) => {
      if (chrome.runtime.lastError || !response) {
        // content scriptが応答しない場合はURLから情報を取得
        const productInfo = {
          url: tab.url,
          title: extractTitleFromTabTitle(tab.title, isAmazonPage),
          addedAt: new Date().toISOString(),
          source: isAmazonPage ? 'amazon' : 'rakuten'
        };
        addProductToQueue(productInfo);
        return;
      }

      if (response.success) {
        const productInfo = response.productInfo;
        // titleが空の場合、tab.titleからフォールバック
        if (!productInfo.title || productInfo.title.trim() === '') {
          productInfo.title = extractTitleFromTabTitle(tab.title, isAmazonPage);
        }
        addProductToQueue(productInfo);
      }
    });
  }

  // タブタイトルから商品名を抽出
  function extractTitleFromTabTitle(tabTitle, isAmazon) {
    if (!tabTitle) return 'Unknown';
    let title = tabTitle;
    if (isAmazon) {
      // 「Amazon.co.jp: 商品名」「Amazon.co.jp：商品名」から商品名を抽出
      title = title.replace(/^Amazon\.co\.jp[：:]\s*/i, '');
      // 「カスタマーレビュー: 商品名」を除去
      title = title.replace(/^カスタマーレビュー[：:]\s*/i, '');
      // 末尾の「 : カテゴリ名」を除去
      title = title.replace(/\s*:\s*[^:]+$/, '');
    }
    return title.trim() || 'Unknown';
  }

  function addProductToQueue(productInfo) {
    chrome.storage.local.get(['queue'], (result) => {
      const queue = result.queue || [];

      // 重複チェック
      const exists = queue.some(item => item.url === productInfo.url);
      if (exists) {
        showMessage('既にキューに追加済みです', 'error');
        chrome.runtime.sendMessage({ action: 'log', text: '既にキューに追加済みです', type: 'error' });
        return;
      }

      queue.push(productInfo);
      chrome.storage.local.set({ queue: queue }, () => {
        showMessage('キューに追加しました', 'success');
        // ログを送信
        const title = productInfo.title || productInfo.url;
        chrome.runtime.sendMessage({ action: 'log', text: `「${title}」をキューに追加しました`, type: 'success' });
        // キュー更新を通知
        chrome.runtime.sendMessage({ action: 'queueUpdated' });
      });
    });
  }

  /**
   * 両方のキューに追加（レビューキュー＋商品情報キュー）
   */
  async function addBothToQueue() {
    await addToQueue();
    await addToProductQueue();
  }

  /**
   * 現在の商品ページを商品情報キュー(batchProductQueue)に追加
   */
  async function addToProductQueue() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    const isAmazonPage = tab.url.includes('amazon.co.jp');
    const isRakutenPage = tab.url.includes('item.rakuten.co.jp');

    if (!isAmazonPage && !isRakutenPage) {
      showMessage('商品ページを開いてください', 'error');
      return;
    }

    let item;
    if (isAmazonPage) {
      // AmazonページからASINを抽出
      const dpMatch = tab.url.match(/\/dp\/([A-Z0-9]{10})/i);
      const gpMatch = tab.url.match(/\/gp\/product\/([A-Z0-9]{10})/i);
      const asin = (dpMatch && dpMatch[1]) || (gpMatch && gpMatch[1]);
      if (!asin) {
        showMessage('ASINが取得できませんでした', 'error');
        return;
      }
      item = asin.toUpperCase();
    } else {
      // 楽天ページのURL（クエリパラメータ除去）
      item = tab.url.split('?')[0];
    }

    chrome.storage.local.get(['batchProductQueue'], (result) => {
      const queue = result.batchProductQueue || [];
      if (queue.includes(item)) {
        showMessage('既に商品キューに追加済みです', 'error');
        return;
      }
      queue.push(item);
      chrome.storage.local.set({ batchProductQueue: queue }, () => {
        showMessage('商品情報キューに追加しました', 'success');
        chrome.runtime.sendMessage({ action: 'batchProductQueueUpdated' });
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
      addRankingBtn.disabled = false;

      if (res && res.success) {
        // 収集中状態を維持（ボタンをリセットしない）
        updateRankingUI(true);
        showRankingMessage(`${itemCount}件の収集を開始しました`, 'success');
      } else {
        // エラー時はボタンを戻す
        updateRankingUI(false);
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
      addRankingBtn.textContent = 'レビューキューに追加';

      if (response && response.success) {
        showRankingMessage(`${response.addedCount}件追加しました`, 'success');
      } else {
        showRankingMessage(response?.error || '追加に失敗しました', 'error');
      }
    });
  }

  async function addRankingToProductQueue() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const count = parseInt(rankingCountInput.value) || 10;

    addRankingProductBtn.disabled = true;
    addRankingBtn.disabled = true;
    startRankingBtn.disabled = true;
    addRankingProductBtn.textContent = '追加中...';

    chrome.runtime.sendMessage({
      action: 'addRankingToProductQueue',
      url: tab.url,
      count: count
    }, (response) => {
      addRankingProductBtn.disabled = false;
      addRankingBtn.disabled = false;
      startRankingBtn.disabled = false;
      addRankingProductBtn.textContent = '商品キューに追加';

      if (response && response.success) {
        showRankingMessage(`商品キューに${response.addedCount}件追加しました`, 'success');
      } else {
        showRankingMessage(response?.error || '追加に失敗しました', 'error');
      }
    });
  }

  /**
   * ランキングから商品キューに追加して、追加分のみ商品情報バッチ収集を開始
   */
  async function startRankingProductCollection() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const count = parseInt(rankingCountInput.value) || 10;

    if (startRankingProductBtn) startRankingProductBtn.disabled = true;
    startRankingBtn.disabled = true;
    addRankingBtn.disabled = true;
    if (addRankingProductBtn) addRankingProductBtn.disabled = true;
    if (startRankingProductBtn) startRankingProductBtn.textContent = '準備中...';

    // まずランキングから商品キューに追加
    chrome.runtime.sendMessage({
      action: 'addRankingToProductQueue',
      url: tab.url,
      count: count
    }, (response) => {
      if (response && response.success && response.addedCount > 0) {
        showRankingMessage(`${response.addedCount}件追加、商品情報収集開始...`, 'success');
        // 追加分のみバッチ収集を開始（キュー全体ではなく今回追加した商品だけ）
        chrome.runtime.sendMessage({
          action: 'startBatchProductCollection',
          items: response.addedItems
        }, (res) => {
          // 他のボタンは戻す
          startRankingBtn.disabled = false;
          addRankingBtn.disabled = false;
          if (addRankingProductBtn) addRankingProductBtn.disabled = false;

          if (res && !res.error) {
            // 収集中状態を維持
            updateRankingProductUI(true);
            showRankingMessage(`${response.addedCount}件の商品情報収集を開始しました`, 'success');
          } else {
            // エラー時はボタンを戻す
            updateRankingProductUI(false);
            showRankingMessage(res?.error || '商品情報収集の開始に失敗しました', 'error');
          }
        });
      } else {
        resetRankingProductBtn();
        showRankingMessage(response?.error || '商品が見つかりませんでした', 'error');
      }
    });
  }

  /**
   * ランキングから「両方」キューに追加（レビューキュー＋商品キュー）
   */
  async function addRankingBothToQueue() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const count = parseInt(rankingCountInput.value) || 10;

    // 全ボタン無効化
    if (addRankingBothBtn) addRankingBothBtn.disabled = true;
    addRankingBtn.disabled = true;
    if (addRankingProductBtn) addRankingProductBtn.disabled = true;
    startRankingBtn.disabled = true;
    if (startRankingProductBtn) startRankingProductBtn.disabled = true;
    if (startRankingBothBtn) startRankingBothBtn.disabled = true;
    if (addRankingBothBtn) addRankingBothBtn.textContent = '追加中...';

    // レビューキューに追加
    chrome.runtime.sendMessage({
      action: 'fetchRanking',
      url: tab.url,
      count: count
    }, (reviewRes) => {
      // 商品キューに追加
      chrome.runtime.sendMessage({
        action: 'addRankingToProductQueue',
        url: tab.url,
        count: count
      }, (productRes) => {
        // ボタン復帰
        if (addRankingBothBtn) { addRankingBothBtn.disabled = false; addRankingBothBtn.textContent = '両方'; }
        addRankingBtn.disabled = false;
        if (addRankingProductBtn) addRankingProductBtn.disabled = false;
        startRankingBtn.disabled = false;
        if (startRankingProductBtn) startRankingProductBtn.disabled = false;
        if (startRankingBothBtn) startRankingBothBtn.disabled = false;

        const reviewCount = (reviewRes && reviewRes.success) ? reviewRes.addedCount : 0;
        const productCount = (productRes && productRes.success) ? productRes.addedCount : 0;
        showRankingMessage(`レビュー${reviewCount}件・商品${productCount}件追加`, 'success');
      });
    });
  }

  /**
   * ランキングから「両方」収集開始（レビュー収集＋商品情報収集を同時に開始）
   */
  async function startRankingBothCollection() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    const count = parseInt(rankingCountInput.value) || 10;

    // 全ボタン無効化
    if (startRankingBothBtn) { startRankingBothBtn.disabled = true; startRankingBothBtn.textContent = '準備中...'; }
    startRankingBtn.disabled = true;
    if (startRankingProductBtn) startRankingProductBtn.disabled = true;
    addRankingBtn.disabled = true;
    if (addRankingProductBtn) addRankingProductBtn.disabled = true;
    if (addRankingBothBtn) addRankingBothBtn.disabled = true;

    // レビューキューに追加
    chrome.runtime.sendMessage({
      action: 'fetchRanking',
      url: tab.url,
      count: count
    }, (reviewRes) => {
      // 商品キューに追加
      chrome.runtime.sendMessage({
        action: 'addRankingToProductQueue',
        url: tab.url,
        count: count
      }, (productRes) => {
        const reviewCount = (reviewRes && reviewRes.success) ? reviewRes.addedCount : 0;
        const productCount = (productRes && productRes.success) ? productRes.addedCount : 0;

        if (reviewCount === 0 && productCount === 0) {
          // 何も追加できなかった
          resetAllRankingBtns();
          showRankingMessage('商品が見つかりませんでした', 'error');
          return;
        }

        // レビュー収集開始
        if (reviewCount > 0 || (reviewRes && reviewRes.success)) {
          chrome.runtime.sendMessage({ action: 'startQueueCollection' }, (res) => {
            if (res && res.success) updateRankingUI(true);
          });
        }

        // 商品情報収集開始
        if (productRes && productRes.success && productRes.addedItems && productRes.addedItems.length > 0) {
          chrome.runtime.sendMessage({
            action: 'startBatchProductCollection',
            items: productRes.addedItems
          }, (res) => {
            if (res && !res.error) updateRankingProductUI(true);
          });
        }

        // ボタン復帰（収集中ボタンは除く）
        addRankingBtn.disabled = false;
        if (addRankingProductBtn) addRankingProductBtn.disabled = false;
        if (addRankingBothBtn) { addRankingBothBtn.disabled = false; addRankingBothBtn.textContent = '両方'; }
        if (startRankingBothBtn) { startRankingBothBtn.disabled = false; startRankingBothBtn.textContent = '両方'; }

        showRankingMessage(`レビュー${reviewCount}件・商品${productCount}件の収集開始`, 'success');
      });
    });
  }

  function resetAllRankingBtns() {
    startRankingBtn.disabled = false;
    if (startRankingProductBtn) startRankingProductBtn.disabled = false;
    if (startRankingBothBtn) { startRankingBothBtn.disabled = false; startRankingBothBtn.textContent = '両方'; }
    addRankingBtn.disabled = false;
    if (addRankingProductBtn) addRankingProductBtn.disabled = false;
    if (addRankingBothBtn) { addRankingBothBtn.disabled = false; addRankingBothBtn.textContent = '両方'; }
  }

  function resetRankingProductBtn() {
    updateRankingProductUI(false);
    startRankingBtn.disabled = false;
    addRankingBtn.disabled = false;
    if (addRankingProductBtn) addRankingProductBtn.disabled = false;
  }

  function openSettings() {
    chrome.runtime.openOptionsPage();
  }

  /**
   * 商品情報を収集（キューに追加してからバッチ収集を開始）
   */
  async function collectProductInfo() {
    const productInfoBtn = document.getElementById('productInfoBtn');
    if (!productInfoBtn) return;

    productInfoBtn.disabled = true;
    productInfoBtn.textContent = '追加中...';

    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

      const isAmazonPage = tab.url.includes('amazon.co.jp');
      const isRakutenPage = tab.url.includes('item.rakuten.co.jp');

      if (!isAmazonPage && !isRakutenPage) {
        showMessage('商品ページを開いてください', 'error');
        resetProductInfoBtn();
        return;
      }

      // キュー用のアイテムを作成
      let item;
      if (isAmazonPage) {
        const dpMatch = tab.url.match(/\/dp\/([A-Z0-9]{10})/i);
        const gpMatch = tab.url.match(/\/gp\/product\/([A-Z0-9]{10})/i);
        const asin = (dpMatch && dpMatch[1]) || (gpMatch && gpMatch[1]);
        if (!asin) {
          showMessage('ASINが取得できませんでした', 'error');
          resetProductInfoBtn();
          return;
        }
        item = asin.toUpperCase();
      } else {
        item = tab.url.split('?')[0];
      }

      // 商品キューに追加
      chrome.storage.local.get(['batchProductQueue'], (result) => {
        const queue = result.batchProductQueue || [];
        const alreadyInQueue = queue.includes(item);

        if (!alreadyInQueue) {
          queue.push(item);
          chrome.storage.local.set({ batchProductQueue: queue });
          chrome.runtime.sendMessage({ action: 'batchProductQueueUpdated' });
        }

        // バッチ収集を開始（この1件だけ）
        chrome.runtime.sendMessage({
          action: 'startBatchProductCollection',
          items: [item]
        }, (res) => {
          if (res && !res.error) {
            productInfoBtn.textContent = '収集中...';
            showMessage('商品情報の収集を開始しました', 'success');
          } else {
            showMessage(res?.error || '収集開始に失敗しました', 'error');
            resetProductInfoBtn();
          }
        });
      });
    } catch (error) {
      showMessage('エラー: ' + error.message, 'error');
      resetProductInfoBtn();
    }
  }

  function resetProductInfoBtn() {
    const productInfoBtn = document.getElementById('productInfoBtn');
    if (productInfoBtn) {
      productInfoBtn.disabled = false;
      productInfoBtn.textContent = '商品情報';
    }
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
      case 'queueCollectionComplete':
        // キュー全件収集完了 — ランキングボタンをリセット
        updateRankingUI(false);
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
      case 'productInfoProgress': {
        // 商品情報収集の進捗
        const progressBar = document.getElementById('productInfoProgressBar');
        const progressText = document.getElementById('productInfoProgressText');
        if (msg.progress) {
          if (msg.progress.phase === 'images' && progressBar && progressText) {
            const pct = Math.round((msg.progress.current / msg.progress.total) * 70) + 10;
            progressBar.style.width = pct + '%';
            progressText.textContent = `メディアをアップロード中... (${msg.progress.current}/${msg.progress.total})`;
          } else if (msg.progress.phase === 'videos' && progressBar && progressText) {
            const pct = Math.round((msg.progress.current / msg.progress.total) * 70) + 10;
            progressBar.style.width = pct + '%';
            progressText.textContent = `動画を処理中... (${msg.progress.current}/${msg.progress.total})`;
          } else if (msg.progress.phase === 'upload' && progressBar && progressText) {
            progressBar.style.width = '90%';
            progressText.textContent = 'JSONを保存中...';
          }
        }
        break;
      }
      case 'batchProductProgressUpdate':
        // 商品バッチ収集の進捗/完了
        if (msg.progress && !msg.progress.isRunning) {
          updateRankingProductUI(false);
          resetProductInfoBtn();
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
