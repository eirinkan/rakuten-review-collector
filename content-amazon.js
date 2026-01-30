/**
 * コンテンツスクリプト - Amazon
 * Amazon.co.jp の商品レビューを収集する
 */

(function() {
  'use strict';

  // ===== ボット対策: 定数 =====
  const MAX_PAGES_PER_SESSION = 999999;  // 1セッションあたりの最大ページ数（実質無制限）
  const MICRO_BREAK_PROBABILITY = 0.05;  // 各ページで休憩する確率（5%）
  const PAGE_WAIT_MEAN_MS = 1500;        // ページ遷移前の平均待機時間（1.5秒、指数分布）
  const PAGE_WAIT_MAX_MS = 8000;         // ページ遷移前の最大待機時間（8秒）
  const ACTIVE_TAB_CHECK_INTERVAL = 2000; // アクティブタブチェック間隔（2秒）
  const READING_SPEED_PER_100CHARS = 400; // 100文字あたりの読み時間（ms）

  // ===== 星評価フィルター: 定数 =====
  // 低評価から順に収集（問題点を優先）
  const STAR_FILTERS = [
    { value: 'one_star', label: '★1' },
    { value: 'two_star', label: '★2' },
    { value: 'three_star', label: '★3' },
    { value: 'four_star', label: '★4' },
    { value: 'five_star', label: '★5' }
  ];
  const MAX_PAGES_PER_FILTER = 999999; // 制限なし（Amazonの実際の制限を検証するため）

  // ===== ボット対策: マウス・スクロール関連関数 =====

  /**
   * ベジェ曲線でポイントを計算（3次ベジェ曲線）
   * 人間らしいマウス軌道を生成するため
   */
  function bezierPoint(t, p0, p1, p2, p3) {
    const u = 1 - t;
    return u * u * u * p0 +
           3 * u * u * t * p1 +
           3 * u * t * t * p2 +
           t * t * t * p3;
  }

  /**
   * ベジェ曲線に沿ってマウスを移動（イベント発火）
   * @param {number} startX - 開始X座標
   * @param {number} startY - 開始Y座標
   * @param {number} endX - 終了X座標
   * @param {number} endY - 終了Y座標
   * @param {number} steps - ステップ数（デフォルト20-40）
   */
  async function moveMouseAlongBezier(startX, startY, endX, endY, steps = null) {
    // ステップ数をランダムに（20-40）
    if (!steps) {
      steps = 20 + Math.floor(Math.random() * 20);
    }

    // 制御点をランダムに生成（曲線の形状を決める）
    const distX = endX - startX;
    const distY = endY - startY;

    // 制御点1: 始点から1/3程度、ランダムにずらす
    const cp1x = startX + distX * 0.3 + (Math.random() - 0.5) * Math.abs(distX) * 0.5;
    const cp1y = startY + distY * 0.3 + (Math.random() - 0.5) * Math.abs(distY) * 0.5;

    // 制御点2: 終点から1/3程度、ランダムにずらす
    const cp2x = startX + distX * 0.7 + (Math.random() - 0.5) * Math.abs(distX) * 0.5;
    const cp2y = startY + distY * 0.7 + (Math.random() - 0.5) * Math.abs(distY) * 0.5;

    for (let i = 0; i <= steps; i++) {
      // フィッツの法則: 最後の方は減速
      let t;
      if (i < steps * 0.7) {
        // 前半70%は通常速度
        t = (i / steps) * 0.7;
      } else {
        // 後半30%は減速（細かく移動）
        const remaining = (i - steps * 0.7) / (steps * 0.3);
        t = 0.7 + remaining * 0.3;
      }

      const x = bezierPoint(t, startX, cp1x, cp2x, endX);
      const y = bezierPoint(t, startY, cp1y, cp2y, endY);

      // わずかな振れ（手のブレ）を追加
      const jitterX = (Math.random() - 0.5) * 2;
      const jitterY = (Math.random() - 0.5) * 2;

      // mousemoveイベントを発火
      const event = new MouseEvent('mousemove', {
        bubbles: true,
        cancelable: true,
        view: window,
        clientX: x + jitterX,
        clientY: y + jitterY
      });
      document.dispatchEvent(event);

      // 間隔をランダムに（5-20ms）
      await sleep(5 + Math.random() * 15);
    }
  }

  /**
   * 要素にマウスを移動してクリック（人間らしく）
   * @param {Element} element - クリック対象の要素
   */
  async function humanClick(element) {
    if (!element) return false;

    const rect = element.getBoundingClientRect();

    // 現在のマウス位置（ページの適当な場所から開始）
    const startX = Math.random() * window.innerWidth;
    const startY = Math.random() * window.innerHeight;

    // ターゲットの中心座標（少しランダムにずらす）
    const targetX = rect.left + rect.width * (0.3 + Math.random() * 0.4);
    const targetY = rect.top + rect.height * (0.3 + Math.random() * 0.4);

    // ベジェ曲線でマウスを移動
    await moveMouseAlongBezier(startX, startY, targetX, targetY);

    // ホバー効果のために少し待機（100-300ms）
    await sleep(100 + Math.random() * 200);

    // mouseenterイベント
    element.dispatchEvent(new MouseEvent('mouseenter', {
      bubbles: true,
      cancelable: true,
      view: window,
      clientX: targetX,
      clientY: targetY
    }));

    // mouseoverイベント
    element.dispatchEvent(new MouseEvent('mouseover', {
      bubbles: true,
      cancelable: true,
      view: window,
      clientX: targetX,
      clientY: targetY
    }));

    // クリック前の微小な待機（50-150ms）
    await sleep(50 + Math.random() * 100);

    // mousedownイベント
    element.dispatchEvent(new MouseEvent('mousedown', {
      bubbles: true,
      cancelable: true,
      view: window,
      button: 0,
      clientX: targetX,
      clientY: targetY
    }));

    // mouseup/click の間隔（50-150ms）
    await sleep(50 + Math.random() * 100);

    // mouseupイベント
    element.dispatchEvent(new MouseEvent('mouseup', {
      bubbles: true,
      cancelable: true,
      view: window,
      button: 0,
      clientX: targetX,
      clientY: targetY
    }));

    // clickイベント
    element.dispatchEvent(new MouseEvent('click', {
      bubbles: true,
      cancelable: true,
      view: window,
      button: 0,
      clientX: targetX,
      clientY: targetY
    }));

    return true;
  }

  /**
   * 人間らしいスクロール動作
   * - 加速・減速パターン
   * - 不規則な間隔
   * - 時々逆方向にスクロール
   */
  async function humanScroll(targetY, options = {}) {
    const {
      maxDuration = 3000,  // 最大スクロール時間
      allowOvershoot = true // 行き過ぎを許可
    } = options;

    const startY = window.scrollY;
    const distance = targetY - startY;
    const direction = distance > 0 ? 1 : -1;
    const absDistance = Math.abs(distance);

    if (absDistance < 50) {
      window.scrollTo({ top: targetY, behavior: 'smooth' });
      return;
    }

    const startTime = Date.now();
    let currentY = startY;
    let velocity = 0;
    const maxVelocity = 15 + Math.random() * 10; // 最大速度（ランダム）

    while (Math.abs(currentY - targetY) > 30 && Date.now() - startTime < maxDuration) {
      if (shouldStop) return;

      const remaining = Math.abs(targetY - currentY);
      const progress = 1 - (remaining / absDistance);

      // 加速・減速パターン
      if (progress < 0.3) {
        // 加速フェーズ
        velocity = Math.min(velocity + 1.5, maxVelocity * progress * 3);
      } else if (progress > 0.7) {
        // 減速フェーズ
        velocity = maxVelocity * (1 - progress) * 2;
      } else {
        // 定速フェーズ（わずかな変動）
        velocity = maxVelocity * (0.8 + Math.random() * 0.4);
      }

      // 最小速度を保証
      velocity = Math.max(velocity, 3);

      // 時々微小な逆方向スクロール（読み返しのシミュレーション）
      if (Math.random() < 0.02 && progress > 0.2 && progress < 0.8) {
        const backScroll = 20 + Math.random() * 50;
        currentY -= direction * backScroll;
        window.scrollTo({ top: currentY, behavior: 'auto' });
        await sleep(200 + Math.random() * 300);
        continue;
      }

      currentY += direction * velocity;
      window.scrollTo({ top: currentY, behavior: 'auto' });

      // 不規則な間隔（10-40ms）
      await sleep(10 + Math.random() * 30);

      // 時々読んでいるような停止（2%の確率）
      if (Math.random() < 0.02) {
        await sleep(300 + Math.random() * 700);
      }
    }

    // 行き過ぎの修正（オプション）
    if (allowOvershoot && Math.random() < 0.3) {
      // 30%の確率で少し行き過ぎる
      const overshoot = 20 + Math.random() * 40;
      window.scrollTo({ top: targetY + direction * overshoot, behavior: 'smooth' });
      await sleep(200 + Math.random() * 200);
      window.scrollTo({ top: targetY, behavior: 'smooth' });
    } else {
      window.scrollTo({ top: targetY, behavior: 'smooth' });
    }
  }

  // 収集状態
  let isCollecting = false;
  let shouldStop = false;
  let currentProductId = ''; // 現在のASIN
  let totalPages = 0; // 総ページ数（ログ表示用）
  let incrementalOnly = false; // 差分取得モード
  let lastCollectedDate = null; // 前回収集日
  let currentQueueName = null; // 定期収集のキュー名
  let autoResumeExecuted = false; // 自動再開が実行済みかどうか
  let collectedReviewKeys = new Set(); // このセッションで収集済みのレビューキー
  let startCollectionLock = false; // 収集開始のロック（重複防止）
  let resumeCollectionLock = false; // 収集再開のロック（重複防止）
  let sessionPageCount = 0; // セッション内のページカウント（ボット対策）
  let consecutiveSkipPages = 0; // 連続してスキップしたページ数（無限ループ防止）
  const MAX_CONSECUTIVE_SKIP = 3; // 連続スキップの上限

  // 星評価フィルター状態
  let useStarFilter = true; // 星評価フィルターを使用するか
  let currentStarFilterIndex = 0; // 現在のフィルターインデックス（0-4）
  let pagesCollectedInCurrentFilter = 0; // 現在のフィルターで収集済みのページ数
  let filterPageCounts = {}; // 各フィルターのページ数を記録（検証用）

  // Amazonセレクター（前回のテストで確認済み）
  const AMAZON_SELECTORS = {
    reviewContainer: '[data-hook="review"]',
    rating: 'i.review-rating span',
    title: '[data-hook="review-title"] span:not([class])',
    body: '[data-hook="review-body"]',
    author: '.a-profile-name',
    date: '[data-hook="review-date"]',
    variation: '[data-hook="format-strip"]',
    helpful: '[data-hook="helpful-vote-statement"]',
    verified: '[data-hook="avp-badge"]',
    vine: 'span.a-color-success.a-text-bold',
    image: '[data-hook="review-image-tile"]',
    nextPage: 'li.a-last a',
    isLastPage: 'li.a-last.a-disabled',
    totalReviews: '[data-hook="cr-filter-info-review-rating-count"]',
    reviewLink: 'a[data-hook="see-all-reviews-link-foot"]',
    productTitle: '#productTitle',
    reviewPageTitle: '[data-hook="product-link"]',  // レビューページでの商品名
    rankingProduct: '[data-asin]',
  };

  // ページタイプを判定
  const isReviewPage = window.location.pathname.includes('/product-reviews/');
  const isProductPage = window.location.pathname.includes('/dp/') || window.location.pathname.includes('/gp/product/');
  const isRankingPage = window.location.pathname.includes('/bestsellers/') || window.location.pathname.includes('/ranking/');

  // メッセージリスナー
  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    switch (message.action) {
      case 'startCollection':
        // 既に収集中の場合
        if (isCollecting) {
          if (message.force) {
            shouldStop = true;
            isCollecting = false;
            log('前の収集を中断して再開始します');
          } else {
            console.log('[Amazonレビュー収集] 既に収集中のためスキップ');
            sendResponse({ success: false, error: '既に収集中' });
            break;
          }
        }

        if (isProductPage) {
          // 商品ページの場合、レビューページに遷移
          const reviewLink = findReviewPageLink();
          if (reviewLink) {
            incrementalOnly = message.incrementalOnly || false;
            lastCollectedDate = message.lastCollectedDate || null;
            currentQueueName = message.queueName || null;
            // productIdが送られてきた場合は設定
            if (message.productId) {
              currentProductId = message.productId;
            }

            const asin = getASIN();
            // currentProductIdを設定（ログ出力で使用）
            if (asin) {
              currentProductId = asin;
            }
            log('レビューページに移動します');

            // 収集状態を設定してからクリックで遷移
            chrome.storage.local.get(['collectionState'], (result) => {
              const state = result.collectionState || {
                isRunning: false,
                reviewCount: 0,
                pageCount: 0,
                totalPages: 0,
                reviews: [],
                logs: []
              };
              state.isRunning = true;
              state.incrementalOnly = incrementalOnly;
              state.lastCollectedDate = lastCollectedDate;
              state.queueName = currentQueueName;
              state.source = 'amazon'; // 販路を追加
              // 重要: startedFromQueueを明示的に保存（レビューページ遷移時に使用）
              // processNextInQueueから開始された場合はtrueが設定されている
              if (state.startedFromQueue === undefined) {
                state.startedFromQueue = false;
              }
              chrome.storage.local.set({ collectionState: state }, async () => {
                // クリックでページ遷移（人間らしいマウス移動＋クリック）
                await humanClick(reviewLink);
              });
            });
            sendResponse({ success: true, redirecting: true });
          } else {
            sendResponse({ success: false, error: 'レビューページが見つかりません' });
          }
        } else if (isReviewPage) {
          // レビューページの場合、収集開始
          incrementalOnly = message.incrementalOnly || false;
          lastCollectedDate = message.lastCollectedDate || null;
          currentQueueName = message.queueName || null;
          startCollection();
          sendResponse({ success: true });
        } else {
          sendResponse({ success: false, error: 'このページでは収集できません' });
        }
        break;

      case 'stopCollection':
        shouldStop = true;
        isCollecting = false;
        chrome.runtime.sendMessage({ action: 'collectionStopped' });
        sendResponse({ success: true });
        break;

      case 'resumeCollection':
        // 同期的なロックチェック（最初に行う - 重複実行防止）
        if (resumeCollectionLock) {
          console.log('[Amazonレビュー収集] resumeCollection実行中のためスキップ');
          sendResponse({ success: false, error: 'resumeCollection実行中' });
          break;
        }
        // isCollectingとstartCollectionLockもチェック
        if (isCollecting || startCollectionLock) {
          console.log('[Amazonレビュー収集] 既に収集中のためスキップ', { isCollecting, startCollectionLock });
          sendResponse({ success: false, error: '既に収集中' });
          break;
        }
        // resumeCollectionLockだけをセット（isCollectingはstartCollection内でセット）
        resumeCollectionLock = true;

        // backgroundからのページ遷移後の収集再開
        const currentPathIsReview = window.location.pathname.includes('/product-reviews/');
        console.log('[Amazonレビュー収集] resumeCollectionメッセージ受信', {
          isReviewPage,
          currentPathIsReview,
          isCollecting,
          startCollectionLock,
          resumeCollectionLock,
          currentUrl: window.location.href.substring(0, 80),
          currentPage: getCurrentPageNumber()
        });
        if (!currentPathIsReview) {
          console.log('[Amazonレビュー収集] レビューページではないためスキップ');
          resumeCollectionLock = false;
          sendResponse({ success: false, error: 'レビューページではありません' });
          break;
        }
        // 収集を再開
        console.log('[Amazonレビュー収集] 収集再開を開始します');
        incrementalOnly = message.incrementalOnly || false;
        lastCollectedDate = message.lastCollectedDate || null;
        currentQueueName = message.queueName || null;
        if (message.productId) {
          currentProductId = message.productId;
        }
        autoResumeExecuted = false;
        const resumePage = getCurrentPageNumber();
        log(`ページ${resumePage}の収集を再開します`);
        startCollection();
        // startCollection完了後にロック解除（非同期処理のため長めに）
        setTimeout(() => { resumeCollectionLock = false; }, 10000);
        sendResponse({ success: true });
        break;

      case 'getProductInfo':
        const productInfo = getProductInfo();
        sendResponse({ success: true, productInfo: productInfo });
        break;

      default:
        sendResponse({ success: false, error: '不明なアクション' });
    }
    return true;
  });

  /**
   * ASINを取得
   */
  function getASIN() {
    const url = window.location.href;

    // /dp/ASIN パターン
    const dpMatch = url.match(/\/dp\/([A-Z0-9]{10})/i);
    if (dpMatch) return dpMatch[1].toUpperCase();

    // /product-reviews/ASIN パターン
    const reviewMatch = url.match(/\/product-reviews\/([A-Z0-9]{10})/i);
    if (reviewMatch) return reviewMatch[1].toUpperCase();

    // /gp/product/ASIN パターン
    const gpMatch = url.match(/\/gp\/product\/([A-Z0-9]{10})/i);
    if (gpMatch) return gpMatch[1].toUpperCase();

    return '';
  }

  /**
   * 商品情報を取得（キュー追加用）
   */
  function getProductInfo() {
    let title = '';
    let url = window.location.href;
    const asin = getASIN();

    // 商品名を取得（複数のセレクターを試行）
    const titleSelectors = [
      '#productTitle',                          // 商品ページの標準
      '[data-hook="product-link"]',             // レビューページ
      '#title',                                 // 一部の商品ページ
      '.product-title-word-break',              // モバイル向け
      'h1.a-size-large',                        // 汎用
      'h1 span#productTitle',                   // 別パターン
    ];

    for (const selector of titleSelectors) {
      const elem = document.querySelector(selector);
      if (elem && elem.textContent.trim()) {
        title = elem.textContent.trim();
        console.log(`[Amazonレビュー収集] 商品名取得成功: "${title.substring(0, 50)}..." (セレクター: ${selector})`);
        break;
      }
    }

    // セレクターで取得できなかった場合、document.titleから抽出
    if (!title || title === asin) {
      const docTitle = document.title || '';
      // "Amazon.co.jp: 商品名" または "Amazon.co.jp: カスタマーレビュー: 商品名" から抽出
      let extracted = docTitle;
      extracted = extracted.replace(/^Amazon\.co\.jp[：:]\s*/i, '');
      extracted = extracted.replace(/^カスタマーレビュー[：:]\s*/i, '');
      extracted = extracted.trim();

      // 抽出した結果がASINでなければ使用
      if (extracted && extracted !== asin && extracted.length > 0) {
        title = extracted;
        console.log(`[Amazonレビュー収集] 商品名をdocument.titleから抽出: "${title.substring(0, 50)}..."`);
      }
    }

    // それでも取得できなかった場合は空文字（ASINのみ表示になる）
    if (!title || title === asin) {
      console.log(`[Amazonレビュー収集] 商品名を取得できませんでした（ASIN: ${asin}）`);
      title = '';
    }

    // レビューページの場合は商品ページURLを構築
    if (isReviewPage && asin) {
      url = `https://www.amazon.co.jp/dp/${asin}`;
    }

    return {
      url: url,
      title: title.substring(0, 100),
      addedAt: new Date().toISOString(),
      source: 'amazon'
    };
  }

  /**
   * 商品ページからレビューページへのリンク要素を取得
   * 重要: アンカーリンク（#で始まるhref）ではなく、product-reviewsページへの実リンクを探す
   */
  function findReviewPageLink() {
    const asin = getASIN();

    // ヘルパー: hrefがproduct-reviewsページへの実リンクかどうか確認
    const isValidReviewLink = (elem) => {
      if (!elem || !elem.href) return false;
      // アンカーリンク（#）やjavascript:は除外
      if (elem.href.startsWith('#') || elem.href.startsWith('javascript:')) return false;
      // product-reviewsを含む実際のURLか確認
      return elem.href.includes('/product-reviews/');
    };

    // 1. 「すべてのレビューを見る」リンク（data-hook属性）
    const reviewLink = document.querySelector(AMAZON_SELECTORS.reviewLink);
    if (isValidReviewLink(reviewLink)) {
      return reviewLink;
    }

    // 2. 「レビューをすべて見る」リンク（a.a-link-emphasis）
    const emphasisLink = document.querySelector('a.a-link-emphasis[href*="product-reviews"]');
    if (isValidReviewLink(emphasisLink)) {
      return emphasisLink;
    }

    // 3. レビューセクションのフッターリンク
    const altLink = document.querySelector('#reviews-medley-footer a[href*="product-reviews"]');
    if (isValidReviewLink(altLink)) {
      return altLink;
    }

    // 4. カスタマーレビューセクション内のリンク
    const crLink = document.querySelector('#customerReviews a[href*="product-reviews"]');
    if (isValidReviewLink(crLink)) {
      return crLink;
    }

    // 5. 現在の商品ASINのproduct-reviewsリンクを検索
    if (asin) {
      const asinReviewLink = document.querySelector(`a[href*="/product-reviews/${asin}"]`);
      if (isValidReviewLink(asinReviewLink)) {
        return asinReviewLink;
      }
    }

    // 6. ページ内のすべてのリンクから、現在の商品のレビューページリンクを探す
    const allLinks = document.querySelectorAll('a[href*="product-reviews"]');
    for (const link of allLinks) {
      // 現在の商品のASINを含むリンクを優先
      if (asin && link.href.includes(`/product-reviews/${asin}`)) {
        if (isValidReviewLink(link)) {
          return link;
        }
      }
    }

    // 7. 見つからない場合、ASINからレビューページURLを生成してダミーリンクを作成
    if (asin) {
      const link = document.createElement('a');
      link.href = `https://www.amazon.co.jp/product-reviews/${asin}/ref=cm_cr_dp_d_show_all_btm?ie=UTF8&reviewerType=all_reviews`;
      document.body.appendChild(link);
      console.log('[Amazonレビュー収集] レビューリンクが見つからないため、ASINからURLを生成:', link.href);
      return link;
    }

    return null;
  }

  /**
   * 収集を開始
   */
  async function startCollection() {
    // 重複実行防止（複数の条件で個別にチェック）
    if (isCollecting) {
      console.log('[Amazonレビュー収集] 既に収集中です（isCollecting=true）- スキップ');
      return;
    }

    if (startCollectionLock) {
      console.log('[Amazonレビュー収集] ロック中です（startCollectionLock=true）- スキップ');
      return;
    }

    // 即座に両方のフラグをロック（競合防止）
    startCollectionLock = true;

    // ===== ボット対策: レート制限チェック =====
    const rateLimit = await checkRateLimit();
    if (!rateLimit.allowed) {
      log(`本日のAmazon収集上限（100ページ）に達しました。明日再試行してください。`, 'error');
      startCollectionLock = false;
      chrome.runtime.sendMessage({ action: 'collectionStopped' });
      return;
    }
    if (rateLimit.remaining <= 10) {
      log(`注意: 本日の残りページ数は${rateLimit.remaining}ページです。`);
    }

    // ===== ボット対策: アクティブタブチェック =====
    if (!await waitForActiveTab()) {
      log('収集を中断しました（タブがアクティブではありません）', 'error');
      startCollectionLock = false;
      return;
    }

    isCollecting = true;
    shouldStop = false;
    autoResumeExecuted = true; // 収集開始したらフラグをセット
    sessionPageCount = 0; // セッションページカウントをリセット

    currentProductId = getASIN();

    // 商品名を取得してbackground.jsに通知（キュー表示の更新用）
    const productInfo = getProductInfo();
    if (productInfo.title && productInfo.title !== currentProductId) {
      chrome.runtime.sendMessage({
        action: 'updateProductTitle',
        productId: currentProductId,
        title: productInfo.title
        // 注意: URLは送信しない（キューに既に正しいURLが設定されているため）
      });
      console.log(`[Amazonレビュー収集] 商品名をbackgroundに通知: "${productInfo.title.substring(0, 50)}..."`);
    }

    // 詳細なデバッグ情報を出力
    console.log('[Amazonレビュー収集] ===== 収集開始 =====');
    console.log('[Amazonレビュー収集] URL:', window.location.href);
    console.log('[Amazonレビュー収集] pathname:', window.location.pathname);
    console.log('[Amazonレビュー収集] isReviewPage:', isReviewPage);
    console.log('[Amazonレビュー収集] isProductPage:', isProductPage);
    console.log('[Amazonレビュー収集] document.title:', document.title);
    console.log('[Amazonレビュー収集] document.readyState:', document.readyState);
    console.log('[Amazonレビュー収集] body長さ:', document.body?.innerHTML?.length || 0);

    // 現在のレビュー要素数を確認
    const initialReviews = document.querySelectorAll(AMAZON_SELECTORS.reviewContainer);
    console.log('[Amazonレビュー収集] 初期レビュー要素数:', initialReviews.length);
    console.log('[Amazonレビュー収集] セレクター:', AMAZON_SELECTORS.reviewContainer);

    // ボット検出ページかどうかチェック
    const robotCheck = document.querySelector('form[action="/errors/validateCaptcha"]');
    if (robotCheck) {
      console.log('[Amazonレビュー収集] ★★★ ボット検出ページを検出！ ★★★');
      log('Amazonがボット検出を行いました。手動で認証が必要です。', 'error');
    }

    // レビュー要素が読み込まれるまで待機（最大10秒）
    const reviewsFound = await waitForReviews(10000, 500);
    if (!reviewsFound) {
      log('レビュー要素の読み込みを待機中...');
      // 追加で5秒待機（ページ読み込みが遅い場合に対応）
      await sleep(5000);

      // 再度確認
      const reviewsAfterWait = document.querySelectorAll(AMAZON_SELECTORS.reviewContainer);
      console.log('[Amazonレビュー収集] 待機後のレビュー要素数:', reviewsAfterWait.length);
      console.log('[Amazonレビュー収集] ページタイトル:', document.title);

      // ページの主要な要素を確認
      console.log('[Amazonレビュー収集] #cm_cr-review_list存在:', !!document.querySelector('#cm_cr-review_list'));
      console.log('[Amazonレビュー収集] .review存在:', document.querySelectorAll('.review').length);
      console.log('[Amazonレビュー収集] [data-hook]要素数:', document.querySelectorAll('[data-hook]').length);
    }

    // 収集済みレビューキーとtotalPagesをストレージから復元
    const storedState = await new Promise(resolve => {
      chrome.storage.local.get(['collectionState'], result => resolve(result.collectionState || {}));
    });

    // 現在のページ番号を取得
    const currentPage = getCurrentPageNumber();
    const storedProductId = storedState.productId || storedState.currentProductId;

    // 同じ商品の継続収集の場合のみ、collectedReviewKeysを復元
    // 条件:
    //   0. 既にcollectedReviewKeysにデータがある場合は復元しない（重複実行対策）
    //   1. ストレージに保存された商品IDと現在の商品IDが一致
    //   2. ページ2以降から再開する場合（ページ1は新規収集）
    //   3. ストレージの収集済みキーが存在する
    if (collectedReviewKeys.size === 0 &&
        storedProductId === currentProductId &&
        currentPage > 1 &&
        storedState.collectedReviewKeys &&
        Array.isArray(storedState.collectedReviewKeys)) {
      collectedReviewKeys = new Set(storedState.collectedReviewKeys);
      console.log(`[Amazonレビュー収集] 継続収集: collectedReviewKeysを復元（${collectedReviewKeys.size}件）`);
    } else if (collectedReviewKeys.size > 0) {
      // 既にキーがある場合は何もしない（重複実行で上書きしない）
      console.log(`[Amazonレビュー収集] 既存キー維持: collectedReviewKeys（${collectedReviewKeys.size}件）`);
    } else {
      // 新規収集または異なる商品: collectedReviewKeysをクリア
      collectedReviewKeys = new Set();
      console.log(`[Amazonレビュー収集] 新規収集: collectedReviewKeysをクリア（ページ${currentPage}、商品: ${currentProductId}）`);
    }

    // totalPagesをストレージから復元（ページ遷移後も維持するため）
    // 同じ商品の継続収集の場合のみ
    if (storedProductId === currentProductId && storedState.totalPages && storedState.totalPages > 0) {
      totalPages = storedState.totalPages;
      console.log(`[Amazonレビュー収集] totalPagesをストレージから復元: ${totalPages}`);
    } else {
      totalPages = 0; // 新規収集の場合はリセット
    }

    // consecutiveSkipPagesをストレージから復元（無限ループ防止カウンタ）
    // 同じ商品の継続収集の場合のみ
    if (storedProductId === currentProductId && typeof storedState.consecutiveSkipPages === 'number') {
      consecutiveSkipPages = storedState.consecutiveSkipPages;
      console.log(`[Amazonレビュー収集] consecutiveSkipPagesをストレージから復元: ${consecutiveSkipPages}`);
    } else {
      consecutiveSkipPages = 0; // 新規収集の場合はリセット
    }

    // 星評価フィルター状態をストレージから復元
    // 同じ商品の継続収集の場合のみ
    if (storedProductId === currentProductId && typeof storedState.currentStarFilterIndex === 'number') {
      useStarFilter = storedState.useStarFilter !== false; // デフォルトtrue
      currentStarFilterIndex = storedState.currentStarFilterIndex;
      pagesCollectedInCurrentFilter = storedState.pagesCollectedInCurrentFilter || 0;
      console.log(`[Amazonレビュー収集] 星評価フィルター状態を復元: index=${currentStarFilterIndex}, pages=${pagesCollectedInCurrentFilter}`);
    } else {
      // 新規収集の場合はリセット
      useStarFilter = true;
      currentStarFilterIndex = 0;
      pagesCollectedInCurrentFilter = 0;
      filterPageCounts = {}; // ページ数カウントもリセット
      console.log('[Amazonレビュー収集] 新規収集: 星評価フィルターを初期化');
    }

    // 新規収集で星評価フィルターが有効な場合、最初のフィルターを適用
    // （ページ1かつフィルターが適用されていない場合のみ）
    if (useStarFilter && currentPage === 1 && !window.location.href.includes('filterByStar')) {
      const firstFilter = STAR_FILTERS[currentStarFilterIndex];
      log(`星評価フィルター収集を開始: ${firstFilter.label}から（最大500件）`);
      await navigateToFilteredPage(currentStarFilterIndex);
      return; // ページ遷移するので、以降の処理は遷移後に行う
    }

    // ===== ボット対策: ページ読み込み後の「見渡し」動作 =====
    await initialPageScan();
    if (shouldStop) {
      isCollecting = false;
      startCollectionLock = false;
      return;
    }

    // レビュー総数を取得
    const expectedTotal = getTotalReviewCount();

    // 星フィルター情報を取得（ログ表示用）
    const currentFilter = useStarFilter && currentStarFilterIndex < STAR_FILTERS.length
      ? STAR_FILTERS[currentStarFilterIndex].label
      : null;

    if (expectedTotal > 0) {
      chrome.storage.local.set({ expectedReviewTotal: expectedTotal });
      totalPages = Math.ceil(expectedTotal / 10); // Amazonは1ページ10件

      // 星フィルター適用時は詳細ログ
      if (currentFilter) {
        log(`${currentFilter}レビュー収集開始（${expectedTotal.toLocaleString()}件、${totalPages}ページ）`);
      } else if (incrementalOnly && lastCollectedDate) {
        log(`差分収集を開始します（前回: ${lastCollectedDate}、全${expectedTotal.toLocaleString()}件中新着のみ）`);
      } else {
        log(`レビュー収集を開始します（全${expectedTotal.toLocaleString()}件、${totalPages}ページ）`);
      }
    } else {
      chrome.storage.local.set({ expectedReviewTotal: 0 });
      // 重要: ストレージから復元したtotalPagesを上書きしない
      // （ページ2以降でgetTotalReviewCount()が0を返す場合に対応）
      if (totalPages === 0) {
        // 新規収集でtotalPagesが不明な場合のみログを出力
        if (currentFilter) {
          log(`${currentFilter}レビュー収集開始`);
        } else if (incrementalOnly && lastCollectedDate) {
          log(`差分収集を開始します（前回: ${lastCollectedDate}）`);
        } else {
          log('レビュー収集を開始します');
        }
      } else {
        // ストレージから復元したtotalPagesを維持
        console.log(`[Amazonレビュー収集] getTotalReviewCount()は0を返しましたが、ストレージから復元したtotalPages(${totalPages})を維持します`);
        if (currentFilter) {
          log(`${currentFilter}レビュー収集継続（${totalPages}ページ）`);
        } else if (incrementalOnly && lastCollectedDate) {
          log(`差分収集を継続します（前回: ${lastCollectedDate}、全${totalPages * 10}件程度）`);
        } else {
          log(`レビュー収集を継続します（全${totalPages * 10}件程度）`);
        }
      }
    }

    // 現在のページからレビューを収集
    const reachedOldReviews = await collectCurrentPage();

    if (reachedOldReviews) {
      log('前回以降の新着レビューの収集が完了しました');
      isCollecting = false;
      if (!shouldStop) {
        chrome.runtime.sendMessage({ action: 'collectionComplete' });
      }
      return;
    }

    // ページネーションがある場合、次のページも収集
    const navigated = await collectAllPages();

    if (navigated) {
      return;
    }

    isCollecting = false;

    if (!shouldStop) {
      chrome.runtime.sendMessage({ action: 'collectionComplete' });
    }
  }

  /**
   * レビューの一意キーを生成
   */
  function generateReviewKey(review) {
    // レビューIDがあればそれを使用（最も確実な一意キー）
    if (review.reviewId) {
      return `amazon_${review.reviewId}`;
    }
    // フォールバック: 本文の先頭100文字 + 著者 + 日付でユニークキーを生成
    return `${(review.body || '').substring(0, 100)}|${review.author || ''}|${review.reviewDate || ''}`;
  }

  /**
   * 現在のページからレビューを収集
   */
  async function collectCurrentPage() {
    // ===== ボット対策: 各レビューを「読む」動作をシミュレート =====
    await simulateReadingReviews();
    if (shouldStop) return false;

    let reviews = extractReviews();
    const currentPage = getCurrentPageNumber();

    if (reviews.length === 0) {
      log('このページにレビューが見つかりませんでした', 'error');
      return false;
    }

    // 収集済みレビューをフィルタリング（セッション内重複チェック）
    const originalCount = reviews.length;
    reviews = reviews.filter(review => {
      const key = generateReviewKey(review);
      if (collectedReviewKeys.has(key)) {
        return false; // 既に収集済み
      }
      return true;
    });

    const skippedBySession = originalCount - reviews.length;
    if (skippedBySession > 0) {
      log(`${skippedBySession}件は既にこのセッションで収集済みのためスキップ`);
    }

    // すべて収集済みの場合
    if (reviews.length === 0) {
      consecutiveSkipPages++;
      const currentPage = getCurrentPageNumber();

      // 連続スキップが上限に達した場合、または最終ページに近い場合は収集完了
      if (consecutiveSkipPages >= MAX_CONSECUTIVE_SKIP) {
        log(`${consecutiveSkipPages}ページ連続でスキップ - 収集完了とします`);
        return true;
      }

      // まだページが残っている場合は次のページに進む
      if (totalPages > 0 && currentPage < totalPages) {
        log(`このページのレビューは全て収集済みです（${currentPage}/${totalPages}ページ）- 次のページに進みます`);
        return false; // 次のページに進む
      }

      // totalPagesが0の場合（フィルター遷移直後など）、「次へ」リンクで判断
      if (totalPages === 0) {
        const nextLink = findNextPageLink();
        if (nextLink) {
          log(`このページのレビューは全て収集済みです（ページ${currentPage}）- 次のページに進みます`);
          return false; // 次のページに進む
        }
      }

      // 最終ページの場合は収集完了
      log('このページのレビューは全て収集済みです（最終ページ）');
      return true;
    }

    // レビューが収集できた場合、連続スキップカウンタをリセット
    consecutiveSkipPages = 0;

    // 差分取得モードの場合、日付でフィルタリング
    let reachedOldReviews = false;
    if (incrementalOnly && lastCollectedDate) {
      const beforeDateFilter = reviews.length;
      reviews = filterReviewsByDate(reviews, lastCollectedDate);
      const filteredCount = beforeDateFilter - reviews.length;

      if (filteredCount > 0) {
        log(`${beforeDateFilter}件中${filteredCount}件は前回収集済み（${lastCollectedDate}より前）`);
      }

      if (reviews.length === 0) {
        log('前回以降の新着レビューがこのページにはありません');
        reachedOldReviews = true;
        return reachedOldReviews;
      }

      if (filteredCount > 0 && reviews.length < beforeDateFilter) {
        reachedOldReviews = true;
      }
    }

    log(`${reviews.length}件のレビューを検出`);

    // 収集したレビューのキーを追跡に追加
    reviews.forEach(review => {
      const key = generateReviewKey(review);
      collectedReviewKeys.add(key);
    });

    // バックグラウンドにデータを送信
    await new Promise((resolve) => {
      chrome.runtime.sendMessage({
        action: 'saveReviews',
        reviews: reviews,
        source: 'amazon'
      }, (response) => {
        resolve(response);
      });
    });

    await updateState(reviews.length);

    return reachedOldReviews;
  }

  /**
   * すべてのページを収集
   */
  async function collectAllPages() {
    while (!shouldStop) {
      // ===== ボット対策: アクティブタブチェック =====
      if (!await waitForActiveTab()) {
        log('収集を中断しました（タブがアクティブではありません）', 'error');
        return false;
      }

      // ===== ボット対策: セッション制限チェック =====
      sessionPageCount++;
      if (sessionPageCount >= MAX_PAGES_PER_SESSION) {
        log(`セッション制限（${MAX_PAGES_PER_SESSION}ページ）に達しました。収集を一時停止します。`, 'error');
        log('続きを収集する場合は、再度「収集開始」を押してください。');
        isCollecting = false;
        chrome.runtime.sendMessage({ action: 'collectionStopped' });
        return false;
      }

      // ===== ボット対策: ランダムなマイクロブレイク（確率ベース） =====
      await maybeHaveBreak();

      // 休憩後もアクティブタブチェック
      if (!await waitForActiveTab()) {
        return false;
      }

      // ===== ボット対策: レート制限チェック =====
      const rateLimit = await checkRateLimit();
      if (!rateLimit.allowed) {
        log(`本日のAmazon収集上限（100ページ）に達しました。明日再試行してください。`, 'error');
        isCollecting = false;
        chrome.runtime.sendMessage({ action: 'collectionStopped' });
        return false;
      }

      // 「次へ」リンクの存在確認
      const canGoNext = hasNextPage();

      // 星評価フィルター: 現在のフィルターでのページ数をインクリメント
      if (useStarFilter) {
        pagesCollectedInCurrentFilter++;
        const currentFilter = getCurrentStarFilter();
        console.log(`[Amazonレビュー収集] フィルター ${currentFilter?.label || '不明'}: ${pagesCollectedInCurrentFilter}/${MAX_PAGES_PER_FILTER}ページ収集`);
      }

      // 10ページ到達または最終ページの場合、次のフィルターに切り替え
      if (!canGoNext || (useStarFilter && pagesCollectedInCurrentFilter >= MAX_PAGES_PER_FILTER)) {
        if (useStarFilter) {
          const currentFilter = getCurrentStarFilter();

          // ページ数を記録（検証用）
          filterPageCounts[currentFilter?.label || '不明'] = pagesCollectedInCurrentFilter;

          // Amazonの実際のページ制限を報告
          log(`【ページ制限検証】${currentFilter?.label || ''}フィルター: ${pagesCollectedInCurrentFilter}ページで終了（Amazonの制限）`);
          console.log(`[ページ制限検証] ${currentFilter?.label || ''}: ${pagesCollectedInCurrentFilter}ページ`);

          // 次のフィルターに切り替え
          if (switchToNextStarFilter()) {
            // 次のフィルターページに遷移
            await navigateToFilteredPage(currentStarFilterIndex);
            return true; // ページ遷移中
          } else {
            // 全フィルター完了 - サマリーを表示
            log('全ての星評価フィルター（★1〜★5）での収集が完了しました');
            log('【ページ制限検証サマリー】');
            let totalPages = 0;
            for (const [filter, pages] of Object.entries(filterPageCounts)) {
              log(`  ${filter}: ${pages}ページ`);
              totalPages += pages;
            }
            log(`  合計: ${totalPages}ページ`);
            console.log('[ページ制限検証] サマリー:', filterPageCounts);
            return false;
          }
        } else {
          log(`最後のページに到達しました（${sessionPageCount}ページ収集）`);
          return false;
        }
      }

      // 人間らしくスクロールして「次へ」ボタンまで移動
      const scrolled = await scrollToNextButton();
      if (!scrolled || shouldStop) {
        return false;
      }

      // ===== ボット対策: ページ遷移前の待機（指数分布、読み終わったらすぐ次へ） =====
      // 人間は読み終わったらすぐ次のページへ移動する
      const waitTime = Math.min(exponentialRandom(PAGE_WAIT_MEAN_MS), PAGE_WAIT_MAX_MS);
      const finalWait = Math.max(500, waitTime); // 最低0.5秒
      log(`次のページへ移動（${Math.round(finalWait / 1000)}秒後）...`);
      await sleep(finalWait);

      // 待機後もアクティブタブチェック
      if (!await waitForActiveTab()) {
        return false;
      }

      // ===== ボット対策: レート制限カウントを増加 =====
      const updated = await incrementRateLimit();
      if (updated.remaining <= 10 && updated.remaining > 0) {
        log(`注意: 本日の残りページ数は${updated.remaining}ページです。`);
      }

      // 差分取得設定と収集済みキーを保存してからページ遷移
      await new Promise((resolve) => {
        chrome.storage.local.get(['collectionState'], (result) => {
          const state = result.collectionState || {};
          state.isRunning = true; // 重要: 収集中フラグを維持
          state.incrementalOnly = incrementalOnly;
          state.lastCollectedDate = lastCollectedDate;
          state.source = 'amazon';
          state.queueName = currentQueueName;
          // 商品IDを保存（継続収集の判定用）
          state.productId = currentProductId;
          // 収集済みレビューキーを保存（次のページでも重複チェックできるように）
          state.collectedReviewKeys = Array.from(collectedReviewKeys);
          state.lastProcessedUrl = window.location.href;
          state.lastProcessedPage = getCurrentPageNumber();
          // 総ページ数を保存（重要：ページ遷移後も維持するため）
          state.totalPages = totalPages;
          // セッションページカウントも保存
          state.sessionPageCount = sessionPageCount;
          // 連続スキップカウントも保存
          state.consecutiveSkipPages = consecutiveSkipPages;
          // 星評価フィルター状態を保存
          state.useStarFilter = useStarFilter;
          state.currentStarFilterIndex = currentStarFilterIndex;
          state.pagesCollectedInCurrentFilter = pagesCollectedInCurrentFilter;
          console.log('[Amazonレビュー収集] 次ページ遷移前の状態保存:', JSON.stringify(state, null, 2));
          chrome.storage.local.set({ collectionState: state }, resolve);
        });
      });

      log('次のページに移動します');

      // 次のページ遷移のために状態をリセット（同一オリジン遷移でスクリプトが再注入されない場合に対応）
      isCollecting = false;
      startCollectionLock = false;
      autoResumeExecuted = false;

      // クリックでページ遷移（人間らしいマウス移動＋クリック）
      await clickNextPage();
      return true; // ページ遷移中
    }

    return false;
  }

  /**
   * 「次へ」リンクを探す（複数の方法でフォールバック）
   * @returns {HTMLElement|null} 見つかった「次へ」リンク、または null
   */
  function findNextPageLink() {
    const currentPage = getCurrentPageNumber();
    const nextPage = currentPage + 1;

    // 方法1: 標準のセレクタ
    const standardSelectors = [
      'li.a-last a',
      '.a-pagination li.a-last a',
    ];

    for (const selector of standardSelectors) {
      try {
        const link = document.querySelector(selector);
        if (link && link.href && link.href.includes('pageNumber')) {
          console.log(`[Amazonレビュー収集] findNextPageLink: セレクタ "${selector}" で発見`);
          return link;
        }
      } catch (e) {
        // セレクタエラーは無視
      }
    }

    // 方法2: 「次へ」テキストを含むリンクを探す（最も確実）
    const allLinks = document.querySelectorAll('a[href*="pageNumber"]');
    for (const link of allLinks) {
      const text = link.textContent || '';
      if (text.includes('次へ') || text.includes('次') || text.includes('→')) {
        console.log(`[Amazonレビュー収集] findNextPageLink: 「次へ」テキストで発見: "${text.trim().substring(0, 20)}"`);
        return link;
      }
    }

    // 方法3: 次ページ番号のリンクを探す
    for (const link of allLinks) {
      try {
        const url = new URL(link.href, window.location.origin);
        const pageNum = parseInt(url.searchParams.get('pageNumber'), 10);
        if (pageNum === nextPage) {
          console.log(`[Amazonレビュー収集] findNextPageLink: ページ番号${nextPage}のリンクで発見`);
          return link;
        }
      } catch (e) {
        // URL解析エラーは無視
      }
    }

    // 方法4: ページネーション内のリンクからテキストで探す
    const paginationLinks = document.querySelectorAll('.a-pagination a, .a-pagination li a');
    for (const link of paginationLinks) {
      const linkText = link.textContent.trim();
      const linkPage = parseInt(linkText, 10);
      if (linkPage === nextPage) {
        console.log(`[Amazonレビュー収集] findNextPageLink: ページネーションの「${linkText}」で発見`);
        return link;
      }
    }

    console.log('[Amazonレビュー収集] findNextPageLink: 次へリンクが見つかりません');
    return null;
  }

  /**
   * 次のページがあるかチェック
   * セレクタだけでなく、ページ番号と総ページ数でも判定
   */
  function hasNextPage() {
    const currentPage = getCurrentPageNumber();

    // 総ページ数が分かっている場合、ページ番号で判定（最も確実）
    if (totalPages > 0 && currentPage < totalPages) {
      console.log(`[Amazonレビュー収集] hasNextPage: ページ${currentPage}/${totalPages} - 次のページあり（ページ番号判定）`);
      return true;
    }

    // 最後のページかどうか確認（無効化された次へボタン）
    const isLastPage = document.querySelector(AMAZON_SELECTORS.isLastPage);
    if (isLastPage) {
      console.log('[Amazonレビュー収集] hasNextPage: li.a-last.a-disabled検出 - 最後のページ');
      return false;
    }

    // 「次へ」リンクを探す
    const nextLink = findNextPageLink();
    if (nextLink) {
      return true;
    }

    // 総ページ数が分かっていて、まだ到達していない場合は次のページありとする
    if (totalPages > 0 && currentPage < totalPages) {
      console.log(`[Amazonレビュー収集] hasNextPage: リンク見つからないが、ページ${currentPage}/${totalPages} - 次のページあり`);
      return true;
    }

    console.log(`[Amazonレビュー収集] hasNextPage: 次のページなし（ページ${currentPage}/${totalPages}）`);
    return false;
  }

  /**
   * 次のページに遷移
   * 重要: Amazonはキャッシュ制御にセッション情報を使用しているため、
   * URL直接操作（pageNumberパラメータのみ変更）では正しいページが表示されない。
   * 「次へ」ボタンを実際にクリックすることで、Amazonの内部ルーティングが正しく動作する。
   *
   * 検証結果（2026年1月）：
   * - URL直接操作 → ページ1のキャッシュが表示される（NG）
   * - element.click() → 正しくページ2に遷移（OK）
   */
  async function clickNextPage() {
    const currentPage = getCurrentPageNumber();
    const nextPage = currentPage + 1;

    console.log(`[Amazonレビュー収集] 次のページに遷移: ページ${currentPage} → ページ${nextPage}`);

    // 「次へ」ボタンを探す
    const nextLink = findNextPageLink();

    if (nextLink) {
      console.log(`[Amazonレビュー収集] 「次へ」ボタンをクリックします: ${nextLink.href}`);
      nextLink.click();
      return;
    }

    // フォールバック: 「次へ」ボタンが見つからない場合はhref属性から遷移を試みる
    console.log('[Amazonレビュー収集] 「次へ」ボタンが見つからないため、ページネーションリンクを探します');
    const paginationLink = document.querySelector('li.a-last a');
    if (paginationLink) {
      const href = paginationLink.getAttribute('href');
      if (href) {
        const url = new URL(href, window.location.origin);
        console.log(`[Amazonレビュー収集] ページネーションリンクで遷移: ${url.toString()}`);
        window.location.href = url.toString();
        return;
      }
    }

    // 最終フォールバック: URL直接操作（動作しない可能性がある）
    console.log('[Amazonレビュー収集] フォールバック: URL直接操作');
    const url = new URL(window.location.href);
    url.searchParams.set('pageNumber', nextPage.toString());
    window.location.href = url.toString();
  }

  // ===== 星評価フィルター関連関数 =====

  /**
   * 現在の星評価フィルター情報を取得
   */
  function getCurrentStarFilter() {
    if (!useStarFilter || currentStarFilterIndex >= STAR_FILTERS.length) {
      return null;
    }
    return STAR_FILTERS[currentStarFilterIndex];
  }

  /**
   * 星評価フィルターを適用したURLを生成
   * @param {string} baseUrl - 基本URL
   * @param {string} filterValue - フィルター値（例: 'one_star'）
   * @returns {string} フィルター適用後のURL
   */
  function buildFilteredUrl(baseUrl, filterValue) {
    const url = new URL(baseUrl);
    // 星評価フィルター
    url.searchParams.set('filterByStar', filterValue);
    // 認証済み購入のみ
    url.searchParams.set('reviewerType', 'avp_only_reviews');
    // トップレビュー順（役に立った順）
    url.searchParams.set('sortBy', 'helpful');
    // ページ1から開始
    url.searchParams.set('pageNumber', '1');
    return url.toString();
  }

  /**
   * 次の星評価フィルターに切り替え
   * @returns {boolean} 切り替え成功したか（まだフィルターが残っているか）
   */
  function switchToNextStarFilter() {
    currentStarFilterIndex++;
    pagesCollectedInCurrentFilter = 0;

    if (currentStarFilterIndex >= STAR_FILTERS.length) {
      log('全ての星評価フィルターでの収集が完了しました');
      return false;
    }

    const nextFilter = STAR_FILTERS[currentStarFilterIndex];
    log(`次のフィルターに切り替え: ${nextFilter.label}`);
    return true;
  }

  /**
   * 星評価フィルターを適用してページに遷移
   * @param {number} filterIndex - フィルターインデックス（0-4）
   */
  async function navigateToFilteredPage(filterIndex) {
    if (filterIndex >= STAR_FILTERS.length) {
      return false;
    }

    const filter = STAR_FILTERS[filterIndex];
    const asin = getASIN();
    const baseUrl = `https://www.amazon.co.jp/product-reviews/${asin}`;
    const filteredUrl = buildFilteredUrl(baseUrl, filter.value);

    log(`${filter.label}フィルターを適用してページに遷移します`);

    // 状態を保存してから遷移
    await new Promise((resolve) => {
      chrome.storage.local.get(['collectionState'], (result) => {
        const state = result.collectionState || {};
        state.isRunning = true;
        state.incrementalOnly = incrementalOnly;
        state.lastCollectedDate = lastCollectedDate;
        state.source = 'amazon';
        state.queueName = currentQueueName;
        state.productId = currentProductId;
        state.collectedReviewKeys = Array.from(collectedReviewKeys);
        // 星評価フィルター状態を保存
        state.useStarFilter = useStarFilter;
        state.currentStarFilterIndex = filterIndex;
        state.pagesCollectedInCurrentFilter = 0;
        // フィルター遷移フラグを設定（background.jsで検出するため）
        state.filterTransitionPending = true;
        state.lastProcessedPage = 0; // ページ番号リセット（新しいフィルターでページ1から開始）
        state.totalPages = 0; // 総ページ数リセット（新しいフィルターで再取得）
        state.consecutiveSkipPages = 0; // 連続スキップカウンタをリセット（重要！）
        console.log(`[Amazonレビュー収集] フィルター遷移前の状態保存: filterIndex=${filterIndex}`);
        chrome.storage.local.set({ collectionState: state }, resolve);
      });
    });

    // 状態をリセット
    isCollecting = false;
    startCollectionLock = false;
    autoResumeExecuted = false;
    totalPages = 0; // フィルター切り替え時に総ページ数をリセット
    consecutiveSkipPages = 0; // 連続スキップカウンタをリセット

    // ページ遷移
    window.location.href = filteredUrl;
    return true;
  }

  /**
   * 要素がビューポート内にあるか確認
   */
  function isElementInViewport(el) {
    const rect = el.getBoundingClientRect();
    return (
      rect.top >= 0 &&
      rect.bottom <= window.innerHeight
    );
  }

  /**
   * 人間らしくスクロールして「次へ」ボタンまで移動
   * - 加速・減速パターン
   * - 途中で2-4回停止（レビューを読んでいるように見せる）
   * - 時々戻り読み
   * - 「次へ」ボタンが見えたら停止
   * - ボタンが見つからない場合はページ下部までスクロール
   */
  async function scrollToNextButton() {
    // findNextPageLink()を使用して「次へ」リンクを探す
    const nextButton = findNextPageLink();

    // ボタンが見つからない場合もページ下部までスクロール（人間らしい動作のため）
    const hasButton = !!nextButton;
    if (!hasButton) {
      console.log('[Amazonレビュー収集] 次へボタンが見つかりません - ページ下部までスクロールします');
    }

    // 途中停止回数（2-4回）
    const pauseCount = 2 + Math.floor(Math.random() * 3);

    // ページの高さとスクロール位置
    const pageHeight = document.body.scrollHeight;
    const viewportHeight = window.innerHeight;
    const maxScroll = pageHeight - viewportHeight;

    // 停止ポイントを計算
    const pausePoints = [];
    for (let i = 1; i <= pauseCount; i++) {
      pausePoints.push(maxScroll * (i / (pauseCount + 1)));
    }

    let pauseIndex = 0;
    let lastScrollY = window.scrollY;
    let loopCount = 0;
    const maxLoops = 50; // 無限ループ防止

    // ボタンがある場合はボタンが見えるまで、ない場合はページ下部まで
    while (loopCount < maxLoops) {
      loopCount++;
      if (shouldStop) return false;

      // ボタンがあり、見えている場合は終了
      if (hasButton && isElementInViewport(nextButton)) {
        break;
      }

      // ボタンがない場合、ページ下部に到達したら終了
      if (!hasButton && window.scrollY >= maxScroll - 50) {
        break;
      }

      // ターゲット位置を計算
      let targetY;
      if (hasButton) {
        const buttonRect = nextButton.getBoundingClientRect();
        const buttonAbsY = window.scrollY + buttonRect.top - viewportHeight + 100;
        if (pauseIndex < pausePoints.length && pausePoints[pauseIndex] < buttonAbsY) {
          targetY = pausePoints[pauseIndex];
        } else {
          targetY = buttonAbsY;
        }
      } else {
        // ボタンがない場合は次の停止ポイントまたはページ下部
        if (pauseIndex < pausePoints.length) {
          targetY = pausePoints[pauseIndex];
        } else {
          targetY = maxScroll;
        }
      }

      // 人間らしいスクロール（加速・減速・時々戻り読み）
      await humanScroll(targetY, { maxDuration: 2000 + Math.random() * 1000 });

      // 停止ポイントに到達したら一時停止
      if (pauseIndex < pausePoints.length && window.scrollY >= pausePoints[pauseIndex] - 50) {
        const pauseTime = 800 + Math.random() * 1500; // 0.8-2.3秒
        await sleep(pauseTime);
        pauseIndex++;
      }

      // 無限ループ防止
      if (Math.abs(window.scrollY - lastScrollY) < 10) {
        // スクロールが進んでいない場合は強制スクロール
        window.scrollTo({ top: window.scrollY + 100, behavior: 'smooth' });
        await sleep(200);
      }
      lastScrollY = window.scrollY;
    }

    // 少し待機
    await sleep(300 + Math.random() * 400);
    return true; // ボタンの有無に関わらずスクロール完了
  }

  /**
   * レビューを抽出
   */
  function extractReviews() {
    const reviews = [];
    const asin = getASIN();

    // 商品名を取得
    let productName = '商品名不明';
    const titleElem = document.querySelector(AMAZON_SELECTORS.productTitle);
    const reviewPageTitleElem = document.querySelector(AMAZON_SELECTORS.reviewPageTitle);
    if (titleElem) {
      productName = titleElem.textContent.trim();
    } else if (reviewPageTitleElem) {
      // レビューページでの商品名（data-hook="product-link"）
      productName = reviewPageTitleElem.textContent.trim();
    } else {
      // フォールバック: ページタイトルから取得
      const pageTitle = document.title;
      if (pageTitle.includes('のカスタマーレビュー')) {
        productName = pageTitle.replace(/のカスタマーレビュー.*$/, '').replace(/^Amazon.*?: /, '');
      }
    }

    // 商品URL
    const productUrl = `https://www.amazon.co.jp/dp/${asin}`;

    // レビュー要素を取得
    const reviewElements = document.querySelectorAll(AMAZON_SELECTORS.reviewContainer);

    reviewElements.forEach((elem, index) => {
      try {
        const review = extractReviewData(elem, productName, productUrl, asin);
        if (review) {
          reviews.push(review);
        }
      } catch (error) {
        console.error(`レビュー ${index + 1} の抽出エラー:`, error);
      }
    });

    return removeDuplicates(reviews);
  }

  /**
   * 個別レビューデータを抽出
   */
  function extractReviewData(elem, productName, productUrl, asin) {
    let rating = 0;
    let reviewDate = '';
    let author = '匿名';
    let body = '';
    let title = '';
    let helpfulCount = 0;
    let variation = '';
    let isVerified = false;  // 認証購入
    let isVine = false;      // Vineレビュー
    let hasImage = false;    // 画像あり
    let country = '日本'; // デフォルトは日本

    // レビューIDを取得（重複判定に使用）
    const reviewId = elem.id || '';

    // 評価を取得
    const ratingElem = elem.querySelector(AMAZON_SELECTORS.rating);
    if (ratingElem) {
      const ratingText = ratingElem.textContent || '';
      // 「5つ星のうち4.0」から評価値（最後の数値）を抽出
      // 注意: 最初の「5」ではなく、「うち」の後の数値を取得する
      const ratingMatch = ratingText.match(/(?:のうち|of\s*)(\d+(?:\.\d+)?)/);
      if (ratingMatch) {
        rating = Math.round(parseFloat(ratingMatch[1]));
      } else {
        // フォールバック: 全ての数値を取得して最後のものを使用
        const allNumbers = ratingText.match(/\d+(?:\.\d+)?/g);
        if (allNumbers && allNumbers.length > 0) {
          rating = Math.round(parseFloat(allNumbers[allNumbers.length - 1]));
        }
      }
    }

    // タイトルを取得
    const titleElem = elem.querySelector(AMAZON_SELECTORS.title);
    if (titleElem) {
      title = titleElem.textContent.trim();
    }

    // 本文を取得（ビデオプレイヤー要素を除外）
    const bodyElem = elem.querySelector(AMAZON_SELECTORS.body);
    if (bodyElem) {
      // クローンを作成してビデオプレイヤー関連要素を除去
      const bodyClone = bodyElem.cloneNode(true);

      // 除去する要素のセレクター（ビデオプレイヤー、スクリプト、スタイル等）
      const removeSelectors = [
        'script',
        'style',
        'video',
        '[class*="vse"]',
        '[class*="video"]',
        '[class*="player"]',
        '[data-video]',
        '.a-icon-alt'  // 星評価のalt text
      ];

      removeSelectors.forEach(selector => {
        const elements = bodyClone.querySelectorAll(selector);
        elements.forEach(el => el.remove());
      });

      // JSON形式のテキストを除去してからテキストを取得
      let cleanedText = bodyClone.textContent.trim();

      // JSONデータ（{で始まり}で終わる長い文字列）を除去
      cleanedText = cleanedText.replace(/\{[^{}]{100,}\}/g, '');

      // ビデオプレイヤーUIテキストのパターンを除去
      cleanedText = cleanedText.replace(/Loaded:\s*[\d.]+%/g, '');
      cleanedText = cleanedText.replace(/Stream Type LIVE.*?全画面表示/g, '');
      cleanedText = cleanedText.replace(/This is a modal window\./g, '');
      cleanedText = cleanedText.replace(/\d+:\d+/g, ''); // タイムスタンプ
      cleanedText = cleanedText.replace(/\d+x/g, ''); // 再生速度
      cleanedText = cleanedText.replace(/Remaining Time -[\d:]+/g, '');

      // 連続する空白を1つに
      cleanedText = cleanedText.replace(/\s+/g, ' ').trim();

      body = cleanedText;
    }

    // 投稿者を取得
    const authorElem = elem.querySelector(AMAZON_SELECTORS.author);
    if (authorElem) {
      author = authorElem.textContent.trim();
    }

    // 日付と国を取得（「2024年1月5日に日本でレビュー済み」形式）
    const dateElem = elem.querySelector(AMAZON_SELECTORS.date);
    if (dateElem) {
      const dateText = dateElem.textContent || '';

      // 日付を抽出
      const dateMatch = dateText.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
      if (dateMatch) {
        const year = dateMatch[1];
        const month = dateMatch[2].padStart(2, '0');
        const day = dateMatch[3].padStart(2, '0');
        reviewDate = `${year}/${month}/${day}`;
      }

      // 国を抽出
      const countryMatch = dateText.match(/に(.+?)でレビュー/);
      if (countryMatch) {
        country = countryMatch[1];
      }
    }

    // 日本以外のレビューはスキップ（海外レビューは収集対象外）
    if (country !== '日本') {
      console.log(`[Amazonレビュー収集] 海外レビューをスキップ: ${country}`);
      return null;
    }

    // バリエーションを取得
    const variationElem = elem.querySelector(AMAZON_SELECTORS.variation);
    if (variationElem) {
      variation = variationElem.textContent.trim();
    }

    // 参考になった数を取得
    const helpfulElem = elem.querySelector(AMAZON_SELECTORS.helpful);
    if (helpfulElem) {
      const helpfulText = helpfulElem.textContent || '';
      // 「123人のお客様がこれが役に立ったと考えています」形式
      const helpfulMatch = helpfulText.match(/(\d+)人/);
      if (helpfulMatch) {
        helpfulCount = parseInt(helpfulMatch[1], 10);
      }
    }

    // 認証購入を確認
    const verifiedElem = elem.querySelector(AMAZON_SELECTORS.verified);
    if (verifiedElem) {
      isVerified = true;
    }

    // Vineレビューを確認
    const vineElem = elem.querySelector(AMAZON_SELECTORS.vine);
    if (vineElem && vineElem.textContent && vineElem.textContent.includes('Vine')) {
      isVine = true;
    }

    // 画像があるか確認
    const imageElem = elem.querySelector(AMAZON_SELECTORS.image);
    if (imageElem) {
      hasImage = true;
    }

    // レビューがない場合はスキップ
    if (!body && !title) {
      return null;
    }

    // 本文が短すぎる場合もスキップ
    if (body.length < 5 && !title) {
      return null;
    }

    // Amazon用17項目（reviewId追加）
    return {
      reviewId: reviewId,        // レビューID（重複判定用）
      reviewDate: reviewDate,
      productId: asin,
      productName: productName,
      productUrl: productUrl,
      rating: rating,
      title: title,
      body: body,
      author: author,
      variation: variation,
      helpfulCount: helpfulCount,
      country: country,
      isVerified: isVerified,
      isVine: isVine,
      hasImage: hasImage,
      pageUrl: window.location.href,
      collectedAt: new Date().toISOString(),
      source: 'amazon'
    };
  }

  /**
   * 重複を除去
   */
  function removeDuplicates(reviews) {
    const seen = new Set();
    return reviews.filter(review => {
      // レビューIDがあればそれを使用（最も確実）
      const key = review.reviewId
        ? `amazon_${review.reviewId}`
        : `${review.body.substring(0, 100)}${review.author}${review.reviewDate}`;
      if (seen.has(key)) {
        return false;
      }
      seen.add(key);
      return true;
    });
  }

  /**
   * 日付を比較可能な形式（YYYY-MM-DD）に正規化
   */
  function normalizeDateString(dateStr) {
    if (!dateStr) return null;

    // YYYY/MM/DD 形式
    const slashMatch = dateStr.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
    if (slashMatch) {
      const year = slashMatch[1];
      const month = slashMatch[2].padStart(2, '0');
      const day = slashMatch[3].padStart(2, '0');
      return `${year}-${month}-${day}`;
    }

    // YYYY-MM-DD 形式
    const dashMatch = dateStr.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (dashMatch) {
      return dateStr.substring(0, 10);
    }

    return null;
  }

  /**
   * レビューを日付でフィルタリング
   */
  function filterReviewsByDate(reviews, afterDate) {
    if (!afterDate) return reviews;

    const normalizedAfterDate = normalizeDateString(afterDate);
    if (!normalizedAfterDate) return reviews;

    return reviews.filter(review => {
      const reviewDateNorm = normalizeDateString(review.reviewDate);
      if (!reviewDateNorm) {
        return true; // 日付がないレビューは含める
      }
      return reviewDateNorm >= normalizedAfterDate;
    });
  }

  /**
   * 現在のページ番号を取得
   */
  function getCurrentPageNumber() {
    const url = new URL(window.location.href);
    return parseInt(url.searchParams.get('pageNumber') || '1', 10);
  }

  /**
   * レビュー総数を取得
   * フィルター適用時も対応
   */
  function getTotalReviewCount() {
    // 1. 標準セレクターで取得を試みる
    const totalElem = document.querySelector(AMAZON_SELECTORS.totalReviews);
    if (totalElem) {
      const text = totalElem.textContent || '';

      // パターン1: 「1,234件のグローバルレーティング」「1,234件中」
      const matchKen = text.match(/([\d,]+)\s*件/);
      if (matchKen) {
        const count = parseInt(matchKen[1].replace(/,/g, ''), 10);
        console.log(`[Amazonレビュー収集] getTotalReviewCount: ${count}件（セレクター: totalReviews - 件パターン）`);
        return count;
      }

      // パターン2: 「1,164一致するカスタマーレビュー」（フィルター適用時）
      const matchItchi = text.match(/([\d,]+)\s*一致/);
      if (matchItchi) {
        const count = parseInt(matchItchi[1].replace(/,/g, ''), 10);
        console.log(`[Amazonレビュー収集] getTotalReviewCount: ${count}件（セレクター: totalReviews - 一致パターン）`);
        return count;
      }
    }

    // 2. フィルター適用時の別要素を試す
    // 「1-10件目 (全50件)」のようなテキストを含む要素
    const filterInfoElems = document.querySelectorAll('[data-hook*="filter"], .a-size-base');
    for (const elem of filterInfoElems) {
      const text = elem.textContent || '';
      // 「全XX件」「XX件中」パターン
      const matchKen = text.match(/(?:全|合計)?[\s]*([\d,]+)\s*件/);
      if (matchKen) {
        const count = parseInt(matchKen[1].replace(/,/g, ''), 10);
        if (count > 0) {
          console.log(`[Amazonレビュー収集] getTotalReviewCount: ${count}件（フォールバック - 件パターン）`);
          return count;
        }
      }
      // 「1,164一致」パターン
      const matchItchi = text.match(/([\d,]+)\s*一致/);
      if (matchItchi) {
        const count = parseInt(matchItchi[1].replace(/,/g, ''), 10);
        if (count > 0) {
          console.log(`[Amazonレビュー収集] getTotalReviewCount: ${count}件（フォールバック - 一致パターン）`);
          return count;
        }
      }
    }

    // 3. ページネーションから推測
    const paginationElems = document.querySelectorAll('.a-pagination li:not(.a-disabled) a');
    let maxPage = 0;
    for (const elem of paginationElems) {
      const pageNum = parseInt(elem.textContent, 10);
      if (!isNaN(pageNum) && pageNum > maxPage) {
        maxPage = pageNum;
      }
    }
    if (maxPage > 0) {
      const estimatedCount = maxPage * 10; // 1ページ10件として推測
      console.log(`[Amazonレビュー収集] getTotalReviewCount: ${estimatedCount}件（ページネーションから推測）`);
      return estimatedCount;
    }

    console.log('[Amazonレビュー収集] getTotalReviewCount: 件数を取得できませんでした');
    return 0;
  }

  /**
   * 状態を更新
   */
  async function updateState(newReviewCount) {
    return new Promise((resolve) => {
      chrome.storage.local.get(['collectionState'], (result) => {
        const state = result.collectionState || {
          isRunning: true,
          reviewCount: 0,
          pageCount: 0,
          totalPages: 1,
          reviews: [],
          logs: []
        };

        state.pageCount = (state.pageCount || 0) + 1;
        state.totalPages = totalPages || Math.ceil(getTotalReviewCount() / 10);
        state.isRunning = isCollecting;
        state.source = 'amazon';
        // 商品IDを保存（継続収集の判定用）
        state.productId = currentProductId;
        // 収集済みレビューキーを保存（ページ遷移後も維持するため）
        state.collectedReviewKeys = Array.from(collectedReviewKeys);
        // 現在のページURLとページ番号を保存
        state.lastProcessedUrl = window.location.href;
        state.lastProcessedPage = getCurrentPageNumber();

        chrome.storage.local.set({ collectionState: state }, () => {
          resolve();
        });
      });
    });
  }

  /**
   * ランダムウェイト時間を取得
   */
  function getRandomWait(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }

  /**
   * スリープ
   */
  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  // ===== ボット対策: 人間らしい行動パターン =====

  /**
   * 指数分布でランダムな値を生成（人間の行動パターンに近い）
   * 短い時間が多く、たまに長い時間が発生する分布
   */
  function exponentialRandom(mean) {
    return -mean * Math.log(1 - Math.random());
  }

  /**
   * ページ読み込み後の「見渡し」動作
   * 人間はページを開いたらまず全体を見渡す
   */
  async function initialPageScan() {
    // ページを開いた直後、まず画面を見渡す動作（1-3秒）
    const scanTime = 1000 + Math.random() * 2000;
    await sleep(scanTime);

    // 軽くスクロールして全体を確認（人間は最初に全体像を把握する）
    const quickScroll = 100 + Math.random() * 200;
    window.scrollTo({ top: quickScroll, behavior: 'smooth' });
    await sleep(500 + Math.random() * 500);
    window.scrollTo({ top: 0, behavior: 'smooth' });
    await sleep(300);
  }

  /**
   * レビュー要素にホバー動作を行う
   */
  async function hoverOnReview(review) {
    const rect = review.getBoundingClientRect();
    if (rect.width === 0 || rect.height === 0) return;

    // レビュー内のランダムな位置にマウスを移動
    const targetX = rect.left + rect.width * (0.1 + Math.random() * 0.7);
    const targetY = rect.top + rect.height * (0.2 + Math.random() * 0.5);

    // 現在位置からベジェ曲線で移動
    const startX = rect.left + Math.random() * 100;
    const startY = rect.top - 50 + Math.random() * 100;
    await moveMouseAlongBezier(startX, startY, targetX, targetY);

    // ホバー状態を維持（人間は興味ある部分をしばらく見る）
    await sleep(300 + Math.random() * 700);

    // 画像があれば見る（30%の確率）
    const image = review.querySelector('[data-hook="review-image-tile"]');
    if (image && Math.random() < 0.3) {
      const imgRect = image.getBoundingClientRect();
      if (imgRect.width > 0 && imgRect.height > 0) {
        await moveMouseAlongBezier(targetX, targetY, imgRect.left + 20, imgRect.top + 20);
        await sleep(500 + Math.random() * 1000); // 画像を見る時間
      }
    }
  }

  /**
   * 各レビューを「読む」動作をシミュレート
   */
  async function simulateReadingReviews() {
    const reviews = document.querySelectorAll(AMAZON_SELECTORS.reviewContainer);

    for (const review of reviews) {
      if (shouldStop) return;

      // レビューまでスクロール（画面中央に表示）
      review.scrollIntoView({ behavior: 'smooth', block: 'center' });
      await sleep(300 + Math.random() * 200); // スクロール完了待ち

      // レビューの長さに応じた読み時間を計算
      const bodyElem = review.querySelector('[data-hook="review-body"]');
      const bodyText = bodyElem?.textContent || '';
      const charCount = bodyText.length;

      // 100文字あたり0.3-0.5秒（人間の読書速度をシミュレート）
      const baseReadTime = (charCount / 100) * (READING_SPEED_PER_100CHARS + Math.random() * 200);
      const readTime = Math.min(Math.max(baseReadTime, 300), 3000); // 0.3-3秒の範囲

      // 時々レビューにホバー（20%の確率）
      if (Math.random() < 0.2) {
        await hoverOnReview(review);
      }

      // 読む時間
      await sleep(readTime);

      // 時々スキップ（つまらないレビューは読み飛ばす: 15%）
      if (Math.random() < 0.15) {
        // スキップ時は早めに次へ
        continue;
      }
    }
  }

  /**
   * ランダムなマイクロブレイク（人間らしい休憩）
   * 固定間隔ではなく、確率ベースで発生
   */
  async function maybeHaveBreak() {
    // 各ページで MICRO_BREAK_PROBABILITY の確率で休憩
    if (Math.random() < MICRO_BREAK_PROBABILITY) {
      const breakType = Math.random();

      if (breakType < 0.6) {
        // 短い休憩（60%）: 3-10秒
        const breakTime = 3000 + Math.random() * 7000;
        log(`少し休憩中...（${Math.round(breakTime / 1000)}秒）`);
        await sleep(breakTime);
      } else if (breakType < 0.9) {
        // 中程度の休憩（30%）: 15-45秒
        const breakTime = 15000 + Math.random() * 30000;
        log(`しばらく休憩中...（${Math.round(breakTime / 1000)}秒）`);
        await sleep(breakTime);
      } else {
        // 長い休憩（10%）: 1-3分
        const breakTime = 60000 + Math.random() * 120000;
        log(`長めの休憩中...（${Math.round(breakTime / 1000)}秒）`);
        await sleep(breakTime);
      }
      return true;
    }
    return false;
  }

  /**
   * ボット対策: タブがアクティブかチェック
   */
  async function isTabActive() {
    return new Promise(resolve => {
      chrome.runtime.sendMessage({ action: 'isTabActive' }, response => {
        resolve(response?.active || false);
      });
    });
  }

  /**
   * ボット対策: タブがアクティブになるまで待機
   */
  async function waitForActiveTab() {
    // アクティブタブチェックを無効化（バックグラウンドでも動作するように）
    return true;
  }

  /**
   * ボット対策: レート制限をチェック
   * 注: 制限を解除（常にOKを返す）
   */
  async function checkRateLimit() {
    // 制限解除: 常に許可を返す
    return { allowed: true, remaining: 999999 };
  }

  /**
   * ボット対策: レート制限カウントを増加
   */
  async function incrementRateLimit() {
    return new Promise(resolve => {
      chrome.runtime.sendMessage({ action: 'incrementAmazonRateLimit' }, response => {
        resolve(response || { count: 0, remaining: 0 });
      });
    });
  }

  /**
   * レビュー要素が表示されるまで待機
   */
  async function waitForReviews(maxWaitMs = 10000, intervalMs = 500) {
    const startTime = Date.now();
    while (Date.now() - startTime < maxWaitMs) {
      const reviews = document.querySelectorAll(AMAZON_SELECTORS.reviewContainer);
      if (reviews.length > 0) {
        console.log(`[Amazonレビュー収集] ${reviews.length}件のレビュー要素を検出（${Date.now() - startTime}ms）`);
        return true;
      }
      await sleep(intervalMs);
    }
    console.log(`[Amazonレビュー収集] レビュー要素が見つかりませんでした（${maxWaitMs}ms待機後）`);
    return false;
  }

  /**
   * ログをポップアップに送信
   */
  function log(text, type = '') {
    // プレフィックスはASINのみ（シンプルに）
    const productPrefix = currentProductId ? `[${currentProductId}] ` : '';
    const currentPage = getCurrentPageNumber();
    const pagePrefix = totalPages > 0 ? `[${currentPage}/${totalPages}] ` : '';
    const fullText = productPrefix + pagePrefix + text;
    console.log(`[Amazonレビュー収集] ${fullText}`);
    chrome.runtime.sendMessage({
      action: 'log',
      text: fullText,
      type: type
    });
  }

  // ページ読み込み時に収集状態を確認し、自動再開
  // background.jsのtabs.onUpdatedと連携し、content script側でも自律的に再開する
  // バックグラウンドタブでも確実に動作するように、複数の方法で初期化を試みる
  if (isReviewPage) {
    console.log('[Amazonレビュー収集] レビューページ検出、自動再開をチェックします');

    // 収集状態を確認して自動再開する関数
    async function checkAndResumeCollection() {
      // 既に収集中または初期化中の場合はスキップ
      if (isCollecting || startCollectionLock || resumeCollectionLock) {
        console.log('[Amazonレビュー収集] 既に処理中のためスキップ:', { isCollecting, startCollectionLock, resumeCollectionLock });
        return;
      }

      try {
        const result = await chrome.storage.local.get(['collectionState']);
        const state = result.collectionState;

        console.log('[Amazonレビュー収集] 自動再開チェック:', {
          isRunning: state?.isRunning,
          source: state?.source,
          productId: state?.productId,
          currentASIN: getASIN(),
          documentReadyState: document.readyState
        });

        // 収集中かつAmazonかつ同じ商品の場合、自動再開
        if (state && state.isRunning && state.source === 'amazon') {
          const storedProductId = state.productId || state.currentProductId;
          const currentASIN = getASIN();

          if (storedProductId === currentASIN) {
            console.log('[Amazonレビュー収集] 収集を自動再開します');

            // 状態を復元
            incrementalOnly = state.incrementalOnly || false;
            lastCollectedDate = state.lastCollectedDate || null;
            currentQueueName = state.queueName || null;
            currentProductId = currentASIN;

            // 収集を開始
            startCollection();
          } else {
            console.log('[Amazonレビュー収集] 商品IDが異なるため自動再開しません:', {
              stored: storedProductId,
              current: currentASIN
            });
          }
        }
      } catch (e) {
        console.log('[Amazonレビュー収集] 自動再開チェックエラー:', e);
      }
    }

    // スクリプト読み込み時点で即座にチェック
    // バックグラウンドタブでは load イベントのコールバックもスロットリングされるため、
    // イベントを待たずに即座に実行する
    // DOM読み込みは startCollection 内の waitForReviews 関数で待機するため問題なし
    checkAndResumeCollection();

    // ===== URL変更監視（SPA対応） =====
    // Amazonは「次へ」クリック時にページをリロードせず、コンテンツを動的に更新することがある
    // その場合、content scriptが再読み込みされないため、URL変更を監視して収集を再開する
    let lastUrl = window.location.href;
    let urlCheckInterval = null;

    function startUrlMonitoring() {
      if (urlCheckInterval) return; // 既に監視中

      urlCheckInterval = setInterval(() => {
        const currentUrl = window.location.href;
        if (currentUrl !== lastUrl) {
          console.log('[Amazonレビュー収集] URL変更を検出:', {
            from: lastUrl,
            to: currentUrl
          });
          lastUrl = currentUrl;

          // URLがレビューページかどうか確認
          const isStillReviewPage = currentUrl.includes('/product-reviews/') ||
                                     currentUrl.includes('/reviews/') ||
                                     currentUrl.includes('reviewerType=');

          if (isStillReviewPage) {
            console.log('[Amazonレビュー収集] レビューページへのURL変更、自動再開をチェックします');
            // フラグをリセットして再開を許可
            isCollecting = false;
            startCollectionLock = false;
            autoResumeExecuted = false;
            // 少し待ってからチェック（DOMが更新されるのを待つ）
            setTimeout(() => {
              checkAndResumeCollection();
            }, 500);
          }
        }
      }, 500); // 500msごとにURL変更をチェック

      console.log('[Amazonレビュー収集] URL変更監視を開始しました');
    }

    // URL監視を開始
    startUrlMonitoring();
  } else {
    console.log('[Amazonレビュー収集] レビューページではない:', window.location.href);
  }

  // テスト用API（メインワールドに注入）
  function injectTestAPI() {
    const script = document.createElement('script');
    script.textContent = `
      window.__AMAZON_REVIEW_EXT__ = {
        version: '1.0.0',
        _sendCommand: function(cmd, data) {
          return new Promise(function(resolve) {
            window.postMessage({ type: 'AMAZON_REVIEW_CMD', cmd: cmd, data: data }, '*');
            window.addEventListener('message', function handler(e) {
              if (e.data && e.data.type === 'AMAZON_REVIEW_RESPONSE' && e.data.cmd === cmd) {
                window.removeEventListener('message', handler);
                resolve(e.data.result);
              }
            });
          });
        },
        openOptions: function() { return this._sendCommand('openOptions'); },
        getStatus: function() { return this._sendCommand('getStatus'); },
        startCollection: function() { return this._sendCommand('startCollection'); },
        stopCollection: function() { return this._sendCommand('stopCollection'); },
        addToQueue: function() { return this._sendCommand('addToQueue'); },
        getProductInfo: function() { return this._sendCommand('getProductInfo'); }
      };
      console.log('[Amazonレビュー収集] テストAPI: window.__AMAZON_REVIEW_EXT__');
    `;
    document.documentElement.appendChild(script);
    script.remove();
  }

  // メインワールドからのコマンドを受信
  window.addEventListener('message', async (e) => {
    if (e.data && e.data.type === 'AMAZON_REVIEW_CMD') {
      const { cmd, data } = e.data;
      let result = null;

      switch (cmd) {
        case 'openOptions':
          chrome.runtime.sendMessage({ action: 'openOptions' });
          result = { success: true };
          break;
        case 'getStatus':
          result = await new Promise(resolve => {
            chrome.storage.local.get(['collectionState', 'queue'], r => {
              resolve({ state: r.collectionState || {}, queue: r.queue || [], isCollecting });
            });
          });
          break;
        case 'startCollection':
          if (isCollecting) {
            shouldStop = true;
            isCollecting = false;
          }
          startCollection();
          result = { success: true };
          break;
        case 'stopCollection':
          shouldStop = true;
          isCollecting = false;
          chrome.runtime.sendMessage({ action: 'collectionStopped' });
          result = { success: true };
          break;
        case 'addToQueue':
          const info = getProductInfo();
          result = await new Promise(resolve => {
            chrome.storage.local.get(['queue'], r => {
              const queue = r.queue || [];
              if (queue.some(item => item.url === info.url)) {
                resolve({ success: false, error: '既に追加済み' });
                return;
              }
              queue.push(info);
              chrome.storage.local.set({ queue }, () => {
                resolve({ success: true, productInfo: info });
              });
            });
          });
          break;
        case 'getProductInfo':
          result = getProductInfo();
          break;
      }

      window.postMessage({ type: 'AMAZON_REVIEW_RESPONSE', cmd, result }, '*');
    }
  });

  // APIを注入
  injectTestAPI();

})();
