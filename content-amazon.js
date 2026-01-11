/**
 * コンテンツスクリプト - Amazon
 * Amazon.co.jp の商品レビューを収集する
 */

(function() {
  'use strict';

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

            const asin = getASIN();
            let prefix = '';
            if (currentQueueName && asin) {
              prefix = `[${currentQueueName}・${asin}] `;
            } else if (asin) {
              prefix = `[${asin}] `;
            }
            log(prefix + 'レビューページに移動します');

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
              chrome.storage.local.set({ collectionState: state }, () => {
                // クリックでページ遷移（window.location.hrefはボット検出される）
                reviewLink.click();
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
        // backgroundからのページ遷移後の収集再開
        console.log('[Amazonレビュー収集] resumeCollectionメッセージ受信', {
          isReviewPage,
          isCollecting,
          startCollectionLock,
          autoResumeExecuted,
          currentUrl: window.location.href.substring(0, 80)
        });
        if (!isReviewPage) {
          console.log('[Amazonレビュー収集] レビューページではないためスキップ');
          sendResponse({ success: false, error: 'レビューページではありません' });
          break;
        }
        if (isCollecting) {
          console.log('[Amazonレビュー収集] 既に収集中のためスキップ');
          sendResponse({ success: false, error: '既に収集中' });
          break;
        }
        // 収集を再開
        console.log('[Amazonレビュー収集] 収集再開を開始します');
        incrementalOnly = message.incrementalOnly || false;
        lastCollectedDate = message.lastCollectedDate || null;
        currentQueueName = message.queueName || null;
        startCollectionLock = false; // ロックをリセット
        autoResumeExecuted = false; // 自動再開フラグをリセット
        log('収集を再開します（background経由）');
        startCollection();
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
    let title = document.title || '';
    let url = window.location.href;
    const asin = getASIN();

    // 商品名を取得
    const titleElem = document.querySelector(AMAZON_SELECTORS.productTitle);
    if (titleElem) {
      title = titleElem.textContent.trim();
    }

    // レビューページの場合は商品ページURLを構築
    if (isReviewPage && asin) {
      url = `https://www.amazon.co.jp/dp/${asin}`;
    }

    return {
      url: url,
      title: title.substring(0, 100),
      addedAt: new Date().toISOString(),
      source: 'amazon' // 販路を追加
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
    // 同期的なロックチェック（重複実行防止）
    if (startCollectionLock || isCollecting) {
      console.log('[Amazonレビュー収集] 既に収集中または開始処理中のためスキップ');
      return;
    }
    startCollectionLock = true; // 即座にロック

    isCollecting = true;
    shouldStop = false;
    autoResumeExecuted = true; // 収集開始したらフラグをセット

    currentProductId = getASIN();

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

    // 収集済みレビューキーをストレージから復元
    const storedState = await new Promise(resolve => {
      chrome.storage.local.get(['collectionState'], result => resolve(result.collectionState || {}));
    });
    if (storedState.collectedReviewKeys && Array.isArray(storedState.collectedReviewKeys)) {
      collectedReviewKeys = new Set(storedState.collectedReviewKeys);
    }

    // レビュー総数を取得
    const expectedTotal = getTotalReviewCount();
    if (expectedTotal > 0) {
      chrome.storage.local.set({ expectedReviewTotal: expectedTotal });
      totalPages = Math.ceil(expectedTotal / 10); // Amazonは1ページ10件
      if (incrementalOnly && lastCollectedDate) {
        log(`差分収集を開始します（前回: ${lastCollectedDate}、全${expectedTotal.toLocaleString()}件中新着のみ）`);
      } else {
        log(`レビュー収集を開始します（全${expectedTotal.toLocaleString()}件）`);
      }
    } else {
      chrome.storage.local.set({ expectedReviewTotal: 0 });
      totalPages = 0;
      if (incrementalOnly && lastCollectedDate) {
        log(`差分収集を開始します（前回: ${lastCollectedDate}）`);
      } else {
        log('レビュー収集を開始します');
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

    // ページネーションがある場合、次のページも収集（fetchベース）
    await collectAllPages();

    // 収集完了
    isCollecting = false;
    startCollectionLock = false;

    if (!shouldStop) {
      chrome.runtime.sendMessage({ action: 'collectionComplete' });
    }
  }

  /**
   * レビューの一意キーを生成
   */
  function generateReviewKey(review) {
    // 本文の先頭100文字 + 著者 + 日付でユニークキーを生成
    return `${(review.body || '').substring(0, 100)}|${review.author || ''}|${review.reviewDate || ''}`;
  }

  /**
   * 現在のページからレビューを収集
   */
  async function collectCurrentPage() {
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

    // すべて収集済みの場合（同じページを2回処理した可能性）
    if (reviews.length === 0) {
      log('このページのレビューは全て収集済みです');
      return false;
    }

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
   * すべてのページを収集（fetchベース - ページ遷移なし）
   */
  async function collectAllPages() {
    // 現在のページに次ページがあるかチェック
    if (!hasNextPage()) {
      log('最後のページに到達しました');
      return false;
    }

    let pageNumber = getCurrentPageNumber() + 1; // 現在ページの次から

    while (!shouldStop) {
      // ランダムウェイト（2-4秒）- fetchは速いので短くできる
      const waitTime = getRandomWait(2000, 4000);
      log(`${(waitTime / 1000).toFixed(1)}秒待機中...`);
      await sleep(waitTime);

      if (shouldStop) {
        return false;
      }

      try {
        log(`ページ${pageNumber}を取得中...`);
        const { reviews, hasNextPage: hasMore } = await fetchReviewPage(pageNumber);

        if (reviews.length === 0) {
          log('最後のページに到達しました');
          return false;
        }

        // 差分取得モードの処理
        let filteredReviews = reviews;
        let reachedOldReviews = false;

        if (incrementalOnly && lastCollectedDate) {
          const beforeFilter = filteredReviews.length;
          filteredReviews = filterReviewsByDate(filteredReviews, lastCollectedDate);
          if (filteredReviews.length < beforeFilter) {
            reachedOldReviews = true;
          }
        }

        // 重複チェック（セッション内）
        filteredReviews = filteredReviews.filter(review => {
          const key = generateReviewKey(review);
          if (collectedReviewKeys.has(key)) {
            return false;
          }
          collectedReviewKeys.add(key);
          return true;
        });

        if (filteredReviews.length > 0) {
          log(`ページ${pageNumber}: ${filteredReviews.length}件のレビューを検出`);

          // バックグラウンドに送信
          await new Promise(resolve => {
            chrome.runtime.sendMessage({
              action: 'saveReviews',
              reviews: filteredReviews,
              source: 'amazon'
            }, resolve);
          });

          await updateState(filteredReviews.length);
        } else {
          log(`ページ${pageNumber}: 新着レビューなし`);
        }

        // 終了条件
        if (reachedOldReviews) {
          log('前回収集済みのレビューに到達しました');
          return false;
        }

        if (!hasMore) {
          log('最後のページに到達しました');
          return false;
        }

        pageNumber++;

      } catch (error) {
        console.error('[Amazonレビュー収集] fetchエラー:', error);
        log(`ページ${pageNumber}の取得に失敗: ${error.message}`, 'error');
        return false;
      }
    }

    return false;
  }

  /**
   * 次のページがあるかチェック
   */
  function hasNextPage() {
    // 最後のページかどうか確認
    const isLastPage = document.querySelector(AMAZON_SELECTORS.isLastPage);
    if (isLastPage) {
      return false;
    }
    // 「次へ」リンクが存在するか
    const nextLink = document.querySelector(AMAZON_SELECTORS.nextPage);
    return !!nextLink;
  }

  /**
   * 「次へ」リンクをクリックしてページ遷移（レガシー - 現在は未使用）
   */
  function clickNextPage() {
    const nextLink = document.querySelector(AMAZON_SELECTORS.nextPage);
    if (nextLink) {
      console.log('[Amazonレビュー収集] 次へリンクをクリック:', nextLink.href);
      const event = new MouseEvent('click', {
        bubbles: true,
        cancelable: true,
        view: window
      });
      nextLink.dispatchEvent(event);
    } else {
      console.log('[Amazonレビュー収集] 次へリンクが見つかりません');
    }
  }

  /**
   * 指定ページのレビューをfetchで取得
   * @param {number} pageNumber - ページ番号
   * @returns {Promise<{reviews: Array, hasNextPage: boolean}>}
   */
  async function fetchReviewPage(pageNumber) {
    const asin = getASIN();
    const url = `https://www.amazon.co.jp/product-reviews/${asin}?pageNumber=${pageNumber}&reviewerType=all_reviews`;

    console.log('[Amazonレビュー収集] fetchでページ取得:', url);

    const response = await fetch(url, {
      credentials: 'include',
      headers: {
        'Accept': 'text/html,application/xhtml+xml'
      }
    });

    if (!response.ok) {
      throw new Error(`HTTP error: ${response.status}`);
    }

    const html = await response.text();
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');

    // ボット検出チェック
    const robotCheck = doc.querySelector('form[action="/errors/validateCaptcha"]');
    if (robotCheck) {
      throw new Error('ボット検出されました。しばらく待ってから再試行してください。');
    }

    // レビューを抽出
    const reviews = extractReviewsFromDocument(doc);

    // 次ページ存在チェック
    const isLastPage = !!doc.querySelector(AMAZON_SELECTORS.isLastPage);
    const hasNextLink = !!doc.querySelector(AMAZON_SELECTORS.nextPage);
    const hasNextPage = !isLastPage && hasNextLink;

    console.log('[Amazonレビュー収集] ページ', pageNumber, '取得完了:', reviews.length, '件, 次ページ:', hasNextPage);

    return { reviews, hasNextPage };
  }

  /**
   * DOMDocumentからレビューを抽出
   * @param {Document} doc - パース済みのDocument
   * @returns {Array} レビュー配列
   */
  function extractReviewsFromDocument(doc) {
    const reviews = [];
    const asin = getASIN();

    // 商品名を取得
    let productName = '商品名不明';
    const titleElem = doc.querySelector(AMAZON_SELECTORS.reviewPageTitle);
    if (titleElem) {
      productName = titleElem.textContent.trim();
    } else {
      // フォールバック: タイトルから取得
      const pageTitle = doc.title || '';
      if (pageTitle.includes('のカスタマーレビュー')) {
        productName = pageTitle.replace(/のカスタマーレビュー.*$/, '').replace(/^Amazon.*?: /, '');
      }
    }

    const productUrl = `https://www.amazon.co.jp/dp/${asin}`;
    const reviewElements = doc.querySelectorAll(AMAZON_SELECTORS.reviewContainer);

    reviewElements.forEach((elem, index) => {
      try {
        const review = extractReviewData(elem, productName, productUrl, asin);
        if (review) {
          reviews.push(review);
        }
      } catch (error) {
        console.error(`[fetch] レビュー ${index + 1} の抽出エラー:`, error);
      }
    });

    return reviews;
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
    let isVerified = false;
    let isVine = false;
    let hasImage = false;
    let country = '日本'; // デフォルトは日本

    // 評価を取得
    const ratingElem = elem.querySelector(AMAZON_SELECTORS.rating);
    if (ratingElem) {
      const ratingText = ratingElem.textContent || '';
      // 「5つ星のうち4.0」から数値を抽出
      const ratingMatch = ratingText.match(/(\d+(?:\.\d+)?)/);
      if (ratingMatch) {
        rating = Math.round(parseFloat(ratingMatch[1]));
      }
    }

    // タイトルを取得
    const titleElem = elem.querySelector(AMAZON_SELECTORS.title);
    if (titleElem) {
      title = titleElem.textContent.trim();
    }

    // 本文を取得
    const bodyElem = elem.querySelector(AMAZON_SELECTORS.body);
    if (bodyElem) {
      body = bodyElem.textContent.trim();
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

    // 検証済み購入を確認
    const verifiedElem = elem.querySelector(AMAZON_SELECTORS.verified);
    if (verifiedElem) {
      isVerified = true;
    }

    // Vineレビューを確認
    const vineElem = elem.querySelector(AMAZON_SELECTORS.vine);
    if (vineElem && vineElem.textContent.includes('Vine')) {
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

    // 楽天と同じ20項目構造を維持（+source, +country で22項目）
    return {
      collectedAt: new Date().toISOString(),
      productId: asin,
      productName: productName,
      productUrl: productUrl,
      rating: rating,
      title: title,
      body: body,
      author: author,
      age: '', // Amazonにはない
      gender: '', // Amazonにはない
      reviewDate: reviewDate,
      orderDate: '', // Amazonにはない
      variation: variation,
      usage: '', // Amazonにはない
      recipient: '', // Amazonにはない
      purchaseCount: '', // Amazonにはない
      helpfulCount: helpfulCount,
      shopReply: '', // Amazonにはない
      shopName: 'Amazon',
      pageUrl: window.location.href,
      // Amazon固有の追加項目
      source: 'amazon',
      country: country,
      isVerified: isVerified,
      isVine: isVine,
      hasImage: hasImage
    };
  }

  /**
   * 重複を除去
   */
  function removeDuplicates(reviews) {
    const seen = new Set();
    return reviews.filter(review => {
      const key = `${review.body.substring(0, 100)}${review.author}${review.reviewDate}`;
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
   */
  function getTotalReviewCount() {
    const totalElem = document.querySelector(AMAZON_SELECTORS.totalReviews);
    if (totalElem) {
      const text = totalElem.textContent || '';
      // 「1,234件のグローバルレーティング」から数値を抽出
      const match = text.match(/([\d,]+)\s*件/);
      if (match) {
        return parseInt(match[1].replace(/,/g, ''), 10);
      }
    }
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
    let productPrefix = '';
    if (currentQueueName && currentProductId) {
      productPrefix = `[${currentQueueName}・${currentProductId}] `;
    } else if (currentProductId) {
      productPrefix = `[${currentProductId}] `;
    }
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

  // ページ読み込み時に収集状態を確認し、自動再開（レビューページのみ）
  if (isReviewPage) {
    console.log('[Amazonレビュー収集] レビューページ検出、自動再開チェック開始');
    chrome.storage.local.get(['collectionState'], (result) => {
      const state = result.collectionState;
      console.log('[Amazonレビュー収集] 状態:', JSON.stringify({
        hasState: !!state,
        isRunning: state?.isRunning,
        source: state?.source,
        isCollecting,
        autoResumeExecuted,
        lastProcessedPage: state?.lastProcessedPage
      }));

      if (state && state.isRunning && state.source === 'amazon' && !isCollecting && !autoResumeExecuted) {
        // 同じURLで再度実行されないようにチェック（ページ遷移が実際に発生したか確認）
        const currentUrl = window.location.href.split('?')[0]; // クエリパラメータを除く
        const lastUrl = (state.lastProcessedUrl || '').split('?')[0];
        const currentPage = getCurrentPageNumber();
        const lastPage = state.lastProcessedPage || 0;

        console.log('[Amazonレビュー収集] ページチェック:', { currentPage, lastPage, shouldResume: currentPage > lastPage });

        // ページ番号が進んでいるか、URLが異なる場合のみ再開
        if (currentPage > lastPage || currentUrl !== lastUrl) {
          incrementalOnly = state.incrementalOnly || false;
          lastCollectedDate = state.lastCollectedDate || null;
          currentQueueName = state.queueName || null;
          autoResumeExecuted = true; // 重複実行防止フラグ

          // DOMが完全に更新されるまで待機（Amazonは動的にレビューをロードするため）
          // 5秒待機に延長（ページ読み込みが遅い場合に対応）
          setTimeout(() => {
            // 再度チェック（タイムアウト中に他の処理で開始された可能性）
            if (!isCollecting) {
              log('収集を再開します');
              startCollection();
            } else {
              console.log('[Amazonレビュー収集] 既に収集中のため再開スキップ');
            }
          }, 5000);
        } else {
          console.log('[Amazonレビュー収集] 同じページのため自動再開をスキップ');
        }
      } else {
        console.log('[Amazonレビュー収集] 自動再開条件を満たさず');
      }
    });
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
