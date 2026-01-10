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
          const reviewUrl = findReviewPageUrl();
          if (reviewUrl) {
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

            // 収集状態を設定してからリダイレクト
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
                window.location.href = reviewUrl;
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
   * 商品ページからレビューページのURLを取得
   */
  function findReviewPageUrl() {
    // 「すべてのレビューを見る」リンクを探す
    const reviewLink = document.querySelector(AMAZON_SELECTORS.reviewLink);
    if (reviewLink && reviewLink.href) {
      return reviewLink.href;
    }

    // ASINからレビューページURLを構築
    const asin = getASIN();
    if (asin) {
      return `https://www.amazon.co.jp/product-reviews/${asin}/ref=cm_cr_dp_d_show_all_btm?ie=UTF8&reviewerType=all_reviews`;
    }

    return null;
  }

  /**
   * 収集を開始
   */
  async function startCollection() {
    if (isCollecting) {
      console.log('[Amazonレビュー収集] 既に収集中のためスキップ');
      return;
    }

    isCollecting = true;
    shouldStop = false;

    currentProductId = getASIN();

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
   * 現在のページからレビューを収集
   */
  async function collectCurrentPage() {
    let reviews = extractReviews();

    if (reviews.length === 0) {
      log('このページにレビューが見つかりませんでした', 'error');
      return false;
    }

    // 差分取得モードの場合、日付でフィルタリング
    let reachedOldReviews = false;
    if (incrementalOnly && lastCollectedDate) {
      const originalCount = reviews.length;
      reviews = filterReviewsByDate(reviews, lastCollectedDate);
      const filteredCount = originalCount - reviews.length;

      if (filteredCount > 0) {
        log(`${originalCount}件中${filteredCount}件は前回収集済み（${lastCollectedDate}より前）`);
      }

      if (reviews.length === 0) {
        log('前回以降の新着レビューがこのページにはありません');
        reachedOldReviews = true;
        return reachedOldReviews;
      }

      if (filteredCount > 0 && reviews.length < originalCount) {
        reachedOldReviews = true;
      }
    }

    log(`${reviews.length}件のレビューを検出`);

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
      const nextPageLink = findNextPageLink();

      if (!nextPageLink) {
        log('最後のページに到達しました');
        return false;
      }

      // ランダムウェイト（6-12秒）- Amazonはボット対策が厳しいため長めに設定
      const waitTime = getRandomWait(6000, 12000);
      log(`${(waitTime / 1000).toFixed(1)}秒待機中...`);
      await sleep(waitTime);

      if (shouldStop) {
        return false;
      }

      // 差分取得設定を保存してからページ遷移
      await new Promise((resolve) => {
        chrome.storage.local.get(['collectionState'], (result) => {
          const state = result.collectionState || {};
          state.incrementalOnly = incrementalOnly;
          state.lastCollectedDate = lastCollectedDate;
          state.source = 'amazon';
          chrome.storage.local.set({ collectionState: state }, resolve);
        });
      });

      window.location.href = nextPageLink;
      return true; // ページ遷移中
    }

    return false;
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
   * 次のページリンクを探す
   */
  function findNextPageLink() {
    // 「次へ」リンクを探す
    const nextLink = document.querySelector(AMAZON_SELECTORS.nextPage);
    if (nextLink && nextLink.href) {
      return nextLink.href;
    }

    // 最後のページかどうか確認
    const isLastPage = document.querySelector(AMAZON_SELECTORS.isLastPage);
    if (isLastPage) {
      return null;
    }

    // URLパラメータから次のページを構築
    const url = new URL(window.location.href);
    const currentPage = parseInt(url.searchParams.get('pageNumber') || '1', 10);
    url.searchParams.set('pageNumber', currentPage + 1);

    return url.toString();
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
    chrome.storage.local.get(['collectionState'], (result) => {
      const state = result.collectionState;
      if (state && state.isRunning && state.source === 'amazon' && !isCollecting) {
        incrementalOnly = state.incrementalOnly || false;
        lastCollectedDate = state.lastCollectedDate || null;
        currentQueueName = state.queueName || null;
        // DOMが完全に更新されるまで待機（Amazonは動的にレビューをロードするため）
        setTimeout(() => {
          log('収集を再開します');
          startCollection();
        }, 2000);
      }
    });
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
