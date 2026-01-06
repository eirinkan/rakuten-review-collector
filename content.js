/**
 * コンテンツスクリプト
 * 楽天市場のレビューページからデータをスクレイピングする
 */

(function() {
  'use strict';

  // 収集状態
  let isCollecting = false;
  let shouldStop = false;

  // メッセージリスナー
  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    switch (message.action) {
      case 'startCollection':
        if (!isCollecting) {
          startCollection();
          sendResponse({ success: true });
        } else {
          sendResponse({ success: false, error: '既に収集中です' });
        }
        break;
      case 'stopCollection':
        shouldStop = true;
        isCollecting = false;
        sendResponse({ success: true });
        break;
      default:
        sendResponse({ success: false, error: '不明なアクション' });
    }
    return true;
  });

  /**
   * 収集を開始
   */
  async function startCollection() {
    isCollecting = true;
    shouldStop = false;

    log('レビュー収集を開始します');

    // 現在のページからレビューを収集
    await collectCurrentPage();

    // ページネーションがある場合、次のページも収集
    await collectAllPages();

    isCollecting = false;

    // 収集完了を通知
    if (!shouldStop) {
      chrome.runtime.sendMessage({ action: 'collectionComplete' });
      log('収集が完了しました', 'success');
    }
  }

  /**
   * 現在のページからレビューを収集
   */
  async function collectCurrentPage() {
    const reviews = extractReviews();

    if (reviews.length === 0) {
      log('このページにレビューが見つかりませんでした', 'error');
      return;
    }

    log(`${reviews.length}件のレビューを検出`);

    // バックグラウンドにデータを送信
    chrome.runtime.sendMessage({
      action: 'saveReviews',
      reviews: reviews
    });

    // 状態を更新
    await updateState(reviews.length);
  }

  /**
   * すべてのページを収集
   */
  async function collectAllPages() {
    while (!shouldStop) {
      // 次のページリンクを探す
      const nextLink = findNextPageLink();

      if (!nextLink) {
        log('最後のページに到達しました');
        break;
      }

      // ランダムウェイト（3-6秒）
      const waitTime = getRandomWait(3000, 6000);
      log(`${(waitTime / 1000).toFixed(1)}秒待機中...`);
      await sleep(waitTime);

      if (shouldStop) break;

      // 次のページに移動
      log('次のページに移動します');
      window.location.href = nextLink;

      // ページ遷移後は新しいコンテンツスクリプトが起動するため、
      // ここで現在のインスタンスは終了
      return;
    }
  }

  /**
   * レビューを抽出
   */
  function extractReviews() {
    const reviews = [];

    // 商品名を取得
    const productNameElem = document.querySelector('.revRvwUserSec .revItemUrl a, .item-name a, h2.revItemTtl a');
    const productName = productNameElem ? productNameElem.textContent.trim() : '商品名不明';

    // 商品URLを取得
    const productUrl = productNameElem ? productNameElem.href : window.location.href;

    // レビュー要素を取得（楽天の構造に対応）
    const reviewElements = document.querySelectorAll('.revRvwUserSec, .review-item, [class*="review"]');

    reviewElements.forEach((elem, index) => {
      try {
        const review = extractReviewData(elem, productName, productUrl);
        if (review) {
          reviews.push(review);
        }
      } catch (error) {
        console.error(`レビュー ${index + 1} の抽出エラー:`, error);
      }
    });

    // 重複排除（同じ内容のレビュー）
    return removeDuplicates(reviews);
  }

  /**
   * 個別レビューデータを抽出
   */
  function extractReviewData(elem, productName, productUrl) {
    // 評価（星の数）
    const ratingElem = elem.querySelector('.revRvwUserEntryStar, .rating, [class*="star"]');
    let rating = 0;
    if (ratingElem) {
      // クラス名から評価を取得（例: revUserRvwStar5）
      const ratingMatch = ratingElem.className.match(/(\d)/);
      if (ratingMatch) {
        rating = parseInt(ratingMatch[1], 10);
      }
      // または aria-label や title から取得
      const ariaLabel = ratingElem.getAttribute('aria-label') || ratingElem.getAttribute('title') || '';
      const ariaMatch = ariaLabel.match(/(\d)/);
      if (ariaMatch && !rating) {
        rating = parseInt(ariaMatch[1], 10);
      }
    }

    // レビュータイトル
    const titleElem = elem.querySelector('.revRvwUserEntryTtl, .review-title, h3, h4');
    const title = titleElem ? titleElem.textContent.trim() : '';

    // レビュー本文
    const bodyElem = elem.querySelector('.revRvwUserEntryCmt, .review-body, .review-text, p');
    const body = bodyElem ? bodyElem.textContent.trim() : '';

    // レビューがない場合はスキップ
    if (!body && !title) {
      return null;
    }

    // 投稿者
    const authorElem = elem.querySelector('.revUserNickname, .reviewer-name, .author');
    const author = authorElem ? authorElem.textContent.trim() : '匿名';

    // 投稿日
    const dateElem = elem.querySelector('.revRvwUserEntryDate, .review-date, .date, time');
    let reviewDate = '';
    if (dateElem) {
      reviewDate = dateElem.textContent.trim();
      // datetime属性があればそちらを優先
      if (dateElem.getAttribute('datetime')) {
        reviewDate = dateElem.getAttribute('datetime');
      }
    }

    // 購入商品の詳細（サイズ、カラーなど）
    const purchaseInfoElem = elem.querySelector('.revRvwUserEntryPurchase, .purchase-info');
    const purchaseInfo = purchaseInfoElem ? purchaseInfoElem.textContent.trim() : '';

    // 参考になった数
    const helpfulElem = elem.querySelector('.revRvwUserEntryHelpful, .helpful-count');
    let helpfulCount = 0;
    if (helpfulElem) {
      const helpfulMatch = helpfulElem.textContent.match(/(\d+)/);
      if (helpfulMatch) {
        helpfulCount = parseInt(helpfulMatch[1], 10);
      }
    }

    return {
      collectedAt: new Date().toISOString(),
      productName: productName,
      productUrl: productUrl,
      rating: rating,
      title: title,
      body: body,
      author: author,
      reviewDate: reviewDate,
      purchaseInfo: purchaseInfo,
      helpfulCount: helpfulCount,
      pageUrl: window.location.href
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
   * 次のページリンクを探す
   */
  function findNextPageLink() {
    // 楽天のページネーションパターン
    const nextSelectors = [
      '.pagination a.next',
      '.pager a.next',
      'a[rel="next"]',
      '.revPagination a:contains("次")',
      'a:contains("次のページ")',
      'a:contains("次へ")',
      '.page-next a',
      '.pagination li.next a'
    ];

    for (const selector of nextSelectors) {
      try {
        // :contains は標準CSSではサポートされていないので、特別処理
        if (selector.includes(':contains')) {
          const text = selector.match(/:contains\("(.+?)"\)/)[1];
          const baseSelector = selector.replace(/:contains\(".+?"\)/, '');
          const elements = document.querySelectorAll(baseSelector || 'a');
          for (const el of elements) {
            if (el.textContent.includes(text) && el.href) {
              return el.href;
            }
          }
        } else {
          const elem = document.querySelector(selector);
          if (elem && elem.href) {
            return elem.href;
          }
        }
      } catch (e) {
        // セレクターエラーは無視
      }
    }

    // ページ番号から次を探す
    const currentPage = getCurrentPageNumber();
    if (currentPage) {
      const pageLinks = document.querySelectorAll('.pagination a, .pager a, [class*="page"] a');
      for (const link of pageLinks) {
        const pageNum = parseInt(link.textContent.trim(), 10);
        if (pageNum === currentPage + 1) {
          return link.href;
        }
      }
    }

    return null;
  }

  /**
   * 現在のページ番号を取得
   */
  function getCurrentPageNumber() {
    // URLから取得
    const urlMatch = window.location.href.match(/[?&]page=(\d+)/);
    if (urlMatch) {
      return parseInt(urlMatch[1], 10);
    }

    // アクティブなページネーション要素から取得
    const activePageElem = document.querySelector('.pagination .active, .pager .current, [class*="page"].active');
    if (activePageElem) {
      const num = parseInt(activePageElem.textContent.trim(), 10);
      if (!isNaN(num)) {
        return num;
      }
    }

    return 1;
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
          reviews: [],
          logs: []
        };

        state.reviewCount = (state.reviewCount || 0) + newReviewCount;
        state.pageCount = (state.pageCount || 0) + 1;
        state.isRunning = isCollecting;

        chrome.storage.local.set({ collectionState: state }, () => {
          // ポップアップに進捗を通知
          chrome.runtime.sendMessage({
            action: 'updateProgress',
            state: state
          });
          resolve();
        });
      });
    });
  }

  /**
   * ランダムウェイト時間を取得（ミリ秒）
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
    console.log(`[楽天レビュー収集] ${text}`);
    chrome.runtime.sendMessage({
      action: 'log',
      text: text,
      type: type
    });
  }

  // ページ読み込み時に収集状態を確認し、自動再開
  chrome.storage.local.get(['collectionState'], (result) => {
    const state = result.collectionState;
    if (state && state.isRunning && !isCollecting) {
      // 前のページからの続きで自動的に収集を再開
      log('収集を再開します');
      startCollection();
    }
  });

})();
