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

    // 商品名を取得（新しい楽天の構造に対応）
    let productName = '商品名不明';
    const productNameSelectors = [
      'a[href*="item.rakuten.co.jp"]',
      '[class*="item-name"] a',
      'h1 a', 'h2 a',
      '.revRvwUserSec .revItemUrl a'
    ];
    for (const selector of productNameSelectors) {
      const elem = document.querySelector(selector);
      if (elem && elem.textContent.trim().length > 5) {
        productName = elem.textContent.trim();
        break;
      }
    }

    // 商品URLを取得
    const productUrlElem = document.querySelector('a[href*="item.rakuten.co.jp"]');
    const productUrl = productUrlElem ? productUrlElem.href : window.location.href;

    // レビュー要素を取得（新しい楽天の構造: ul > li で「購入者さん」を含む）
    const allListItems = document.querySelectorAll('li');
    const reviewElements = [];

    allListItems.forEach(li => {
      const text = li.textContent;
      // レビューの特徴: 「購入者さん」または日付パターン、そして十分な長さ
      if ((text.includes('購入者さん') || text.includes('注文日')) && text.length > 50) {
        reviewElements.push(li);
      }
    });

    // 旧構造にも対応
    if (reviewElements.length === 0) {
      const oldElements = document.querySelectorAll('.revRvwUserSec, .review-item, [class*="review-entry"]');
      oldElements.forEach(elem => reviewElements.push(elem));
    }

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
   * 新しい楽天の構造（CSS module）と旧構造の両方に対応
   */
  function extractReviewData(elem, productName, productUrl) {
    const text = elem.textContent || '';

    let rating = 0;
    let reviewDate = '';
    let author = '匿名';
    let body = '';
    let title = '';
    let purchaseInfo = '';
    let helpfulCount = 0;

    // 新構造の場合: CSS module クラス名で要素を探す（楽天の現在の構造）
    const ratingElem = elem.querySelector('[class*="number-wrapper"]');
    const reviewerElem = elem.querySelector('[class*="reviewer-name"]');
    const bodyElem = elem.querySelector('[class*="word-break-break-all"]');

    if (ratingElem || reviewerElem || bodyElem) {
      // 新構造: CSS moduleクラスで要素がある場合

      // 評価を取得（number-wrapper内のテキスト）
      if (ratingElem) {
        const ratingText = ratingElem.textContent.trim();
        const r = parseInt(ratingText, 10);
        if (r >= 1 && r <= 5) {
          rating = r;
        }
      }

      // 投稿者を取得
      if (reviewerElem) {
        author = reviewerElem.textContent.trim() || '匿名';
      }

      // 本文を取得
      if (bodyElem) {
        body = bodyElem.textContent.trim();
      }

      // 注文日を取得（テキストから抽出）
      const orderDateMatch = text.match(/注文日[：:]\s*(\d{4}\/\d{1,2}\/\d{1,2})/);
      if (orderDateMatch) {
        reviewDate = orderDateMatch[1];
      }
    } else {
      // テキストベースで抽出（フォールバック）

      // 評価: テキスト先頭の数字（1-5）
      const ratingMatch = text.match(/^(\d)/);
      if (ratingMatch) {
        const r = parseInt(ratingMatch[1], 10);
        if (r >= 1 && r <= 5) {
          rating = r;
        }
      }

      // 日付: YYYY/MM/DD 形式
      const dateMatch = text.match(/(\d{4}\/\d{1,2}\/\d{1,2})/);
      if (dateMatch) {
        reviewDate = dateMatch[1];
      }

      // 投稿者: 「購入者さん」や「○○さん」
      if (text.includes('購入者さん')) {
        author = '購入者さん';
      } else {
        const authorMatch = text.match(/(\S+さん)/);
        if (authorMatch) {
          author = authorMatch[1];
        }
      }

      // 本文: 投稿者名の後から「注文日」や「参考になった」の前まで
      let bodyText = text;

      // 先頭の評価と日付を除去
      bodyText = bodyText.replace(/^\d\d{4}\/\d{1,2}\/\d{1,2}/, '');

      // 投稿者名を除去
      bodyText = bodyText.replace(/購入者さん/, '');
      bodyText = bodyText.replace(/\S+さん/, '');

      // 末尾の不要な部分を除去
      bodyText = bodyText.replace(/注文日[：:].+$/, '');
      bodyText = bodyText.replace(/参考になった.*$/, '');
      bodyText = bodyText.replace(/不適切レビュー報告.*$/, '');

      body = bodyText.trim();
    }

    // 旧構造のセレクターも試す（フォールバック）
    if (!body && !rating) {
      const oldRatingElem = elem.querySelector('.revRvwUserEntryStar, .rating');
      if (oldRatingElem) {
        const ratingMatch = oldRatingElem.className.match(/(\d)/);
        if (ratingMatch) {
          rating = parseInt(ratingMatch[1], 10);
        }
      }

      const oldTitleElem = elem.querySelector('.revRvwUserEntryTtl, .review-title, h3, h4');
      title = oldTitleElem ? oldTitleElem.textContent.trim() : '';

      const oldBodyElem = elem.querySelector('.revRvwUserEntryCmt, .review-body, .review-text, p');
      body = oldBodyElem ? oldBodyElem.textContent.trim() : '';

      const oldAuthorElem = elem.querySelector('.revUserNickname, .reviewer-name, .author');
      author = oldAuthorElem ? oldAuthorElem.textContent.trim() : '匿名';

      const oldDateElem = elem.querySelector('.revRvwUserEntryDate, .review-date, .date, time');
      if (oldDateElem) {
        reviewDate = oldDateElem.textContent.trim();
        if (oldDateElem.getAttribute('datetime')) {
          reviewDate = oldDateElem.getAttribute('datetime');
        }
      }

      const oldPurchaseElem = elem.querySelector('.revRvwUserEntryPurchase, .purchase-info');
      purchaseInfo = oldPurchaseElem ? oldPurchaseElem.textContent.trim() : '';

      const oldHelpfulElem = elem.querySelector('.revRvwUserEntryHelpful, .helpful-count');
      if (oldHelpfulElem) {
        const helpfulMatch = oldHelpfulElem.textContent.match(/(\d+)/);
        if (helpfulMatch) {
          helpfulCount = parseInt(helpfulMatch[1], 10);
        }
      }
    }

    // レビューがない場合はスキップ
    if (!body && !title) {
      return null;
    }

    // 本文が短すぎる場合もスキップ（ノイズ除去）
    if (body.length < 10 && !title) {
      return null;
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
