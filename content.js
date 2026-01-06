/**
 * コンテンツスクリプト
 * 楽天市場のレビューページからデータをスクレイピングする
 */

(function() {
  'use strict';

  // 収集状態
  let isCollecting = false;
  let shouldStop = false;

  // ページタイプを判定
  const isReviewPage = window.location.hostname === 'review.rakuten.co.jp';
  const isItemPage = window.location.hostname === 'item.rakuten.co.jp';

  // メッセージリスナー
  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    switch (message.action) {
      case 'startCollection':
        if (!isCollecting) {
          if (isItemPage) {
            // 商品ページの場合、レビューページに遷移
            const reviewUrl = findReviewPageUrl();
            if (reviewUrl) {
              log('レビューページに移動します');
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
                chrome.storage.local.set({ collectionState: state }, () => {
                  window.location.href = reviewUrl;
                });
              });
              sendResponse({ success: true, redirecting: true });
            } else {
              sendResponse({ success: false, error: 'レビューページが見つかりません' });
            }
          } else {
            // レビューページの場合、収集開始
            startCollection();
            sendResponse({ success: true });
          }
        } else {
          sendResponse({ success: false, error: '既に収集中です' });
        }
        break;
      case 'stopCollection':
        shouldStop = true;
        isCollecting = false;
        sendResponse({ success: true });
        break;
      case 'getProductInfo':
        // 商品情報を取得
        const productInfo = getProductInfo();
        sendResponse({ success: true, productInfo: productInfo });
        break;
      default:
        sendResponse({ success: false, error: '不明なアクション' });
    }
    return true;
  });

  /**
   * 商品情報を取得（キュー追加用）
   */
  function getProductInfo() {
    let title = document.title || '';
    let url = window.location.href;

    // 商品名を取得
    const titleSelectors = [
      '.item_name',
      '[class*="itemName"]',
      'h1',
      '.product-name'
    ];

    for (const selector of titleSelectors) {
      const elem = document.querySelector(selector);
      if (elem && elem.textContent.trim().length > 5) {
        title = elem.textContent.trim();
        break;
      }
    }

    // レビューページの場合は商品ページURLを取得
    if (isReviewPage) {
      const productLink = document.querySelector('a[href*="item.rakuten.co.jp"]');
      if (productLink) {
        url = productLink.href;
      }
    }

    return {
      url: url,
      title: title.substring(0, 100),
      addedAt: new Date().toISOString()
    };
  }

  /**
   * 商品ページからレビューページのURLを取得
   */
  function findReviewPageUrl() {
    // review.rakuten.co.jp へのリンクを探す
    const reviewLinks = document.querySelectorAll('a[href*="review.rakuten.co.jp/item"]');
    for (const link of reviewLinks) {
      // レビューページへの直接リンクを優先
      if (link.href.includes('/item/') && !link.href.includes('/wd/')) {
        return link.href;
      }
    }

    // 見つからない場合、全てのreview.rakuten.co.jpリンクから探す
    const allReviewLinks = document.querySelectorAll('a[href*="review.rakuten.co.jp"]');
    for (const link of allReviewLinks) {
      if (link.href.includes('/item/')) {
        return link.href;
      }
    }

    return null;
  }

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

    // 商品管理番号を取得
    let productId = '';
    // ページ内の「商品番号：」から取得
    const pageText = document.body.textContent || '';
    const productIdMatch = pageText.match(/商品番号[：:]\s*([^\s\n]+)/);
    if (productIdMatch) {
      productId = productIdMatch[1].trim();
    }
    // URLから取得（フォールバック）
    if (!productId && productUrl) {
      const urlMatch = productUrl.match(/item\.rakuten\.co\.jp\/[^\/]+\/([^\/\?]+)/);
      if (urlMatch) {
        productId = urlMatch[1];
      }
    }

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
        const review = extractReviewData(elem, productName, productUrl, productId);
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
  function extractReviewData(elem, productName, productUrl, productId) {
    const text = elem.textContent || '';

    let rating = 0;
    let reviewDate = '';
    let orderDate = ''; // 注文日
    let author = '匿名';
    let body = '';
    let title = '';
    let purchaseInfo = '';
    let helpfulCount = 0;
    let variation = ''; // バリエーション（サイズ、カラーなど）
    let age = ''; // 年代
    let gender = ''; // 性別
    let usage = ''; // 用途（実用品・普段使い、プレゼント等）
    let recipient = ''; // 贈り先（自分用、家族へ等）
    let purchaseCount = ''; // 購入回数（はじめて、リピート等）
    let shopName = ''; // ショップ名

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

      // タイトルを取得（新構造）
      const titleElem = elem.querySelector('[class*="review-title"], [class*="title"], h3, h4, [class*="heading"]');
      if (titleElem) {
        title = titleElem.textContent.trim();
      }

      // 本文を取得
      if (bodyElem) {
        body = bodyElem.textContent.trim();
      }

      // 注文日を取得（テキストから抽出）
      const orderDateMatch = text.match(/注文日[：:]\s*(\d{4}\/\d{1,2}\/\d{1,2})/);
      if (orderDateMatch) {
        orderDate = orderDateMatch[1];
      }

      // レビュー投稿日を取得
      const reviewDateMatch = text.match(/(\d{4}\/\d{1,2}\/\d{1,2})/);
      if (reviewDateMatch && !orderDate) {
        // 注文日ではない日付をレビュー日として取得
        reviewDate = reviewDateMatch[1];
      } else if (reviewDateMatch && orderDate !== reviewDateMatch[1]) {
        // 注文日とは別の日付がある場合
        reviewDate = reviewDateMatch[1];
      }

      // バリエーション（サイズ、カラー等）を取得
      const variationPatterns = [
        /カラー[：:]\s*([^\s、,]+)/,
        /サイズ[：:]\s*([^\s、,]+)/,
        /color[：:]\s*([^\s、,]+)/i,
        /size[：:]\s*([^\s、,]+)/i,
        /種類[：:]\s*([^\s、,]+)/,
        /タイプ[：:]\s*([^\s、,]+)/
      ];

      const variationParts = [];
      for (const pattern of variationPatterns) {
        const match = text.match(pattern);
        if (match) {
          variationParts.push(match[0]);
        }
      }
      if (variationParts.length > 0) {
        variation = variationParts.join(' / ');
      }

      // バリエーション要素を探す（新構造）
      if (!variation) {
        const varElem = elem.querySelector('[class*="variation"], [class*="option"], [class*="sku"]');
        if (varElem) {
          variation = varElem.textContent.trim();
        }
      }

      // 年代を取得（10代、20代、30代、40代、50代、60代、70代以上など）
      const ageMatch = text.match(/(10代|20代|30代|40代|50代|60代|70代以上|70代)/);
      if (ageMatch) {
        age = ageMatch[1];
      }

      // 性別を取得
      const genderMatch = text.match(/(男性|女性)/);
      if (genderMatch) {
        gender = genderMatch[1];
      }

      // 用途を取得
      const usagePatterns = [
        '実用品・普段使い', '趣味', '仕事', 'プレゼント', 'イベント',
        'ビジネス', 'スポーツ', 'アウトドア', '旅行', '通勤',
        '普段使い', '日常使い', '実用', 'ギフト'
      ];
      for (const pattern of usagePatterns) {
        if (text.includes(pattern)) {
          usage = pattern;
          break;
        }
      }

      // 贈り先を取得
      const recipientPatterns = [
        '自分用', '家族へ', '親戚へ', '友人へ', '知人へ',
        '仕事関係へ', '子供へ', '男性へ', '女性へ', '恋人へ',
        '配偶者へ', '子どもへ', '親へ', '祖父母へ'
      ];
      for (const pattern of recipientPatterns) {
        if (text.includes(pattern)) {
          recipient = pattern;
          break;
        }
      }

      // 購入回数を取得
      if (text.includes('はじめて')) {
        purchaseCount = 'はじめて';
      } else if (text.includes('リピート')) {
        purchaseCount = 'リピート';
      } else {
        const countMatch = text.match(/(\d+)回目/);
        if (countMatch) {
          purchaseCount = countMatch[0];
        }
      }

      // ショップ名を取得
      const shopElem = elem.querySelector('[class*="shop-name"], [class*="store-name"]');
      if (shopElem) {
        shopName = shopElem.textContent.trim();
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
      productId: productId,
      productName: productName,
      productUrl: productUrl,
      rating: rating,
      title: title,
      body: body,
      author: author,
      age: age,
      gender: gender,
      reviewDate: reviewDate,
      orderDate: orderDate,
      variation: variation,
      usage: usage,
      recipient: recipient,
      purchaseCount: purchaseCount,
      purchaseInfo: purchaseInfo,
      helpfulCount: helpfulCount,
      shopName: shopName,
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
    // 楽天の新しいUI構造に対応したセレクター
    const allLinks = document.querySelectorAll('a');

    // 「次へ」「>」「»」などのテキストを含むリンクを探す
    for (const link of allLinks) {
      const text = link.textContent.trim();
      if ((text === '次へ' || text === '>' || text === '»' || text === '次' || text.includes('次のページ')) && link.href) {
        // 同じページへのリンクでないことを確認
        if (link.href !== window.location.href && link.href.includes('review.rakuten.co.jp')) {
          return link.href;
        }
      }
    }

    // CSSセレクターで探す
    const nextSelectors = [
      'a[rel="next"]',
      '.pagination a.next',
      '.pager a.next',
      '.page-next a',
      '.pagination li.next a',
      '[class*="pagination"] a[class*="next"]',
      '[class*="pager"] a[class*="next"]'
    ];

    for (const selector of nextSelectors) {
      try {
        const elem = document.querySelector(selector);
        if (elem && elem.href && elem.href !== window.location.href) {
          return elem.href;
        }
      } catch (e) {
        // セレクターエラーは無視
      }
    }

    // ページ番号から次を探す
    const currentPage = getCurrentPageNumber();
    if (currentPage) {
      const pageLinks = document.querySelectorAll('a');
      for (const link of pageLinks) {
        const text = link.textContent.trim();
        const pageNum = parseInt(text, 10);
        if (pageNum === currentPage + 1 && link.href && link.href.includes('review.rakuten.co.jp')) {
          return link.href;
        }
      }

      // URLパラメータで次のページを構築
      const url = new URL(window.location.href);
      const nextPage = currentPage + 1;
      url.searchParams.set('page', nextPage);

      // 次のページが存在するか確認（ページネーション要素があれば存在する可能性が高い）
      const paginationExists = document.querySelector('[class*="pagination"], [class*="pager"], [class*="page-nav"]');
      if (paginationExists) {
        return url.toString();
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
   * 総ページ数を取得
   */
  function getTotalPages() {
    // ページネーション内のリンクから最大ページ番号を探す
    const pageLinks = document.querySelectorAll('a');
    let maxPage = 1;

    for (const link of pageLinks) {
      const text = link.textContent.trim();
      const num = parseInt(text, 10);
      // 数字のみのリンクでページ番号らしいもの
      if (!isNaN(num) && num > 0 && num <= 1000 && link.href && link.href.includes('review.rakuten.co.jp')) {
        if (num > maxPage) {
          maxPage = num;
        }
      }
    }

    // ページネーション要素内のテキストからも探す
    const paginationElems = document.querySelectorAll('[class*="pagination"], [class*="pager"], [class*="page-nav"]');
    for (const elem of paginationElems) {
      const matches = elem.textContent.match(/(\d+)/g);
      if (matches) {
        for (const match of matches) {
          const num = parseInt(match, 10);
          if (num > maxPage && num <= 1000) {
            maxPage = num;
          }
        }
      }
    }

    return maxPage;
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

        state.reviewCount = (state.reviewCount || 0) + newReviewCount;
        state.pageCount = (state.pageCount || 0) + 1;
        state.totalPages = getTotalPages();
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

  // ページ読み込み時に収集状態を確認し、自動再開（レビューページのみ）
  if (isReviewPage) {
    chrome.storage.local.get(['collectionState'], (result) => {
      const state = result.collectionState;
      if (state && state.isRunning && !isCollecting) {
        // 前のページからの続きで自動的に収集を再開
        log('収集を再開します');
        startCollection();
      }
    });
  }

})();
