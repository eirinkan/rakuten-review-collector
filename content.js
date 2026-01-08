/**
 * コンテンツスクリプト
 * 楽天市場のレビューページからデータをスクレイピングする
 */

(function() {
  'use strict';

  // 収集状態
  let isCollecting = false;
  let shouldStop = false;
  let currentProductId = ''; // 現在の商品管理番号

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
              // 商品IDを取得してログに表示
              const itemUrlMatch = window.location.href.match(/item\.rakuten\.co\.jp\/[^\/]+\/([^\/\?]+)/);
              const itemProductId = itemUrlMatch ? itemUrlMatch[1] : '';
              const prefix = itemProductId ? `[${itemProductId}] ` : '';
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
   * 商品管理番号を取得
   */
  function getProductId() {
    // 商品URLを取得
    const productUrlElem = document.querySelector('a[href*="item.rakuten.co.jp"]');
    const productUrl = productUrlElem ? productUrlElem.href : window.location.href;

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

    return productId;
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

    // 商品管理番号を取得
    currentProductId = getProductId();

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
    const textDisplayElements = elem.querySelectorAll('[class*="text-display"]');

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

      // タイトルを取得（style-boldクラスを持つ要素、ショップコメント除外）
      for (const el of textDisplayElements) {
        const className = el.className || '';
        const elText = el.textContent.trim();
        if (className.includes('style-bold') && elText.length >= 5 && elText.length < 100 &&
            !elText.includes('ショップからのコメント')) {
          title = elText;
          break;
        }
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

      // バリエーション（サイズ、カラー等）を取得 - 短いテキストから特定行を抽出
      for (const el of textDisplayElements) {
        const className = el.className || '';
        const elText = el.textContent.trim();
        if (className.includes('size-small') && elText.length < 100 &&
            (elText.includes('種類:') || elText.includes('カラー:') ||
             elText.includes('サイズ:') || elText.includes('タイプ:'))) {
          const lines = elText.split('\n');
          const varLine = lines.find(line =>
            line.includes('種類:') || line.includes('カラー:') ||
            line.includes('サイズ:') || line.includes('タイプ:')
          );
          if (varLine) {
            variation = varLine.trim();
          }
          break;
        }
      }

      // フォールバック: テキストからパターンマッチング
      if (!variation) {
        const variationPatterns = [
          /カラー[：:]\s*([^\s、,\n]+)/,
          /サイズ[：:]\s*([^\s、,\n]+)/,
          /種類[：:]\s*([^\s、,\n]+)/,
          /タイプ[：:]\s*([^\s、,\n]+)/
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
      }

      // ショップ名を商品URLから取得
      if (productUrl) {
        const shopMatch = productUrl.match(/item\.rakuten\.co\.jp\/([^\/]+)/);
        if (shopMatch) {
          shopName = shopMatch[1];
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
    const currentPage = getCurrentPageNumber();
    const totalPages = getTotalPages();

    log(`現在のページ: ${currentPage} / ${totalPages}`);

    // 最終ページに到達している場合
    if (currentPage >= totalPages) {
      return null;
    }

    // 方法1: 「次へ」リンクを探す
    const allLinks = document.querySelectorAll('a');
    for (const link of allLinks) {
      const text = link.textContent.trim();
      if ((text === '次へ' || text === '>' || text === '»' || text === '次' || text.includes('次のページ')) && link.href) {
        if (link.href !== window.location.href && link.href.includes('review.rakuten.co.jp')) {
          return link.href;
        }
      }
    }

    // 方法2: 次のページ番号のリンクを探す
    const nextPageNum = currentPage + 1;
    for (const link of allLinks) {
      const text = link.textContent.trim();
      if (text === String(nextPageNum) && link.href && link.href.includes('review.rakuten.co.jp')) {
        return link.href;
      }
    }

    // 方法3: URLパラメータで次のページを構築
    // 楽天レビューページのURL形式: /item/1/SHOP_ID/ITEM_ID/PAGE.SORT/
    const currentUrl = window.location.href;

    // URL形式1: /1.1/ → /2.1/ のパターン
    const pagePattern = /\/(\d+)\.(\d+)\/?(\?.*)?$/;
    const pageMatch = currentUrl.match(pagePattern);
    if (pageMatch) {
      const pageNum = parseInt(pageMatch[1], 10);
      const sortNum = pageMatch[2];
      const query = pageMatch[3] || '';
      const nextUrl = currentUrl.replace(pagePattern, `/${pageNum + 1}.${sortNum}/${query}`);
      return nextUrl;
    }

    // URL形式2: ?page=X パラメータ
    const url = new URL(currentUrl);
    if (url.searchParams.has('page')) {
      url.searchParams.set('page', nextPageNum);
      return url.toString();
    }

    return null;
  }

  /**
   * 現在のページ番号を取得
   */
  function getCurrentPageNumber() {
    const url = window.location.href;

    // 楽天レビューページのURL形式: /PAGE.SORT/ (例: /1.1/, /2.1/)
    const pagePattern = /\/(\d+)\.\d+\/?(\?.*)?$/;
    const pageMatch = url.match(pagePattern);
    if (pageMatch) {
      return parseInt(pageMatch[1], 10);
    }

    // ?page=X パラメータ
    const urlParamMatch = url.match(/[?&]page=(\d+)/);
    if (urlParamMatch) {
      return parseInt(urlParamMatch[1], 10);
    }

    // アクティブなページネーション要素から取得
    const activeSelectors = [
      '[class*="pagination"] [class*="active"]',
      '[class*="pagination"] [class*="current"]',
      '[class*="pager"] [class*="active"]',
      '.pagination .active',
      '.pager .current'
    ];

    for (const selector of activeSelectors) {
      try {
        const elem = document.querySelector(selector);
        if (elem) {
          const num = parseInt(elem.textContent.trim(), 10);
          if (!isNaN(num) && num > 0) {
            return num;
          }
        }
      } catch (e) {
        // セレクターエラーは無視
      }
    }

    return 1;
  }

  /**
   * 総ページ数を取得
   */
  function getTotalPages() {
    let maxPage = 1;

    // 方法1: レビュー件数から計算（1ページあたり約15件）
    const reviewCountText = document.body.textContent;
    const reviewCountMatch = reviewCountText.match(/(\d+)件/);
    if (reviewCountMatch) {
      const totalReviews = parseInt(reviewCountMatch[1], 10);
      if (totalReviews > 0) {
        const calculatedPages = Math.ceil(totalReviews / 15);
        if (calculatedPages > maxPage) {
          maxPage = calculatedPages;
        }
      }
    }

    // 方法2: ページネーションリンクから最大ページを探す
    const pageLinks = document.querySelectorAll('a');
    for (const link of pageLinks) {
      if (!link.href || !link.href.includes('review.rakuten.co.jp')) continue;

      const text = link.textContent.trim();
      // 数字のみのリンク
      if (/^\d+$/.test(text)) {
        const num = parseInt(text, 10);
        if (num > maxPage && num <= 1000) {
          maxPage = num;
        }
      }

      // URLからページ番号を抽出 (/N.1/ 形式)
      const urlPageMatch = link.href.match(/\/(\d+)\.\d+\/?(\?.*)?$/);
      if (urlPageMatch) {
        const num = parseInt(urlPageMatch[1], 10);
        if (num > maxPage && num <= 1000) {
          maxPage = num;
        }
      }
    }

    // 方法3: 「最後」「最終」リンクのURLから取得
    for (const link of pageLinks) {
      const text = link.textContent.trim();
      if ((text === '最後' || text === '最終' || text === '»»' || text.includes('最終ページ')) && link.href) {
        const lastPageMatch = link.href.match(/\/(\d+)\.\d+\/?(\?.*)?$/);
        if (lastPageMatch) {
          const num = parseInt(lastPageMatch[1], 10);
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
   * ログをポップアップに送信（商品管理番号を自動プレフィックス）
   */
  function log(text, type = '') {
    const prefix = currentProductId ? `[${currentProductId}] ` : '';
    const fullText = prefix + text;
    console.log(`[楽天レビュー収集] ${fullText}`);
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
      if (state && state.isRunning && !isCollecting) {
        // 前のページからの続きで自動的に収集を再開
        log('収集を再開します');
        startCollection();
      }
    });
  }

  // テスト用API（メインワールドに注入）
  function injectTestAPI() {
    const script = document.createElement('script');
    script.textContent = `
      window.__RAKUTEN_REVIEW_EXT__ = {
        version: '1.3.2',
        _sendCommand: function(cmd, data) {
          return new Promise(function(resolve) {
            window.postMessage({ type: 'RAKUTEN_REVIEW_CMD', cmd: cmd, data: data }, '*');
            window.addEventListener('message', function handler(e) {
              if (e.data && e.data.type === 'RAKUTEN_REVIEW_RESPONSE' && e.data.cmd === cmd) {
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
      console.log('[楽天レビュー収集] テストAPI: window.__RAKUTEN_REVIEW_EXT__');
    `;
    document.documentElement.appendChild(script);
    script.remove();
  }

  // メインワールドからのコマンドを受信
  window.addEventListener('message', async (e) => {
    if (e.data && e.data.type === 'RAKUTEN_REVIEW_CMD') {
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
          if (!isCollecting) {
            startCollection();
            result = { success: true };
          } else {
            result = { success: false, error: '既に収集中' };
          }
          break;
        case 'stopCollection':
          shouldStop = true;
          isCollecting = false;
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

      window.postMessage({ type: 'RAKUTEN_REVIEW_RESPONSE', cmd, result }, '*');
    }
  });

  // APIを注入
  injectTestAPI();

})();
