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
        // 既に収集中の場合
        if (isCollecting) {
          // 強制フラグがある場合のみ再開始（ユーザーが明示的に押した場合）
          if (message.force) {
            shouldStop = true;
            isCollecting = false;
            log('前の収集を中断して再開始します');
          } else {
            // キュー収集などの自動収集では無視
            console.log('[楽天レビュー収集] 既に収集中のためスキップ');
            sendResponse({ success: false, error: '既に収集中' });
            break;
          }
        }

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
        break;
      case 'stopCollection':
        shouldStop = true;
        isCollecting = false;
        // background.jsに停止を通知（collectingItemsからの削除用）
        chrome.runtime.sendMessage({ action: 'collectionStopped' });
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
    // 既に収集中の場合は何もしない（重複防止）
    if (isCollecting) {
      console.log('[楽天レビュー収集] 既に収集中のためスキップ');
      return;
    }

    isCollecting = true;
    shouldStop = false;

    // 商品管理番号を取得
    currentProductId = getProductId();

    // レビュー総数を取得して保存（収集完了時の検証用）
    const expectedTotal = getTotalReviewCount();
    if (expectedTotal > 0) {
      chrome.storage.local.set({ expectedReviewTotal: expectedTotal });
      log(`レビュー収集を開始します（全${expectedTotal.toLocaleString()}件）`);
    } else {
      chrome.storage.local.set({ expectedReviewTotal: 0 });
      log('レビュー収集を開始します');
    }

    // 現在のページからレビューを収集
    await collectCurrentPage();

    // ページネーションがある場合、次のページも収集
    // navigatedがtrueの場合、ページ遷移中なので完了通知を送らない
    const navigated = await collectAllPages();

    if (navigated) {
      // ページ遷移中は何もしない（新しいページで収集を継続）
      return;
    }

    isCollecting = false;

    // 収集完了を通知（ログはbackground.jsで出力）
    if (!shouldStop) {
      chrome.runtime.sendMessage({ action: 'collectionComplete' });
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

    // バックグラウンドにデータを送信（完了を待つ）
    await new Promise((resolve) => {
      chrome.runtime.sendMessage({
        action: 'saveReviews',
        reviews: reviews
      }, (response) => {
        resolve(response);
      });
    });

    // 状態を更新（saveReviews完了後に実行）
    await updateState(reviews.length);
  }

  /**
   * すべてのページを収集
   * @returns {boolean} ページ遷移した場合はtrue、完了した場合はfalse
   */
  async function collectAllPages() {
    while (!shouldStop) {
      // 次のページへ移動する方法を探す
      const nextPage = findNextPage();

      if (!nextPage) {
        log('最後のページに到達しました');
        return false; // 完了
      }

      // ランダムウェイト（3-6秒）
      const waitTime = getRandomWait(3000, 6000);
      log(`${(waitTime / 1000).toFixed(1)}秒待機中...`);
      await sleep(waitTime);

      if (shouldStop) {
        return false; // 停止された
      }

      // 次のページに移動
      if (nextPage.type === 'button') {
        // ボタンクリック方式（新UI）
        // 現在のレビュー要素を記録（変更検知用）
        const currentReviewCount = document.querySelectorAll('li').length;

        // ボタンをクリック
        nextPage.element.click();

        // コンテンツの更新を待つ
        const contentUpdated = await waitForContentUpdate(currentReviewCount);

        if (!contentUpdated) {
          log('最終ページに到達しました');
          return false;
        }

        // 新しいページのレビューを収集
        await collectCurrentPage();

        // 次のループで次のページを探す
        continue;

      } else if (nextPage.type === 'link' || nextPage.type === 'url') {
        // リンク/URL方式（旧UI）
        window.location.href = nextPage.url;

        // ページ遷移後は新しいコンテンツスクリプトが起動するため、
        // ここで現在のインスタンスは終了
        return true; // ページ遷移中
      }
    }

    return false; // shouldStopで終了
  }

  /**
   * コンテンツの更新を待つ（AJAX読み込み用）
   * @param {number} previousCount 更新前のレビュー要素数
   * @param {number} timeout タイムアウト（ミリ秒）
   * @returns {boolean} 更新されたらtrue、タイムアウトしたらfalse
   */
  async function waitForContentUpdate(previousCount, timeout = 10000) {
    const startTime = Date.now();
    const checkInterval = 500; // 500msごとにチェック

    while (Date.now() - startTime < timeout) {
      await sleep(checkInterval);

      // レビュー要素の変化をチェック
      const currentReviewElements = document.querySelectorAll('li');
      let reviewLiCount = 0;
      currentReviewElements.forEach(li => {
        const text = li.textContent;
        if ((text.includes('購入者さん') || text.includes('注文日')) && text.length > 50) {
          reviewLiCount++;
        }
      });

      // URLの変化をチェック（ページ番号が変わったか）
      const currentPage = getCurrentPageNumber();

      // スクロール位置の変化をチェック
      const scrollY = window.scrollY;

      // ローディング要素が消えたかチェック
      const loadingElem = document.querySelector('[class*="loading"], [class*="spinner"]');

      // 何らかの変化があれば更新完了と判断
      // （レビュー内容が変わっている、またはローディングが終わっている）
      if (!loadingElem || (Date.now() - startTime > 2000)) {
        // 少し追加で待ってから完了
        await sleep(1000);
        return true;
      }
    }

    return false; // タイムアウト
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
    let shopReply = ''; // ショップからの返信

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

      // 参考になった数を取得（新構造）
      // 「参考になった X人」「参考になった（X人）」などのパターン
      const helpfulPatterns = [
        /参考になった[：:\s]*(\d+)\s*人/,
        /参考になった[（(](\d+)[）)]/,
        /(\d+)\s*人が参考になった/,
        /参考になった\s*(\d+)/
      ];
      for (const pattern of helpfulPatterns) {
        const helpfulMatch = text.match(pattern);
        if (helpfulMatch) {
          helpfulCount = parseInt(helpfulMatch[1], 10);
          break;
        }
      }

      // ショップからの返信を取得（新構造）
      // 「ショップからのコメント」の後に続くテキストを抽出
      const shopReplyMatch = text.match(/ショップからのコメント[：:\s]*(.+?)(?=(?:参考になった|不適切レビュー|$))/s);
      if (shopReplyMatch) {
        shopReply = shopReplyMatch[1].trim();
        // 改行や余分な空白を整理
        shopReply = shopReply.replace(/\s+/g, ' ').trim();
        // 長すぎる場合は切り詰め
        if (shopReply.length > 500) {
          shopReply = shopReply.substring(0, 500) + '...';
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

      // ショップ返信（旧構造）
      const oldShopReplyElem = elem.querySelector('.revRvwShopComment, .shop-reply, [class*="shop-comment"]');
      if (oldShopReplyElem) {
        shopReply = oldShopReplyElem.textContent.trim();
        shopReply = shopReply.replace(/^ショップからのコメント[：:\s]*/i, '');
        if (shopReply.length > 500) {
          shopReply = shopReply.substring(0, 500) + '...';
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
      shopReply: shopReply,
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
   * 次のページへ移動する方法を探す
   * @returns {Object|null} { type: 'link'|'button', element: Element, url?: string } または null
   */
  function findNextPage() {
    const currentPage = getCurrentPageNumber();

    // === 方法1: 「次へ」ボタンを探す（新しい楽天UI - 最優先） ===
    // 新UIではページネーションがbutton要素で実装されている
    const nextButtonSelectors = [
      'button.navigation-button-right--3z-F_',
      'button[class*="navigation-button-right"]',
      'button[class*="next"]',
      '[class*="pagination"] button:last-child'
    ];

    for (const selector of nextButtonSelectors) {
      try {
        const button = document.querySelector(selector);
        if (button && !button.disabled) {
          const buttonText = button.textContent.trim();
          if (buttonText.includes('次へ') || buttonText === '>' || buttonText === '»') {
            return { type: 'button', element: button };
          }
        }
      } catch (e) {
        // セレクターエラーは無視
      }
    }

    // テキストで「次へ」を含むボタンを探す
    const allButtons = document.querySelectorAll('button');
    for (const button of allButtons) {
      if (button.disabled) continue;
      const text = button.textContent.trim();
      if (text === '次へ' || text.includes('次へ') || text === '>' || text === '»') {
        return { type: 'button', element: button };
      }
    }

    // === 方法2: 次のページ番号ボタンを探す ===
    const nextPageNum = currentPage + 1;
    for (const button of allButtons) {
      if (button.disabled) continue;
      const text = button.textContent.trim();
      if (text === String(nextPageNum)) {
        return { type: 'button', element: button };
      }
    }

    // === 方法3: 「次へ」リンクを探す（旧UI対応） ===
    const allLinks = document.querySelectorAll('a');
    for (const link of allLinks) {
      const text = link.textContent.trim();
      if ((text === '次へ' || text === '>' || text === '»' || text === '次' || text.includes('次のページ') || text === '>>') && link.href) {
        if (link.href !== window.location.href && link.href.includes('review.rakuten.co.jp')) {
          return { type: 'link', element: link, url: link.href };
        }
      }
    }

    // === 方法4: 次のページ番号のリンクを探す ===
    for (const link of allLinks) {
      const text = link.textContent.trim();
      if (text === String(nextPageNum) && link.href && link.href.includes('review.rakuten.co.jp')) {
        return { type: 'link', element: link, url: link.href };
      }
    }

    // === 方法5: URLパターンで次のページを構築（最後の手段） ===
    const currentUrl = window.location.href;
    const pagePattern = /\/(\d+)\.(\d+)\/?(\?.*)?$/;
    const pageMatch = currentUrl.match(pagePattern);
    if (pageMatch) {
      const pageNum = parseInt(pageMatch[1], 10);
      const sortNum = pageMatch[2];
      const query = pageMatch[3] || '';

      const totalPages = getTotalPages();

      if (pageNum < totalPages) {
        const nextUrl = currentUrl.replace(pagePattern, `/${pageNum + 1}.${sortNum}/${query}`);
        return { type: 'url', url: nextUrl };
      }
    }

    return null;
  }

  /**
   * 次のページリンクを探す（後方互換性のため残す）
   * @deprecated findNextPage() を使用してください
   */
  function findNextPageLink() {
    const next = findNextPage();
    if (!next) return null;
    if (next.type === 'link' || next.type === 'url') {
      return next.url;
    }
    // ボタンの場合はnullを返す（旧方式では対応不可）
    return null;
  }

  /**
   * 現在のページ番号を取得
   */
  function getCurrentPageNumber() {
    const url = window.location.href;

    // 新しい楽天レビューページのURL形式: ?p=N (例: ?p=2, ?p=3)
    const newPageMatch = url.match(/[?&]p=(\d+)/);
    if (newPageMatch) {
      return parseInt(newPageMatch[1], 10);
    }

    // 旧楽天レビューページのURL形式: /PAGE.SORT/ (例: /1.1/, /2.1/)
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

    // アクティブなページネーション要素から取得（新UI - aria-current="page"）
    const activeButton = document.querySelector('button[aria-current="page"]');
    if (activeButton) {
      const num = parseInt(activeButton.textContent.trim(), 10);
      if (!isNaN(num) && num > 0) {
        return num;
      }
    }

    // アクティブなページネーション要素から取得（旧UI）
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
   * レビュー総数を取得
   * @returns {number} レビュー総数（取得できない場合は0）
   */
  function getTotalReviewCount() {
    // 方法1: 新UI - メイン商品エリアの「(○○件)」を探す
    // クラス名 text-container を持つspan要素の最初のもの（メインコンテンツ）
    const textContainers = document.querySelectorAll('span[class*="text-container"]');
    for (const elem of textContainers) {
      const text = elem.textContent.trim();
      // 「(1,318件)」のような形式にマッチ（[0-9]を使用）
      const match = text.match(/([0-9,]+)件/);
      if (match && text.startsWith('(') && text.endsWith(')')) {
        const count = parseInt(match[1].replace(/,/g, ''), 10);
        if (count > 0 && count < 100000) {
          return count;
        }
      }
    }

    // 方法2: 旧UI - ページ上部の「○○件のレビュー」を探す
    const headerElements = document.querySelectorAll('h1, h2, h3, .review-count');
    for (const elem of headerElements) {
      const text = elem.textContent || '';
      const match = text.match(/([0-9,]+)\s*件/);
      if (match) {
        const count = parseInt(match[1].replace(/,/g, ''), 10);
        if (count > 0 && count < 100000) {
          return count;
        }
      }
    }

    return 0; // 取得できなかった
  }

  /**
   * 総ページ数を取得
   */
  function getTotalPages() {
    let maxPage = 1;

    // 方法1: レビュー件数から計算（1ページあたり約15件）
    const totalReviews = getTotalReviewCount();
    if (totalReviews > 0) {
      const calculatedPages = Math.ceil(totalReviews / 15);
      if (calculatedPages > maxPage) {
        maxPage = calculatedPages;
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
   * 状態を更新（pageCountとtotalPagesのみ。reviewCountとreviewsはbackground.jsで管理）
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

        // reviewsとreviewCountは変更しない（background.jsのsaveToLocalStorageで管理）
        // pageCountとtotalPagesのみ更新
        state.pageCount = (state.pageCount || 0) + 1;
        state.totalPages = getTotalPages();
        state.isRunning = isCollecting;

        chrome.storage.local.set({ collectionState: state }, () => {
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
          // 既に収集中の場合は停止してから再開始
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
          // background.jsに停止を通知（collectingItemsからの削除用）
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

      window.postMessage({ type: 'RAKUTEN_REVIEW_RESPONSE', cmd, result }, '*');
    }
  });

  // APIを注入
  injectTestAPI();

})();
