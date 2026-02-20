/**
 * コンテンツスクリプト
 * 楽天市場のレビューページからデータをスクレイピングする
 */

(function() {
  'use strict';

  // バージョン（manifest.jsonと同期）
  const VERSION = '2.0.32';
  console.log(`[楽天レビュー収集] v${VERSION} 読み込み完了 - URL: ${window.location.href.substring(0, 80)}...`);

  // 収集状態
  let isCollecting = false;
  let shouldStop = false;
  let currentProductId = ''; // 現在の商品管理番号
  let totalPages = 0; // 総ページ数（ログ表示用）
  let incrementalOnly = false; // 差分取得モード
  let lastCollectedDate = null; // 前回収集日
  let currentQueueName = null; // 定期収集のキュー名
  let enableDateFilter = false; // 期間指定フィルター
  let dateFilterFrom = null; // 期間指定：開始日
  let dateFilterTo = null; // 期間指定：終了日

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
            // 差分取得パラメータを設定
            incrementalOnly = message.incrementalOnly || false;
            lastCollectedDate = message.lastCollectedDate || null;
            currentQueueName = message.queueName || null;

            // 商品IDを取得してログに表示
            const itemUrlMatch = window.location.href.match(/item\.rakuten\.co\.jp\/[^\/]+\/([^\/\?]+)/);
            const itemProductId = itemUrlMatch ? itemUrlMatch[1] : '';
            // 定期収集の場合は[キュー名・商品ID]形式
            let prefix = '';
            if (currentQueueName && itemProductId) {
              prefix = `[${currentQueueName}・${itemProductId}] `;
            } else if (itemProductId) {
              prefix = `[${itemProductId}] `;
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
              // 差分取得設定を保存
              state.incrementalOnly = incrementalOnly;
              state.lastCollectedDate = lastCollectedDate;
              state.queueName = currentQueueName;
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
          // 差分取得パラメータを設定
          incrementalOnly = message.incrementalOnly || false;
          lastCollectedDate = message.lastCollectedDate || null;
          currentQueueName = message.queueName || null;
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
        // 商品情報を取得（キュー追加用・最小限）
        const productInfo = getProductInfo();
        sendResponse({ success: true, productInfo: productInfo });
        break;
      case 'collectProductInfo':
        // 商品情報を詳細に収集（Google Drive保存用）
        if (isItemPage) {
          // ページの主要コンテンツが読み込まれるまで待機してから収集
          waitForRakutenPageReady().then(() => {
            try {
              const data = collectRakutenProductInfo();
              sendResponse({ success: true, data });
            } catch (error) {
              sendResponse({ success: false, error: error.message });
            }
          });
        } else {
          sendResponse({ success: false, error: '楽天の商品ページ（item.rakuten.co.jp）で実行してください' });
        }
        break;
      default:
        return;
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
   * 楽天商品ページの主要コンテンツが読み込まれるまで待機
   */
  async function waitForRakutenPageReady(maxWaitMs = 20000) {
    if (document.readyState === 'loading') {
      await new Promise(resolve => {
        document.addEventListener('DOMContentLoaded', resolve, { once: true });
      });
    }

    const startTime = Date.now();
    const checkInterval = 300;

    while (Date.now() - startTime < maxWaitMs) {
      // og:title があれば基本的なページ情報は読み込み済み
      const hasTitle = !!document.querySelector('meta[property="og:title"]');
      // 画像が1つ以上表示されているか
      const hasImage = document.querySelectorAll('img[src*="r10s.jp"], img[src*="rakuten"]').length > 0;

      if (hasTitle && hasImage) {
        // 遅延読み込みの画像やJS描画コンテンツを少し待つ
        await new Promise(r => setTimeout(r, 500));
        console.log(`[楽天商品情報収集] ページ準備完了（${Date.now() - startTime}ms）`);
        return;
      }

      await new Promise(r => setTimeout(r, checkInterval));
    }

    console.warn(`[楽天商品情報収集] ${maxWaitMs}ms待機後もページが完全に読み込まれていない可能性があります`);
  }

  /**
   * 楽天商品ページから詳細な商品情報を収集（Google Drive保存用）
   * セレクターは docs/楽天商品情報収集ロジック.md に基づく
   */
  function collectRakutenProductInfo() {
    // 商品名: og:titleから加工
    const ogTitle = document.querySelector('meta[property="og:title"]');
    let title = ogTitle ? ogTitle.content : document.title;
    title = title.replace(/^【楽天市場】/, '').replace(/：[^：]+$/, '').trim();

    // URLパスからショップID・商品管理番号を抽出
    const pathParts = location.pathname.split('/').filter(Boolean);
    const shopSlug = pathParts[0] || '';
    const itemSlug = pathParts[1] || '';

    // 販売価格
    let sellingPrice = '';
    const priceEl = document.querySelector(
      '[class*="number-display--"][class*="color-crimson--"][class*="size-l--"] [class*="number--"][class*="primary--"]'
    );
    if (priceEl) {
      sellingPrice = priceEl.textContent.trim();
    } else {
      const priceFb = document.querySelector(
        '[class*="number-display--"][class*="color-crimson--"] [class*="number--"][class*="primary--"]'
      );
      if (priceFb) sellingPrice = priceFb.textContent.trim();
    }

    // 元価格（メーカー希望小売価格）
    const origEl = document.querySelector('[class*="item-original-price--"] [class*="value--"]');
    const originalPrice = origEl ? origEl.textContent.trim().replace(/円$/, '') : '';

    // レビュー情報（優先度順にフォールバック）
    // 1. AggregateRating内のitemprop（個別レビューの評価と区別するため親を限定）
    const aggRatingEl = document.querySelector('[itemtype*="AggregateRating"] [itemprop="ratingValue"]');
    const aggCountEl = document.querySelector('[itemtype*="AggregateRating"] [itemprop="reviewCount"]');
    // 2. 新UIのCSS moduleセレクタ
    const cssScoreEl = document.querySelector('[class*="review-score--"]');
    const cssTotalEl = document.querySelector('[class*="review-total--"]');

    // rating: AggregateRating itemprop → CSS module → 空
    const scoreEl = (aggRatingEl?.content)
      ? { textContent: aggRatingEl.content }
      : cssScoreEl || null;

    // reviewCount: AggregateRating itemprop → CSS module → 旧UI件数リンク → ショップレビュー
    let totalEl = (aggCountEl?.content)
      ? { textContent: aggCountEl.content + '件' }
      : cssTotalEl || null;
    if (!totalEl) {
      const reviewLink = document.querySelector('.normal-reserve-review a:not([aria-label="レビューを書く"])');
      if (reviewLink) {
        totalEl = { textContent: reviewLink.textContent.trim() };
      }
    }
    if (!totalEl) {
      const rnkReview = document.querySelector('.rnkInShopItemReview span');
      if (rnkReview) {
        totalEl = rnkReview;
      }
    }

    // ショップ名
    const titleParts = document.title.split('：');
    const shopName = titleParts.length > 1 ? titleParts[titleParts.length - 1].trim() : '';

    // カテゴリ（JSON-LD BreadcrumbListから）
    let categories = [];
    try {
      const ldScripts = document.querySelectorAll('script[type="application/ld+json"]');
      for (const ldScript of ldScripts) {
        const data = JSON.parse(ldScript.textContent);
        if (data.itemListElement) {
          categories = data.itemListElement
            .map(item => item.item?.name || '')
            .filter(name => name && name !== '楽天市場');
          break;
        }
      }
    } catch (e) {}

    // バリエーション
    const varButtons = document.querySelectorAll('[class*="button-multiline--"]');
    const variations = Array.from(varButtons)
      .map(b => b.textContent.trim().replace(/\s+/g, ' '))
      .filter(t => t.length > 0 && t.length < 50 && t !== '―')
      // 末尾の価格パターンを除去（例: "【XL】27-30cm2,480円" → "【XL】27-30cm"）
      .map(t => t.replace(/\d{1,3}(,\d{3})*円$/, '').trim());

    // 商品説明テキスト（.item_desc → .sale_desc → og:description のフォールバック）
    function extractDescText(container) {
      if (!container) return '';
      const clone = container.cloneNode(true);
      clone.querySelectorAll('style, script').forEach(el => el.remove());
      return clone.textContent.trim()
        .replace(/<!--[\s\S]*?-->/g, '')
        .replace(/\s{3,}/g, '\n')
        .trim();
    }
    let description = extractDescText(document.querySelector('.item_desc'));
    if (!description) {
      description = extractDescText(document.querySelector('.sale_desc'));
    }
    if (!description) {
      const ogDesc = document.querySelector('meta[property="og:description"]');
      if (ogDesc?.content) description = ogDesc.content.trim();
    }

    // 商品画像の収集
    const images = [];
    const seen = new Set();

    // 高解像度URLに変換するヘルパー（すべてshop.r10s.jpに統一、クエリ除去）
    function toHighResUrl(src) {
      return src.split('?')[0]
        .replace('tshop.r10s.jp', 'shop.r10s.jp')
        .replace('image.rakuten.co.jp', 'shop.r10s.jp')
        .replace('thumbnail.image.rakuten.co.jp', 'shop.r10s.jp');
    }

    // 1. og:image（メイン画像）
    const ogImage = document.querySelector('meta[property="og:image"]');
    if (ogImage?.content) {
      const url = toHighResUrl(ogImage.content);
      images.push({ url, type: 'main' });
      seen.add(url);
    }

    // 2. ギャラリー画像（商品画像ギャラリーコンテナ内のみ）
    // ページ上部の商品画像スライダー内の画像だけを収集（おすすめ・ランキング等を除外）
    document.querySelectorAll('[class*="r-image--"] img, [class*="image-wrapper--"] img').forEach(img => {
      const src = img.src || '';
      if (!src || src.includes('data:image')) return;
      const url = toHighResUrl(src);
      if (seen.has(url)) return;
      if (img.naturalWidth > 0 && img.naturalWidth <= 2) return;
      images.push({ url, type: 'gallery' });
      seen.add(url);
    });

    // 3. 商品説明欄の画像（.item_desc / .sale_desc）
    const descContainers = document.querySelectorAll('.item_desc, .sale_desc');
    descContainers.forEach(container => {
      container.querySelectorAll('img').forEach(img => {
        const src = img.src || '';
        if (!src || src.includes('data:image')) return;
        const url = toHighResUrl(src);
        if (seen.has(url)) return;
        if (img.naturalWidth > 0 && img.naturalWidth <= 2) return;
        images.push({ url, type: 'description' });
        seen.add(url);
      });
    });

    // 送料情報
    let shipping = '';
    document.querySelectorAll('[class*="text-display--"]').forEach(el => {
      if (el.textContent.trim() === '送料無料') shipping = '送料無料';
    });

    // 動画の収集（楽天の商品ページにはvideo要素やiframe埋め込みがある場合がある）
    const videos = [];
    const videoSeen = new Set();

    // video要素
    document.querySelectorAll('video source, video[src]').forEach(el => {
      const src = el.src || el.getAttribute('src') || '';
      if (src && !videoSeen.has(src)) {
        videos.push({ url: src, type: src.includes('.m3u8') ? 'hls' : 'mp4', source: 'video-element' });
        videoSeen.add(src);
      }
    });

    // 商品説明欄内の動画（.item_desc / .sale_desc）
    document.querySelectorAll('.item_desc video source, .item_desc video[src], .sale_desc video source, .sale_desc video[src]').forEach(el => {
      const src = el.src || el.getAttribute('src') || '';
      if (src && !videoSeen.has(src)) {
        videos.push({ url: src, type: src.includes('.m3u8') ? 'hls' : 'mp4', source: 'description' });
        videoSeen.add(src);
      }
    });

    // スクリプトタグ内のMP4 URLを探す（blob: URLの代わりに実際のURLを取得）
    document.querySelectorAll('script:not([src])').forEach(script => {
      const text = script.textContent || '';
      if (text.length > 500000) return;
      const mp4Regex = /https?:\/\/[^"'\s,]+\.mp4[^"'\s,]*/g;
      let match;
      while ((match = mp4Regex.exec(text)) !== null) {
        const url = match[0];
        if (!videoSeen.has(url)) {
          videos.push({ url, type: 'mp4', source: 'script' });
          videoSeen.add(url);
        }
      }
    });

    // 動画サムネイル（楽天は動画ボタン付きの画像がある場合）
    const videoThumbnails = [];
    document.querySelectorAll('[class*="video"] img, [data-video] img').forEach(img => {
      const src = img.src || '';
      if (src && !src.includes('data:image')) {
        videoThumbnails.push(toHighResUrl(src));
      }
    });

    return {
      source: 'rakuten',
      title,
      shopName,
      shopSlug,
      itemSlug,
      url: location.href.split('?')[0],
      price: sellingPrice ? sellingPrice + '円' : '',
      listPrice: originalPrice ? originalPrice + '円' : '',
      rating: scoreEl ? scoreEl.textContent.trim() : '',
      reviewCount: totalEl ? totalEl.textContent.trim().replace(/[()（）]/g, '') : '',
      categories,
      variations,
      description: description.substring(0, 5000),
      images,
      videos: videos.length > 0 ? videos : undefined,
      videoThumbnails: videoThumbnails.length > 0 ? videoThumbnails : undefined,
      shipping,
      collectedAt: new Date().toISOString()
    };
  }

  /**
   * 新着順ソートを確認・設定
   * URLにsort=6パラメータを追加して直接遷移（セレクトボックス操作より確実）
   * @returns {boolean} ソートを変更した場合はtrue（ページ遷移が発生する）
   */
  async function ensureNewestSort() {
    const currentUrl = new URL(window.location.href);
    const currentSort = currentUrl.searchParams.get('sort');

    // 既にsort=6の場合は何もしない
    if (currentSort === '6') {
      console.log('[楽天レビュー収集] URLで既に新着順です（sort=6）');
      return false;
    }

    console.log('[楽天レビュー収集] 新着順に変更します（現在のsort: ' + (currentSort || 'なし') + '）');

    // 収集状態を保存（ページ遷移後に再開するため）
    await new Promise(resolve => {
      chrome.storage.local.get(['collectionState'], (result) => {
        const state = result.collectionState || {};
        state.isRunning = true;
        state.incrementalOnly = incrementalOnly;
        state.lastCollectedDate = lastCollectedDate;
        state.queueName = currentQueueName;
        state.source = 'rakuten';
        chrome.storage.local.set({ collectionState: state }, resolve);
      });
    });

    log('新着順に並び替えています...');

    // URLにsort=6を追加して遷移（確実にページ遷移が発生する）
    currentUrl.searchParams.set('sort', '6');
    window.location.href = currentUrl.toString();

    // ページ遷移が発生するので、新しいページで収集が再開される
    return true;
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

    // 設定を読み込み
    const settings = await new Promise(resolve => {
      chrome.storage.sync.get(['enableDateFilter', 'dateFilterFrom', 'dateFilterTo'], resolve);
    });

    // 期間指定フィルター設定をグローバル変数に反映
    enableDateFilter = settings.enableDateFilter || false;
    dateFilterFrom = settings.dateFilterFrom || null;
    dateFilterTo = settings.dateFilterTo || null;

    console.log('[楽天レビュー収集] 期間指定設定:', { enableDateFilter, dateFilterFrom, dateFilterTo });

    // 常に新着順ソートを使用（ページ上のセレクトボックスを操作）
    const sortChanged = await ensureNewestSort();
    if (sortChanged) {
      // ソート変更後、ページが更新されるので収集は新しいインスタンスで再開される
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
      // 総ページ数を計算（1ページ30件）
      totalPages = Math.ceil(expectedTotal / 30);
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

    // 差分取得で古いレビューに到達した場合は早期終了
    if (reachedOldReviews) {
      log('前回以降の新着レビューの収集が完了しました');
      isCollecting = false;
      if (!shouldStop) {
        chrome.runtime.sendMessage({ action: 'collectionComplete' });
      }
      return;
    }

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
   * @returns {boolean} 差分取得で古いレビューに到達した場合、または期間外に到達した場合はtrue
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
      reviews = filterReviewsByDate(reviews, lastCollectedDate, null);
      const filteredCount = originalCount - reviews.length;

      if (filteredCount > 0) {
        log(`${originalCount}件中${filteredCount}件は前回収集済み（${lastCollectedDate}より前）`);
      }

      // 全てのレビューが古い場合、これ以上ページを進める必要がない
      if (reviews.length === 0) {
        log('前回以降の新着レビューがこのページにはありません');
        reachedOldReviews = true;
        return reachedOldReviews;
      }

      // 一部のレビューが古い場合、次のページはさらに古いので終了フラグを立てる
      if (filteredCount > 0 && reviews.length < originalCount) {
        reachedOldReviews = true;
      }
    }

    // 期間指定フィルターが有効な場合
    if (enableDateFilter && (dateFilterFrom || dateFilterTo)) {
      const originalCount = reviews.length;
      reviews = filterReviewsByDateRange(reviews, dateFilterFrom, dateFilterTo);
      const filteredCount = originalCount - reviews.length;

      // 新着順かつ開始日指定ありの場合、期間外レビューが1件でも出てきたら早期終了
      // （新着順なので、これ以降のページは全て開始日より古いため）
      if (filteredCount > 0 && dateFilterFrom) {
        const currentUrl = new URL(window.location.href);
        const isNewestFirst = currentUrl.searchParams.get('sort') === '6';
        if (isNewestFirst) {
          const rangeText = dateFilterTo
            ? `${dateFilterFrom}〜${dateFilterTo}`
            : `${dateFilterFrom}以降`;
          if (reviews.length > 0) {
            log(`${originalCount}件中${reviews.length}件が期間内（${rangeText}）- このページで収集を終了`);
          } else {
            log(`指定期間（${rangeText}）より古いレビューに到達 - 収集を終了`);
          }
          reachedOldReviews = true;
        }
      }
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

    return reachedOldReviews;
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
        // 現在のページ番号を記録（変更検知用）
        const currentPageNum = getCurrentPageNumber();

        // ボタンをクリック
        nextPage.element.click();

        // コンテンツの更新を待つ（ページ番号が変わるまで）
        const contentUpdated = await waitForContentUpdate(currentPageNum);

        if (!contentUpdated) {
          log('最終ページに到達しました');
          return false;
        }

        // 新しいページのレビューを収集
        const reachedOldReviews = await collectCurrentPage();

        // 差分取得で古いレビューに到達した場合は早期終了
        if (reachedOldReviews) {
          log('前回以降の新着レビューの収集が完了しました');
          return false;
        }

        // 次のループで次のページを探す
        continue;

      } else if (nextPage.type === 'link' || nextPage.type === 'url') {
        // リンク/URL方式（旧UI）
        // 差分取得設定を保存してからページ遷移
        await new Promise((resolve) => {
          chrome.storage.local.get(['collectionState'], (result) => {
            const state = result.collectionState || {};
            state.incrementalOnly = incrementalOnly;
            state.lastCollectedDate = lastCollectedDate;
            chrome.storage.local.set({ collectionState: state }, resolve);
          });
        });

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
   * @param {number} initialPageNum 更新前のページ番号
   * @param {number} timeout タイムアウト（ミリ秒）
   * @returns {boolean} 更新されたらtrue、タイムアウトしたらfalse
   */
  async function waitForContentUpdate(initialPageNum, timeout = 8000) {
    const startTime = Date.now();
    const checkInterval = 300; // 300msごとにチェック

    while (Date.now() - startTime < timeout) {
      await sleep(checkInterval);

      // ページ番号の変化をチェック（最も確実な方法）
      const currentPage = getCurrentPageNumber();
      if (currentPage > initialPageNum) {
        // ページ番号が増えた = 次のページに移動した
        await sleep(500); // コンテンツ読み込み完了を少し待つ
        return true;
      }

      // URLの変化もチェック
      const urlPageMatch = window.location.href.match(/[?&]p=([0-9]+)/);
      const urlPage = urlPageMatch ? parseInt(urlPageMatch[1], 10) : 1;
      if (urlPage > initialPageNum) {
        await sleep(500);
        return true;
      }
    }

    // タイムアウト = 最終ページ（ページ番号が変わらなかった）
    return false;
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
   * 日付を比較可能な形式（YYYY-MM-DD）に正規化
   * 2024/1/5 → 2024-01-05
   */
  function normalizeDateString(dateStr) {
    if (!dateStr) return null;

    // YYYY/M/D または YYYY/MM/DD 形式を YYYY-MM-DD に変換
    const match = dateStr.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
    if (match) {
      const year = match[1];
      const month = match[2].padStart(2, '0');
      const day = match[3].padStart(2, '0');
      return `${year}-${month}-${day}`;
    }

    // すでに YYYY-MM-DD 形式の場合
    const isoMatch = dateStr.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) {
      return dateStr.substring(0, 10);
    }

    return null;
  }

  /**
   * レビューを日付でフィルタリング（差分取得用）
   * @param {Array} reviews - レビュー配列
   * @param {string} afterDate - この日付より後のレビューのみ取得（YYYY-MM-DD形式）
   * @returns {Array} フィルタリングされたレビュー
   */
  function filterReviewsByDate(reviews, afterDate) {
    if (!afterDate) return reviews;

    const normalizedAfterDate = normalizeDateString(afterDate);
    if (!normalizedAfterDate) return reviews;

    return reviews.filter(review => {
      const reviewDateNorm = normalizeDateString(review.reviewDate);
      if (!reviewDateNorm) {
        // 日付がないレビューは含める（安全側に倒す）
        return true;
      }
      // レビュー日が前回収集日以降の場合のみ含める（同日も含む）
      return reviewDateNorm >= normalizedAfterDate;
    });
  }

  /**
   * レビューを期間でフィルタリング（期間指定用）
   * @param {Array} reviews - レビュー配列
   * @param {string|null} fromDate - 開始日（YYYY-MM-DD形式、nullの場合は制限なし）
   * @param {string|null} toDate - 終了日（YYYY-MM-DD形式、nullの場合は制限なし）
   * @returns {Array} フィルタリングされたレビュー
   */
  function filterReviewsByDateRange(reviews, fromDate, toDate) {
    if (!fromDate && !toDate) return reviews;

    const normalizedFromDate = fromDate ? normalizeDateString(fromDate) : null;
    const normalizedToDate = toDate ? normalizeDateString(toDate) : null;

    return reviews.filter(review => {
      const reviewDateNorm = normalizeDateString(review.reviewDate);
      if (!reviewDateNorm) {
        // 日付がないレビューは除外（期間指定時は日付不明は収集しない）
        return false;
      }

      // 開始日チェック
      if (normalizedFromDate && reviewDateNorm < normalizedFromDate) {
        return false;
      }

      // 終了日チェック
      if (normalizedToDate && reviewDateNorm > normalizedToDate) {
        return false;
      }

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
   * ログをポップアップに送信（商品管理番号・ページ番号を自動プレフィックス）
   */
  function log(text, type = '') {
    // 定期収集の場合は[キュー名・商品ID]形式
    let productPrefix = '';
    if (currentQueueName && currentProductId) {
      productPrefix = `[${currentQueueName}・${currentProductId}] `;
    } else if (currentProductId) {
      productPrefix = `[${currentProductId}] `;
    }
    // ページ情報を追加（総ページ数が設定されている場合のみ）
    const currentPage = getCurrentPageNumber();
    const pagePrefix = totalPages > 0 ? `[${currentPage}/${totalPages}] ` : '';
    const fullText = productPrefix + pagePrefix + text;
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
        // 差分取得設定を復元
        incrementalOnly = state.incrementalOnly || false;
        lastCollectedDate = state.lastCollectedDate || null;
        currentQueueName = state.queueName || null;
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
