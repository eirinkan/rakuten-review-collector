/**
 * Amazon商品ページから商品情報を収集するコンテンツスクリプト
 * 既存のcontent-amazon.js（レビュー収集）とは独立して動作
 *
 * 収集データ: 商品名、価格、画像URL、箇条書き、商品説明、A+コンテンツ、技術仕様、カテゴリ等
 */

(() => {
  'use strict';

  // 多重登録防止
  if (window.__amazonProductCollectorLoaded) return;
  window.__amazonProductCollectorLoaded = true;

  const VERSION = '1.0.0';

  // ===== セレクター定義（フォールバック付き） =====
  const SELECTORS = {
    // 商品名
    productTitle: [
      '#productTitle',
      '#title span',
      'h1.product-title-word-break'
    ],
    // 価格
    price: [
      '#corePriceDisplay_desktop_feature_div .a-price .a-offscreen',
      '#corePrice_feature_div .a-price .a-offscreen',
      '#tp_price_block_total_price_ww .a-price .a-offscreen',
      '#priceblock_ourprice',
      '#priceblock_dealprice',
      '.a-color-price',
      '.a-price .a-offscreen'
    ],
    // セール価格（元価格）
    listPrice: [
      '.a-price[data-a-strike="true"] .a-offscreen',
      '#listPrice',
      '.priceBlockStrikePriceString'
    ],
    // メイン画像
    mainImage: [
      '#landingImage',
      '#imgBlkFront',
      '#main-image'
    ],
    // ギャラリーサムネイル
    galleryThumbs: [
      '#altImages li.a-spacing-small.item img',
      '.imageThumbnail img'
    ],
    // ブランド
    brand: [
      '#bylineInfo',
      '.po-brand .a-span9 .a-size-base'
    ],
    // 箇条書き（Feature Bullets）
    featureBullets: [
      '#feature-bullets .a-list-item',
      '#feature-bullets li'
    ],
    // 商品説明
    productDescription: [
      '#productDescription',
      '#productDescription_feature_div'
    ],
    // A+コンテンツ
    aplus: [
      '#aplus_feature_div',
      '#aplus',
      '#aplusProductDescription'
    ],
    // 技術仕様テーブル
    techSpecs: [
      '#productDetails_techSpec_section_1',
      '#prodDetails table.a-keyvalue',
      '#technicalSpecifications_section_1'
    ],
    // 商品詳細（箇条書き形式）
    detailBullets: [
      '#detailBullets_feature_div .a-list-item',
      '#detail-bullets .content li'
    ],
    // カテゴリ（パンくずリスト）
    breadcrumbs: [
      '#wayfinding-breadcrumbs_feature_div a.a-link-normal'
    ],
    // 評価
    rating: [
      '#acrPopover .a-icon-alt',
      '.a-icon-star .a-icon-alt'
    ],
    // レビュー数
    reviewCount: [
      '#acrCustomerReviewText'
    ],
    // バリエーション
    variations: [
      '#variation_color_name .selection',
      '#variation_size_name .selection',
      '#twister .a-button-selected .a-button-text'
    ]
  };

  /**
   * セレクターリストから最初に見つかった要素を返す
   */
  function queryFirst(selectorList) {
    for (const sel of selectorList) {
      const el = document.querySelector(sel);
      if (el) return el;
    }
    return null;
  }

  /**
   * セレクターリストからテキスト内容がある最初の要素を返す
   * 価格など、要素は存在するがテキストが空のケースに対応
   */
  function queryFirstWithText(selectorList) {
    for (const sel of selectorList) {
      const el = document.querySelector(sel);
      if (el && el.textContent.trim()) return el;
    }
    return null;
  }

  /**
   * セレクターリストから全要素を返す
   */
  function queryAllFirst(selectorList) {
    for (const sel of selectorList) {
      const els = document.querySelectorAll(sel);
      if (els.length > 0) return Array.from(els);
    }
    return [];
  }

  /**
   * URLからASINを抽出
   */
  function extractASIN(url) {
    if (!url) url = location.href;
    const dpMatch = url.match(/\/dp\/([A-Z0-9]{10})/i);
    if (dpMatch) return dpMatch[1].toUpperCase();
    const reviewMatch = url.match(/\/product-reviews\/([A-Z0-9]{10})/i);
    if (reviewMatch) return reviewMatch[1].toUpperCase();
    const gpMatch = url.match(/\/gp\/product\/([A-Z0-9]{10})/i);
    if (gpMatch) return gpMatch[1].toUpperCase();
    return '';
  }

  /**
   * サムネイルURLを高解像度URLに変換
   * Amazon画像URLのサイズ指定部分を大サイズに置換
   * 例: ._AC_US100_.jpg → ._AC_SL1500_.jpg
   *     ._SX300_.jpg → ._SL1500_.jpg
   *     .__CR0,0,970,600_PT0_SX970_V1___.jpg → ._SL1500_.jpg
   */
  function toHighResUrl(url) {
    if (!url) return '';
    // ファイル名中の「.」から拡張子前の「.」までのサイズ指定部分を丸ごと置換
    // 例: 51zKH8u9AlL._AC_US100_.jpg → 51zKH8u9AlL._SL1500_.jpg
    return url.replace(/\._{1,2}[A-Z]{2}[^.]+_*\./, '._SL1500_.');
  }

  /**
   * 商品ページからすべての画像URLを収集
   */
  function collectImageUrls() {
    const images = [];
    const seen = new Set();

    // 1. メイン画像（data-old-hires > data-a-dynamic-image > src）
    const mainImg = queryFirst(SELECTORS.mainImage);
    if (mainImg) {
      // data-old-hires が最高解像度
      const hiRes = mainImg.getAttribute('data-old-hires');
      if (hiRes && !seen.has(hiRes)) {
        images.push({ url: hiRes, type: 'main' });
        seen.add(hiRes);
      }

      // data-a-dynamic-image からも取得（複数サイズ）
      const dynamicAttr = mainImg.getAttribute('data-a-dynamic-image');
      if (dynamicAttr) {
        try {
          const dynamicImages = JSON.parse(dynamicAttr);
          // 最大サイズのURLを選択
          let maxUrl = '';
          let maxSize = 0;
          for (const [url, dims] of Object.entries(dynamicImages)) {
            const size = Array.isArray(dims) ? dims[0] * dims[1] : 0;
            if (size > maxSize) {
              maxSize = size;
              maxUrl = url;
            }
          }
          if (maxUrl && !seen.has(maxUrl) && !hiRes) {
            images.push({ url: maxUrl, type: 'main' });
            seen.add(maxUrl);
          }
        } catch (e) {
          // JSON解析エラーは無視
        }
      }

      // フォールバック: src
      if (images.length === 0 && mainImg.src && !seen.has(mainImg.src)) {
        images.push({ url: mainImg.src, type: 'main' });
        seen.add(mainImg.src);
      }
    }

    // 2. ギャラリー画像（サムネイルを高解像度に変換）
    const thumbs = queryAllFirst(SELECTORS.galleryThumbs);
    for (const thumb of thumbs) {
      const thumbSrc = thumb.src || '';
      // 動画のサムネイルはスキップ（動画アイコンが含まれる場合）
      const parentLi = thumb.closest('li');
      if (parentLi && parentLi.classList.contains('videoThumbnail')) continue;

      const hiRes = toHighResUrl(thumbSrc);
      if (hiRes && !seen.has(hiRes)) {
        images.push({ url: hiRes, type: 'gallery' });
        seen.add(hiRes);
      }
    }

    return images;
  }

  /**
   * 商品ページから動画URL・サムネイルを収集
   * Amazon動画はHLS/DASHストリーミングの場合が多く、直接ダウンロードできないケースがある
   */
  function collectVideoUrls() {
    const videos = [];
    const seen = new Set();

    // 方法1: ページ内スクリプトからMP4 URLを探す
    // Amazonはvideoデータをscriptタグ内のJSONに埋め込むことがある
    const scripts = document.querySelectorAll('script:not([src])');
    for (const script of scripts) {
      const text = script.textContent || '';
      if (text.length > 500000) continue; // 大きすぎるスクリプトはスキップ

      // mp4 URLを抽出
      const mp4Regex = /https?:\/\/[^"'\s,]+\.mp4[^"'\s,]*/g;
      let match;
      while ((match = mp4Regex.exec(text)) !== null) {
        const url = match[0].replace(/\\u002F/g, '/').replace(/\\/g, '');
        if (!seen.has(url) && !url.includes('thumbnail') && !url.includes('preview')) {
          videos.push({ url, type: 'mp4', source: 'script' });
          seen.add(url);
        }
      }

      // m3u8 (HLS) URLも記録
      const hlsRegex = /https?:\/\/[^"'\s,]+\.m3u8[^"'\s,]*/g;
      while ((match = hlsRegex.exec(text)) !== null) {
        const url = match[0].replace(/\\u002F/g, '/').replace(/\\/g, '');
        if (!seen.has(url)) {
          videos.push({ url, type: 'hls', source: 'script' });
          seen.add(url);
        }
      }
    }

    // 方法2: video要素から直接取得
    const videoElements = document.querySelectorAll('video source, video[src]');
    for (const el of videoElements) {
      const src = el.src || el.getAttribute('src') || '';
      if (src && !seen.has(src)) {
        const type = src.includes('.m3u8') ? 'hls' : 'mp4';
        videos.push({ url: src, type, source: 'video-element' });
        seen.add(src);
      }
    }

    // 動画サムネイル画像を収集（動画がダウンロードできない場合のフォールバック用）
    const videoThumbs = document.querySelectorAll('li.videoThumbnail img');
    const thumbnails = [];
    for (const thumb of videoThumbs) {
      const src = thumb.src || '';
      if (src) {
        thumbnails.push(toHighResUrl(src));
      }
    }

    return { videos, thumbnails };
  }

  /**
   * A+コンテンツから画像URLとテキストを収集
   */
  function collectAplusContent() {
    const aplusEl = queryFirst(SELECTORS.aplus);
    if (!aplusEl || aplusEl.textContent.trim().length < 10) {
      return null;
    }

    const result = {
      text: '',
      images: []
    };

    // テキスト収集（重要なテキスト要素のみ）
    const textElements = aplusEl.querySelectorAll('h1, h2, h3, h4, h5, p, li, td, span.a-text-bold');
    const texts = [];
    const seenTexts = new Set();
    for (const el of textElements) {
      const text = el.textContent.trim();
      if (text && text.length > 2 && !seenTexts.has(text)) {
        texts.push(text);
        seenTexts.add(text);
      }
    }
    result.text = texts.join('\n');

    // 画像収集
    const imgs = aplusEl.querySelectorAll('img');
    const seen = new Set();
    for (const img of imgs) {
      // data-src（遅延読み込み）またはsrc
      const src = img.getAttribute('data-src') || img.src || '';
      if (!src || src.includes('data:image') || src.includes('transparent-pixel')) continue;

      const hiRes = toHighResUrl(src);
      if (hiRes && !seen.has(hiRes)) {
        result.images.push(hiRes);
        seen.add(hiRes);
      }
    }

    return result;
  }

  /**
   * 技術仕様を収集
   */
  function collectTechSpecs() {
    const specs = {};

    // テーブル形式の仕様
    const techTable = queryFirst(SELECTORS.techSpecs);
    if (techTable) {
      const rows = techTable.querySelectorAll('tr');
      for (const row of rows) {
        const th = row.querySelector('th');
        const td = row.querySelector('td');
        if (th && td) {
          const key = th.textContent.trim().replace(/\s+/g, ' ');
          const value = td.textContent.trim().replace(/\s+/g, ' ');
          if (key && value) {
            specs[key] = value;
          }
        }
      }
    }

    // 箇条書き形式の詳細情報
    const detailItems = queryAllFirst(SELECTORS.detailBullets);
    for (const item of detailItems) {
      const text = item.textContent.trim().replace(/\s+/g, ' ');
      // 「ラベル : 値」形式を分割
      const colonIndex = text.indexOf(':');
      if (colonIndex > 0 && colonIndex < text.length - 1) {
        const key = text.substring(0, colonIndex).trim();
        const value = text.substring(colonIndex + 1).trim();
        if (key && value && !key.includes('カスタマーレビュー')) {
          specs[key] = value;
        }
      }
    }

    return Object.keys(specs).length > 0 ? specs : null;
  }

  /**
   * 商品ページからすべての情報を収集
   */
  function collectProductInfo() {
    console.log(`[Amazon商品情報収集 v${VERSION}] 情報収集を開始`);

    const asin = extractASIN();
    if (!asin) {
      console.error('[Amazon商品情報収集] ASINを取得できませんでした');
      return null;
    }

    // 商品名
    const titleEl = queryFirst(SELECTORS.productTitle);
    const title = titleEl ? titleEl.textContent.trim() : '';

    // 価格（テキストが空の要素をスキップ）
    const priceEl = queryFirstWithText(SELECTORS.price);
    const price = priceEl ? priceEl.textContent.trim() : '';

    // 元価格（セール時）
    const listPriceEl = queryFirstWithText(SELECTORS.listPrice);
    const listPrice = listPriceEl ? listPriceEl.textContent.trim() : '';

    // ブランド
    const brandEl = queryFirst(SELECTORS.brand);
    let brand = brandEl ? brandEl.textContent.trim() : '';
    // 「ブランド: xxx」「ストア: xxx」からブランド名のみ抽出
    brand = brand.replace(/^(ブランド|ストア|Brand)[：:]\s*/i, '').replace(/のストアを表示$/, '').trim();

    // 箇条書き
    const bulletEls = queryAllFirst(SELECTORS.featureBullets);
    const bullets = bulletEls
      .map(el => el.textContent.trim())
      .filter(text => text.length > 0 && !text.includes('この商品について'));

    // 商品説明
    const descEl = queryFirst(SELECTORS.productDescription);
    const description = descEl ? descEl.textContent.trim() : '';

    // A+コンテンツ
    const aplusContent = collectAplusContent();

    // 技術仕様
    const techSpecs = collectTechSpecs();

    // カテゴリ
    const breadcrumbEls = queryAllFirst(SELECTORS.breadcrumbs);
    const categories = breadcrumbEls.map(el => el.textContent.trim()).filter(t => t.length > 0);

    // 評価
    const ratingEl = queryFirst(SELECTORS.rating);
    const rating = ratingEl ? ratingEl.textContent.trim() : '';

    // レビュー数
    const reviewCountEl = queryFirst(SELECTORS.reviewCount);
    const reviewCount = reviewCountEl ? reviewCountEl.textContent.trim() : '';

    // バリエーション
    const variationEls = queryAllFirst(SELECTORS.variations);
    const variations = variationEls.map(el => el.textContent.trim()).filter(t => t.length > 0);

    // 画像URL一覧
    const images = collectImageUrls();

    // A+コンテンツの画像もimagesに含める
    const aplusImages = aplusContent?.images || [];

    // 動画URL
    const videoData = collectVideoUrls();

    const productData = {
      asin,
      title,
      url: `https://www.amazon.co.jp/dp/${asin}`,
      price,
      listPrice: listPrice || undefined,
      brand: brand || undefined,
      categories: categories.length > 0 ? categories : undefined,
      rating: rating || undefined,
      reviewCount: reviewCount || undefined,
      variations: variations.length > 0 ? variations : undefined,
      bullets: bullets.length > 0 ? bullets : undefined,
      description: description || undefined,
      aplusText: aplusContent?.text || undefined,
      techSpecs: techSpecs || undefined,
      images,
      aplusImages: aplusImages.length > 0 ? aplusImages : undefined,
      videos: videoData.videos.length > 0 ? videoData.videos : undefined,
      videoThumbnails: videoData.thumbnails.length > 0 ? videoData.thumbnails : undefined,
      collectedAt: new Date().toISOString()
    };

    // undefinedのキーを除去
    const cleanData = JSON.parse(JSON.stringify(productData));

    console.log(`[Amazon商品情報収集] 収集完了: ${title?.substring(0, 50)}... (画像: ${images.length}枚, A+画像: ${aplusImages.length}枚, 動画: ${videoData.videos.length}件)`);

    return cleanData;
  }

  // ===== メッセージハンドラー =====
  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (!message || !message.action) return;

    switch (message.action) {
      case 'collectProductInfo': {
        try {
          const data = collectProductInfo();
          if (data) {
            sendResponse({ success: true, data });
          } else {
            sendResponse({ success: false, error: 'Amazon商品ページではないか、情報を取得できませんでした' });
          }
        } catch (error) {
          console.error('[Amazon商品情報収集] エラー:', error);
          sendResponse({ success: false, error: error.message });
        }
        return true; // 非同期レスポンス
      }

      case 'isProductPage': {
        // 商品ページかどうかを判定
        const url = location.href;
        const isProduct = url.includes('amazon.co.jp') && (
          url.includes('/dp/') ||
          url.includes('/gp/product/')
        );
        sendResponse({ isProductPage: isProduct, asin: extractASIN() });
        return false;
      }
    }
  });

  console.log(`[Amazon商品情報収集 v${VERSION}] コンテンツスクリプト読み込み完了`);
})();
