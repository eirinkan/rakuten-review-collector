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

  const VERSION = '1.1.0';

  /**
   * スマホ版ページかどうかを判定
   * AmazonはUAに基づいてサーバーサイドで異なるHTMLを返す
   */
  function isMobilePage() {
    // PC版にのみ存在する要素を先にチェック
    if (document.querySelector('#dp-container')) return false;
    if (document.querySelector('#landingImage')) return false;
    if (document.querySelector('#corePriceDisplay_desktop_feature_div')) return false;
    // スマホ版にのみ存在する要素（IDバリエーションも含む）
    if (document.querySelector('#corePriceDisplay_mobile_feature_div')) return true;
    if (document.querySelector('#corePrice_mobile_feature_div')) return true;
    if (document.querySelector('#mobile_buybox')) return true;
    if (document.querySelector('#mobile_buybox_feature_div')) return true;
    if (document.querySelector('#immersive-main')) return true;
    return false;
  }

  // ===== セレクター定義（PC版・スマホ版両対応、フォールバック付き） =====
  const SELECTORS = {
    // 商品名（スマホ版: span#title、PC版: #productTitle）
    productTitle: [
      '#productTitle',
      'span#title',
      '#title span',
      'h1.product-title-word-break'
    ],
    // 価格（スマホ版: mobile用div、PC版: desktop用div）
    price: [
      '#corePriceDisplay_desktop_feature_div .a-price .a-offscreen',
      '#corePriceDisplay_mobile_feature_div .a-price .a-offscreen',
      '#corePrice_mobile_feature_div .a-price .a-offscreen',
      '#corePrice_feature_div .a-price .a-offscreen',
      '#tp_price_block_total_price_ww .a-price .a-offscreen',
      '#mobile_buybox .a-price .a-offscreen',
      '#mobile_buybox_feature_div .a-price .a-offscreen',
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
    // メイン画像（スマホ版: immersive-view、PC版: landingImage）
    mainImage: [
      '#landingImage',
      '#imgBlkFront',
      '#main-image',
      '#image-block-iv-product-image-0'
    ],
    // ギャラリーサムネイル（スマホ版: imageGallery、PC版: altImages）
    galleryThumbs: [
      '#altImages li.a-spacing-small.item img',
      '.imageThumbnail img',
      '#imageGallery img.product-image',
      '#imageGallery_feature_div img[src*="images-amazon.com"]',
      '#imageGallery_feature_div img[src*="m.media-amazon.com"]',
      '[id^="image-block-product-image-"] img'
    ],
    // ブランド
    brand: [
      '#bylineInfo',
      '.po-brand .a-span9 .a-size-base'
    ],
    // 箇条書き（Feature Bullets）— PC/スマホ共通
    featureBullets: [
      '#feature-bullets .a-list-item',
      '#feature-bullets li'
    ],
    // 商品説明
    productDescription: [
      '#productDescription',
      '#productDescription_feature_div',
      '#productDescription_RI'
    ],
    // A+コンテンツ — PC/スマホ共通（classが異なるがIDは同じ）
    aplus: [
      '#aplus_feature_div',
      '#aplus',
      '#aplusProductDescription'
    ],
    // 技術仕様テーブル（スマホ版: additionalProductDetails、PC版: productDetails_techSpec）
    techSpecs: [
      '#productDetails_techSpec_section_1',
      '#prodDetails table.a-keyvalue',
      '#technicalSpecifications_section_1',
      '#additionalProductDetails-content table'
    ],
    // 商品詳細（箇条書き形式）
    detailBullets: [
      '#detailBullets_feature_div .a-list-item',
      '#detail-bullets .content li',
      '#additionalProductDetails-content .a-list-item'
    ],
    // カテゴリ（パンくずリスト — スマホ版: #breadcrumb、PC版: #wayfinding-breadcrumbs）
    breadcrumbs: [
      '#wayfinding-breadcrumbs_feature_div a.a-link-normal',
      '#breadcrumb_feature_div a.a-link-normal',
      '#breadcrumb a'
    ],
    // 評価（スマホ版: averageCustomerReviews内、PC版: acrPopover内）
    rating: [
      '#acrPopover .a-icon-alt',
      '#averageCustomerReviews_feature_div .a-icon-alt',
      '.a-icon-star .a-icon-alt'
    ],
    // レビュー数
    reviewCount: [
      '#acrCustomerReviewText',
      '#averageCustomerReviews_feature_div .a-size-small.a-color-secondary'
    ],
    // バリエーション（スマホ版: inline-twister、PC版: twister）
    variations: [
      '#variation_color_name .selection',
      '#variation_size_name .selection',
      '#twister .a-button-selected .a-button-text',
      '#inline-twister-row-color_name .a-button-selected .a-button-text',
      '#inline-twister-row-size_name .a-button-selected .a-button-text'
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
    //      51cvOwC5rkL.SX38_SY50_CR,...__.jpg → 51cvOwC5rkL._SL1500_.jpg
    return url.replace(/\._{0,2}[A-Z]{2}[^.]+_*\./, '._SL1500_.');
  }

  /**
   * 商品ページからすべての画像URLを収集（PC版・スマホ版両対応）
   */
  function collectImageUrls() {
    const images = [];
    const seen = new Set();
    const mobile = isMobilePage();

    if (mobile) {
      // === スマホ版の画像収集 ===
      // 1. immersive-view / image-block内の商品画像
      const ivImages = document.querySelectorAll('#immersive-main img, [id^="image-block-iv-product-image-"] img, [id^="image-block-product-image-"] img');
      if (ivImages.length > 0) {
        let first = true;
        for (const img of ivImages) {
          const src = img.src || img.getAttribute('data-src') || '';
          if (!src || src.includes('data:image') || src.includes('transparent-pixel') || src.includes('grey-pixel')) continue;
          const hiRes = toHighResUrl(src);
          if (hiRes && !seen.has(hiRes)) {
            images.push({ url: hiRes, type: first ? 'main' : 'gallery' });
            seen.add(hiRes);
            first = false;
          }
        }
      }

      // 2. imageGallery内の画像
      const galleryImgs = document.querySelectorAll('#imageGallery img.product-image, #imageGallery_feature_div img[src*="images-amazon.com"], #imageGallery_feature_div img[src*="m.media-amazon.com"]');
      for (const img of galleryImgs) {
        const src = img.src || img.getAttribute('data-src') || '';
        if (!src || src.includes('data:image') || src.includes('transparent-pixel') || src.includes('grey-pixel')) continue;
        const hiRes = toHighResUrl(src);
        if (hiRes && !seen.has(hiRes)) {
          images.push({ url: hiRes, type: images.length === 0 ? 'main' : 'gallery' });
          seen.add(hiRes);
        }
      }

      // 3. フォールバック: imageBlock内の全画像
      if (images.length === 0) {
        const blockImgs = document.querySelectorAll('#imageBlock_feature_div img[src*="images-amazon.com"], #imageBlock_feature_div img[src*="m.media-amazon.com"]');
        let first = true;
        for (const img of blockImgs) {
          const src = img.src || '';
          if (!src || src.includes('data:image') || src.includes('transparent-pixel') || src.includes('grey-pixel')) continue;
          // 小さすぎるアイコンを除外
          if (img.naturalWidth > 0 && img.naturalWidth <= 30) continue;
          const hiRes = toHighResUrl(src);
          if (hiRes && !seen.has(hiRes)) {
            images.push({ url: hiRes, type: first ? 'main' : 'gallery' });
            seen.add(hiRes);
            first = false;
          }
        }
      }
    } else {
      // === PC版の画像収集（従来のロジック） ===
      // 1. メイン画像（data-old-hires > data-a-dynamic-image > src）
      const mainImg = queryFirst(SELECTORS.mainImage);
      if (mainImg) {
        const hiRes = mainImg.getAttribute('data-old-hires');
        if (hiRes && !seen.has(hiRes)) {
          images.push({ url: hiRes, type: 'main' });
          seen.add(hiRes);
        }

        const dynamicAttr = mainImg.getAttribute('data-a-dynamic-image');
        if (dynamicAttr) {
          try {
            const dynamicImages = JSON.parse(dynamicAttr);
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

        if (images.length === 0 && mainImg.src && !seen.has(mainImg.src)) {
          images.push({ url: mainImg.src, type: 'main' });
          seen.add(mainImg.src);
        }
      }

      // 2. ギャラリー画像（サムネイルを高解像度に変換）
      const thumbs = queryAllFirst(SELECTORS.galleryThumbs);
      for (const thumb of thumbs) {
        const thumbSrc = thumb.src || '';
        const parentLi = thumb.closest('li');
        if (parentLi && parentLi.classList.contains('videoThumbnail')) continue;

        const hiRes = toHighResUrl(thumbSrc);
        if (hiRes && !seen.has(hiRes)) {
          images.push({ url: hiRes, type: 'gallery' });
          seen.add(hiRes);
        }
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

    // テキスト収集（重要なテキスト要素のみ、script/styleタグを除外）
    // 除外対象: 比較テーブル（standard/premium）、カルーセル（売れ筋商品等）
    const textElements = [...aplusEl.querySelectorAll('h1, h2, h3, h4, h5, p, li, td, span.a-text-bold')]
      .filter(el => !el.closest('.apm-tablemodule, .comparison-table, [class*="carousel"]'));
    const texts = [];
    const seenTexts = new Set();
    for (const el of textElements) {
      const text = getCleanText(el);
      if (text && text.length > 2 && !seenTexts.has(text)) {
        texts.push(text);
        seenTexts.add(text);
      }
    }
    result.text = texts.join('\n');

    // 画像収集（比較テーブル・カルーセル内の画像を除外）
    const imgs = [...aplusEl.querySelectorAll('img')]
      .filter(img => !img.closest('.apm-tablemodule, .comparison-table, [class*="carousel"]'));
    const seen = new Set();
    for (const img of imgs) {
      // data-src（遅延読み込み）またはsrc
      const src = img.getAttribute('data-src') || img.src || '';
      if (!src || src.includes('data:image') || src.includes('transparent-pixel') || src.includes('grey-pixel')) continue;

      const hiRes = toHighResUrl(src);
      if (hiRes && !seen.has(hiRes)) {
        result.images.push(hiRes);
        seen.add(hiRes);
      }
    }

    return result;
  }

  /**
   * 要素内のテキストをscript/style要素を除外して取得
   */
  function getCleanText(element) {
    const clone = element.cloneNode(true);
    clone.querySelectorAll('script, style').forEach(el => el.remove());
    return clone.textContent.trim().replace(/\s+/g, ' ');
  }

  /**
   * 技術仕様を収集
   */
  function collectTechSpecs() {
    const specs = {};
    const mobile = isMobilePage();

    // テーブル形式の仕様
    const techTable = queryFirst(SELECTORS.techSpecs);
    if (techTable) {
      const rows = techTable.querySelectorAll('tr');
      for (const row of rows) {
        const th = row.querySelector('th');
        const td = row.querySelector('td');
        if (th && td) {
          const key = getCleanText(th);
          const value = getCleanText(td);
          if (key && value) {
            specs[key] = value;
          }
        }
      }
    }

    // スマホ版: additionalProductDetails からテキスト形式で抽出
    if (mobile && Object.keys(specs).length === 0) {
      const detailContent = document.querySelector('#additionalProductDetails-content');
      if (detailContent) {
        const items = detailContent.querySelectorAll('.a-list-item, li');
        for (const item of items) {
          const text = getCleanText(item);
          const colonIndex = text.indexOf(':');
          if (colonIndex > 0 && colonIndex < text.length - 1) {
            const key = text.substring(0, colonIndex).trim();
            const value = text.substring(colonIndex + 1).trim();
            if (key && value && !key.includes('カスタマーレビュー')) {
              specs[key] = value;
            }
          }
        }
      }
    }

    // 箇条書き形式の詳細情報
    const detailItems = queryAllFirst(SELECTORS.detailBullets);
    for (const item of detailItems) {
      const text = getCleanText(item);
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
    const mobile = isMobilePage();
    console.log(`[Amazon商品情報収集 v${VERSION}] 情報収集を開始（${mobile ? 'スマホ版' : 'PC版'}）`);

    const asin = extractASIN();
    if (!asin) {
      console.error('[Amazon商品情報収集] ASINを取得できませんでした');
      return null;
    }

    // 商品名
    const titleEl = queryFirst(SELECTORS.productTitle);
    const title = titleEl ? titleEl.textContent.trim() : '';

    // 単価（"￥1 / ml" など）かどうかを判定
    function isUnitPrice(el) {
      const container = el.closest('.a-price')?.parentElement || el.parentElement;
      return container && /\/\s*(ml|mL|g|kg|個|本|枚|100|l|L)/i.test(container.textContent);
    }

    // 価格（テキストが空の要素をスキップ、単価・ポイント表示を除外）
    const priceEl = queryFirstWithText(SELECTORS.price);
    let price = priceEl ? priceEl.textContent.trim() : '';
    // 価格として有効かチェック（価格形式であること、かつ単価でないこと）
    const isValidPrice = price &&
      (/[￥¥]\s*[\d,]+/.test(price) || /^\d[\d,]+円/.test(price)) &&
      !(priceEl && isUnitPrice(priceEl));
    if (!isValidPrice) {
      price = '';
      for (const sel of SELECTORS.price) {
        const els = document.querySelectorAll(sel);
        for (const el of els) {
          const text = el.textContent.trim();
          if (/^[￥¥]\s*[\d,]+/.test(text) && !isUnitPrice(el)) {
            price = text;
            break;
          }
        }
        if (price) break;
      }
    }

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
      .map(el => getCleanText(el))
      .filter(text => text.length > 0 && !text.includes('この商品について'));

    // 商品説明
    const descEl = queryFirst(SELECTORS.productDescription);
    const description = descEl ? getCleanText(descEl) : '';

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

    // レビュー数（PC版セレクタ → モバイル版フォールバック）
    let reviewCount = '';
    const reviewCountEl = queryFirst(SELECTORS.reviewCount);
    if (reviewCountEl) {
      reviewCount = reviewCountEl.textContent.trim();
    } else {
      // モバイル版: #acrCustomerReviewLink内のテキストから括弧内の数字を抽出
      // HTML例: <span aria-label="44,362 レビュー">(44,362)</span>
      const reviewLink = document.querySelector('#acrCustomerReviewLink');
      if (reviewLink) {
        const match = reviewLink.textContent.match(/\(([\d,]+)\)/);
        if (match) reviewCount = match[1] + '件';
      }
      // さらにフォールバック: aria-label属性から取得
      if (!reviewCount) {
        const ariaEl = document.querySelector('[aria-label*="レビュー"]');
        if (ariaEl) {
          const ariaMatch = ariaEl.getAttribute('aria-label').match(/([\d,]+)/);
          if (ariaMatch) reviewCount = ariaMatch[1] + '件';
        }
      }
    }

    // バリエーション（ボタン内のバリエーション名のみ抽出、価格・在庫情報を除外）
    const variationEls = queryAllFirst(SELECTORS.variations);
    const variations = variationEls.map(el => {
      // swatch-title-text があればそこからバリエーション名のみ取得
      const nameEl = el.querySelector('.swatch-title-text, .swatch-title-text-display');
      if (nameEl) return nameEl.textContent.trim();
      // フォールバック: getCleanTextから価格・在庫パターンを除去
      return getCleanText(el)
        .replace(/[￥¥][\d,]+.*$/, '')  // 価格以降を除去
        .replace(/在庫.*$/, '')          // 在庫情報を除去
        .trim();
    }).filter(t => t.length > 0);

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

  /**
   * ページの主要コンテンツが読み込まれるまで待機
   * Amazon商品ページは遅延読み込みが多いため、主要要素の出現を確認してから収集する
   */
  async function waitForPageReady(maxWaitMs = 20000) {
    // すでにDOMContentLoadedが完了しているか確認
    if (document.readyState === 'loading') {
      await new Promise(resolve => {
        document.addEventListener('DOMContentLoaded', resolve, { once: true });
      });
    }

    // 主要要素が出現するまでポーリング（最大maxWaitMs）
    const startTime = Date.now();
    const checkInterval = 300;

    while (Date.now() - startTime < maxWaitMs) {
      const hasTitle = !!queryFirst(SELECTORS.productTitle);
      const hasPrice = !!queryFirstWithText(SELECTORS.price);
      const hasMainImage = !!queryFirst(SELECTORS.mainImage);
      // スマホ版の追加チェック: immersive-view内の画像
      const hasMobileImage = !!document.querySelector('#immersive-main img, #imageGallery img');

      // タイトル＋（価格または画像）が揃えば準備完了
      if (hasTitle && (hasPrice || hasMainImage || hasMobileImage)) {
        // A+コンテンツや画像データ属性が遅延設定されるのを少し待つ
        await new Promise(r => setTimeout(r, 500));
        console.log(`[Amazon商品情報収集] ページ準備完了（${Date.now() - startTime}ms）`);
        return;
      }

      await new Promise(r => setTimeout(r, checkInterval));
    }

    console.warn(`[Amazon商品情報収集] ${maxWaitMs}ms待機後もページが完全に読み込まれていない可能性があります。現在の状態で収集を開始します`);
  }

  // ===== メッセージハンドラー =====
  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (!message || !message.action) return;

    switch (message.action) {
      case 'collectProductInfo': {
        // ページの準備完了を待ってから収集開始
        waitForPageReady().then(() => {
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
        }).catch(error => {
          console.error('[Amazon商品情報収集] ページ待機エラー:', error);
          sendResponse({ success: false, error: error.message });
        });
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
