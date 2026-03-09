/**
 * 楽天検索結果ページで各商品に「キューに追加」ボタンを注入するコンテンツスクリプト
 * 対象: https://search.rakuten.co.jp/search/mall/...
 */

(() => {
  'use strict';

  if (window.__rakutenSearchQueueLoaded) return;
  window.__rakutenSearchQueueLoaded = true;

  // スタイルを注入
  const style = document.createElement('style');
  style.textContent = `
    .shushu-queue-btn {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 4px 10px;
      margin-top: 4px;
      background: #BF0000;
      color: #fff;
      border: none;
      border-radius: 4px;
      font-size: 12px;
      cursor: pointer;
      transition: background 0.2s;
    }
    .shushu-queue-btn:hover {
      background: #E53935;
    }
    .shushu-queue-btn.added {
      background: #067d62;
      cursor: default;
    }
    .shushu-queue-btn.added:hover {
      background: #067d62;
    }
  `;
  document.head.appendChild(style);

  // 商品カードからURLを取得
  function getProductUrl(card) {
    const link = card.querySelector('a[href*="item.rakuten.co.jp"]');
    if (!link) return '';
    try {
      const url = new URL(link.href);
      return `${url.origin}${url.pathname}`.replace(/\/$/, '');
    } catch (e) {
      return '';
    }
  }

  // 商品カードからタイトルを取得
  function getTitle(card) {
    const el = card.querySelector('.searchresultitem .content.title a, .dui-card .title-link--3Ho6z, a[href*="item.rakuten.co.jp"]');
    return el ? el.textContent.trim().substring(0, 100) : '';
  }

  // キューに追加
  function addToQueue(url, title, btn) {
    if (!url) return;

    chrome.storage.local.get(['queue'], (result) => {
      const queue = result.queue || [];

      // 重複チェック
      if (queue.some(item => item.url === url)) {
        btn.textContent = '追加済み';
        btn.classList.add('added');
        return;
      }

      queue.push({
        id: Date.now().toString() + '_' + Math.random().toString(36).substr(2, 6),
        url: url,
        title: title || url,
        source: 'rakuten',
        addedAt: new Date().toISOString()
      });

      chrome.storage.local.set({ queue }, () => {
        btn.textContent = '✓ 追加済み';
        btn.classList.add('added');
      });
    });
  }

  // 商品カードにボタンを注入
  function injectButtons() {
    // 楽天検索結果の商品カードセレクタ
    const cards = document.querySelectorAll('.searchresultitem:not(.shushu-processed), .dui-card:not(.shushu-processed), [data-ratid]:not(.shushu-processed)');

    chrome.storage.local.get(['queue'], (result) => {
      const queue = result.queue || [];
      const queuedUrls = new Set(
        queue.filter(item => item.source === 'rakuten')
             .map(item => item.url)
      );

      cards.forEach(card => {
        card.classList.add('shushu-processed');

        const url = getProductUrl(card);
        if (!url) return;

        const title = getTitle(card);

        // ボタンの挿入先を探す
        const priceArea = card.querySelector('.price, .important, .content.price') ||
                          card.querySelector('a[href*="item.rakuten.co.jp"]');
        if (!priceArea) return;

        const btn = document.createElement('button');
        btn.className = 'shushu-queue-btn';

        if (queuedUrls.has(url)) {
          btn.textContent = '✓ 追加済み';
          btn.classList.add('added');
        } else {
          btn.textContent = '+ キューに追加';
          btn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            addToQueue(url, title, btn);
          });
        }

        priceArea.parentNode.insertBefore(btn, priceArea.nextSibling);
      });
    });
  }

  // 初回実行
  injectButtons();

  // 無限スクロール・ページ切替に対応
  const observer = new MutationObserver(() => {
    const unprocessed = document.querySelectorAll('.searchresultitem:not(.shushu-processed), .dui-card:not(.shushu-processed), [data-ratid]:not(.shushu-processed)');
    if (unprocessed.length > 0) {
      injectButtons();
    }
  });

  observer.observe(document.body, { childList: true, subtree: true });
})();
