/**
 * Amazon検索結果ページで各商品に「キューに追加」ボタンを注入するコンテンツスクリプト
 * 対象: https://www.amazon.co.jp/s?k=...
 */

(() => {
  'use strict';

  if (window.__amazonSearchQueueLoaded) return;
  window.__amazonSearchQueueLoaded = true;

  // 検索結果ページかどうか判定
  function isSearchPage() {
    return location.pathname === '/s' || location.pathname.startsWith('/s/');
  }

  if (!isSearchPage()) return;

  // スタイルを注入
  const style = document.createElement('style');
  style.textContent = `
    .shushu-queue-btn {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 4px 10px;
      margin-top: 4px;
      background: #232f3e;
      color: #fff;
      border: none;
      border-radius: 4px;
      font-size: 12px;
      cursor: pointer;
      transition: background 0.2s;
    }
    .shushu-queue-btn:hover {
      background: #37475a;
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

  // 商品カードからASINを取得
  function getAsin(card) {
    return card.getAttribute('data-asin') || '';
  }

  // 商品カードからタイトルを取得
  function getTitle(card) {
    const el = card.querySelector('h2 a span, h2 span');
    return el ? el.textContent.trim().substring(0, 100) : '';
  }

  // キューに追加
  function addToQueue(asin, title, btn) {
    if (!asin) return;

    const url = `https://www.amazon.co.jp/dp/${asin}`;

    chrome.storage.local.get(['queue'], (result) => {
      const queue = result.queue || [];

      // 重複チェック
      if (queue.some(item => item.url === url || (item.productId && item.productId === asin))) {
        btn.textContent = '追加済み';
        btn.classList.add('added');
        return;
      }

      queue.push({
        id: Date.now().toString() + '_' + asin,
        url: url,
        title: title || asin,
        productId: asin,
        source: 'amazon',
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
    const cards = document.querySelectorAll('[data-asin]:not([data-asin=""]):not(.shushu-processed)');

    // 現在のキューを取得して重複判定
    chrome.storage.local.get(['queue'], (result) => {
      const queue = result.queue || [];
      const queuedAsins = new Set(
        queue.filter(item => item.source === 'amazon' && item.productId)
             .map(item => item.productId)
      );

      cards.forEach(card => {
        card.classList.add('shushu-processed');

        const asin = getAsin(card);
        if (!asin || asin.length !== 10) return;

        // 広告・スポンサー枠はスキップ（必要に応じて収集したい場合もあるのでスキップしない）
        const title = getTitle(card);

        // ボタンの挿入先を探す
        const priceSection = card.querySelector('.a-price') ||
                             card.querySelector('.a-row.a-size-base') ||
                             card.querySelector('h2');
        if (!priceSection) return;

        const btn = document.createElement('button');
        btn.className = 'shushu-queue-btn';

        if (queuedAsins.has(asin)) {
          btn.textContent = '✓ 追加済み';
          btn.classList.add('added');
        } else {
          btn.textContent = '+ キューに追加';
          btn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            addToQueue(asin, title, btn);
          });
        }

        // 価格の後ろに挿入
        priceSection.parentNode.insertBefore(btn, priceSection.nextSibling);
      });
    });
  }

  // 初回実行
  injectButtons();

  // 無限スクロール・ページ切替に対応
  const observer = new MutationObserver(() => {
    const unprocessed = document.querySelectorAll('[data-asin]:not([data-asin=""]):not(.shushu-processed)');
    if (unprocessed.length > 0) {
      injectButtons();
    }
  });

  observer.observe(document.body, { childList: true, subtree: true });
})();
