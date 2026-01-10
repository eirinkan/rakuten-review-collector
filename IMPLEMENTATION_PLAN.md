# レビュー収集 Chrome拡張 - Amazon対応 実装プラン

## 概要
既存の楽天レビュー収集機能に、Amazonの商品レビュー収集機能を追加する。

---

## 定義済み項目

| カテゴリ | 項目 | 定義内容 |
|---------|------|---------|
| 名称 | 拡張機能名 | 「レビュー収集」 |
| スプレッドシート | Amazon用シート名 | ASIN |
| スプレッドシート | Amazonヘッダー文字色 | #ff9900 |
| スプレッドシート | Amazonヘッダー背景色 | #000000 |
| 設定 | スプレッドシート分離 | チェックボックスで選択可能 |
| 設定 | CSV販路別出力 | チェックボックスで選択可能 |
| URL | Amazonランキング | https://www.amazon.co.jp/gp/bestsellers/ref=zg_bsms_tab_bs |
| キュー | 販路判別 | Amazon/楽天を識別可能に |
| データ | 項目数 | 22項目（source + country追加） |

---

## URL判定ロジック

```
楽天商品ページ:    item.rakuten.co.jp/*
楽天レビューページ: review.rakuten.co.jp/*
楽天ランキング:    ranking.rakuten.co.jp/*

Amazon商品ページ:  amazon.co.jp/dp/* または amazon.co.jp/gp/product/*
Amazonレビュー:    amazon.co.jp/product-reviews/*
Amazonランキング:  amazon.co.jp/gp/bestsellers/* または amazon.co.jp/ranking/*
```

---

## Amazonランキング機能

### ランキングページURL
- `https://www.amazon.co.jp/gp/bestsellers/`
- `https://www.amazon.co.jp/gp/bestsellers/electronics/`
- `https://www.amazon.co.jp/gp/new-releases/`

### 取得方法
- ランキングページのDOM（商品リスト）から `[data-asin]` 属性でASINを抽出
- 各商品のレビューページURLを構築: `/product-reviews/[ASIN]/`

### UIフロー（楽天と同じ）
1. ランキングページを開く
2. 「収集開始」ボタンで上位N件をキューに追加＆収集開始
3. 「キューに追加」ボタンで追加のみ

---

## データ項目（22項目）

| # | 項目名 | 楽天 | Amazon | 備考 |
|---|--------|------|--------|------|
| 1 | collectedAt | ✓ | ✓ | 収集日時 |
| 2 | productId | 商品ID | ASIN | 商品識別子 |
| 3 | productName | ✓ | ✓ | 商品名 |
| 4 | productUrl | ✓ | ✓ | 商品URL |
| 5 | rating | ✓ | ✓ | 評価（1-5） |
| 6 | title | ✓ | ✓ | レビュータイトル |
| 7 | body | ✓ | ✓ | レビュー本文 |
| 8 | author | ✓ | ✓ | 投稿者名 |
| 9 | age | ✓ | （空） | 年代 |
| 10 | gender | ✓ | （空） | 性別 |
| 11 | reviewDate | ✓ | ✓ | レビュー日 |
| 12 | orderDate | ✓ | （空） | 注文日 |
| 13 | variation | ✓ | ✓ | バリエーション |
| 14 | usage | ✓ | （空） | 用途 |
| 15 | recipient | ✓ | （空） | 贈り先 |
| 16 | purchaseCount | ✓ | （空） | 購入回数 |
| 17 | helpfulCount | ✓ | ✓ | 参考になった数 |
| 18 | shopReply | ✓ | （空） | ショップ返信 |
| 19 | shopName | 店舗名 | Amazon | ショップ名 |
| 20 | pageUrl | ✓ | ✓ | レビューページURL |
| 21 | source | rakuten | amazon | 販路識別 |
| 22 | country | （空） | ✓ | 投稿国（Amazon用） |

---

## ファイル構成

| ファイル | 変更内容 |
|---------|---------|
| `manifest.json` | Amazon権限追加、content-amazon.js追加 |
| `content-amazon.js` | **新規作成** - Amazonレビュー収集 |
| `background.js` | Amazon URL判定、ランキング取得、ASIN抽出 |
| `popup.js` | Amazonページ検出 |
| `popup.html` | 警告メッセージ更新 |
| `options.js` | Amazon URL対応（キュー追加時） |
| `options.html` | Amazon用スプレッドシート設定 |

---

## content-amazon.js 主要関数

```javascript
- extractReviews()        : レビューリストから全レビューを抽出
- extractReviewData()     : 個別レビューのデータ抽出（22項目）
- findNextPage()          : 「次へ」リンクを探す
- getProductInfo()        : 商品情報取得（ASIN、商品名、URL）
- startCollection()       : 収集開始
- collectCurrentPage()    : 現在ページのレビュー収集
- collectAllPages()       : 全ページ収集（ページネーション対応）
```

---

## Amazonレビューセレクター

```javascript
// レビューコンテナ
'[data-hook="review"]'
'#cm-cr-dp-review-list .review'

// 評価
'[data-hook="review-star-rating"] span'
'.a-icon-star span'

// タイトル
'[data-hook="review-title"] span:not([class])'

// 本文
'[data-hook="review-body"] span'

// 投稿者
'.a-profile-name'

// 日付
'[data-hook="review-date"]'

// 参考になった数
'[data-hook="helpful-vote-statement"]'

// バリエーション
'[data-hook="format-strip"]'

// ページネーション
'.a-pagination .a-last a'
```

---

## 警告メッセージリンク先

| リンク | URL |
|--------|-----|
| 楽天商品ページ | https://www.rakuten.co.jp/ |
| 楽天ランキング | https://ranking.rakuten.co.jp/sitemap/ |
| Amazon商品ページ | https://www.amazon.co.jp/ |
| Amazonランキング | https://www.amazon.co.jp/gp/bestsellers/ |

---

## 実装ステータス

### 完了
- [x] manifest.json にAmazon権限追加
- [x] content-amazon.js 新規作成（商品ページ・レビュー収集）
- [x] background.js でAmazon URL判定
- [x] popup.js/popup.html でAmazonページ検出・表示
- [x] options.js/options.html でAmazon用スプレッドシート設定
- [x] 警告メッセージに4つのリンク追加
- [x] ヘッダー色をネイビーに変更
- [x] スプレッドシートタイトル表示機能

### 完了 (追加実装)
- [x] **Amazonランキングからの商品取得機能** (2026-01-10)
  - background.js に fetchAmazonRankingProducts() 関数追加
  - popup.js でAmazonランキングページ検出・ランキングモード有効化
  - ランキングページから data-asin 属性と /dp/ASIN パターンでASIN抽出
  - 商品URL・ASIN をキューに追加

---

## 注意事項

- Amazonは頻繁にDOM構造を変更するため、セレクターの更新が必要になる可能性
- 2024年11月以降、ログインなしでは少数のレビューしか表示されない
- ボット対策が厳しいため、適切な待機時間を設定（5-10秒）
- data-hook属性を優先、フォールバック多めに設計
