/**
 * pdf-generator.js
 * 商品画像からPDFを生成するモジュール
 * pdf-lib (PDFLib global) に依存
 *
 * 使い方:
 *   importScripts('pdf-lib.min.js', 'pdf-generator.js');
 *   const pdfBytes = await generateProductImagesPDF(images, options);
 */

/**
 * 商品画像をPDFにまとめる
 * @param {Array} images - 画像データの配列
 *   [{data: Uint8Array, mimeType: 'image/jpeg'|'image/png', section: 'main'|'gallery'|..., order: 1}]
 * @param {Object} options
 *   {productId: string, source: 'rakuten'|'amazon', baseName: string, mode?: 'desktop'|'mobile'}
 * @returns {Promise<Uint8Array>} PDFバイナリ
 */
async function generateProductImagesPDF(images, options) {
  const { PDFDocument, StandardFonts, rgb } = PDFLib;

  const A4_WIDTH = 595.28;
  const A4_HEIGHT = 841.89;
  const MARGIN = 40;

  const pdfDoc = await PDFDocument.create();
  const helvetica = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const helveticaBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  // セクション定義
  const sectionOrder = ['main', 'gallery', 'product', 'description', 'aplus'];
  const sectionLabels = {
    main: 'Main Image',
    gallery: 'Gallery Images',
    product: 'Product Images',
    description: 'Description Images',
    aplus: 'A+ Content Images'
  };

  // 画像をセクション別にグループ化
  const groupedImages = {};
  for (const img of images) {
    const section = img.section || 'product';
    if (!groupedImages[section]) groupedImages[section] = [];
    groupedImages[section].push(img);
  }

  // 有効なセクション数と画像総数を計算
  const activeSections = sectionOrder.filter(s => groupedImages[s] && groupedImages[s].length > 0);
  const totalImages = images.length;

  // --- 表紙ページ ---
  const coverPage = pdfDoc.addPage([A4_WIDTH, A4_HEIGHT]);

  // タイトル
  coverPage.drawText('Product Images', {
    x: MARGIN,
    y: A4_HEIGHT - 160,
    size: 36,
    font: helveticaBold,
    color: rgb(0.15, 0.15, 0.15),
  });

  // 区切り線
  coverPage.drawRectangle({
    x: MARGIN,
    y: A4_HEIGHT - 175,
    width: A4_WIDTH - MARGIN * 2,
    height: 2,
    color: rgb(0.8, 0.8, 0.8),
  });

  // 商品ID
  coverPage.drawText(options.productId, {
    x: MARGIN,
    y: A4_HEIGHT - 210,
    size: 20,
    font: helvetica,
    color: rgb(0.3, 0.3, 0.3),
  });

  // ソース & モード
  const sourceLabel = options.source === 'rakuten' ? 'Rakuten' : 'Amazon';
  const modeLabel = options.mode === 'mobile' ? ' (SP)' : options.source === 'rakuten' ? ' (PC)' : '';
  coverPage.drawText(`${sourceLabel}${modeLabel}`, {
    x: MARGIN,
    y: A4_HEIGHT - 240,
    size: 14,
    font: helvetica,
    color: rgb(0.5, 0.5, 0.5),
  });

  // 画像数
  coverPage.drawText(`${totalImages} images`, {
    x: MARGIN,
    y: A4_HEIGHT - 270,
    size: 14,
    font: helvetica,
    color: rgb(0.5, 0.5, 0.5),
  });

  // セクション一覧
  let yPos = A4_HEIGHT - 330;
  for (const section of activeSections) {
    const count = groupedImages[section].length;
    const label = sectionLabels[section] || section;
    coverPage.drawText(`${label}: ${count}`, {
      x: MARGIN + 20,
      y: yPos,
      size: 12,
      font: helvetica,
      color: rgb(0.4, 0.4, 0.4),
    });
    yPos -= 24;
  }

  // 日付
  const dateStr = new Date().toISOString().split('T')[0];
  coverPage.drawText(dateStr, {
    x: MARGIN,
    y: 30,
    size: 10,
    font: helvetica,
    color: rgb(0.6, 0.6, 0.6),
  });

  // --- セクション別にページ追加 ---
  let pageNumber = 1; // 表紙を除いたページ番号

  for (const section of activeSections) {
    const sectionImages = groupedImages[section];
    const label = sectionLabels[section] || section;

    // セクション見出しページ
    const headerPage = pdfDoc.addPage([A4_WIDTH, A4_HEIGHT]);
    pageNumber++;

    // セクション名（ページ中央やや上）
    const labelWidth = helveticaBold.widthOfTextAtSize(label, 32);
    headerPage.drawText(label, {
      x: (A4_WIDTH - labelWidth) / 2,
      y: A4_HEIGHT / 2 + 30,
      size: 32,
      font: helveticaBold,
      color: rgb(0.2, 0.2, 0.2),
    });

    // 枚数
    const countText = `${sectionImages.length} image${sectionImages.length > 1 ? 's' : ''}`;
    const countWidth = helvetica.widthOfTextAtSize(countText, 18);
    headerPage.drawText(countText, {
      x: (A4_WIDTH - countWidth) / 2,
      y: A4_HEIGHT / 2 - 15,
      size: 18,
      font: helvetica,
      color: rgb(0.5, 0.5, 0.5),
    });

    // 各画像ページ
    for (let i = 0; i < sectionImages.length; i++) {
      const img = sectionImages[i];
      const page = pdfDoc.addPage([A4_WIDTH, A4_HEIGHT]);
      pageNumber++;

      try {
        // 画像をPDFに埋め込み
        let pdfImage;
        if (img.mimeType === 'image/jpeg' || img.mimeType === 'image/jpg') {
          pdfImage = await pdfDoc.embedJpg(img.data);
        } else if (img.mimeType === 'image/png') {
          pdfImage = await pdfDoc.embedPng(img.data);
        } else {
          // サポート外のフォーマット（offscreenで変換されるはずだが念のため）
          page.drawText(`Unsupported format: ${img.mimeType}`, {
            x: MARGIN,
            y: A4_HEIGHT / 2,
            size: 14,
            font: helvetica,
            color: rgb(0.8, 0, 0),
          });
          continue;
        }

        // 画像をページに収まるようスケール（拡大はしない）
        const availableWidth = A4_WIDTH - MARGIN * 2;
        const availableHeight = A4_HEIGHT - MARGIN * 2 - 30; // フッター用スペース

        const imgWidth = pdfImage.width;
        const imgHeight = pdfImage.height;

        const scaleX = availableWidth / imgWidth;
        const scaleY = availableHeight / imgHeight;
        const scale = Math.min(scaleX, scaleY, 1);

        const drawWidth = imgWidth * scale;
        const drawHeight = imgHeight * scale;

        // ページ中央に配置（フッター分を考慮して少し上に）
        const x = (A4_WIDTH - drawWidth) / 2;
        const y = (A4_HEIGHT - drawHeight) / 2 + 10;

        page.drawImage(pdfImage, {
          x,
          y,
          width: drawWidth,
          height: drawHeight,
        });
      } catch (embedError) {
        // 画像埋め込みエラー: エラーページを表示
        const errMsg = `Error embedding image: ${embedError.message}`;
        page.drawText(errMsg, {
          x: MARGIN,
          y: A4_HEIGHT / 2,
          size: 11,
          font: helvetica,
          color: rgb(0.7, 0, 0),
        });
        if (img.originalUrl) {
          page.drawText(`URL: ${img.originalUrl.substring(0, 80)}`, {
            x: MARGIN,
            y: A4_HEIGHT / 2 - 20,
            size: 9,
            font: helvetica,
            color: rgb(0.5, 0.5, 0.5),
          });
        }
      }

      // フッター: セクション名 - 番号
      const footerText = `${label} - ${i + 1}/${sectionImages.length}`;
      const footerWidth = helvetica.widthOfTextAtSize(footerText, 9);
      page.drawText(footerText, {
        x: (A4_WIDTH - footerWidth) / 2,
        y: 25,
        size: 9,
        font: helvetica,
        color: rgb(0.6, 0.6, 0.6),
      });
    }
  }

  return await pdfDoc.save();
}
