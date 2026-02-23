/**
 * offscreen.js
 * WebP→JPG変換 / 画像リサイズ用オフスクリーンドキュメント
 *
 * Service Worker (background.js) からメッセージを受け取り、
 * Canvas APIで画像を処理して結果を返す。
 *
 * 対応処理:
 * - WebP画像をJPEGに変換
 * - 幅1500pxを超える画像を幅1200pxにリサイズ
 * - 上記に該当しないJPG/PNGはそのまま返す
 */

const MAX_WIDTH = 1500;      // この幅を超えたらリサイズ
const TARGET_WIDTH = 1200;   // リサイズ後の幅
const JPEG_QUALITY = 0.92;   // JPEG出力品質

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'processImageForPdf') {
    processImage(message.imageBase64, message.mimeType)
      .then(result => sendResponse(result))
      .catch(error => {
        console.error('[offscreen] 画像処理エラー:', error);
        sendResponse({ error: error.message });
      });
    return true; // 非同期レスポンス
  }
});

/**
 * 画像を処理（WebP変換 / リサイズ）
 * @param {string} base64 - Base64エンコードされた画像データ
 * @param {string} mimeType - 元のMIMEタイプ
 * @returns {Promise<{base64: string, mimeType: string, width: number, height: number, processed: boolean}>}
 */
async function processImage(base64, mimeType) {
  // base64 → Blob
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  const blob = new Blob([bytes], { type: mimeType });

  // 画像をデコードしてサイズを取得
  const bitmap = await createImageBitmap(blob);
  const { width, height } = bitmap;

  const isWebP = mimeType === 'image/webp';
  const needsResize = width > MAX_WIDTH;

  // JPG/PNGで、リサイズ不要ならそのまま返す
  if (!isWebP && !needsResize) {
    bitmap.close();
    return {
      base64,
      mimeType,
      width,
      height,
      processed: false
    };
  }

  // Canvas描画（リサイズ or WebP変換）
  let targetWidth = width;
  let targetHeight = height;
  if (needsResize) {
    targetWidth = TARGET_WIDTH;
    targetHeight = Math.round(height * (TARGET_WIDTH / width));
  }

  const canvas = document.getElementById('canvas');
  canvas.width = targetWidth;
  canvas.height = targetHeight;
  const ctx = canvas.getContext('2d');
  ctx.drawImage(bitmap, 0, 0, targetWidth, targetHeight);
  bitmap.close();

  // JPEGに変換（WebPの場合もJPEGにする）
  const outputBlob = await new Promise(resolve => {
    canvas.toBlob(resolve, 'image/jpeg', JPEG_QUALITY);
  });

  // Blob → base64
  const outputBase64 = await blobToBase64(outputBlob);

  return {
    base64: outputBase64,
    mimeType: 'image/jpeg',
    width: targetWidth,
    height: targetHeight,
    processed: true
  };
}

/**
 * Blob → base64文字列（data:接頭辞なし）
 */
function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result;
      // "data:image/jpeg;base64," の後の部分を取得
      const base64 = dataUrl.split(',')[1];
      resolve(base64);
    };
    reader.onerror = () => reject(new Error('FileReader error'));
    reader.readAsDataURL(blob);
  });
}
