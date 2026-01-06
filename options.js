/**
 * 設定画面のスクリプト
 * GAS Web App URLの保存・読み込み・接続テストを行う
 */

document.addEventListener('DOMContentLoaded', () => {
  const gasUrlInput = document.getElementById('gasUrl');
  const saveBtn = document.getElementById('saveBtn');
  const testBtn = document.getElementById('testBtn');
  const statusDiv = document.getElementById('status');
  const spreadsheetLinkDiv = document.getElementById('spreadsheetLink');
  const spreadsheetUrlLink = document.getElementById('spreadsheetUrl');

  // 保存済みの設定を読み込む
  loadSettings();

  // 保存ボタンのクリックイベント
  saveBtn.addEventListener('click', saveSettings);

  // 接続テストボタンのクリックイベント
  testBtn.addEventListener('click', testConnection);

  /**
   * 保存済みの設定を読み込む
   */
  function loadSettings() {
    chrome.storage.sync.get(['gasUrl'], (result) => {
      if (result.gasUrl) {
        gasUrlInput.value = result.gasUrl;
      }
    });
  }

  /**
   * 設定を保存する
   */
  function saveSettings() {
    const gasUrl = gasUrlInput.value.trim();

    // URLの簡易バリデーション
    if (gasUrl && !isValidGasUrl(gasUrl)) {
      showStatus('error', 'URLの形式が正しくありません。Google Apps ScriptのWeb App URLを入力してください。');
      return;
    }

    chrome.storage.sync.set({ gasUrl }, () => {
      if (chrome.runtime.lastError) {
        showStatus('error', '保存に失敗しました: ' + chrome.runtime.lastError.message);
      } else {
        if (gasUrl) {
          showStatus('success', '設定を保存しました。レビューはスプレッドシートに自動保存されます。');
        } else {
          showStatus('success', '設定を保存しました。レビューはCSVファイルとしてダウンロードされます。');
        }
      }
    });
  }

  /**
   * GAS URLの簡易バリデーション
   */
  function isValidGasUrl(url) {
    return url.startsWith('https://script.google.com/macros/s/') && url.includes('/exec');
  }

  /**
   * GASへの接続テスト
   */
  async function testConnection() {
    const gasUrl = gasUrlInput.value.trim();

    if (!gasUrl) {
      showStatus('error', 'URLを入力してください。');
      return;
    }

    if (!isValidGasUrl(gasUrl)) {
      showStatus('error', 'URLの形式が正しくありません。');
      return;
    }

    showStatus('', 'テスト中...');
    statusDiv.style.display = 'block';
    statusDiv.style.background = '#e2e3e5';
    statusDiv.style.color = '#383d41';
    spreadsheetLinkDiv.style.display = 'none';

    try {
      // GETリクエストでスプレッドシート情報を取得
      const response = await fetch(gasUrl, {
        method: 'GET',
        mode: 'cors'
      });

      const data = await response.json();

      if (data.success) {
        showStatus('success', '接続テスト成功。GASへの通信が確認できました。');

        // スプレッドシートURLを表示
        if (data.spreadsheetUrl) {
          spreadsheetUrlLink.href = data.spreadsheetUrl;
          spreadsheetUrlLink.textContent = data.spreadsheetUrl;
          spreadsheetLinkDiv.style.display = 'block';

          // スプレッドシートURLを保存
          chrome.storage.sync.set({ spreadsheetUrl: data.spreadsheetUrl });
        }
      } else {
        showStatus('error', '接続テスト失敗: ' + (data.error || '不明なエラー'));
      }
    } catch (error) {
      // GETが失敗した場合、POSTでno-corsモードを試す
      try {
        await fetch(gasUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            test: true,
            timestamp: new Date().toISOString()
          })
        });
        showStatus('success', '接続テスト成功（レスポンス取得不可）。スプレッドシートURLは手動で確認してください。');
      } catch (postError) {
        showStatus('error', '接続テスト失敗: ' + error.message);
      }
    }
  }

  /**
   * ステータスメッセージを表示
   */
  function showStatus(type, message) {
    statusDiv.textContent = message;
    statusDiv.className = 'status';
    if (type) {
      statusDiv.classList.add(type);
    }
    statusDiv.style.display = 'block';

    // 成功メッセージは3秒後に消す
    if (type === 'success') {
      setTimeout(() => {
        statusDiv.style.display = 'none';
      }, 3000);
    }
  }
});
