/**
 * ユーザー管理スプレッドシート用 Google Apps Script
 * Google連絡先からユーザーを追加する機能
 * A列: 名前、B列: メールアドレス
 */

/**
 * スプレッドシートを開いたときにカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('👥 連絡先から追加')
    .addItem('追加する', 'showContactPicker')
    .addToUi();
}

/**
 * 連絡先選択ダイアログを表示
 */
function showContactPicker() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; }
      h3 { color: #BF0000; margin-bottom: 15px; }
      .contact-list { max-height: 420px; overflow-y: auto; border: 1px solid #ddd; padding: 10px; border-radius: 8px; }
      .contact-item { padding: 8px; margin: 4px 0; background: #f5f5f5; border-radius: 4px; cursor: pointer; }
      .contact-item:hover { background: #e0e0e0; }
      .contact-item input { margin-right: 10px; }
      .btn { padding: 10px 20px; margin: 5px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; }
      .btn-primary { background: #BF0000; color: white; }
      .btn-primary:hover { background: #8B0000; }
      .btn-secondary { background: #666; color: white; }
      .loading { text-align: center; padding: 20px; color: #666; }
      .search-box { width: 100%; padding: 10px; margin-bottom: 10px; border: 1px solid #ddd; border-radius: 6px; box-sizing: border-box; }
    </style>
    <input type="text" class="search-box" id="search" placeholder="検索..." onkeyup="filterContacts()">
    <div class="contact-list" id="contactList">
      <div class="loading">連絡先を読み込み中...</div>
    </div>
    <div style="margin-top: 15px; text-align: right;">
      <button class="btn btn-secondary" onclick="google.script.host.close()">キャンセル</button>
      <button class="btn btn-primary" onclick="addSelected()">選択したユーザーを追加</button>
    </div>
    <script>
      let allContacts = [];

      // 連絡先を読み込む
      google.script.run
        .withSuccessHandler(function(contacts) {
          allContacts = contacts;
          renderContacts(contacts);
        })
        .withFailureHandler(function(error) {
          document.getElementById('contactList').innerHTML =
            '<div style="color:red;">エラー: ' + error.message + '</div>';
        })
        .getGoogleContacts();

      function renderContacts(contacts) {
        const list = document.getElementById('contactList');
        if (contacts.length === 0) {
          list.innerHTML = '<div style="color:#666;">連絡先が見つかりません</div>';
          return;
        }
        list.innerHTML = contacts.map((c, i) =>
          '<div class="contact-item">' +
          '<input type="checkbox" id="contact_' + i + '" data-name="' + escapeHtml(c.name || '') + '" data-email="' + escapeHtml(c.email) + '">' +
          '<label for="contact_' + i + '">' + escapeHtml(c.name || c.email) + '</label>' +
          '</div>'
        ).join('');
      }

      function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
      }

      function filterContacts() {
        const query = document.getElementById('search').value.toLowerCase();
        const filtered = allContacts.filter(c =>
          (c.name && c.name.toLowerCase().includes(query)) ||
          c.email.toLowerCase().includes(query)
        );
        renderContacts(filtered);
      }

      function addSelected() {
        const checkboxes = document.querySelectorAll('input[type="checkbox"]:checked');
        const users = Array.from(checkboxes).map(cb => ({
          name: cb.getAttribute('data-name') || '',
          email: cb.getAttribute('data-email')
        }));
        if (users.length === 0) {
          showMessage('ユーザーを選択してください', 'error');
          return;
        }
        // 即座にフィードバック表示
        showMessage(users.length + '人を追加しています...', 'loading');
        document.querySelector('.btn-primary').disabled = true;

        google.script.run
          .withSuccessHandler(function(result) {
            showMessage('✓ ' + result.message, 'success');
            setTimeout(function() { google.script.host.close(); }, 1200);
          })
          .withFailureHandler(function(error) {
            showMessage('エラー: ' + error.message, 'error');
            document.querySelector('.btn-primary').disabled = false;
          })
          .addUsersFromContacts(users);
      }

      function showMessage(text, type) {
        const list = document.getElementById('contactList');
        const color = type === 'error' ? '#c00' : type === 'loading' ? '#666' : '#080';
        list.innerHTML = '<div style="text-align:center;padding:40px;color:' + color + ';font-size:16px;">' + text + '</div>';
      }
    </script>
  `)
  .setWidth(500)
  .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Google連絡先から追加');
}

/**
 * Googleコンタクトからメールアドレスを取得
 * @returns {Array} 連絡先の配列 [{name, email}]
 */
function getGoogleContacts() {
  try {
    const contacts = [];
    const seenEmails = new Set();
    const seenNames = new Set();
    const people = People.People.Connections.list('people/me', {
      personFields: 'names,emailAddresses',
      pageSize: 1000
    });

    if (people.connections) {
      people.connections.forEach(person => {
        if (person.emailAddresses && person.emailAddresses.length > 0) {
          const name = person.names && person.names.length > 0
            ? person.names[0].displayName
            : '';
          const email = person.emailAddresses[0].value.toLowerCase();
          const nameKey = name.toLowerCase().trim();

          // メールと名前の両方で重複チェック
          if (!seenEmails.has(email) && (!nameKey || !seenNames.has(nameKey))) {
            seenEmails.add(email);
            if (nameKey) seenNames.add(nameKey);
            contacts.push({
              name: name,
              email: person.emailAddresses[0].value
            });
          }
        }
      });
    }

    // 名前でソート
    contacts.sort((a, b) => (a.name || a.email).localeCompare(b.name || b.email));

    return contacts;
  } catch (error) {
    console.error('連絡先取得エラー:', error);
    throw new Error('連絡先の取得に失敗しました。People APIが有効になっているか確認してください。');
  }
}

/**
 * 選択されたユーザーをスプレッドシートに追加
 * @param {Array} users - 追加するユーザーの配列 [{name, email}]
 */
function addUsersFromContacts(users) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('シート1') || ss.getSheets()[0];

  // 既存のメールアドレスを取得（B列）
  const existingEmails = new Set();
  if (sheet.getLastRow() > 1) {
    const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    data.forEach(row => {
      if (row[0]) existingEmails.add(row[0].toString().toLowerCase().trim());
    });
  }

  // 新しいユーザーを追加（A列: 名前、B列: メール）
  let addedCount = 0;
  users.forEach(user => {
    if (!existingEmails.has(user.email.toLowerCase().trim())) {
      sheet.appendRow([user.name || '', user.email]);
      addedCount++;
    }
  });

  const skippedCount = users.length - addedCount;
  let message = addedCount + '人のユーザーを追加しました。';
  if (skippedCount > 0) {
    message += '\n(' + skippedCount + '人は既に登録済みのためスキップ)';
  }

  return { success: true, message: message, added: addedCount, skipped: skippedCount };
}

