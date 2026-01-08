/**
 * ユーザー管理スプレッドシート用 Google Apps Script
 * Google連絡先からユーザーを追加する機能
 */

/**
 * スプレッドシートを開いたときにカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('👥 ユーザー管理')
    .addItem('📇 Google連絡先から追加', 'showContactPicker')
    .addItem('✅ 重複を削除', 'removeDuplicateEmails')
    .addSeparator()
    .addItem('📊 ユーザー数を確認', 'showUserCount')
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
      .contact-list { max-height: 300px; overflow-y: auto; border: 1px solid #ddd; padding: 10px; border-radius: 8px; }
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
    <h3>📇 Google連絡先から追加</h3>
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
          '<div class="contact-item" data-email="' + c.email + '">' +
          '<input type="checkbox" id="contact_' + i + '" value="' + c.email + '">' +
          '<label for="contact_' + i + '">' + (c.name || c.email) + ' &lt;' + c.email + '&gt;</label>' +
          '</div>'
        ).join('');
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
        const emails = Array.from(checkboxes).map(cb => cb.value);
        if (emails.length === 0) {
          alert('ユーザーを選択してください');
          return;
        }
        google.script.run
          .withSuccessHandler(function(result) {
            alert(result.message);
            google.script.host.close();
          })
          .withFailureHandler(function(error) {
            alert('エラー: ' + error.message);
          })
          .addUsersFromContacts(emails);
      }
    </script>
  `)
  .setWidth(500)
  .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, '連絡先から追加');
}

/**
 * Googleコンタクトからメールアドレスを取得
 * @returns {Array} 連絡先の配列 [{name, email}]
 */
function getGoogleContacts() {
  try {
    const contacts = [];
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
          person.emailAddresses.forEach(email => {
            contacts.push({
              name: name,
              email: email.value
            });
          });
        }
      });
    }

    // メールアドレスでソート
    contacts.sort((a, b) => a.email.localeCompare(b.email));

    return contacts;
  } catch (error) {
    console.error('連絡先取得エラー:', error);
    throw new Error('連絡先の取得に失敗しました。People APIが有効になっているか確認してください。');
  }
}

/**
 * 選択されたユーザーをスプレッドシートに追加
 * @param {Array} emails - 追加するメールアドレスの配列
 */
function addUsersFromContacts(emails) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('シート1') || ss.getSheets()[0];

  // 既存のメールアドレスを取得
  const existingEmails = new Set();
  if (sheet.getLastRow() > 1) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    data.forEach(row => {
      if (row[0]) existingEmails.add(row[0].toString().toLowerCase().trim());
    });
  }

  // 新しいメールアドレスを追加
  let addedCount = 0;
  emails.forEach(email => {
    if (!existingEmails.has(email.toLowerCase().trim())) {
      sheet.appendRow([email]);
      addedCount++;
    }
  });

  const skippedCount = emails.length - addedCount;
  let message = addedCount + '人のユーザーを追加しました。';
  if (skippedCount > 0) {
    message += '\n(' + skippedCount + '人は既に登録済みのためスキップ)';
  }

  return { success: true, message: message, added: addedCount, skipped: skippedCount };
}

/**
 * 重複メールアドレスを削除
 */
function removeDuplicateEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('シート1') || ss.getSheets()[0];
  const ui = SpreadsheetApp.getUi();

  if (sheet.getLastRow() < 2) {
    ui.alert('ユーザーが登録されていません。');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);

  const seen = new Set();
  const uniqueRows = [];

  rows.forEach(row => {
    const email = row[0].toString().toLowerCase().trim();
    if (email && !seen.has(email)) {
      seen.add(email);
      uniqueRows.push(row);
    }
  });

  const removedCount = rows.length - uniqueRows.length;

  if (removedCount > 0) {
    sheet.clear();
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    if (uniqueRows.length > 0) {
      sheet.getRange(2, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
    }
    ui.alert('✅ 完了', removedCount + '件の重複を削除しました。', ui.ButtonSet.OK);
  } else {
    ui.alert('重複はありませんでした。');
  }
}

/**
 * ユーザー数を表示
 */
function showUserCount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('シート1') || ss.getSheets()[0];
  const ui = SpreadsheetApp.getUi();

  const count = Math.max(0, sheet.getLastRow() - 1);
  ui.alert('📊 ユーザー数', '現在 ' + count + ' 人のユーザーが登録されています。', ui.ButtonSet.OK);
}
