function fetchInquiryEmails() {
  // 1. スプレッドシートの準備（現在開いているシートを取得）
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 2. Gmailから未読の対象メールを検索（件名に【お問い合わせ】を含む未読メール）
  const query = 'is:unread subject:"【お問い合わせ】"';
  const threads = GmailApp.search(query, 0, 10); // 最新10件を取得

  // 3. 見つかったメールを1件ずつ処理するループ
  for (let i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();

    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];

      // 未読の場合のみ処理を実行
      if (message.isUnread()) {
        const body = message.getPlainBody(); // メールの本文を取得
        const date = message.getDate();      // 受信日時を取得

        // 4. 本文からデータを抽出（正規表現という技術を使います）
        const nameMatch = body.match(/氏名：(.*)/);
        const emailMatch = body.match(/メールアドレス：(.*)/);
        const contentMatch = body.match(/問い合わせ内容：(.*)/);

        // 抽出できたか確認し、データを取り出す（見つからなければ空欄にする）
        const name = nameMatch ? nameMatch[1].trim() : "";
        const email = emailMatch ? emailMatch[1].trim() : "";
        const content = contentMatch ? contentMatch[1].trim() : "";

        // 5. スプレッドシートの最終行にデータを追記
       // 氏名が抽出できた場合（空欄ではない場合）のみ、シートに書き込む
        if (name !== "") {
          sheet.appendRow([date, name, email, content]);
        }

        // 6. 処理したメールを既読にして、次回以降スキップする（二重登録防止）
        message.markRead();
      }
    }
  }
}
