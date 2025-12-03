function setupRecipients(){
    try{
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const existing = ss.getSheetByName('配信先情報');
        // 同名のシートがあれば削除して初期化
        if (existing) ss.deleteSheet(existing);
        // なければ作成
        const sheet = ss.insertSheet('配信先情報');
        // メールデータを2次元配列にする
        const data = [
            ['田中太郎', 'tanaka@company.com'],
            ['佐藤花子', 'sato@company.com'],
            ['山田次郎', 'yamada@company.com']
        ];
        // 配列にしてヘッダー出力
        sheet.getRange('A1:B1').setValues([['名前','メールアドレス']]);
        // ⭐️ (1,1,data.length, data[0].length) の書き方にすれば配列の要素数を気にせず書ける
        sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

        SpreadsheetApp.getUi().alert('配信先設定完了！')
       
    }catch(e){}
}

// 配信先取得
function getRecipients(){
    /**
     * ⭐️ ss = SpreadsheetApp.getActiveSpreadsheet()
     *     sheet = ss.getSheetByName('配信先情報')
     * と書いても同じこと、要は可読性の問題
     */
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('配信先情報');
    // 配信先情報がない場合は[]を返して処理終了
    if (!sheet) return [];
    // 配信先情報のすべてのデータを取得格納
    const data = sheet.getDataRange().getValues();

        // 配信先情報のヘッダーをSliceしてデータ取得　⚠️これは実践の書き方で書いてるけど
        // 要件にあった配列形式で取得とは少し違う　配列をオブジェクトにしてるやり方
        // なので呼び出し側でもオブジェクトの呼び出し方で使う　data[0].email,
        return data.slice(1).map(row => ({
            name: row[0],
            email: row[1]
        }));
}

// メール送信関数
function sendTestEmail() {
    // 配信先取得関数呼び出して格納
    const recipients = getRecipients();

    // 配信先取得にデータがない場合はお知らせして処理終了
    if (recipients.length === 0) {
        SpreadsheetApp.getUi().alert('先に配信先を設定してください ');
        return;
    }
    // 基本集計関数を呼び出して日付、合計金額、件数をメールの件名、本文に出力
    const summary = calculateSummary();
    // ⭐️ ``バッククォートで囲むと「テンプレートリテラル」になる
    // テンプレートリテラルの中で ${} を使って変数や式を埋め込める
    const subject = `【テスト】${summary.period} 売上レポート`;
    const body = `${summary.period} の売上実績\n\n総売上: ${summary.totalSales}円\n件数: ${summary.totalCount}件`;
    // メール送信　配信先取得関数から取得したオブジェクトを使ってメールアドレスを添付
    GmailApp.sendEmail(recipients[0].email, subject, body);
    SpreadsheetApp.getUi().alert(`テストメール送信完了\n送信先: ${recipients[0].email}`);
}

/**
 * ⭐️ メールが送れるか配信テスト
 * ⭐️ メニューに配信先情報設定とテストメール機能追加
 * ⭐️ メール内容と配信先が性格に設定されているか確認
 */
/**
 * ⭐️ 配列の要素数を気にせず配列を set するには
 * ⭐️ (開始行, 開始列, data.length, data[0].length) で書く
 * ⭐️ 複数のセルに値を入力するには .setValues() を使う
 * ⭐️ メールの送信は GmailApp.sendEmail()
 */