// ⭐️ 毎月末日に自動実行するトリガー設定
function setupAutoExecution(){
    // ⭐️ 既存のトリガーを削除（同じものが何個も作られないように）
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach (t => {
        // 💡 getHandlerFunction() これはトリガーのメンテナンス系で使う
        if (t.getHandlerFunction() === 'executeMonthlyReport') {
            ScriptApp.deleteTrigger(t);
        }
    });

    // ⭐️ 新しいトリガーを作成
    // 💡 .newTrigger() これは新たにトリガーを作成する
    ScriptApp.newTrigger('executeMonthlyReport')
        .timeBased()
        .onMonthDay(20)
        .atHour(9) // ⭐️ 毎月20日の9：00に設定
        .create();

    SpreadsheetApp.getUi().alert('毎月20日の自動トリガーを作成しました！')
}

// 全工程自動実行関数
function executeMonthlyReport(){
    try{
        // ⭐️ 部門別レポート作成関数
        createDepartmentSummary();

        // ⭐️ 一斉メール配信関数
        sendReportWithPDF();

        // 💡 Session は今このスクリプトを操作している人の情報を取るためのオブジェクト
        const myEmail = Session.getActiveUser().getEmail();
        const subject = '【お知らせ】レポート一斉メール配信'
        const body = 'ご報告\nおつかれさまです\nレポートの一斉メール配信が完了しました！'
        // メール送信
        GmailApp.sendEmail(
            myEmail,
            subject,
            body,
        );
        
        Logger.log('自動処理が正常に完了しました');
    // try-catchで全体のエラーをキャッチ
    }catch(e){
        Logger.log('エラーが発生しました' + e );
    }
}

// 全工程手動テスト
function testFullProcess(){
    const confirm = SpreadsheetApp.getUi().alert(
        '全工程テスト',
        '🗒️レポート作成→📝PDF変換→📨配信を実行します。よろしいですか？',
        SpreadsheetApp.getUi().ButtonSet.YES_NO_CANCEL
    );

    if (confirm === SpreadsheetApp.getUi().Button.YES){
        executeMonthlyReport();
        SpreadsheetApp.getUi().alert('全工程テスト完了！');
    };
}

/**
 * ⚠️ 全工程テスト： 🙆‍♂️
 * ⚠️ エラーハンドリングを全行程に組む？？：🙆‍♂️
 * ⚠️ 確認ダイアログを表示して、誤実行を防ぐ仕組み？？
 * ⚠️ 最終メニューを整理して、全機能アクセス？
 * ⚠️ 完全自動化できてるか確認
 * ⚠️ 本格運用に向けた最終チェックを行い、システムの完成
 */

/**
 * 💡 .getHandlerFunction() これはトリガーのメンテナンス系で使う
 * 💡 .newTrigger() これは新たにトリガーを作成する
 * 💡 Session は今このスクリプトを操作している人の情報を取るためのオブジェクト
 * 💡 全工程を実行するマスタ関数を作って、それを手動でテストを行う関数を作るのはベストプラクティス
 */