function sendReportWithPDF(){
    // 配信先データ取得関数呼び出し、格納
    const recipients = getRecipients();

    // ⭐️ まずはこれを入れておけばデータない場合ここで処理を止められるからここ大事
    if (recipients.length === 0){
        SpreadsheetApp.getUi().alert('配信先データがありません');
        return;
    }
    // シートをPDFに変換して保存する関数を呼び出し
    const pdfFile = convertToPdf(); // ⭐️ convertToPdf()の最後にreturnをしてなくてForEachの中でundefinedを繰り返してた
    // 基本集計から合計金額と件数と日付を取ってくる
    const summary = calculateSummary();
    const subject = `【テスト】${summary.period} 売上レポート`; // 件名
    const body = `${summary.period} の売上実績\n\n総売上: ${summary.totalSales}円\n件数: ${summary.totalCount}件\n\n詳細は添付のPDFをご確認下さい`; //本文
    // メール送信件数を数える変数
    let successCount = 0;
    // 配信先データをループしてPDFファイルを添付して配信する処理
    recipients.forEach(person => {
        try{
            GmailApp.sendEmail(
                person.email, // 呼び出した関数からメールアドレスを添付
                subject,
                body,{
                    attachments:[pdfFile.getBlob()] // ファイル添付
            });
            successCount++ // メール送信成功回数のカウンタ
            Utilities.sleep(1000); // メール送信の間隔　1,000 = 1秒
        }catch(e){
            Logger.log(`送信失敗: ${recipients.email}`);
        }
    });
    // 送信成功回数とデータの数の比較をお知らせ
    SpreadsheetApp.getUi().alert(`配信完了！\n成功: ${successCount}/${recipients.length}件`);
}
// ⚠️ このままだと動かない：直した
// 送信間隔を調整ができてない：送信間隔の調整は Utilities.sleep(1000); これで送信間隔１秒　

/**
 * ⚠️送信成功、失敗のカウント結果を管理する機能がわからない：successCount変数作ってこれをカウンタにした
 * エラーが発生しても処理を継続する方法は？：try-catch
 * 送信後に結果を残す：getUi().alertでお知らせ　
 * 機能の動作確認：convertToPdf()が作成したPDFファイルを返していなかったのでメール送信でエラーが出てた
 * メニューに機能追加
 * 配信が正確に行われるか確認：動作確認済み
 */
/**
 * ⭐️ 最初にデータがあるかシートがあるかを if(!sheet) return;のようなガード条件で中断するのが一番シンプル
 * ⭐️ 今回はループの中でファイルを返す関数のミスでメール送信でバグが出た
 * ⭐️ エラーが起きやすいとこに単発try‐catchを入れよう
 * ⭐️ ループの中、外部API呼び出し、ファイル操作(Drive)など
 * ⭐️ 大きな処理を丸ごとtry-catchで包むのはNG
 */