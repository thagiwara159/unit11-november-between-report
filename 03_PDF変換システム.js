// ⭐️ PDF変換機能の関数
function convertToPdf(){
    try{
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        // ⭐️ Google Export APIを使ってシートをPDFに変換
        // ⭐️ URLFetchAppでAPIを呼び出し認証トークンを設定 blobを使うと思う
        // ⭐️ レスポンスからPDFのバイナリデータを取得
        const sheet = ss.getSheetByName('部門別レポート');
        // 部門別レポートがない場合はお知らせ
        if (!sheet){
            SpreadsheetApp.getUi().alert('先に部門別レポートを作成して下さい');
            return;
        }
        // ⭐️ ここは簡単にPDFのオプションを付与して、変換したいシートのIDとオプションを合体
        // ⭐️ gridlinesはtrueが罫線あり、falseが罫線なし
        const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` + 
                    `exportFormat=pdf&gridlines=false&gid=${sheet.getSheetId()}`;
        // ⭐️ 認証トークン取得
        const token = ScriptApp.getOAuthToken();
        // ⭐️ 承認処理　⚠️ここはあとで調べる
        const response = UrlFetchApp.fetch(url,{
            headers: { Authorization: 'Bearer ' + token } // ⭐️ Bearer のあとに半角スペースを入れる！！毎回ミスってる
        });
        // ⭐️ ここでフォルダ作ってそのままフォルダ操作
        const folder = createReportFolder('月次レポート');
        // ⭐️ バイナリデータにファイル名を付けて、PDFファイルとして格納
        const blob = response.getBlob().setName('月次レポート.PDF');
        // ⭐️ 指定のフォルダにPDFファイルを保存
        folder.createFile(blob);

    }catch(e){}
}

// ⭐️Google Driveと連携する関数　呼び出し側で引数を設定してる
function createReportFolder(folderName){
    try{
        const folder = DriveApp.getFoldersByName(folderName);
        /**
         * ⭐️ DriveAppを使ってPDFファイルをGoogleDriveに保存　DriveApp.createFile()
         * ⭐️ 月次レポート フォルダを自動作成して管理？する仕組み
         * ⭐️ 適切なファイル名でPDFを保存し重複を避ける　タイムスタンプかな
         */
        // ドライブから月次レポートを検索してある場合はそれを使うない場合は作成
        if (folder.hasNext()){
            return folder.next(); // ⚠️ ここ後で調べる
        }
        return DriveApp.createFolder(folderName); 
        
    }catch(e){}
}